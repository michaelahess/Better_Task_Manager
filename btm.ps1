#Requires -Modules ScheduledTasks
<#
.SYNOPSIS  Interactive Task Scheduler Manager
.NOTES     Run as Administrator for full functionality.  PowerShell 5.1+
#>

$ErrorActionPreference = 'Continue'

# warn if not elevated — listing works, but create/edit/delete need admin
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning 'Not running as Administrator — some operations may fail.'
    Start-Sleep -Seconds 2
}

# ═══════════════════════════════════════════════════════════════════
#  CORE HELPERS
# ═══════════════════════════════════════════════════════════════════

function Write-Header ([string]$Sub = '') {
    Clear-Host
    Write-Host '  ╔═══════════════════════════════════════════════╗' -ForegroundColor Cyan
    Write-Host '  ║              Better Task Manager              ║' -ForegroundColor Cyan
    Write-Host '  ╚═══════════════════════════════════════════════╝' -ForegroundColor Cyan
    if ($Sub) { Write-Host "`n  $Sub`n" -ForegroundColor Yellow }
    else      { Write-Host }
}

function Pause-ForUser { Read-Host "`n  Press Enter to continue" | Out-Null }

function Read-Valid ([string]$Prompt, [string[]]$Valid) {
    do { $v = (Read-Host "  $Prompt").Trim() } while ($v -notin $Valid)
    return $v
}

function Show-Paged {
    param(
        [AllowEmptyCollection()][object[]]$Rows,
        [int]$PageSize = 20
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-Host '  (no results found)' -ForegroundColor Yellow
        Pause-ForUser; return
    }

    $total = $Rows.Count
    $pages = [Math]::Ceiling($total / $PageSize)
    $page  = 0

    do {
        Clear-Host
        $s = $page * $PageSize
        $e = [Math]::Min($s + $PageSize - 1, $total - 1)
        $Rows[$s..$e] | Format-Table -AutoSize | Out-String | Write-Host
        Write-Host "  Page $($page + 1)/$pages  ·  $total total item(s)" -ForegroundColor DarkCyan

        if ($pages -le 1) { Pause-ForUser; return }

        $nav = @()
        if ($page -gt 0)           { $nav += '[P]rev' }
        if ($page -lt $pages - 1)  { $nav += '[N]ext' }
        $nav += '[B]ack'
        Write-Host ("  " + ($nav -join '   ')) -ForegroundColor Gray

        switch ((Read-Host '  >').Trim().ToUpper()) {
            'N' { if ($page -lt $pages - 1) { $page++ } }
            'P' { if ($page -gt 0)          { $page-- } }
            'B' { return }
        }
    } while ($true)
}

function Show-TaskListPaged {
    param(
        [AllowEmptyCollection()][Microsoft.Management.Infrastructure.CimInstance[]]$Tasks,
        [AllowEmptyCollection()][object[]]$Rows,
        [int]$PageSize = 20
    )

    if (-not $Tasks -or $Tasks.Count -eq 0) {
        Write-Host '  (no results found)' -ForegroundColor Yellow
        Pause-ForUser; return
    }

    $total = $Tasks.Count
    $pages = [Math]::Ceiling($total / $PageSize)
    $page  = 0

    do {
        Clear-Host
        $s = $page * $PageSize
        $e = [Math]::Min($s + $PageSize - 1, $total - 1)

        $pageRows = for ($i = $s; $i -le $e; $i++) {
            $props = [ordered]@{ '#' = $i + 1 }
            $Rows[$i].PSObject.Properties | ForEach-Object { $props[$_.Name] = $_.Value }
            [PSCustomObject]$props
        }
        $pageRows | Format-Table -AutoSize | Out-String | Write-Host
        Write-Host "  Page $($page + 1)/$pages  ·  $total total item(s)" -ForegroundColor DarkCyan

        $nav = @()
        if ($page -gt 0)           { $nav += '[P]rev' }
        if ($page -lt $pages - 1)  { $nav += '[N]ext' }
        $nav += '[B]ack'
        Write-Host ("  " + ($nav -join '   ')) -ForegroundColor Gray
        Write-Host '  Row actions: [#] Detail  [#E]dit  [#N]able  [#D]isable  [#R]un  [#S]top  [#X] Delete  [#H]istory' -ForegroundColor DarkGray

        $raw = (Read-Host '  >').Trim()
        $up  = $raw.ToUpper()

        if ($up -eq 'N' -and $page -lt $pages - 1) { $page++; continue }
        if ($up -eq 'P' -and $page -gt 0)          { $page--; continue }
        if ($up -eq 'B') { return }

        # Accept "<number>" or "<number><letter>" (with optional space)
        if ($raw -match '^(\d+)\s*([A-Za-z]?)$') {
            $n          = [int]$Matches[1]
            $actionChar = $Matches[2].ToUpper()
            if ($n -ge 1 -and $n -le $total) {
                if ($actionChar -eq '') {
                    Show-TaskDetail -Task $Tasks[$n - 1]
                } else {
                    Invoke-QuickAction -Task $Tasks[$n - 1] -Action $actionChar
                }
            }
        }
    } while ($true)
}

function Invoke-QuickAction {
    param(
        [Microsoft.Management.Infrastructure.CimInstance]$Task,
        [string]$Action
    )

    $fullName = "$($Task.TaskPath)$($Task.TaskName)"

    switch ($Action) {
        'E' { Edit-TaskWizard -Task $Task; return }
        'H' { Show-TaskHistory -Task $Task; return }
        'N' {
            Write-Host "`n  Enable: $fullName" -ForegroundColor Yellow
            if ((Read-Valid 'Enable this task? [Y/N]' @('Y','y','N','n')) -in 'Y','y') {
                try { Enable-ScheduledTask -TaskName $Task.TaskName -TaskPath $Task.TaskPath | Out-Null
                      Write-Host '  Task enabled.' -ForegroundColor Green }
                catch { Write-Host "  Error: $_" -ForegroundColor Red }
            } else { Write-Host '  Cancelled.' -ForegroundColor Gray }
            Pause-ForUser
        }
        'D' {
            Write-Host "`n  Disable: $fullName" -ForegroundColor Yellow
            if ((Read-Valid 'Disable this task? [Y/N]' @('Y','y','N','n')) -in 'Y','y') {
                try { Disable-ScheduledTask -TaskName $Task.TaskName -TaskPath $Task.TaskPath | Out-Null
                      Write-Host '  Task disabled.' -ForegroundColor Green }
                catch { Write-Host "  Error: $_" -ForegroundColor Red }
            } else { Write-Host '  Cancelled.' -ForegroundColor Gray }
            Pause-ForUser
        }
        'R' {
            Write-Host "`n  Run now: $fullName" -ForegroundColor Yellow
            if ((Read-Valid 'Run this task now? [Y/N]' @('Y','y','N','n')) -in 'Y','y') {
                try { Start-ScheduledTask -TaskName $Task.TaskName -TaskPath $Task.TaskPath
                      Write-Host '  Task started.' -ForegroundColor Green }
                catch { Write-Host "  Error: $_" -ForegroundColor Red }
            } else { Write-Host '  Cancelled.' -ForegroundColor Gray }
            Pause-ForUser
        }
        'S' {
            Write-Host "`n  Stop: $fullName" -ForegroundColor Yellow
            if ((Read-Valid 'Stop this task? [Y/N]' @('Y','y','N','n')) -in 'Y','y') {
                try { Stop-ScheduledTask -TaskName $Task.TaskName -TaskPath $Task.TaskPath
                      Write-Host '  Task stopped.' -ForegroundColor Green }
                catch { Write-Host "  Error: $_" -ForegroundColor Red }
            } else { Write-Host '  Cancelled.' -ForegroundColor Gray }
            Pause-ForUser
        }
        'X' {
            Write-Host "`n  DELETE: $fullName" -ForegroundColor Red
            Write-Host '  WARNING: This cannot be undone.' -ForegroundColor Red
            Write-Host '  Type  DELETE  to confirm (anything else = cancel):'
            if ((Read-Host '  >').Trim() -ceq 'DELETE') {
                try { Unregister-ScheduledTask -TaskName $Task.TaskName -TaskPath $Task.TaskPath -Confirm:$false
                      Write-Host '  Task deleted.' -ForegroundColor Green }
                catch { Write-Host "  Error: $_" -ForegroundColor Red }
            } else { Write-Host '  Cancelled.' -ForegroundColor Gray }
            Pause-ForUser
        }
        default { Show-TaskDetail -Task $Task }
    }
}

function Show-TaskDetail {
    param([Microsoft.Management.Infrastructure.CimInstance]$Task)

    Clear-Host
    Write-Host "  ── $($Task.TaskPath)$($Task.TaskName) ──`n" -ForegroundColor Cyan
    Write-Host "  Description : $(if ($Task.Description) { $Task.Description } else { '(none)' })"
    Write-Host "  State       : $($Task.State)"
    Write-Host "  Run As      : $($Task.Principal.UserId)  [$($Task.Principal.RunLevel)]"

    Write-Host "`n  Trigger(s)  :"
    if ($Task.Triggers -and @($Task.Triggers).Count -gt 0) {
        foreach ($tr in @($Task.Triggers)) {
            Write-Host "              · $(Format-TriggerSummary $tr)  [Enabled: $($tr.Enabled)]"
        }
    } else { Write-Host '              (none)' }

    Write-Host "`n  Action(s)   :"
    foreach ($a in @($Task.Actions)) {
        Write-Host "              · Execute : $($a.Execute)"
        if ($a.Arguments)        { Write-Host "                Args    : $($a.Arguments)" }
        if ($a.WorkingDirectory) { Write-Host "                WorkDir : $($a.WorkingDirectory)" }
    }

    $s = $Task.Settings
    if ($s) {
        Write-Host "`n  Settings    :"
        Write-Host "              · Exec time limit    : $($s.ExecutionTimeLimit)"
        Write-Host "              · Multiple instances : $($s.MultipleInstances)"
        Write-Host "              · Start when missed  : $($s.StartWhenAvailable)"
        Write-Host "              · Hidden             : $($s.Hidden)"
    }

    try {
        $info = Get-ScheduledTaskInfo -TaskName $Task.TaskName -TaskPath $Task.TaskPath -EA Stop
        Write-Host "`n  Last Run    : $(if ($info.LastRunTime -gt [DateTime]::MinValue) { $info.LastRunTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Never' })"
        Write-Host "  Next Run    : $(if ($info.NextRunTime -gt [DateTime]::MinValue) { $info.NextRunTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' })"
        $rc    = [uint32]$info.LastTaskResult
        $rcStr = switch ($rc) {
            0          { 'Success'        }
            267009     { 'Still running'  }
            267011     { 'Never run'      }
            2147942402 { 'File not found' }
            default    { ''               }
        }
        Write-Host "  Last Result : 0x$($rc.ToString('X8'))$(if ($rcStr) { "  ($rcStr)" })"
    } catch {}

    Write-Host
    Write-Host '  [E] Edit this task   [Enter] Back' -ForegroundColor Gray
    if ((Read-Host '  >').Trim().ToUpper() -eq 'E') {
        Edit-TaskWizard -Task $Task
    }
}

function Get-TaskRows {
    param(
        [Microsoft.Management.Infrastructure.CimInstance[]]$Tasks,
        [switch]$WithInfo
    )
    if (-not $Tasks -or $Tasks.Count -eq 0) { return @() }

    $result = foreach ($t in $Tasks) {
        if ($WithInfo) {
            try   { $i = Get-ScheduledTaskInfo -TaskName $t.TaskName -TaskPath $t.TaskPath -EA Stop }
            catch { $i = $null }

            [PSCustomObject]@{
                Name       = $t.TaskName
                Folder     = $t.TaskPath
                State      = $t.State
                'Last Run'  = if ($i -and $i.LastRunTime  -gt [DateTime]::MinValue) { $i.LastRunTime.ToString('yyyy-MM-dd HH:mm')  } else { 'Never' }
                'Next Run'  = if ($i -and $i.NextRunTime  -gt [DateTime]::MinValue) { $i.NextRunTime.ToString('yyyy-MM-dd HH:mm')  } else { 'N/A'   }
                'Exit Code' = if ($i) { '0x{0:X8}' -f [uint32]$i.LastTaskResult } else { 'N/A' }
            }
        } else {
            [PSCustomObject]@{ Name = $t.TaskName; Folder = $t.TaskPath; State = $t.State }
        }
    }
    return @($result)
}

function Select-Task ([string]$Verb = 'select') {
    do {
        $kw = (Read-Host '  Filter keyword (blank = show all)').Trim()
        Write-Host '  Searching...' -ForegroundColor DarkGray

        $found = @(Get-ScheduledTask -EA SilentlyContinue | Where-Object {
            !$kw -or $_.TaskName -like "*$kw*" -or $_.TaskPath -like "*$kw*"
        })

        if ($found.Count -eq 0) {
            Write-Host '  No tasks found.' -ForegroundColor Yellow
            if ((Read-Valid 'Search again? [Y/N]' @('Y','y','N','n')) -in 'N','n') { return $null }
            continue
        }

        Write-Host "`n  $($found.Count) task(s) found:`n" -ForegroundColor Cyan
        for ($i = 0; $i -lt $found.Count; $i++) {
            $col = switch ($found[$i].State) {
                'Ready'    { 'Green'    }
                'Disabled' { 'DarkGray' }
                'Running'  { 'Yellow'   }
                default    { 'White'    }
            }
            Write-Host "  [$($i + 1)] $($found[$i].TaskPath)$($found[$i].TaskName) " -NoNewline
            Write-Host "[$($found[$i].State)]" -ForegroundColor $col
        }
        Write-Host '  [0] Cancel / Back'

        do {
            $raw = (Read-Host "`n  Number to $Verb").Trim()
            if ($raw -eq '0') { return $null }
            $n = 0
            $ok = [int]::TryParse($raw, [ref]$n) -and $n -ge 1 -and $n -le $found.Count
        } while (!$ok)

        return $found[$n - 1]
    } while ($true)
}

function Format-TriggerSummary ([object]$Trigger) {
    $cn = $Trigger.CimClass.CimClassName
    switch -Wildcard ($cn) {
        '*Startup*' { 'At startup' }
        '*Daily*'   { "Daily  @  $($Trigger.StartBoundary)" }
        '*Weekly*'  { "Weekly  @  $($Trigger.StartBoundary)" }
        '*Logon*'   { "At logon$(if ($Trigger.UserId) { " ($($Trigger.UserId))" })" }
        '*Time*'    { "Once  @  $($Trigger.StartBoundary)" }
        '*Event*'   { "On Event" }
        default     { $cn -replace 'MSFT_Task', '' }
    }
}

# ═══════════════════════════════════════════════════════════════════
#  WIZARD INPUT HELPERS  (shared by Create & Edit)
# ═══════════════════════════════════════════════════════════════════

function Get-TriggerFromUser {
    Write-Host "`n  Trigger type:" -ForegroundColor Yellow
    Write-Host '   [1] At system startup'
    Write-Host '   [2] Daily'
    Write-Host '   [3] Weekly'
    Write-Host '   [4] At user logon'
    Write-Host '   [5] Once (specific date/time)'
    Write-Host '   [6] On Event Log entry  (Server 2016+ / Win10+)'
    Write-Host '   [B] Cancel'

    $c = Read-Valid 'Trigger [1-6 or B]' @('1','2','3','4','5','6','B','b')
    if ($c -in 'B','b') { return $null }

    switch ($c) {
        '1' { return New-ScheduledTaskTrigger -AtStartup }

        '2' {
            $t = (Read-Host '  Daily at what time? (e.g. 03:00)').Trim()
            return New-ScheduledTaskTrigger -Daily -At $t
        }

        '3' {
            Write-Host '  Days — comma-separated (e.g. Monday,Wednesday,Friday)'
            $days = (Read-Host '  Days').Trim() -split '\s*,\s*'
            $t    = (Read-Host '  At what time? (e.g. 09:00)').Trim()
            return New-ScheduledTaskTrigger -Weekly -DaysOfWeek $days -At $t
        }

        '4' {
            $u = (Read-Host '  Username to watch (blank = any user)').Trim()
            if ($u) { return New-ScheduledTaskTrigger -AtLogOn -User $u }
            else    { return New-ScheduledTaskTrigger -AtLogOn }
        }

        '5' {
            $dt = (Read-Host '  Date/time (e.g. 2026-06-01 08:00)').Trim()
            return New-ScheduledTaskTrigger -Once -At $dt
        }

        '6' {
            $log = (Read-Host '  Event log   (e.g. System)').Trim()
            $src = (Read-Host '  Source/provider  (e.g. Service Control Manager)').Trim()
            $id  = [int](Read-Host '  Event ID').Trim()
            return New-ScheduledTaskTrigger -OnEvent -Log $log -Source $src -EventId $id
        }
    }
}

function Get-ActionFromUser {
    Write-Host "`n  Action — what to run:" -ForegroundColor Yellow
    $exe = (Read-Host '  Executable or script path').Trim()
    if (!$exe) { return $null }

    $arg  = (Read-Host '  Arguments (optional)').Trim()
    $wdir = (Read-Host '  Working directory (optional)').Trim()

    $p = @{ Execute = $exe }
    if ($arg)  { $p.Argument         = $arg  }
    if ($wdir) { $p.WorkingDirectory = $wdir }
    return New-ScheduledTaskAction @p
}

function Get-PrincipalFromUser {
    # Returns a hashtable: @{ Principal=...; [User=...]; [Password=...] }
    Write-Host "`n  Run as account:" -ForegroundColor Yellow
    Write-Host '   [1] SYSTEM            (no password, highest privilege)'
    Write-Host '   [2] NETWORK SERVICE   (limited network access)'
    Write-Host '   [3] LOCAL SERVICE     (limited local access)'
    Write-Host '   [4] Specific user     (will prompt for password)'
    Write-Host '   [B] Cancel'

    $c = Read-Valid 'Account [1-4 or B]' @('1','2','3','4','B','b')
    if ($c -in 'B','b') { return $null }

    switch ($c) {
        '1' { return @{ Principal = (New-ScheduledTaskPrincipal -UserId 'SYSTEM'          -LogonType ServiceAccount -RunLevel Highest) } }
        '2' { return @{ Principal = (New-ScheduledTaskPrincipal -UserId 'NETWORK SERVICE' -LogonType ServiceAccount) } }
        '3' { return @{ Principal = (New-ScheduledTaskPrincipal -UserId 'LOCAL SERVICE'   -LogonType ServiceAccount) } }

        '4' {
            $user  = (Read-Host '  Username (e.g. DOMAIN\user  or  .\localadmin)').Trim()
            $secpw = Read-Host '  Password' -AsSecureString
            $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
                         [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secpw))
            $principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Password -RunLevel Highest
            return @{ Principal = $principal; User = $user; Password = $plain }
        }
    }
}

# ═══════════════════════════════════════════════════════════════════
#  1-3  LIST FUNCTIONS
# ═══════════════════════════════════════════════════════════════════

function Show-EnabledTasks {
    Write-Header 'Enabled Tasks'
    $kw = (Read-Host '  Filter keyword (blank = show all)').Trim()
    Write-Host '  Loading — fetching run info for each task, may take a moment...' -ForegroundColor DarkGray
    $tasks = @(Get-ScheduledTask -EA SilentlyContinue | Where-Object {
        $_.State -eq 'Ready' -and (!$kw -or $_.TaskName -like "*$kw*" -or $_.TaskPath -like "*$kw*")
    })
    Show-TaskListPaged -Tasks $tasks -Rows (Get-TaskRows -Tasks $tasks -WithInfo)
}

function Show-DisabledTasks {
    Write-Header 'Disabled Tasks'
    $kw = (Read-Host '  Filter keyword (blank = show all)').Trim()
    Write-Host '  Loading — fetching run info for each task, may take a moment...' -ForegroundColor DarkGray
    $tasks = @(Get-ScheduledTask -EA SilentlyContinue | Where-Object {
        $_.State -eq 'Disabled' -and (!$kw -or $_.TaskName -like "*$kw*" -or $_.TaskPath -like "*$kw*")
    })
    Show-TaskListPaged -Tasks $tasks -Rows (Get-TaskRows -Tasks $tasks -WithInfo)
}

function Show-RunningTasks {
    do {
        Write-Header 'Running Tasks'
        $kw = (Read-Host '  Filter keyword (blank = show all)').Trim()
        $tasks = @(Get-ScheduledTask -EA SilentlyContinue | Where-Object {
            $_.State -eq 'Running' -and (!$kw -or $_.TaskName -like "*$kw*" -or $_.TaskPath -like "*$kw*")
        })
        Show-TaskListPaged -Tasks $tasks -Rows (Get-TaskRows -Tasks $tasks)
        Write-Host '  [R] Refresh   [Enter] Back' -ForegroundColor Gray
        $r = (Read-Host '  >').Trim().ToUpper()
    } while ($r -eq 'R')
}

# ═══════════════════════════════════════════════════════════════════
#  4-6  ENABLE / DISABLE / DELETE
# ═══════════════════════════════════════════════════════════════════

function Invoke-EnableTask {
    Write-Header 'Enable a Task'
    $t = Select-Task 'enable'
    if (!$t) { return }

    Write-Host "`n  $($t.TaskPath)$($t.TaskName)  [$($t.State)]" -ForegroundColor Yellow
    if ((Read-Valid 'Enable this task? [Y/N]' @('Y','y','N','n')) -in 'Y','y') {
        try {
            Enable-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath | Out-Null
            Write-Host '  Task enabled.' -ForegroundColor Green
        } catch { Write-Host "  Error: $_" -ForegroundColor Red }
    } else { Write-Host '  Cancelled.' -ForegroundColor Gray }
    Pause-ForUser
}

function Invoke-DisableTask {
    Write-Header 'Disable a Task'
    $t = Select-Task 'disable'
    if (!$t) { return }

    Write-Host "`n  $($t.TaskPath)$($t.TaskName)  [$($t.State)]" -ForegroundColor Yellow
    if ((Read-Valid 'Disable this task? [Y/N]' @('Y','y','N','n')) -in 'Y','y') {
        try {
            Disable-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath | Out-Null
            Write-Host '  Task disabled.' -ForegroundColor Green
        } catch { Write-Host "  Error: $_" -ForegroundColor Red }
    } else { Write-Host '  Cancelled.' -ForegroundColor Gray }
    Pause-ForUser
}

function Invoke-DeleteTask {
    Write-Header 'Delete a Task'
    $t = Select-Task 'delete'
    if (!$t) { return }

    Write-Host "`n  $($t.TaskPath)$($t.TaskName)" -ForegroundColor Red
    Write-Host '  WARNING: This cannot be undone.' -ForegroundColor Red
    Write-Host "  Type  DELETE  to confirm (anything else = cancel):"
    $confirm = (Read-Host '  >').Trim()

    if ($confirm -ceq 'DELETE') {
        try {
            Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Confirm:$false
            Write-Host '  Task deleted.' -ForegroundColor Green
        } catch { Write-Host "  Error: $_" -ForegroundColor Red }
    } else { Write-Host '  Cancelled.' -ForegroundColor Gray }
    Pause-ForUser
}

function Show-TaskHistory {
    param([Microsoft.Management.Infrastructure.CimInstance]$Task)
    Write-Header 'Task Run History'
    $t = if ($Task) { $Task } else { Select-Task 'view history for' }
    if (!$t) { return }

    Write-Host '  Reading event log...' -ForegroundColor DarkGray

    $taskFullPath = $t.TaskPath + $t.TaskName

    $categoryMap = @{
        100 = 'Task Started'
        101 = 'Task Start Failed'
        102 = 'Task completed'
        103 = 'Action failed'
        106 = 'Task registered'
        107 = 'Task triggered on schedule'
        110 = 'Task triggered by user'
        111 = 'Task terminated'
        119 = 'Task triggered on idle'
        129 = 'Created Task Process'
        140 = 'Task registration updated'
        141 = 'Task deleted'
        200 = 'Action started'
        201 = 'Action completed'
        202 = 'Action failed to start'
        203 = 'Action completed (failed)'
        330 = 'Task stopping due to user request'
        331 = 'Task stopping'
    }

    try {
        $allEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'Microsoft-Windows-TaskScheduler/Operational'
            Id      = @(100,101,102,103,106,107,110,111,119,129,140,141,200,201,202,203,330,331)
        } -EA Stop | Where-Object {
            $_.Properties.Count -gt 0 -and $_.Properties[0].Value -eq $taskFullPath
        } | Select-Object -First 500

        if ($allEvents.Count -eq 0) {
            Write-Host "  No history found for '$($t.TaskName)'." -ForegroundColor Yellow
            Write-Host '  Task Scheduler history may not be enabled.' -ForegroundColor DarkGray
            Write-Host '  To enable: open Task Scheduler > Action menu > Enable All Tasks History' -ForegroundColor DarkGray
            Pause-ForUser; return
        }

        $rows = $allEvents | Sort-Object TimeCreated -Descending | ForEach-Object {
            $corr = if ($_.ActivityId) {
                $g = "$($_.ActivityId)"
                if ($g.Length -gt 18) { $g.Substring(0, 18) + '...' } else { $g }
            } else { '' }

            $op = $_.OpcodeDisplayName
            if (!$op -or $op -eq 'Info') {
                $op = if ($_.Opcode -gt 0) { "($($_.Opcode))" } else { 'Info' }
            }

            [PSCustomObject]@{
                'Date and Time'  = $_.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
                'E'              = $_.Id
                'Task Category'  = if ($categoryMap.ContainsKey($_.Id)) { $categoryMap[$_.Id] } else { "Event $($_.Id)" }
                'Operational'    = $op
                'Correlation Id' = $corr
            }
        }

        Clear-Host
        Write-Host "  History — $taskFullPath  ($($rows.Count) event(s))`n" -ForegroundColor Cyan
        Show-Paged -Rows $rows
        return

    } catch {
        if ($_.Exception.Message -match 'No events' -or $_.FullyQualifiedErrorId -match 'NoMatchingEventsFound') {
            Write-Host "  No history found for '$($t.TaskName)'." -ForegroundColor Yellow
            Write-Host '  Task Scheduler history may not be enabled.' -ForegroundColor DarkGray
            Write-Host '  To enable: open Task Scheduler > Action menu > Enable All Tasks History' -ForegroundColor DarkGray
        } else {
            Write-Host "  Error reading event log: $_" -ForegroundColor Red
        }
    }
    Pause-ForUser
}

function Invoke-StopTask {
    Write-Header 'Stop a Running Task'
    $running = @(Get-ScheduledTask -EA SilentlyContinue | Where-Object State -eq 'Running')

    if ($running.Count -eq 0) {
        Write-Host '  No tasks are currently running.' -ForegroundColor Yellow
        Pause-ForUser; return
    }

    Write-Host "`n  $($running.Count) running task(s):`n" -ForegroundColor Cyan
    for ($i = 0; $i -lt $running.Count; $i++) {
        Write-Host "  [$($i + 1)] $($running[$i].TaskPath)$($running[$i].TaskName)"
    }
    Write-Host '  [0] Cancel / Back'

    do {
        $raw = (Read-Host "`n  Number to stop").Trim()
        if ($raw -eq '0') { return }
        $n = 0
        $ok = [int]::TryParse($raw, [ref]$n) -and $n -ge 1 -and $n -le $running.Count
    } while (!$ok)

    $t = $running[$n - 1]
    Write-Host "`n  $($t.TaskPath)$($t.TaskName)" -ForegroundColor Yellow
    if ((Read-Valid 'Stop this task? [Y/N]' @('Y','y','N','n')) -in 'Y','y') {
        try {
            Stop-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath
            Write-Host '  Task stopped.' -ForegroundColor Green
        } catch { Write-Host "  Error: $_" -ForegroundColor Red }
    } else { Write-Host '  Cancelled.' -ForegroundColor Gray }
    Pause-ForUser
}

# ═══════════════════════════════════════════════════════════════════
#  7  CREATE WIZARD
# ═══════════════════════════════════════════════════════════════════

function New-TaskWizard {
    Write-Header 'Create a New Task'
    Write-Host '  Type B and Enter at any prompt to cancel.' -ForegroundColor DarkGray

    # ── Identity ────────────────────────────────────────────────────
    $name = (Read-Host "`n  Task name").Trim()
    if (!$name -or $name -in 'B','b') { return }

    $folder = (Read-Host '  Folder / path (default: \)').Trim()
    if ($folder -in 'B','b') { return }
    if (!$folder) { $folder = '\' }
    if (!$folder.StartsWith('\')) { $folder = "\$folder" }
    if (!$folder.EndsWith('\'))   { $folder = "$folder\" }

    $desc = (Read-Host '  Description (optional)').Trim()

    # ── Trigger ──────────────────────────────────────────────────────
    $trigger = Get-TriggerFromUser
    if (!$trigger) { Write-Host '  Cancelled.' -ForegroundColor Gray; Pause-ForUser; return }

    # ── Action ───────────────────────────────────────────────────────
    $action = Get-ActionFromUser
    if (!$action) { Write-Host '  Cancelled.' -ForegroundColor Gray; Pause-ForUser; return }

    # ── Run As ───────────────────────────────────────────────────────
    $pi = Get-PrincipalFromUser
    if (!$pi) { Write-Host '  Cancelled.' -ForegroundColor Gray; Pause-ForUser; return }

    # ── Settings (sensible defaults) ─────────────────────────────────
    $settings = New-ScheduledTaskSettingsSet `
        -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
        -MultipleInstances  IgnoreNew `
        -StartWhenAvailable

    # ── Review ───────────────────────────────────────────────────────
    Clear-Host
    Write-Host '  ── Review ─────────────────────────────────────────────' -ForegroundColor Cyan
    Write-Host "  Name        : $name"
    Write-Host "  Folder      : $folder"
    if ($desc) { Write-Host "  Description : $desc" }
    Write-Host "  Trigger     : $(Format-TriggerSummary $trigger)"
    Write-Host "  Execute     : $($action.Execute)"
    if ($action.Arguments)        { Write-Host "  Arguments   : $($action.Arguments)" }
    if ($action.WorkingDirectory) { Write-Host "  Working dir : $($action.WorkingDirectory)" }
    Write-Host "  Run As      : $($pi.Principal.UserId)  [$($pi.Principal.RunLevel)]"
    Write-Host

    if ((Read-Valid 'Create this task? [Y/N]' @('Y','y','N','n')) -in 'N','n') {
        Write-Host '  Cancelled.' -ForegroundColor Gray; Pause-ForUser; return
    }

    # ── Register ─────────────────────────────────────────────────────
    try {
        $def = New-ScheduledTask -Action $action -Trigger $trigger -Principal $pi.Principal -Settings $settings

        $reg = @{ TaskName = $name; TaskPath = $folder; InputObject = $def }
        if ($pi.User)     { $reg.User     = $pi.User }
        if ($pi.Password) { $reg.Password = $pi.Password }

        Register-ScheduledTask @reg -EA Stop | Out-Null

        if ($desc) {
            $svc = New-Object -ComObject 'Schedule.Service'
            $svc.Connect()
            $fp  = if ($folder -eq '\') { '\' } else { $folder.TrimEnd('\') }
            $def = $svc.GetFolder($fp).GetTask($name).Definition
            $def.RegistrationInfo.Description = $desc
            # Pass credentials through — required when task runs as a stored-password user account
            $descLogon = $def.Principal.LogonType
            $descUser  = $def.Principal.UserId
            $descPw    = if ($pi.Password) { $pi.Password } else { $null }
            $svc.GetFolder($fp).RegisterTaskDefinition($name, $def, 4, $descUser, $descPw, $descLogon) | Out-Null
        }

        Write-Host "  Task '$name' created successfully!" -ForegroundColor Green
    } catch {
        Write-Host "  Error: $_" -ForegroundColor Red
    }
    Pause-ForUser
}

# ═══════════════════════════════════════════════════════════════════
#  8  EDIT WIZARD
# ═══════════════════════════════════════════════════════════════════

function Edit-TaskWizard {
    param([Microsoft.Management.Infrastructure.CimInstance]$Task)

    Write-Header 'Edit a Task'
    $t = if ($Task) { $Task } else { Select-Task 'edit' }
    if (!$t) { return }

    do {
        Clear-Host
        Write-Host "  ── $($t.TaskPath)$($t.TaskName) ──`n" -ForegroundColor Cyan
        Write-Host "  Description : $(if ($t.Description) { $t.Description } else { '(none)' })"
        Write-Host "  State       : $($t.State)"
        Write-Host "  Run As      : $($t.Principal.UserId)  [$($t.Principal.RunLevel)]"

        Write-Host '  Trigger(s)  :' -NoNewline
        if ($t.Triggers -and @($t.Triggers).Count -gt 0) {
            Write-Host
            foreach ($tr in @($t.Triggers)) { Write-Host "              · $(Format-TriggerSummary $tr)" }
        } else { Write-Host ' (none)' }

        Write-Host '  Action(s)   :'
        foreach ($a in @($t.Actions)) {
            $line = "              · $($a.Execute)"
            if ($a.Arguments) { $line += "  $($a.Arguments)" }
            Write-Host $line
        }

        Write-Host "`n  What would you like to edit?" -ForegroundColor Yellow
        Write-Host '   [1] Replace trigger'
        Write-Host '   [2] Replace action'
        Write-Host '   [3] Change run-as account'
        Write-Host '   [4] Update description'
        Write-Host '   [5] Change execution time limit'
        Write-Host '   [B] Back to main menu'

        $c = Read-Valid 'Choice [1-5 or B]' @('1','2','3','4','5','B','b')
        if ($c -in 'B','b') { return }

        switch ($c) {
            '1' {
                $newTrigger = Get-TriggerFromUser
                if ($newTrigger) {
                    try {
                        Set-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Trigger $newTrigger | Out-Null
                        $t = Get-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath
                        Write-Host '  Trigger updated.' -ForegroundColor Green
                    } catch { Write-Host "  Error: $_" -ForegroundColor Red }
                    Pause-ForUser
                }
            }
            '2' {
                $newAction = Get-ActionFromUser
                if ($newAction) {
                    try {
                        Set-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Action $newAction | Out-Null
                        $t = Get-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath
                        Write-Host '  Action updated.' -ForegroundColor Green
                    } catch { Write-Host "  Error: $_" -ForegroundColor Red }
                    Pause-ForUser
                }
            }
            '3' {
                $pi = Get-PrincipalFromUser
                if ($pi) {
                    try {
                        $params = @{ TaskName = $t.TaskName; TaskPath = $t.TaskPath; Principal = $pi.Principal }
                        if ($pi.User)     { $params.User     = $pi.User }
                        if ($pi.Password) { $params.Password = $pi.Password }
                        Set-ScheduledTask @params | Out-Null
                        $t = Get-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath
                        Write-Host '  Run-as account updated.' -ForegroundColor Green
                    } catch { Write-Host "  Error: $_" -ForegroundColor Red }
                    Pause-ForUser
                }
            }
            '4' {
                $newDesc = (Read-Host '  New description (blank to clear)').Trim()
                try {
                    $svc = New-Object -ComObject 'Schedule.Service'
                    $svc.Connect()
                    $fp  = if ($t.TaskPath -eq '\') { '\' } else { $t.TaskPath.TrimEnd('\') }
                    $def = $svc.GetFolder($fp).GetTask($t.TaskName).Definition
                    $def.RegistrationInfo.Description = $newDesc

                    $userId    = $def.Principal.UserId
                    $logonType = $def.Principal.LogonType
                    $password  = $null
                    # LogonType 1 = Password, 6 = InteractiveTokenOrPassword — both require credentials
                    if ($logonType -in 1, 6) {
                        Write-Host "  Task runs as '$userId' with a stored password." -ForegroundColor DarkGray
                        $secpw    = Read-Host '  Enter password to save changes' -AsSecureString
                        $password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
                                        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secpw))
                    }

                    $svc.GetFolder($fp).RegisterTaskDefinition($t.TaskName, $def, 4, $userId, $password, $logonType) | Out-Null
                    $t = Get-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath
                    Write-Host '  Description updated.' -ForegroundColor Green
                } catch { Write-Host "  Error: $_" -ForegroundColor Red }
                Pause-ForUser
            }
            '5' {
                $curLimit = $t.Settings.ExecutionTimeLimit
                Write-Host "  Current limit: $(if ($curLimit) { $curLimit } else { '(none / unlimited)' })"
                Write-Host '  Enter the new limit. Both 0 = no limit.' -ForegroundColor DarkGray
                $hRaw = (Read-Host '  Hours   (0-999)').Trim()
                $mRaw = (Read-Host '  Minutes (0-59) ').Trim()
                $h = 0; $m = 0
                [int]::TryParse($hRaw, [ref]$h) | Out-Null
                [int]::TryParse($mRaw, [ref]$m) | Out-Null
                $h = [Math]::Max(0, $h)
                $m = [Math]::Max(0, [Math]::Min(59, $m))
                try {
                    $svc = New-Object -ComObject 'Schedule.Service'
                    $svc.Connect()
                    $fp  = if ($t.TaskPath -eq '\') { '\' } else { $t.TaskPath.TrimEnd('\') }
                    $def = $svc.GetFolder($fp).GetTask($t.TaskName).Definition
                    $def.Settings.ExecutionTimeLimit = if ($h -eq 0 -and $m -eq 0) {
                        'PT0S'
                    } else {
                        "PT${h}H${m}M"
                    }
                    $userId    = $def.Principal.UserId
                    $logonType = $def.Principal.LogonType
                    $password  = $null
                    if ($logonType -in 1, 6) {
                        Write-Host "  Task runs as '$userId' with a stored password." -ForegroundColor DarkGray
                        $secpw    = Read-Host '  Enter password to save changes' -AsSecureString
                        $password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
                                        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secpw))
                    }
                    $svc.GetFolder($fp).RegisterTaskDefinition(
                        $t.TaskName, $def, 4, $userId, $password, $logonType) | Out-Null
                    $t = Get-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath
                    $display = if ($h -eq 0 -and $m -eq 0) { 'unlimited' } else { "${h}h ${m}m" }
                    Write-Host "  Execution time limit set to $display." -ForegroundColor Green
                } catch { Write-Host "  Error: $_" -ForegroundColor Red }
                Pause-ForUser
            }
        }
    } while ($true)
}

# ═══════════════════════════════════════════════════════════════════
#  RUN / EXPORT / IMPORT / CLONE
# ═══════════════════════════════════════════════════════════════════

function Invoke-RunTask {
    Write-Header 'Run a Task Now'
    $t = Select-Task 'run'
    if (!$t) { return }

    Write-Host "`n  $($t.TaskPath)$($t.TaskName)  [$($t.State)]" -ForegroundColor Yellow
    if ((Read-Valid 'Run this task now? [Y/N]' @('Y','y','N','n')) -in 'Y','y') {
        try {
            Start-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath
            Write-Host '  Task started.' -ForegroundColor Green
        } catch { Write-Host "  Error: $_" -ForegroundColor Red }
    } else { Write-Host '  Cancelled.' -ForegroundColor Gray }
    Pause-ForUser
}

function Export-TaskToXml {
    Write-Header 'Export Task to XML'
    $t = Select-Task 'export'
    if (!$t) { return }

    $default = "$env:USERPROFILE\Desktop\$($t.TaskName).xml"
    Write-Host "  Default save path: $default" -ForegroundColor DarkGray
    $path = (Read-Host '  Save path (blank = Desktop)').Trim()
    if (!$path) { $path = $default }

    try {
        Export-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath |
            Out-File -FilePath $path -Encoding UTF8
        Write-Host "  Exported to: $path" -ForegroundColor Green
    } catch { Write-Host "  Error: $_" -ForegroundColor Red }
    Pause-ForUser
}

function Import-TaskFromXml {
    Write-Header 'Import Task from XML'
    $path = (Read-Host '  Path to XML file').Trim()
    if (!$path) { Write-Host '  Cancelled.' -ForegroundColor Gray; Pause-ForUser; return }
    if (!(Test-Path $path)) {
        Write-Host '  File not found.' -ForegroundColor Red; Pause-ForUser; return
    }

    $xmlText = Get-Content $path -Raw
    $name    = (Read-Host '  New task name').Trim()
    if (!$name) { Write-Host '  Cancelled.' -ForegroundColor Gray; Pause-ForUser; return }

    $folder = (Read-Host '  Folder / path (default: \)').Trim()
    if (!$folder) { $folder = '\' }
    if (!$folder.StartsWith('\')) { $folder = "\$folder" }
    if ($folder -ne '\' -and !$folder.EndsWith('\')) { $folder = "$folder\" }

    # Parse XML to detect whether the task uses stored user credentials.
    # LogonType 'Password'(1) and 'InteractiveTokenOrPassword'(6) require a password at registration.
    try { $xmlDoc = [xml]$xmlText } catch {
        Write-Host "  Invalid XML: $_" -ForegroundColor Red; Pause-ForUser; return
    }
    $ns          = @{ ts = 'http://schemas.microsoft.com/windows/2004/02/mit/task' }
    $ltNode      = Select-Xml -Xml $xmlDoc -XPath '//ts:LogonType' -Namespace $ns
    $logonTypeStr = if ($ltNode) { $ltNode.Node.InnerText } else { 'Interactive' }
    $uidNode     = Select-Xml -Xml $xmlDoc -XPath '//ts:UserId'    -Namespace $ns
    $userId      = if ($uidNode) { $uidNode.Node.InnerText } else { '' }

    # Map XML string to COM enum value
    $logonTypeMap = @{
        'Password'                   = 1
        'S4U'                        = 2
        'InteractiveToken'           = 3
        'Group'                      = 4
        'ServiceAccount'             = 5
        'InteractiveTokenOrPassword' = 6
    }
    $logonTypeNum = if ($logonTypeMap.ContainsKey($logonTypeStr)) { $logonTypeMap[$logonTypeStr] } else { 3 }
    $needsCreds   = $logonTypeStr -in 'Password', 'InteractiveTokenOrPassword'

    try {
        if ($needsCreds) {
            Write-Host "  Task runs as '$userId' with stored credentials." -ForegroundColor DarkGray
            $secpw    = Read-Host '  Enter password for this account' -AsSecureString
            $password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
                            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secpw))
            $svc = New-Object -ComObject 'Schedule.Service'
            $svc.Connect()
            $fp = if ($folder -eq '\') { '\' } else { $folder.TrimEnd('\') }
            # RegisterTask(path, xml, flags=6 CREATE_OR_UPDATE, userId, password, logonType, sddl)
            $svc.GetFolder($fp).RegisterTask($name, $xmlText, 6, $userId, $password, $logonTypeNum, $null) | Out-Null
        } else {
            Register-ScheduledTask -TaskName $name -TaskPath $folder -Xml $xmlText -EA Stop | Out-Null
        }
        Write-Host "  Task '$name' imported successfully." -ForegroundColor Green
    } catch { Write-Host "  Error: $_" -ForegroundColor Red }
    Pause-ForUser
}

function Invoke-CloneTask {
    Write-Header 'Clone / Copy a Task'
    $t = Select-Task 'clone'
    if (!$t) { return }

    Write-Host "`n  Source: $($t.TaskPath)$($t.TaskName)" -ForegroundColor Yellow
    $newName = (Read-Host '  New task name').Trim()
    if (!$newName) { Write-Host '  Cancelled.' -ForegroundColor Gray; Pause-ForUser; return }

    $folder = (Read-Host "  Destination folder (blank = $($t.TaskPath))").Trim()
    if (!$folder) { $folder = $t.TaskPath }
    if (!$folder.StartsWith('\')) { $folder = "\$folder" }
    if ($folder -ne '\' -and !$folder.EndsWith('\')) { $folder = "$folder\" }

    # CIM may return LogonType as a string name ('Password') or an integer ('1') depending on
    # the Windows version — resolve to a numeric value so the credential check is always reliable.
    $ltRaw = "$($t.Principal.LogonType)"
    $ltMap = @{ 'None'=0; 'Password'=1; 'S4U'=2; 'InteractiveToken'=3; 'Group'=4; 'ServiceAccount'=5; 'InteractiveTokenOrPassword'=6 }
    $logonTypeNum = 0
    if (![int]::TryParse($ltRaw, [ref]$logonTypeNum)) {
        $logonTypeNum = if ($ltMap.ContainsKey($ltRaw)) { $ltMap[$ltRaw] } else { 3 }
    }
    $needsCreds = $logonTypeNum -in 1, 6

    try {
        $xmlText = Export-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath
        if ($needsCreds) {
            $userId   = $t.Principal.UserId
            Write-Host "  Task runs as '$userId' with stored credentials." -ForegroundColor DarkGray
            $secpw    = Read-Host '  Enter password for this account' -AsSecureString
            $password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
                            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secpw))
            $svc = New-Object -ComObject 'Schedule.Service'
            $svc.Connect()
            $fp = if ($folder -eq '\') { '\' } else { $folder.TrimEnd('\') }
            $svc.GetFolder($fp).RegisterTask($newName, $xmlText, 6, $userId, $password, $logonTypeNum, $null) | Out-Null
        } else {
            Register-ScheduledTask -TaskName $newName -TaskPath $folder -Xml $xmlText -EA Stop | Out-Null
        }
        Write-Host "  Task cloned as '$newName' in $folder." -ForegroundColor Green
    } catch { Write-Host "  Error: $_" -ForegroundColor Red }
    Pause-ForUser
}

# ═══════════════════════════════════════════════════════════════════
#  BULK ENABLE / DISABLE
# ═══════════════════════════════════════════════════════════════════

function Invoke-BulkToggle {
    param([string]$TargetState)   # 'Ready' = enable, 'Disabled' = disable

    $verb      = if ($TargetState -eq 'Ready') { 'enable'             } else { 'disable'              }
    $title     = if ($TargetState -eq 'Ready') { 'Bulk Enable Tasks'  } else { 'Bulk Disable Tasks'   }
    $fromState = if ($TargetState -eq 'Ready') { 'Disabled'           } else { 'Ready'                }

    Write-Header $title
    $kw = (Read-Host '  Filter keyword (blank = show all)').Trim()
    Write-Host '  Searching...' -ForegroundColor DarkGray

    $found = @(Get-ScheduledTask -EA SilentlyContinue | Where-Object {
        $_.State -eq $fromState -and (!$kw -or $_.TaskName -like "*$kw*" -or $_.TaskPath -like "*$kw*")
    })

    if ($found.Count -eq 0) {
        Write-Host "  No tasks found to $verb." -ForegroundColor Yellow
        Pause-ForUser; return
    }

    Write-Host "`n  $($found.Count) task(s) eligible:`n" -ForegroundColor Cyan
    for ($i = 0; $i -lt $found.Count; $i++) {
        Write-Host "  [$($i + 1)] $($found[$i].TaskPath)$($found[$i].TaskName)"
    }

    Write-Host "`n  Enter numbers to $verb (e.g.  1,3,5  or  2-6  or  ALL):" -ForegroundColor Yellow
    $raw = (Read-Host '  Selection').Trim()

    if (!$raw -or $raw -in '0','B','b') {
        Write-Host '  Cancelled.' -ForegroundColor Gray; Pause-ForUser; return
    }

    $indices = @()
    if ($raw -ieq 'ALL') {
        $indices = 1..$found.Count
    } else {
        foreach ($part in ($raw -split ',')) {
            $part = $part.Trim()
            if ($part -match '^(\d+)-(\d+)$') {
                $indices += [int]$Matches[1]..[int]$Matches[2]
            } elseif ($part -match '^\d+$') {
                $indices += [int]$part
            }
        }
    }
    $indices = $indices | Where-Object { $_ -ge 1 -and $_ -le $found.Count } | Sort-Object -Unique

    if ($indices.Count -eq 0) {
        Write-Host '  No valid selections.' -ForegroundColor Yellow; Pause-ForUser; return
    }

    if ((Read-Valid "$verb $($indices.Count) task(s)? [Y/N]" @('Y','y','N','n')) -in 'N','n') {
        Write-Host '  Cancelled.' -ForegroundColor Gray; Pause-ForUser; return
    }

    $ok = 0; $fail = 0
    foreach ($idx in $indices) {
        $task = $found[$idx - 1]
        try {
            if ($TargetState -eq 'Ready') {
                Enable-ScheduledTask  -TaskName $task.TaskName -TaskPath $task.TaskPath | Out-Null
            } else {
                Disable-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath | Out-Null
            }
            $ok++
        } catch {
            Write-Host "  Failed [$($task.TaskName)]: $_" -ForegroundColor Red
            $fail++
        }
    }
    $color = if ($fail -eq 0) { 'Green' } else { 'Yellow' }
    Write-Host "  Complete: $ok succeeded, $fail failed." -ForegroundColor $color
    Pause-ForUser
}

function Invoke-BulkEnable  { Invoke-BulkToggle 'Ready'    }
function Invoke-BulkDisable { Invoke-BulkToggle 'Disabled' }

# ═══════════════════════════════════════════════════════════════════
#  RUN-AS AUDIT
# ═══════════════════════════════════════════════════════════════════

function Show-RunAsAudit {
    Write-Header 'Run-As Account Audit'
    $kw = (Read-Host '  Filter keyword (blank = all tasks)').Trim()
    Write-Host '  Loading tasks...' -ForegroundColor DarkGray

    $tasks = @(Get-ScheduledTask -EA SilentlyContinue | Where-Object {
        !$kw -or $_.TaskName -like "*$kw*" -or $_.TaskPath -like "*$kw*"
    })

    if ($tasks.Count -eq 0) {
        Write-Host '  No tasks found.' -ForegroundColor Yellow
        Pause-ForUser; return
    }

    $rows = $tasks |
        Sort-Object { $_.Principal.UserId }, { $_.TaskPath }, { $_.TaskName } |
        ForEach-Object {
            [PSCustomObject]@{
                'Run As'    = $_.Principal.UserId
                'LogonType' = $_.Principal.LogonType
                'RunLevel'  = $_.Principal.RunLevel
                'State'     = $_.State
                'Task'      = $_.TaskName
                'Folder'    = $_.TaskPath
            }
        }

    Show-Paged -Rows $rows
}

# ═══════════════════════════════════════════════════════════════════
#  MAIN MENU LOOP
# ═══════════════════════════════════════════════════════════════════

do {
    Write-Header

    Write-Host '  LIST' -ForegroundColor DarkCyan
    Write-Host '   [1]  Enabled tasks'
    Write-Host '   [2]  Disabled tasks'
    Write-Host '   [3]  Running tasks'
    Write-Host
    Write-Host '  MANAGE' -ForegroundColor DarkCyan
    Write-Host '   [4]  Enable a task'
    Write-Host '   [5]  Disable a task'
    Write-Host '   [6]  Bulk enable tasks'
    Write-Host '   [7]  Bulk disable tasks'
    Write-Host '   [8]  Run a task now'
    Write-Host '   [9]  Stop a running task'
    Write-Host '   [10] Delete a task'
    Write-Host
    Write-Host '  CONFIGURE' -ForegroundColor DarkCyan
    Write-Host '   [11] Create a task'
    Write-Host '   [12] Edit a task'
    Write-Host '   [13] Clone / copy a task'
    Write-Host '   [14] Export task to XML'
    Write-Host '   [15] Import task from XML'
    Write-Host
    Write-Host '  INSPECT' -ForegroundColor DarkCyan
    Write-Host '   [16] Run history'
    Write-Host '   [17] Run-as account audit'
    Write-Host
    Write-Host '   [Q]  Quit'
    Write-Host

    $choice = Read-Valid 'Choice [1-17 or Q]' @(
        '1','2','3','4','5','6','7','8','9','10',
        '11','12','13','14','15','16','17','Q','q'
    )

    switch ($choice) {
        '1'  { Show-EnabledTasks   }
        '2'  { Show-DisabledTasks  }
        '3'  { Show-RunningTasks   }
        '4'  { Invoke-EnableTask   }
        '5'  { Invoke-DisableTask  }
        '6'  { Invoke-BulkEnable   }
        '7'  { Invoke-BulkDisable  }
        '8'  { Invoke-RunTask      }
        '9'  { Invoke-StopTask     }
        '10' { Invoke-DeleteTask   }
        '11' { New-TaskWizard      }
        '12' { Edit-TaskWizard     }
        '13' { Invoke-CloneTask    }
        '14' { Export-TaskToXml    }
        '15' { Import-TaskFromXml  }
        '16' { Show-TaskHistory    }
        '17' { Show-RunAsAudit     }
    }
} while ($choice -notin 'Q', 'q')

Write-Host "`n  Goodbye.`n" -ForegroundColor DarkGray
