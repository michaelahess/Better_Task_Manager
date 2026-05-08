# Better Task Manager
Faster, more focused, BETTER, management of Windows Tasks.

An interactive, menu-driven Windows Task Scheduler manager for PowerShell 5.1+. Replaces clicking through the Task Scheduler GUI with a fast keyboard-driven terminal interface — list, search, create, edit, clone, export/import, and act on tasks without leaving the console. Now with network support!

> **Requires:** PowerShell 5.1+, Windows 8.1 / Server 2012 R2 or newer. Run as Administrator for full functionality.

I despise Microsoft's Task Scheduler. It's slow, laggy, hard to read arguments. It just sucks.

This is an easier to use (in my opinion) way of doing task management.

---

## Quick Start

```powershell
# Right-click PowerShell → Run as Administrator
.\TaskSchedulerManager.ps1
```

---

## Main Menu

```
  ╔═══════════════════════════════════════════════╗
  ║              Better Task Manager              ║
  ╚═══════════════════════════════════════════════╝

  LIST
   [1]  Enabled tasks
   [2]  Disabled tasks
   [3]  Running tasks

  MANAGE
   [4]  Enable a task          [8]  Run a task now
   [5]  Disable a task         [9]  Stop a running task
   [6]  Bulk enable tasks      [10] Delete a task
   [7]  Bulk disable tasks

  CONFIGURE
   [11] Create a task          [14] Export task to XML
   [12] Edit a task            [15] Import task from XML
   [13] Clone / copy a task

  INSPECT
   [16] Run history
   [17] Run-as account audit

  REMOTE
   [18] Connect to remote machine

   [Q]  Quit
```

---

## List Views

All three list views (Enabled, Disabled, Running) support a **keyword filter** at load time and full **inline row actions** — no need to return to the main menu.

```
  #  Name                 Folder       State    Last Run          Next Run          Exit Code
  -  ----                 ------       -----    --------          --------          ---------
  1  DailyBackupJob       \Corp\       Ready    2026-05-07 22:00  2026-05-08 22:00  0x00000000
  2  WeeklyReportExport   \Corp\       Ready    2026-05-04 06:00  2026-05-11 06:00  0x00000000
  3  NightlyDBCleanup     \Corp\       Ready    2026-05-07 03:00  2026-05-08 03:00  0x00041306
  4  ADSyncPulse          \Corp\IT\    Ready    2026-05-08 08:15  2026-05-08 09:15  0x00000000
  5  CertRenewalCheck     \Corp\IT\    Ready    2026-05-01 00:00  2026-06-01 00:00  0x00000000

  Page 1/1  ·  5 total item(s)
  [B]ack
  Row actions: [#] Detail  [#E]dit  [#N]able  [#D]isable  [#R]un  [#S]top  [#X] Delete  [#H]istory
  >
```

### Row Action Reference

| Input | Action |
|-------|--------|
| `3` | View full detail for row 3 |
| `3E` | Open Edit wizard for row 3 |
| `3N` | Enable task on row 3 |
| `3D` | Disable task on row 3 |
| `3R` | Run task on row 3 immediately |
| `3S` | Stop task on row 3 |
| `3X` | Delete task on row 3 (requires typing `DELETE`) |
| `3H` | Show run history for row 3 |

Spacing and case are flexible — `3 E` and `3E` both work.

---

## Task Detail

Drill into any row with `[#]` to see the full task definition, then jump directly to Edit with `[E]`.

```
  ── \Corp\DailyBackupJob ──

  Description : Backs up Corp file shares to NAS nightly at 10 PM
  State       : Ready
  Run As      : CORP\svc_backup  [Highest]

  Trigger(s)  :
              · Daily  @  2026-01-01T22:00:00  [Enabled: True]

  Action(s)   :
              · Execute : C:\Scripts\Backup-FileShares.ps1
                Args    : -Target \\nas01\backups -Compress

  Settings    :
              · Exec time limit    : PT4H0M
              · Multiple instances : IgnoreNew
              · Start when missed  : True
              · Hidden             : False

  Last Run    : 2026-05-07 22:00:11
  Next Run    : 2026-05-08 22:00:00
  Last Result : 0x00000000  (Success)

  [E] Edit this task   [Enter] Back
```

---

## Run History

Matches the layout of Event Viewer's History tab — one row per event, sortable by Correlation Id to trace a single run end-to-end.

```
  History — \Corp\DailyBackupJob  (12 event(s))

  Date and Time         E    Task Category               Operational  Correlation Id
  -------------         -    -------------               -----------  --------------
  2026-05-07 22:04:51   201  Action completed            Info         {3f2a1c08-7b...
  2026-05-07 22:04:51   102  Task completed              Info         {3f2a1c08-7b...
  2026-05-07 22:00:14   200  Action started              Info         {3f2a1c08-7b...
  2026-05-07 22:00:12   129  Created Task Process        Info         {3f2a1c08-7b...
  2026-05-07 22:00:11   100  Task Started                Info         {3f2a1c08-7b...
  2026-05-07 22:00:11   107  Task triggered on schedule  Info         {3f2a1c08-7b...
  2026-05-06 22:03:27   201  Action completed            Info         {9d0e4f21-2a...
  ...
```

> History requires the Task Scheduler operational log to be enabled:  
> **Task Scheduler → Action → Enable All Tasks History**

---

## Edit Wizard

Reached via menu option `[12]`, the `[E]` shortcut from any detail screen, or `[#E]` directly from a list row.

```
  ── \Corp\WeeklyReportExport ──

  What would you like to edit?
   [1] Replace trigger
   [2] Replace action
   [3] Change run-as account
   [4] Update description
   [5] Change execution time limit
   [B] Back to main menu
```

---

## Bulk Enable / Disable

Options `[6]` and `[7]` let you filter tasks and act on multiple at once.

```
  Bulk Enable Tasks

  Filter keyword (blank = show all): Report

  14 task(s) eligible:

  [1] \Corp\MonthlyFinanceReport
  [2] \Corp\WeeklyReportExport
  [3] \Corp\IT\ADGroupAuditReport
  ...

  Enter numbers to enable (e.g.  1,3,5  or  2-6  or  ALL):
  > 1-3

  enable 3 task(s)? [Y/N]: Y
  Complete: 3 succeeded, 0 failed.
```

---

## Export / Import / Clone

| Option | Description |
|--------|-------------|
| **[14] Export to XML** | Saves the task definition as XML (defaults to Desktop) |
| **[15] Import from XML** | Registers a task from an XML file; prompts for credentials if the task uses stored user credentials |
| **[13] Clone / Copy** | Duplicates a task under a new name/folder; handles stored credentials automatically |

---

## Run-As Account Audit

Option `[17]` lists every task alongside its run-as account, logon type, and run level — useful for security reviews or identifying service account dependencies.

```
  Run As              LogonType        RunLevel  State     Task                   Folder
  ------              ---------        --------  -----     ----                   ------
  CORP\svc_backup     Password         Highest   Ready     DailyBackupJob         \Corp\
  CORP\svc_reports    Password         Highest   Ready     WeeklyReportExport     \Corp\
  NETWORK SERVICE     ServiceAccount   Limited   Ready     CertRenewalCheck       \Corp\IT\
  SYSTEM              ServiceAccount   Highest   Ready     NightlyDBCleanup       \Corp\
```

---

## Connect to Remote Machine

Option `[18]` allows you to connect to a remote machine and perform all the same tasks on it.

Enter a full FQDN or IP address to connect and either authenticate as current user, or use alternative credentials.

```
Remote Connection

  Target machine (FQDN or IP address): testserver.domain.com
  Use alternate credentials? [Y/N]: y
  Username (DOMAIN\user or user@domain): domain\itsmemario
  Password: ****************
  Connecting to testserver.domain.com...
  Connected to testserver.domain.com
```
The home page will now update:

```
Connected to: testserver.domain.com

  LIST
   [1]  Enabled tasks
   [2]  Disabled tasks
   [3]  Running tasks

  MANAGE
   [4]  Enable a task          [8]  Run a task now
   [5]  Disable a task         [9]  Stop a running task
   [6]  Bulk enable tasks      [10] Delete a task
   [7]  Bulk disable tasks

  CONFIGURE
   [11] Create a task          [14] Export task to XML
   [12] Edit a task            [15] Import task from XML
   [13] Clone / copy a task

  INSPECT
   [16] Run history
   [17] Run-as account audit

  REMOTE
   [18] Remote connection     (connected: testserver.domain.com)

   [Q]  Quit

  Choice [1-18 or Q]:
```

Now if you go back to option ```18```:

```
Connected to: testserver.domain.com

  Remote Connection

  Currently connected to: testserver.domain.com
   [1] Connect to a different machine
   [2] Disconnect (return to local)
   [B] Back
  Choice [1-2 or B]:
```

---

## Compatibility

| Windows Version | Support |
|----------------|---------|
| Windows 10 / 11 | Full |
| Server 2016 / 2019 / 2022 | Full |
| Windows 8.1 / Server 2012 R2 | Supported (less tested) |
| Windows 7 / Server 2008 R2 | Not supported (`ScheduledTasks` module unavailable) |

---

## Notes

- All credential-sensitive operations (create, edit, import, clone) use the Windows Task Scheduler COM API (`Schedule.Service`) to pass stored credentials safely — plaintext is never written to disk.
- The running task list includes an `[R] Refresh` option to re-poll without returning to the main menu.
- After inline row actions (enable/disable/run/stop/delete) the list data is not automatically refreshed — press `B` and re-enter the list to see updated state.
