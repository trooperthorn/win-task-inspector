# Task Inspector

A viewer for every Windows Scheduled Task on this computer or a remote one.

Columns: folder, name, state, last run, last result, next run, a plain-English
trigger summary (daily/weekly/logon/boot/event/idle, repetition interval, delay,
automatic-maintenance period), the action (executables, with COM-handler actions
resolved to their DLL or hosting service through the registry), the file's description
and publisher, and any URLs or hostnames found in the arguments or task XML.

## Run

Run `dist\TaskInspector.exe` from the repo or a Release. To rebuild after changes:

    powershell -NoProfile -ExecutionPolicy Bypass -File Build.ps1

`Build.ps1` compiles `src\Program.cs` with the C# compiler that ships in .NET
Framework 4.x (no SDK or module to install), embeds `TaskInspector.ps1` as a resource,
and stamps the icon, manifest, and version info. The exe hosts PowerShell in-process on
an STA thread: no console window, no child powershell.exe, and "Relaunch as Admin"
re-launches the exe itself. Rebuild after editing the script. The built exe is committed
in `dist\` and attached to every GitHub Release, so you can also just download it.

The script still runs on its own for development, with the same arguments:

    powershell -NoProfile -ExecutionPolicy Bypass -STA -File TaskInspector.ps1

or double-click `TaskInspector.vbs`. Requires Windows PowerShell 5.1 or later.

## Local and remote

The app opens against this computer. Type a hostname, FQDN, or IP address in the
**Computer** box and press Connect (or Enter) to switch to another machine, as the
Task Scheduler MMC snap-in does. The connection is a CIM session over WinRM with a
DCOM fallback, so the target needs WinRM enabled (`Enable-PSRemoting`) or DCOM open in
its firewall, and you need administrator rights there. If the first attempt fails you
are asked once for credentials; if that fails too, the app **falls back to the local
computer** and explains why in the status bar (hover it for the full error). Nothing
blocks on a dialog.

Everything works remotely: the inventory, the Problem checks (files through
`CIM_DataFile`, COM classes and services through `StdRegProv`, SIDs through
`Win32_SID`), enable/disable/run/stop/delete, Properties including History, and
"Open file location" (via the target's admin share). Two differences: the Authenticode
signer is not checked remotely, and task actions that use user-scoped variables such as
`%LOCALAPPDATA%` cannot be expanded on another machine, so they show as "not checked"
rather than as a problem.

## Command line

    TaskInspector.exe                              window, this computer
    TaskInspector.exe <host>                       window, connected to <host>
    TaskInspector.exe <host> <out.csv>             headless: write the inventory to CSV, exit
    TaskInspector.exe . <out.csv>                  headless export of this computer
    TaskInspector.exe -ComputerName <host> -Csv <out.csv> [-Credential DOMAIN\user]
    TaskInspector.exe -Help

Headless mode opens no window. The CSV has a leading `Computer` column plus every grid
column. The exit code is 0 on success, 2 if the target cannot be reached, 1 on any
other failure. Output is written to the console the exe was started from when there is
one; when captured by a script, rely on the exit code and the CSV. With `-Credential`
only the user name is given; the password is prompted for in a secure dialog.

## Controls

- **Left pane** picks a view: all, enabled, disabled, running, or by trigger type
  (logon, startup, idle, schedule, event, on demand), plus Problems and Non-Microsoft.
  Below it, **Task Scheduler Library** is the folder tree. Selecting a folder shows its
  tasks and every subfolder's tasks by default; the vertical-dots button beside the
  heading switches to showing only the selected folder, like Task Scheduler. Counts in
  the tree include subfolders. Drag the divider to resize the pane.
- **Properties** (right-click, or double-click a row) opens a read-only window with the
  same tabs as Task Scheduler: General, Triggers, Actions, Conditions, Settings,
  History (from the TaskScheduler/Operational log, if history is enabled), and XML.
- Column sort order and the current selection survive Refresh and every action.
- **Filter** matches folder, name, action, arguments, publisher, accesses, triggers, author.
- **Hide \Microsoft\ tasks** is an extra filter, off by default; the tree is the main scope.
- **Only with URLs/hosts** narrows to tasks whose definition names a network target.
- **Disable / Enable / Run now** act on the selected rows (multi-select with Ctrl/Shift).
  SYSTEM-owned tasks need an elevated session; use **Relaunch as Admin**.
- **Right-click a row** for Enable, Disable, Run now, Stop, Open file location (Explorer
  with the action's executable selected, or the task definition under
  `C:\Windows\System32\Tasks` if the file is gone), Copy task path, Copy command line,
  and **Delete task** (asks for confirmation; cannot be undone).
- **Red rows** have a problem: the action's executable, DLL, or COM class no longer
  exists, or the task name, folder, or run-as principal carries a SID that no longer
  resolves to an account (a leftover from a deleted user). The Problem column says
  which. The **Problems** view on the left filters to those rows.
- **Result column** shows the last run's return code; hover it for the meaning. Task
  Scheduler codes (0x41301, 0x80041324, ...) and common Windows codes (0x80070005,
  0x800710E0, ...) come from a built-in table; anything else is looked up through the
  Windows message table, and small values are also explained as process exit codes.
- **Export CSV** writes the full inventory.
- The bottom pane shows details and the raw task XML for the selected row.

## Documentation

- [docs/HOW-IT-WORKS.md](docs/HOW-IT-WORKS.md): runtime requirements, everything the tool
  can do, and how each part works (process model, data collection, remote transport,
  result decoding, build).
- The `build` workflow compiles the exe on every push and uploads it as an artifact;
  pushing a `v*` tag attaches it to a GitHub Release.

## Error log

Every failure is appended to `%LOCALAPPDATA%\TaskInspector\TaskInspector.log`: connection
attempts and why they failed, per-task read errors, enable/disable/run/delete failures,
history lookups, headless export errors, and any fatal error from the host. Each run adds
a Start line with the user, elevation, PowerShell version, and arguments, and an Exit
line. The file rotates to `.log.1` at 2 MB. The **Log** button on the toolbar opens it.

## Limits

"Accesses" is static: it reports targets named in the task definition, not what the
program connects to at run time. A telemetry task whose executable holds its endpoints
internally will show none. The Publisher column (signer and file description) is the
better hint for those.
