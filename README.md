# Task Inspector

A single-file WPF viewer for every Windows Scheduled Task on the machine.

Columns: folder, name, state, last run, last result, next run, a plain-English
trigger summary (daily/weekly/logon/boot/event/idle, repetition interval, delay,
automatic-maintenance period), the action (executables, with COM-handler actions
resolved to their DLL through the registry), the file's description and
Authenticode signer, and any URLs or hostnames found in the arguments or task XML.

## Run

Double-click `TaskInspector.vbs` (no console window), or:

    powershell -NoProfile -ExecutionPolicy Bypass -STA -File TaskInspector.ps1

Requires Windows PowerShell 5.1 or later. No modules, no install.

## Controls

- **Left pane** picks a view: all, enabled, disabled, running, or by trigger type
  (logon, startup, idle, schedule, event, on demand), plus Problems and Non-Microsoft.
- **Properties** (right-click, or double-click a row) opens a read-only window with the
  same tabs as Task Scheduler: General, Triggers, Actions, Conditions, Settings,
  History (from the TaskScheduler/Operational log, if history is enabled), and XML.
- Column sort order and the current selection survive Refresh and every action.
- **Filter** matches folder, name, action, arguments, publisher, accesses, triggers, author.
- **Hide \Microsoft\ tasks** is on by default; untick to see the full tree.
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
- **Export CSV** writes the full inventory.
- The bottom pane shows details and the raw task XML for the selected row.

## Limits

"Accesses" is static: it reports targets named in the task definition, not what the
program connects to at run time. A telemetry task whose executable holds its endpoints
internally will show none. The Publisher column (signer and file description) is the
better hint for those.
