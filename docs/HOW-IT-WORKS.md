# Task Inspector: runtime, capabilities, and internals

## Runtime requirements

| Item | Requirement |
|---|---|
| Operating system | Windows 10 or 11, Windows Server 2016 or later |
| PowerShell | Windows PowerShell 5.1 (preinstalled on every supported Windows). PowerShell 7 is not used. |
| .NET | .NET Framework 4.x (preinstalled). WPF, `System.Management.Automation`, and the C# compiler all come from it. |
| Build | `%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe`. No SDK, Visual Studio, or NuGet packages. |
| Rights, local | Reading works as a standard user. Changing SYSTEM-owned tasks needs elevation (the app offers Relaunch as Admin). |
| Rights, remote | Local administrator on the target, reachable over WinRM (TCP 5985/5986) or DCOM/RPC (TCP 135 plus dynamic ports). |
| Task history | Only if the target's `Microsoft-Windows-TaskScheduler/Operational` log is enabled. |

The exe is about 115 KB. It is unsigned; SmartScreen may prompt when it is run from a
download or network share. There is no installer. The only files it writes are CSV
exports you ask for and the error log under `%LOCALAPPDATA%\TaskInspector`.

## What it can do

- List every scheduled task in every folder, on this computer or a remote one.
- Show last run time, result code with a decoded meaning, next run time, and a
  plain-English summary of each trigger including repetition, delay, and maintenance
  windows.
- Show what each task actually launches: the executable and arguments, or for
  COM-handler tasks the DLL or the Windows service that hosts the class.
- Flag problems in red: an action whose file, DLL, or COM class no longer exists, and
  SIDs in the task name, folder, or principal that no longer resolve to an account.
- Filter by view (enabled, disabled, running, trigger type, problems, non-Microsoft),
  by folder in a Task Scheduler Library tree, and by free text.
- Enable, disable, run, stop, and delete tasks, singly or in bulk, from the toolbar or a
  right-click menu. Sort order and selection are preserved across actions.
- Open a read-only Properties window with the same tabs as Task Scheduler, including
  History from the event log and the raw task XML.
- Open the action's file location in Explorer (through the admin share when remote).
- Export the inventory to CSV, interactively or headless from the command line for
  scripting and scheduled collection.

## How it works

### Process model

`TaskInspector.exe` is a tiny C# program (`src/Program.cs`). It carries
`TaskInspector.ps1` as an embedded resource and, on start, creates an in-process
PowerShell runspace on an STA thread and runs the script inside it. There is no
`powershell.exe` child process and no console window. The script's `param()` block
receives the command-line arguments because the host builds a `& { script } args`
invocation that quotes values and passes `-Name` tokens through as parameter names.

In headless mode the host attaches to the parent console so the summary line is
visible when run from a terminal. Because the exe is a GUI-subsystem program, that
output cannot be captured by a pipe; automation should use the exit code and the CSV.
The script records its intended exit code in `$global:TaskInspectorExit`, which the host
returns, because `exit` inside a hosted runspace does not propagate on its own.

### Data collection

Each refresh calls `Get-ScheduledTask`, then per task `Get-ScheduledTaskInfo` (last and
next run, result) and `Export-ScheduledTask` (XML). All three accept a `CimSession`, so
the same code serves local and remote targets; `$script:CimArgs` is empty for local and
`@{CimSession=...}` for remote, and every task cmdlet is splatted with it.

For every action the script resolves what will run:

- **Exec actions**: environment variables are expanded (against the target's environment
  when remote), doubled backslashes collapsed, and bare names such as `powershell.exe`
  resolved through PATH locally or System32 remotely. File existence and version-info
  publisher come from the file system locally or `CIM_DataFile` remotely.
- **COM-handler actions**: the CLSID is looked up under `HKLM\SOFTWARE\Classes\CLSID`
  and the WOW6432Node mirror. `InprocServer32` or `LocalServer32` gives the DLL or exe.
  If neither exists but the class has an `AppID` with a `LocalService`, the class is
  service-hosted and the service's `ImagePath` is shown instead; that is normal, not a
  defect. Registry reads use `Microsoft.Win32.Registry` locally (the `HKLM:` provider is
  about 0.2 s per miss under the Classes hive, which made the first version take 44 s)
  and `StdRegProv` over CIM remotely.
- **SIDs**: any `S-1-...` in the task name, folder, or principal is translated with
  `SecurityIdentifier.Translate` locally or `Win32_SID` remotely. Untranslatable SIDs
  are flagged.

Authenticode signatures are checked only on demand in the details pane, and only
locally, because each check costs about half a second.

### Result codes

The Result column's tooltip comes from `Get-ResultDetail`. A built-in table covers the
Task Scheduler `SCHED_S_*` and `SCHED_E_*` codes and the Windows codes that the system
message table describes poorly. Anything else goes through `Win32Exception` (for
`0x8007xxxx` and small values) or `Marshal.GetExceptionForHR` (for other HRESULTs).
Small values are also described as process exit codes, since a program returning 1 is
stored the same way as Win32 error 1.

### Remote connection

`Connect-Target` tries `New-CimSession` over WSMan first, then DCOM, verifying each with
a `Win32_OperatingSystem` query. On success it also reads the target's Windows directory,
system drive, and SYSTEM-scope environment variables so paths like `%SystemRoot%` expand
correctly. On failure the UI asks once for credentials, then falls back to the local
computer and reports the reason in the status bar. Headless mode does not fall back; it
exits with code 2 so a collection script cannot silently record the wrong machine.

### User interface

The window is WPF loaded from inline XAML: a toolbar, a left pane with the views list
and the folder tree, a `DataGrid`, and a details/XML pane. Filtering rebuilds the grid's
`ObservableCollection` but first captures the current `SortDescriptions` and selected
task paths and reapplies them, which is why actions do not reorder the grid. Problem
rows use a `DataTrigger` on `HasProblem`. The Properties window is a second XAML
window populated from the raw CIM task object.

### Build

`Build.ps1` generates `src/TaskInspector.ico` on first run (drawn with System.Drawing),
then invokes `csc.exe` with `/target:winexe`, the icon, the manifest
(`asInvoker`, per-monitor DPI aware), the embedded script resource, and references to
`System.Management.Automation` and `System.Windows.Forms`. The compiler is C# 5, so the
host avoids newer language features. Output goes to `dist\`, which is not committed;
the GitHub Actions workflow builds it on every push and publishes the exe as an
artifact.

### Logging

`Write-Log` appends to `%LOCALAPPDATA%\TaskInspector\TaskInspector.log` with a timestamp,
level (INFO, WARN, ERROR), message, the exception text, and the script line. It is
called from every catch block that matters: connection attempts per protocol, per-task
`Get-ScheduledTaskInfo` and `Export-ScheduledTask` failures, task actions, history
lookups, headless export, the initial load, and an outer guard around the window. The
C# host writes FATAL lines to the same file for errors that escape the runspace. The log
rotates once at 2 MB.

## Known limits

- "Accesses" is static: it lists URLs and hostnames written into the task definition,
  not what the program connects to at run time.
- Signer is not checked on remote computers.
- Task actions using user-scoped variables such as `%LOCALAPPDATA%` cannot be expanded
  on a remote computer and are shown as "not checked" rather than as a problem.
- History depends on the target's Task Scheduler operational log being enabled.
- Remote mode was verified against a machine's own IP over WinRM; a different target
  may expose environment differences not seen in testing.
