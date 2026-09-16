// Task Inspector host. Embeds TaskInspector.ps1 as a resource and runs it in an
// in-process PowerShell runspace on an STA thread (WPF needs STA). No console of its
// own; in headless mode (-Csv) it attaches to the parent console so output is visible.
using System;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

[assembly: AssemblyTitle("Task Inspector")]
[assembly: AssemblyDescription("Viewer for Windows Scheduled Tasks, local or remote")]
[assembly: AssemblyCompany("Sean T. Martinez")]
[assembly: AssemblyProduct("Task Inspector")]
[assembly: AssemblyCopyright("MIT")]
[assembly: AssemblyVersion("1.1.0.0")]
[assembly: AssemblyFileVersion("1.1.0.0")]

internal static class Program
{
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool AttachConsole(int dwProcessId);
    private const int ATTACH_PARENT_PROCESS = -1;

    [STAThread]
    private static int Main(string[] args)
    {
        bool headless = false;
        foreach (var a in args)
        {
            var l = a.ToLowerInvariant();
            if (l == "-csv" || l == "-help" || l == "-?" || l == "/?" || l == "--help") headless = true;
        }
        if (args.Length >= 2 && !args[0].StartsWith("-") && !args[1].StartsWith("-")) headless = true; // <host> <out.csv>
        if (headless && AttachConsole(ATTACH_PARENT_PROCESS))
        {
            Console.SetOut(new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true });
            Console.SetError(new StreamWriter(Console.OpenStandardError()) { AutoFlush = true });
        }

        int exitCode = 0;
        var t = new Thread(() => { exitCode = Run(args, headless); }, 16 * 1024 * 1024);
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();
        return exitCode;
    }

    private static int Run(string[] args, bool headless)
    {
        string script;
        var asm = Assembly.GetExecutingAssembly();
        using (var s = asm.GetManifestResourceStream("TaskInspector.ps1"))
        {
            if (s == null)
            {
                Fail("Embedded script missing. Rebuild with Build.ps1.", headless);
                return 2;
            }
            using (var r = new StreamReader(s)) script = r.ReadToEnd();
        }

        var iss = InitialSessionState.CreateDefault();
        iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;
        iss.ApartmentState = ApartmentState.STA;
        iss.ThreadOptions = PSThreadOptions.UseCurrentThread;

        using (var rs = RunspaceFactory.CreateRunspace(iss))
        {
            rs.Open();
            rs.SessionStateProxy.SetVariable("TaskInspectorExe", asm.Location);
            rs.SessionStateProxy.SetVariable("TaskInspectorScript", script);
            using (var ps = PowerShell.Create())
            {
                ps.Runspace = rs;
                // Invoke the script as a script block so its param() block binds the
                // command-line arguments the way powershell.exe -File does: tokens that
                // look like -Name are parameter names, everything else is a quoted value.
                var cmd = new System.Text.StringBuilder("& ([scriptblock]::Create($TaskInspectorScript))");
                foreach (var a in args)
                {
                    if (System.Text.RegularExpressions.Regex.IsMatch(a, @"^-[A-Za-z?][A-Za-z0-9?]*:?$") || a == "/?")
                        cmd.Append(' ').Append(a == "/?" ? "-Help" : a);
                    else
                        cmd.Append(" '").Append(a.Replace("'", "''")).Append('\'');
                }
                ps.AddScript(cmd.ToString(), false);
                try
                {
                    ps.Invoke();
                    int code = ExitCodeFrom(rs);
                    if (ps.HadErrors && ps.Streams.Error.Count > 0)
                    {
                        var sb = new System.Text.StringBuilder();
                        foreach (var e in ps.Streams.Error)
                        {
                            sb.AppendLine(e.ToString());
                            if (e.InvocationInfo != null) sb.AppendLine("  at line " + e.InvocationInfo.ScriptLineNumber);
                        }
                        Fail(sb.ToString(), headless);
                        return 1;
                    }
                    return code;
                }
                catch (Exception ex)
                {
                    // 'exit N' inside the script surfaces here as an ExitException-derived error
                    var inner = ex.InnerException ?? ex;
                    if (inner.GetType().Name == "ExitException") return ExitCodeFrom(rs);
                    Fail(ex.ToString(), headless);
                    return 1;
                }
            }
        }
    }

    private static int ExitCodeFrom(Runspace rs)
    {
        // The script sets $global:TaskInspectorExit before calling exit.
        var v = rs.SessionStateProxy.PSVariable.GetValue("TaskInspectorExit");
        if (v is int) return (int)v;
        var l = rs.SessionStateProxy.PSVariable.GetValue("LASTEXITCODE");
        return l is int ? (int)l : 0;
    }

    private static void Fail(string message, bool headless)
    {
        if (headless)
        {
            try { Console.Error.WriteLine("Task Inspector: " + message); } catch { }
        }
        else
        {
            MessageBox.Show(message, "Task Inspector", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
