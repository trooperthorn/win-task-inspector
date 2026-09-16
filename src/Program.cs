// Task Inspector host. Embeds TaskInspector.ps1 as a resource and runs it in an
// in-process PowerShell runspace on an STA thread (WPF needs STA). No console.
using System;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

[assembly: AssemblyTitle("Task Inspector")]
[assembly: AssemblyDescription("Viewer for Windows Scheduled Tasks")]
[assembly: AssemblyCompany("Sean T. Martinez")]
[assembly: AssemblyProduct("Task Inspector")]
[assembly: AssemblyCopyright("MIT")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        int exitCode = 0;
        // Run on an explicit STA thread with a large stack; the WPF DataGrid and
        // the PowerShell parser both like headroom.
        var t = new Thread(() => { exitCode = Run(args); }, 16 * 1024 * 1024);
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();
        return exitCode;
    }

    private static int Run(string[] args)
    {
        string script;
        var asm = Assembly.GetExecutingAssembly();
        using (var s = asm.GetManifestResourceStream("TaskInspector.ps1"))
        {
            if (s == null)
            {
                MessageBox.Show("Embedded script missing. Rebuild with Build.ps1.", "Task Inspector", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            // Let the script find its own exe (used by "Relaunch as Admin").
            rs.SessionStateProxy.SetVariable("TaskInspectorExe", asm.Location);
            using (var ps = PowerShell.Create())
            {
                ps.Runspace = rs;
                ps.AddScript(script, false);
                try
                {
                    ps.Invoke();
                    if (ps.HadErrors && ps.Streams.Error.Count > 0)
                    {
                        var sb = new System.Text.StringBuilder();
                        foreach (var e in ps.Streams.Error)
                        {
                            sb.AppendLine(e.ToString());
                            if (e.InvocationInfo != null) sb.AppendLine("  at line " + e.InvocationInfo.ScriptLineNumber);
                        }
                        MessageBox.Show(sb.ToString(), "Task Inspector: script errors", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return 1;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Task Inspector: fatal error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 1;
                }
            }
        }
        return 0;
    }
}
