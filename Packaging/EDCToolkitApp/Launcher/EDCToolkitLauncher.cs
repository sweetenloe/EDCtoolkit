using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class EDCToolkitLauncher
{
    [STAThread]
    private static void Main()
    {
        try
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var scriptPath = Path.Combine(baseDir, "Scripts", "EDCtoolkit", "EDCtoolkit.GUI.ps1");
            if (!File.Exists(scriptPath))
            {
                MessageBox.Show("Missing GUI script: " + scriptPath, "EDC Toolkit", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var powerShellPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
            if (!File.Exists(powerShellPath))
            {
                MessageBox.Show("Windows PowerShell was not found.", "EDC Toolkit", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var args = "-NoLogo -NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\" -Theme Dark";
            var psi = new ProcessStartInfo
            {
                FileName = powerShellPath,
                Arguments = args,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = baseDir
            };

            Process.Start(psi);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "EDC Toolkit", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
