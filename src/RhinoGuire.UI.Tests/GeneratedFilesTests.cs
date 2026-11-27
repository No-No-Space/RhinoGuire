using System.Diagnostics;

namespace RhinoGuire.UI.Tests;

/// <summary>Generated files match their sources (the generators' own <c>--check</c> mode).</summary>
public class GeneratedFilesTests
{
    [Theory]
    [InlineData("tools/build_aaj_skin.py")]
    [InlineData("tools/build_lucide_icons.py")]
    public void Generated_files_are_up_to_date(string script)
    {
        var (code, output) = RunPython(Repo.Path_(script.Split('/')), "--check");
        Assert.True(code == 0, $"python {script} --check failed (rerun the script and commit its output):\n{output}");
    }

    private static (int Code, string Output) RunPython(string script, string args)
    {
        foreach (string exe in new[] { "python", "py", "python3" })
        {
            try
            {
                var psi = new ProcessStartInfo(exe, $"\"{script}\" {args}")
                {
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    WorkingDirectory = Repo.Root,
                };
                using var p = Process.Start(psi)!;
                string output = p.StandardOutput.ReadToEnd() + p.StandardError.ReadToEnd();
                p.WaitForExit();
                return (p.ExitCode, output);
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // not on PATH, try the next name
            }
        }
        throw new InvalidOperationException("Python 3 is needed for this test and was not found on PATH.");
    }
}
