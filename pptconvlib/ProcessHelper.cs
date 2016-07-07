namespace pptconv
{
    using System.Diagnostics;

    internal class ProcessHelper
    {
        public static void ExecuteScript(string script_path, string command_line = null, bool headless = false)
        {
            using (Process process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = script_path,
                    Arguments = command_line,
                    CreateNoWindow = headless,
                    UseShellExecute = true,
                }
            })
            {
                process.Start();
                process.WaitForExit();
            }
        }

        public static void TaskKill(string name)
        {
            ExecuteScript("taskkill", string.Format("/F /T /IM \"{0}\"", name), true);
        }
    }
}
