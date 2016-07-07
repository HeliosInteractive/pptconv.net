namespace pptconv
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Threading;
    using System.Diagnostics;

    public class LibreOfficeProcess
    {
        private readonly LibreOfficeFinder finder;

        public LibreOfficeProcess(LibreOfficeFinder sofficeFinder)
        {
            finder = sofficeFinder;
        }

        public bool Convert(string presentation, string outdir, TimeSpan timeout)
        {
            if (finder == null || !finder.Found)
                return false;

            if (!File.Exists(presentation))
                return false;

            if (!Directory.Exists(outdir))
                return false;

            string output = Path.Combine(outdir, Path.GetFileName(presentation));

            if (File.Exists(output))
                File.Delete(output);

            if (IsRunning)
                LibreOfficeProcessHelper.Kill();

            string[] arguments = new string[]
            {
                "--nologo",
                "--headless",
                "--invisible",
                "--norestore",
                "--nodefault",
                "--nocrashreport",
                "--nofirststartwizard",
                string.Format("--outdir \"{0}\"", outdir),
                string.Format("--convert-to pdf \"{0}\"", presentation)
            };

            ProcessStartInfo sinfo = new ProcessStartInfo
            {
                UseShellExecute = false,
                FileName = finder.SOfficeBinaryPath,
                Arguments = string.Join(" ", arguments),
                WorkingDirectory = Path.GetDirectoryName(finder.SOfficeBinaryPath),
            };

            using (Process sofficeProcess = new Process { StartInfo = sinfo })
            {
                sofficeProcess.Start();
                sofficeProcess.WaitForExit((int)timeout.TotalMilliseconds);
            }

            Thread.Sleep(75);
            return File.Exists(output);
        }

        public static bool IsRunning
        {
            get { return Process.GetProcessesByName("soffice.bin").Count() > 0; }
        }
    }
}
