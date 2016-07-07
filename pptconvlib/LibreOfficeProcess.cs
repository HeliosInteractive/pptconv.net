namespace pptconv
{
    using System;
    using System.IO;
    using System.Threading;
    using System.Diagnostics;

    public class LibreOfficeProcess
    {
        private readonly LibreOfficeFinder finder;

        public LibreOfficeProcess(LibreOfficeFinder sofficeFinder)
        {
            finder = sofficeFinder;
        }

        public LibreOfficeProcess()
            : this(new LibreOfficeFinder())
        {
            finder.Search();
        }

        public bool Convert(string presentation, string outdir, TimeSpan timeout)
        {
            if (finder == null || !finder.Found)
                return false;

            if (!File.Exists(presentation))
                return false;

            if (!Directory.Exists(outdir))
                return false;

            string outfile = string.Format("{0}.pdf", Path.GetFileNameWithoutExtension(presentation));
            string output = Path.Combine(outdir, outfile);

            if (File.Exists(output))
                File.Delete(output);

            if (LibreOfficeProcessHelper.IsRunning)
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
                string.Format("--convert-to pdf --outdir \"{0}\" \"{1}\"", outdir, presentation)
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
    }
}
