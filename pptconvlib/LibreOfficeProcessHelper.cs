namespace pptconv
{
    using System;
    using System.Linq;
    using System.Management;
    using System.Diagnostics;

    public static class LibreOfficeProcessHelper
    {
        public static void KillProcess(int pid)
        {
            string query = string.Format("Select * From Win32_Process Where ParentProcessID={0}", pid);

            using (ManagementObjectSearcher processSearcher = new ManagementObjectSearcher(query))
            using (ManagementObjectCollection processCollection = processSearcher.Get())
            {
                try
                {
                    Process proc = Process.GetProcessById(pid);
                    if (!proc.HasExited) proc.Kill();
                }
                catch (ArgumentException)
                {
                    // Process already exited.
                }

                if (processCollection != null)
                {
                    foreach (ManagementObject mo in processCollection)
                    {
                        KillProcess(Convert.ToInt32(mo["ProcessID"]));
                    }
                }
            }
        }

        public static void KillProcess(string name)
        {
            Process.GetProcessesByName(name).ToList().ForEach(proc =>
            {
                try { KillProcess(proc.Id); }
                catch (InvalidOperationException) { /* zombie process */ }
            });
        }

        public static void Kill()
        {
            KillProcess("soffice.bin");
        }
    }
}
