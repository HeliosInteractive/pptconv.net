namespace pptconvtest
{
    using pptconv;

    using System.Diagnostics;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class LibreOfficeProcessTest
    {
        [TestMethod]
        public void RunningInstanceTest()
        {
            LibreOfficeFinder finder = new LibreOfficeFinder();

            finder.Search();

            Assert.IsTrue(finder.Found);
            Assert.IsFalse(LibreOfficeProcess.IsRunning);

            using (Process process = new Process { StartInfo = new ProcessStartInfo { FileName = finder.SOfficeBinaryPath } })
            {
                process.Start();
                process.WaitForInputIdle();

                Assert.IsTrue(LibreOfficeProcess.IsRunning);

                LibreOfficeProcessHelper.Kill();

                Assert.IsFalse(LibreOfficeProcess.IsRunning);
            }
        }
    }
}
