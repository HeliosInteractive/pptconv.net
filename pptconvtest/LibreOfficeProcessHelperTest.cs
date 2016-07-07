namespace pptconvtest
{
    using pptconv;

    using System.Diagnostics;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class LibreOfficeProcessHelperTest
    {
        [TestMethod]
        public void RunningInstanceTest()
        {
            LibreOfficeFinder finder = new LibreOfficeFinder();

            finder.Search();

            Assert.IsTrue(finder.Found);
            Assert.IsFalse(LibreOfficeProcessHelper.IsRunning);

            using (Process process = new Process { StartInfo = new ProcessStartInfo { FileName = finder.SOfficeBinaryPath } })
            {
                process.Start();
                process.WaitForInputIdle();

                Assert.IsTrue(LibreOfficeProcessHelper.IsRunning);

                LibreOfficeProcessHelper.Kill();

                Assert.IsFalse(LibreOfficeProcessHelper.IsRunning);
            }
        }
    }
}
