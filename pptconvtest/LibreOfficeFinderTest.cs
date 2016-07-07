namespace pptconvtest
{
    using pptconv;

    using System;
    using System.IO;

    using Microsoft.QualityTools.Testing.Fakes;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class LibreOfficeFinderTest
    {
        static void ClearEnvironment()
        {
            foreach (string key in Environment.GetEnvironmentVariables().Keys)
                Environment.SetEnvironmentVariable(key, string.Empty);
        }

        [TestMethod]
        public void SearchTest()
        {
            string programDir = Path.GetTempPath();

            string libreOfficeDir = Path.Combine(programDir, "LibreOffice x");
            string libreOfficeProgramDir = Path.Combine(libreOfficeDir, "program");
            string sofficePath = Path.Combine(libreOfficeProgramDir, "soffice.exe");

            try { Directory.Delete(libreOfficeDir, true); }
            catch { /* no-op */ }

            Directory.CreateDirectory(libreOfficeDir);
            Directory.CreateDirectory(Path.Combine(libreOfficeDir, "sub1"));
            Directory.CreateDirectory(Path.Combine(libreOfficeDir, "sub2"));
            Directory.CreateDirectory(libreOfficeProgramDir);
            File.Create(sofficePath).Dispose();


            // search by command line
            using (ShimsContext.Create())
            {
                LibreOfficeFinder finder = new LibreOfficeFinder(LibreOfficeFinder.SearchOptions.CommandLine);

                System.Fakes.ShimEnvironment.GetCommandLineArgs = () =>
                {
                    return new string[] { string.Format("{0}=\"{1}\"", LibreOfficeFinder.SOfficeCommandLineSwitch, sofficePath) };
                };
                finder.Search();
                Assert.IsTrue(finder.Found);

                System.Fakes.ShimEnvironment.GetCommandLineArgs = () =>
                {
                    return new string[] { string.Format("{0}=\"{1}\"", LibreOfficeFinder.SOfficeCommandLineSwitch, libreOfficeDir) };
                };
                finder.Search();
                Assert.IsTrue(finder.Found);
            }

            // search by program directory
            {
                LibreOfficeFinder finder = new LibreOfficeFinder(LibreOfficeFinder.SearchOptions.ProgramDirectory);

                ClearEnvironment();
                finder.Search();

                Assert.IsFalse(finder.Found);

                Environment.SetEnvironmentVariable("PROGRAMFILES", programDir);
                finder.Search();

                Assert.IsTrue(finder.Found);
                Assert.AreEqual(finder.SOfficeBinaryPath, sofficePath);
                Assert.AreEqual(finder.SOfficeInstallPath, libreOfficeDir);

                ClearEnvironment();
                finder.Search();

                Assert.IsFalse(finder.Found);

                Environment.SetEnvironmentVariable("PROGRAMFILES(x86)", programDir);
                finder.Search();

                Assert.IsTrue(finder.Found);
                Assert.AreEqual(finder.SOfficeBinaryPath, sofficePath);
                Assert.AreEqual(finder.SOfficeInstallPath, libreOfficeDir);

                ClearEnvironment();
                finder.Search();

                Assert.IsFalse(finder.Found);

                Environment.SetEnvironmentVariable("PROGRAMW6432", programDir);
                finder.Search();

                Assert.IsTrue(finder.Found);
                Assert.AreEqual(finder.SOfficeBinaryPath, sofficePath);
                Assert.AreEqual(finder.SOfficeInstallPath, libreOfficeDir);

                ClearEnvironment();
                finder.Search();

                Assert.IsFalse(finder.Found);
            }

            // search by environment variable
            {
                LibreOfficeFinder finder = new LibreOfficeFinder(LibreOfficeFinder.SearchOptions.EnvironmentVariable);

                ClearEnvironment();
                finder.Search();

                Assert.IsFalse(finder.Found);

                Environment.SetEnvironmentVariable(LibreOfficeFinder.SOfficeEnvironmentVariableName, libreOfficeDir);
                finder.Search();

                Assert.IsTrue(finder.Found);
                Assert.AreEqual(finder.SOfficeBinaryPath, sofficePath);
                Assert.AreEqual(finder.SOfficeInstallPath, libreOfficeDir);

                ClearEnvironment();
                finder.Search();

                Assert.IsFalse(finder.Found);

                Environment.SetEnvironmentVariable(LibreOfficeFinder.SOfficeEnvironmentVariableName, sofficePath);
                finder.Search();

                Assert.IsTrue(finder.Found);
                Assert.AreEqual(finder.SOfficeBinaryPath, sofficePath);
                Assert.AreEqual(finder.SOfficeInstallPath, libreOfficeDir);
            }

            // search by user path
            {
                LibreOfficeFinder finder = new LibreOfficeFinder(LibreOfficeFinder.SearchOptions.UserPath);

                ClearEnvironment();
                finder.Search();

                Assert.IsFalse(finder.Found);

                finder.Search(sofficePath);

                Assert.IsTrue(finder.Found);
                Assert.AreEqual(finder.SOfficeBinaryPath, sofficePath);
                Assert.AreEqual(finder.SOfficeInstallPath, libreOfficeDir);

                ClearEnvironment();
                finder.Search();

                Assert.IsFalse(finder.Found);

                finder.Search(libreOfficeDir);

                Assert.IsTrue(finder.Found);
                Assert.AreEqual(finder.SOfficeBinaryPath, sofficePath);
                Assert.AreEqual(finder.SOfficeInstallPath, libreOfficeDir);
            }
        }
    }
}
