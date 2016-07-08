namespace pptconvtest
{
    using pptconv;

    using System;
    using System.IO;
    using System.Diagnostics;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class LibreOfficeProcessTest
    {
        public static string Outdir
        {
            get
            {
                string outdir = Path.Combine(Path.GetTempPath(), "outdir");

                if (!Directory.Exists(outdir))
                    Directory.CreateDirectory(outdir);

                return outdir;
            }
        }

        public static string Slides
        {
            get
            {
                string slides = Path.Combine(Path.GetTempPath(), "test.pptx");

                if (!File.Exists(slides))
                    File.WriteAllBytes(slides, Properties.Resources.test);

                return slides;
            }
        }

        public static TimeSpan TestTimeout
        {
            get { return TimeSpan.FromMinutes(1); }
        }

        [TestMethod]
        public void OfficeNotFoundConvertTest()
        {
            LibreOfficeFinder finder = new LibreOfficeFinder();
            LibreOfficeProcess soffice = new LibreOfficeProcess(finder);

            Assert.IsFalse(soffice.Convert(Slides, Outdir, TestTimeout));

            finder.Search();
            Assert.IsTrue(soffice.Convert(Slides, Outdir, TestTimeout));
        }

        [TestMethod]
        public void ResourceSlideConvertTest()
        {
            LibreOfficeFinder finder = new LibreOfficeFinder();
            LibreOfficeProcess soffice = new LibreOfficeProcess(finder);
            finder.Search();

            Assert.IsTrue(finder.Found);
            Assert.IsTrue(soffice.Convert(Slides, Outdir, TestTimeout));
        }

        [TestMethod]
        public void NonExistingSlideConvertTest()
        {
            LibreOfficeFinder finder = new LibreOfficeFinder();
            LibreOfficeProcess soffice = new LibreOfficeProcess(finder);
            finder.Search();

            Assert.IsTrue(finder.Found);
            string randomSlidePath = Path.Combine(Path.GetTempPath(), "someRandomSlide.pptx");
            Assert.IsFalse(soffice.Convert(randomSlidePath, Outdir, TestTimeout));
        }

        [TestMethod]
        public void NonExistingFinderConvertTest()
        {
            LibreOfficeProcess soffice = new LibreOfficeProcess(null);
            Assert.IsFalse(soffice.Convert(Slides, Outdir, TestTimeout));
        }

        [TestMethod]
        public void NonExistingOutDirConvertTest()
        {
            LibreOfficeProcess soffice = new LibreOfficeProcess();

            string outdir = Outdir;
            Directory.Delete(outdir, true);

            Assert.IsFalse(Directory.Exists(outdir));
            Assert.IsFalse(soffice.Convert(Slides, outdir, TestTimeout));
        }

        [TestMethod]
        public void ZombieProcessConvertTest()
        {
            LibreOfficeFinder finder = new LibreOfficeFinder();
            finder.Search();

            Process process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = finder.SOfficeBinaryPath
                }
            };

            try
            {
                process.Start();
                process.WaitForInputIdle();

                Assert.IsTrue(LibreOfficeProcessHelper.IsRunning);

                LibreOfficeProcess soffice = new LibreOfficeProcess();
                Assert.IsTrue(soffice.Convert(Slides, Outdir, TestTimeout));

                Assert.IsFalse(LibreOfficeProcessHelper.IsRunning);
            }
            finally
            {
                process.Dispose();
            }
        }
    }
}
