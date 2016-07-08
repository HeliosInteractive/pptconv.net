namespace pptconvtest
{
    using pptconv;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class LibreOfficeConverterTest
    {
        [TestMethod]
        public void QueueResourceSlideTest()
        {
            bool done = false;

            using (LibreOfficeConverter converter = new LibreOfficeConverter())
            {
                converter.ConversionSucceed += (file) => { if (file == LibreOfficeProcessTest.Slides) done = true; };
                converter.Queue(LibreOfficeProcessTest.Slides, LibreOfficeProcessTest.TestTimeout);
                converter.Flush();
            }

            Assert.IsTrue(done);
        }

        [TestMethod]
        public void QueueAfterDisposeTest()
        {
            bool done = false;
            bool failed = false;

            LibreOfficeConverter converter = new LibreOfficeConverter();
            try
            {
                converter.ConversionSucceed += (file) => { if (file == LibreOfficeProcessTest.Slides) done = true; };
                converter.Queue(LibreOfficeProcessTest.Slides, LibreOfficeProcessTest.TestTimeout);
                converter.Flush();
            }
            finally
            {
                converter.Dispose();
                converter.ConversionFailed += (file) => { if (file == LibreOfficeProcessTest.Slides) failed = true; };
                converter.Queue(LibreOfficeProcessTest.Slides, LibreOfficeProcessTest.TestTimeout);
            }

            Assert.IsTrue(done);
            Assert.IsTrue(failed);
        }
    }
}
