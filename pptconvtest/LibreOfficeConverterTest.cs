namespace pptconvtest
{
    using pptconv;

    using System;
    using System.IO;
    using System.Threading;
    using System.Collections.Generic;

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

        [TestMethod]
        public void BatchQueueFlushTest()
        {
            int count = 5;
            int converted = 0;
            List<string> batch = new List<string>();

            string testDir = Path.GetDirectoryName(LibreOfficeProcessTest.Slides);

            for (int i = 0; i < count; ++i)
            {
                string name = Path.Combine(testDir, string.Format("test_{0}.pptx", i));
                if (!File.Exists(name))
                    File.WriteAllBytes(name, Properties.Resources.test);
                batch.Add(name);
            }

            using (LibreOfficeConverter converter = new LibreOfficeConverter())
            {
                converter.ConversionSucceed += (file) => { converted++; };
                batch.ForEach(name => { converter.Queue(name, LibreOfficeProcessTest.TestTimeout); });
                converter.Flush();
            }

            Assert.AreEqual(converted, count);
        }

        [TestMethod]
        public void BatchQueueWaitTest()
        {
            int count = 5;
            int converted = 0;
            List<string> batch = new List<string>();

            int seconds = 5;
            TimeSpan timeout = TimeSpan.FromSeconds(seconds);
            TimeSpan waittime = TimeSpan.FromSeconds(count * seconds);

            string testDir = Path.GetDirectoryName(LibreOfficeProcessTest.Slides);

            for (int i = 0; i < count; ++i)
            {
                string name = Path.Combine(testDir, string.Format("test_{0}.pptx", i));
                if (!File.Exists(name))
                    File.WriteAllBytes(name, Properties.Resources.test);
                batch.Add(name);
            }

            using (LibreOfficeConverter converter = new LibreOfficeConverter())
            {
                converter.ConversionSucceed += (file) => { converted++; };
                batch.ForEach(name => { converter.Queue(name, LibreOfficeProcessTest.TestTimeout); });
                Thread.Sleep(waittime);
            }

            Assert.AreEqual(converted, count);
        }
    }
}
