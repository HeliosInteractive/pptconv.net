namespace pptconv
{
    using System;
    using System.IO;
    using System.Threading;
    using System.Collections.Concurrent;
    using ConvertData = System.Tuple<string, System.TimeSpan>;

    public class LibreOfficeConverter : IDisposable
    {
        public Action<string> ConversionSucceed;
        public Action<string> ConversionFailed;

        private Thread convertThread;
        private readonly string outdir;
        private ManualResetEventSlim resetEvent;
        private readonly LibreOfficeFinder finder;
        private readonly LibreOfficeProcess officeProcess;
        private readonly BlockingCollection<ConvertData> queue;

        public LibreOfficeConverter(LibreOfficeFinder officeFinder, string outputDirictory)
        {
            finder = officeFinder;
            outdir = outputDirictory;
            queue = new BlockingCollection<ConvertData>();
            officeProcess = new LibreOfficeProcess(finder);
            convertThread = new Thread(ConvertThreadMain);
            resetEvent = new ManualResetEventSlim(false);
            convertThread.Start();
        }

        public LibreOfficeConverter()
            : this(new LibreOfficeFinder(), Directory.GetCurrentDirectory())
        { }

        public void Queue(string slides, TimeSpan timeout)
        {
            if (IsDisposed && ConversionFailed != null)
                ConversionFailed(slides);

            queue.Add(new Tuple<string, TimeSpan>(slides, timeout));

            if (!resetEvent.IsSet)
                resetEvent.Set();
        }

        private void ConvertThreadMain()
        {
            while (!queue.IsAddingCompleted)
            {
                resetEvent.Wait();

                while(!queue.IsCompleted)
                {
                    var data = queue.Take();

                    if (officeProcess.Convert(data.Item1, outdir, data.Item2))
                        if (ConversionSucceed != null) ConversionSucceed(data.Item1);
                    else
                        if (ConversionFailed != null) ConversionFailed(data.Item1);
                }

                if (resetEvent.IsSet)
                    resetEvent.Reset();
            }
        }

        #region IDisposable Support
        private volatile bool IsDisposed = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!IsDisposed)
            {
                IsDisposed = true;

                if (disposing)
                {
                    queue.CompleteAdding();
                    resetEvent.Set();
                    convertThread.Join();

                    resetEvent.Dispose();
                    queue.Dispose();
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }
        #endregion
    }
}
