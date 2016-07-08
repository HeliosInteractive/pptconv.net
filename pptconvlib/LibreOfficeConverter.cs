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
        private TimeSpan maxJoinTimeout;
        private ManualResetEventSlim resetEvent;
        private readonly LibreOfficeFinder finder;
        private readonly LibreOfficeProcess officeProcess;
        private readonly BlockingCollection<ConvertData> queue;

        public LibreOfficeConverter(LibreOfficeFinder officeFinder, string outputDirictory)
        {
            finder = officeFinder;
            outdir = outputDirictory;
            officeProcess = new LibreOfficeProcess(finder);
            queue = new BlockingCollection<ConvertData>();
            convertThread = new Thread(ConvertThreadMain);
            resetEvent = new ManualResetEventSlim(false);
            maxJoinTimeout = TimeSpan.FromSeconds(10);
            convertThread.Start();
        }

        public LibreOfficeConverter()
            : this(new LibreOfficeFinder(), Directory.GetCurrentDirectory())
        {
            finder.Search();
        }

        public void Queue(string slides, TimeSpan timeout)
        {
            if (IsDisposed)
            {
                if (ConversionFailed != null)
                    ConversionFailed(slides);
                return;
            }

            if (maxJoinTimeout < timeout)
                maxJoinTimeout = timeout;

            queue.Add(new Tuple<string, TimeSpan>(slides, timeout));

            if (!resetEvent.IsSet)
                resetEvent.Set();
        }

        public void Flush()
        {
            lock(this)
            {
                while (!queue.IsCompleted && queue.Count > 0)
                {
                    if (queue.IsAddingCompleted)
                        return;
                    else
                    {
                        ConvertData data = queue.Take();

                        if (officeProcess.Convert(data.Item1, outdir, data.Item2))
                        {
                            if (ConversionSucceed != null) ConversionSucceed(data.Item1);
                            else if (ConversionFailed != null) ConversionFailed(data.Item1);
                        }
                    }
                }
            }
        }

        private void ConvertThreadMain()
        {
            while (!queue.IsAddingCompleted)
            {
                resetEvent.Wait();
                Flush();
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
                    convertThread.Join(maxJoinTimeout);

                    if (convertThread.IsAlive)
                        convertThread.Abort();

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
