namespace pptconv
{
    using System;
    using System.IO;
    using System.Linq;

    class Program
    {
        static int Main(string[] args)
        {
            LibreOfficeFinder finder = new LibreOfficeFinder();
            finder.Search();

            if (!finder.Found)
            {
                Console.WriteLine("Could not find LibreOffice. Please install LibreOffice and try again.");
                return 1;
            }

            Console.WriteLine("Writing output to: {0}", Directory.GetCurrentDirectory());

            using (LibreOfficeConverter converter = new LibreOfficeConverter(finder, Directory.GetCurrentDirectory()))
            {
                converter.ConversionSucceed += file => Console.WriteLine("SUCCESS: {0}", file);
                converter.ConversionFailed += file => Console.WriteLine("FAILED: {0}", file);

                args
                    .Select(arg => { if (Path.IsPathRooted(arg)) return arg; else return Path.GetFullPath(arg); })
                    .Where(arg => arg != args[0])
                    .Where(arg => File.Exists(arg))
                    .ToList()
                    .ForEach(file =>
                {
                    // 10 seconds per MB (kinda)
                    TimeSpan timeout = TimeSpan.FromSeconds((Math.Ceiling((double)(new FileInfo(file).Length / (1024 * 1024))) + 1) * 10);
                    Console.WriteLine("Queuing file: {0} with timeout: {1} seconds", file, timeout.TotalSeconds);
                    converter.Queue(file, timeout);
                });

                Console.WriteLine("Waiting for all files to be converted.");
                converter.Flush();
            }

            return 0;
        }
    }
}
