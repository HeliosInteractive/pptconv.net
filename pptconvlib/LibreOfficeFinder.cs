namespace pptconv
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Collections.Generic;

    public class LibreOfficeFinder
    {
        [Flags]
        public enum SearchOptions
        {
            None = 0,
            UserPath = 2,
            CommandLine = 4,
            ProgramDirectory = 8,
            EnvironmentVariable = 16
        }

        private SearchOptions options;
        private bool executablesFound;
        private string officeBasePath;
        private string officeUnoPath;

        public LibreOfficeFinder(SearchOptions opts)
        {
            executablesFound = false;
            options = opts;
        }

        public LibreOfficeFinder()
            : this(
                  SearchOptions.UserPath |
                  SearchOptions.CommandLine |
                  SearchOptions.ProgramDirectory |
                  SearchOptions.EnvironmentVariable)
        { }

        public void Search(string userPath = "")
        {
            Found = false;
            SOfficeBinaryPath = string.Empty;
            SOfficeInstallPath = string.Empty;
            List<string> extrapaths = new List<string>();

            if (options.HasFlag(SearchOptions.UserPath))
            {
                if (File.Exists(userPath))
                {
                    extrapaths.Add(Path.GetDirectoryName(Path.GetDirectoryName(userPath)));
                }
                else
                {
                    extrapaths.Add(userPath);
                }
            }

            if (options.HasFlag(SearchOptions.CommandLine))
            {
                string[] command_lines = Environment.GetCommandLineArgs();

                if (command_lines != null)
                {
                    foreach (string command_line in command_lines)
                    {
                        string cmd_var_name = string.Format("{0}=", SOfficeCommandLineSwitch);

                        if (command_line.StartsWith(cmd_var_name))
                        {
                            string path = command_line.Replace(cmd_var_name, string.Empty).Trim('\"').Trim('\'');

                            if (File.Exists(path))
                            {
                                extrapaths.Add(Path.GetDirectoryName(Path.GetDirectoryName(path)));
                            }
                            else
                            {
                                extrapaths.Add(path);
                            }
                        }
                    }
                }
            }

            if (options.HasFlag(SearchOptions.ProgramDirectory))
            {
                List<string> programDirectories = new List<string>();
                List<string> programDirectoryNames = new List<string>();
                var environment_variables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                foreach (string key in Environment.GetEnvironmentVariables().Keys)
                    environment_variables.Add(key, Environment.GetEnvironmentVariable(key));

                string[] programDirVarNames = new string[]
                {
                    "ProgramFiles",
                    "ProgramW6432",
                    "ProgramFiles(x86)"
                };

                foreach (string programDirVarName in programDirVarNames)
                {
                    if (environment_variables.ContainsKey(programDirVarName))
                    {
                        string basePath = Environment.GetEnvironmentVariable(programDirVarName);

                        if (!programDirectories.Contains(basePath))
                        {
                            programDirectories.Add(basePath);

                            if (!string.IsNullOrEmpty(Path.GetPathRoot(basePath)))
                            {
                                string baseName = basePath.Replace(Path.GetPathRoot(basePath), string.Empty);

                                if (!programDirectoryNames.Contains(baseName))
                                {
                                    programDirectoryNames.Add(baseName);
                                }
                            }
                        }
                    }
                }

                var driveLetters = DriveInfo.GetDrives()
                    .Where(di => di.DriveType == DriveType.Fixed)
                    .Select(di => di.Name)
                    .ToList();

                foreach (string programDirectoryName in programDirectoryNames)
                {
                    driveLetters.ForEach(dl =>
                    {
                        string programDirectory = string.Format("{0}{1}", dl, programDirectoryName);

                        if (Directory.Exists(programDirectory))
                        {
                            programDirectories.Add(programDirectory);
                        }
                    });
                }

                programDirectories.ForEach(path =>
                {
                    foreach (string directory in Directory.GetDirectories(path, "LibreOffice*"))
                    {
                        extrapaths.Add(directory);
                    }
                });
            }

            if (options.HasFlag(SearchOptions.EnvironmentVariable))
            {
                if (Environment.GetEnvironmentVariables().Contains(SOfficeEnvironmentVariableName))
                {
                    string path = Environment.GetEnvironmentVariable(SOfficeEnvironmentVariableName);

                    if (File.Exists(path))
                    {
                        extrapaths.Add(Path.GetDirectoryName(Path.GetDirectoryName(path)));
                    }
                    else
                    {
                        extrapaths.Add(path);
                    }
                }
            }

            extrapaths
                .Where(ep => Directory.Exists(ep))
                .ToList()
                .ForEach(ep =>
                {
                    string executable_name = "soffice.exe";

                    Directory.GetDirectories(ep).ToList().ForEach(subdir =>
                    {
                        string possible_path = Path.Combine(subdir, executable_name);

                        if (File.Exists(possible_path))
                        {
                            Found = true;
                            SOfficeBinaryPath = possible_path;
                            SOfficeInstallPath = ep;
                            return;
                        }
                    });
                });
        }

        public bool Found
        {
            get { return executablesFound; }
            private set { executablesFound = value; }
        }

        public string SOfficeBinaryPath
        {
            get { return officeUnoPath; }
            private set { officeUnoPath = value; }
        }

        public string SOfficeInstallPath
        {
            get { return officeBasePath; }
            private set { officeBasePath = value; }
        }

        public static string SOfficeEnvironmentVariableName
        {
            get { return "PPTCONV_LIBREOFFICE_PATH"; }
        }

        public static string SOfficeCommandLineSwitch
        {
            get { return "soffice"; }
        }
    }
}
