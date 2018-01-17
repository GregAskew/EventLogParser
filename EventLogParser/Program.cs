namespace EventLogParser {

    #region Usings
    using Extensions;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    #endregion

    class Program {
        static void Main(string[] args) {

            if ((args.Length == 0) || (args[0] == "-?")) {
                PrintUsage();
                return;
            }

            var argsList = new List<string>();
            foreach (var arg in args) {
                argsList.Add(arg.ToLowerInvariant());
            }

            #region Check for required parameters
            if (args.Any(x => string.Equals(x, "-file", StringComparison.OrdinalIgnoreCase)) && args.Any(x => string.Equals(x, "-computerFqdn", StringComparison.OrdinalIgnoreCase))) {
                Console.WriteLine("Must specify either -file or -computerFqdn, but not both.");
                PrintUsage();
                return;
            }

            if (!args.Any(x => string.Equals(x, "-file", StringComparison.OrdinalIgnoreCase)) && !args.Any(x => string.Equals(x, "-computerFqdn", StringComparison.OrdinalIgnoreCase))) {
                Console.WriteLine("Must specify either -file or -computerFqdn.");
                PrintUsage();
                return;
            }

            if (args.Any(x => string.Equals(x, "-computerFqdn", StringComparison.OrdinalIgnoreCase)) && !args.Any(x => string.Equals(x, "-logName", StringComparison.OrdinalIgnoreCase))) {
                Console.WriteLine("Must specify -logName with -computerFqdn.");
                PrintUsage();
                return;
            }
            #endregion

            var file = string.Empty;
            var computerFqdn = string.Empty;
            var logName = string.Empty;
            var format = ReportFormat.CSV;

            #region Get file path
            int fileIndex = argsList.IndexOf("-file");
            if (fileIndex > -1) {
                if (argsList.Count > (fileIndex + 1)) {
                    file = argsList[fileIndex + 1];
                }

                if (string.IsNullOrWhiteSpace(file)) {
                    Console.WriteLine("-file not specified.");
                    return;
                }

                if (!File.Exists(file)) {
                    Console.WriteLine($"-file does not exist: {file}");
                    return;
                }
            }
            #endregion

            #region Get computerFqdn and logName
            else {
                int computerFqdnIndex = argsList.IndexOf("-computerfqdn");
                if (computerFqdnIndex > -1) {
                    if (argsList.Count > (computerFqdnIndex + 1)) {
                        computerFqdn = argsList[computerFqdnIndex + 1];
                    }

                    int dotCount = computerFqdn.Count(x => x == '.');
                    if (dotCount < 2) {
                        Console.WriteLine("-computerFqdn must be in fqdn format: computername.company.com");
                        return;
                    }
                }

                if (string.IsNullOrWhiteSpace(computerFqdn)) {
                    Console.WriteLine("-computerFqdn not specified.");
                    return;
                }

                int logNameIndex = argsList.IndexOf("-logname");
                if (logNameIndex > -1) {
                    if (argsList.Count > (logNameIndex + 1)) {
                        logName = argsList[logNameIndex + 1];
                    }
                }

                if (string.IsNullOrWhiteSpace(logName)) {
                    Console.WriteLine("-logName not specified.");
                    return;
                }
            }
            #endregion

            #region Get event Ids
            var eventIds = new List<int>();
            int eventsIndex = argsList.IndexOf("-events");
            if (eventsIndex > -1) {
                if (argsList.Count > (eventsIndex + 1)) {
                    var eventIdsText = argsList[eventsIndex + 1].Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                    if (eventIdsText.Length > 0) {
                        foreach (var eventIdText in eventIdsText) {
                            int eventId = -1;
                            int.TryParse(eventIdText, out eventId);
                            if (eventId > -1) {
                                eventIds.Add(eventId);
                            }
                        }
                    }
                }

                if (eventIds.Count == 0) {
                    Console.WriteLine("-events specified without required events parameter.");
                    return;
                }
            }
            #endregion

            #region Start date
            var startDate = new DateTime(2010, 1, 1);
            int startDateIndex = argsList.IndexOf("-startdate");
            if (startDateIndex > -1) {
                if (argsList.Count > (startDateIndex + 1)) {
                    if (!DateTime.TryParseExact(argsList[startDateIndex + 1], "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out startDate)) {
                        Console.WriteLine($"Invalid start date: {argsList[startDateIndex + 1]}");
                        return;
                    }
                }
                else {
                    Console.WriteLine("-startdate specified without required parameter.");
                    return;
                }
            }
            #endregion

            #region End date
            var endDate = DateTime.UtcNow;
            int endDateIndex = argsList.IndexOf("-enddate");
            if (endDateIndex > -1) {
                if (argsList.Count > (endDateIndex + 1)) {
                    if (!DateTime.TryParseExact(argsList[endDateIndex + 1], "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out endDate)) {
                        Console.WriteLine($"Invalid end date: {argsList[endDateIndex + 1]}");
                        return;
                    }
                }
                else {
                    Console.WriteLine("-enddate specified without required parameter.");
                    return;
                }
            }

            if (endDate < startDate) {
                Console.WriteLine($"Start date: {startDate.YMDFriendly()} must be less than end date: {endDate.YMDFriendly()}");
                return;
            }
            #endregion

            #region Report Format
            int formatIndex = argsList.IndexOf("-format");
            if (formatIndex > -1) {
                if (argsList.Count > (formatIndex + 1)) {
                    if (Enum.TryParse<ReportFormat>(argsList[formatIndex + 1], ignoreCase: true, result: out format)) {
                        if (!Enum.IsDefined(typeof(ReportFormat), value: format)) {
                            Console.WriteLine($"Invalid report format specified: {argsList[formatIndex + 1]}");
                            return;
                        }
                    }
                    else {
                        Console.WriteLine($"Invalid report format specified: {argsList[formatIndex + 1]}");
                        return;
                    }
                }
                else {
                    Console.WriteLine("-format specified without required parameter.");
                    return;
                }
            } 
            #endregion

            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
            TaskProcessor.DoWork(file, computerFqdn, logName, eventIds, startDate, endDate, format);
        }

        private static void PrintUsage() {
            Console.WriteLine("Usage: EventLogParser.exe -file <pathToEvtx> | -computerFqdn <computerName.company.com> -logName <logName> [-events EventId,EventId,EventId] [-startdate yyyy-MM-dd] [-enddate yyyy-MM-dd] [-format CSV | XML]");
        }

        /// <summary>
        /// Unhandled Exception Handler
        /// </summary>
        /// <param name="sender">Sender</param>
        /// <param name="e">Event Args</param>
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e) {
            var exception = e.ExceptionObject as Exception;
            EventLog.WriteEntry("Application", exception.VerboseExceptionString(), EventLogEntryType.Error);
        }
    }
}
