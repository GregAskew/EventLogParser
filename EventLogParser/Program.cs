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

            if ((args.Length == 0) || (args[0] == "-?") || (args[0] == "/?")) {
                PrintUsage();
                return;
            }

            #region Check for required parameters
            if (args.Any(x => x.IndexOf("/FilePath:", StringComparison.OrdinalIgnoreCase) > -1)
                && args.Any(x => x.IndexOf("/ComputerFqdn:", StringComparison.OrdinalIgnoreCase) > -1)) {
                Console.WriteLine("Must specify either /FilePath: or /ComputerFqdn:, but not both.");
                PrintUsage();
                return;
            }

            if (!args.Any(x => x.IndexOf("/FilePath:", StringComparison.OrdinalIgnoreCase) > -1)
                && !args.Any(x => x.IndexOf("/ComputerFqdn:", StringComparison.OrdinalIgnoreCase) > -1)) {
                Console.WriteLine("Must specify either /FilePath: or /ComputerFqdn:.");
                PrintUsage();
                return;
            }

            if (args.Any(x => x.IndexOf("/ComputerFqdn:", StringComparison.OrdinalIgnoreCase) > -1)
                && !args.Any(x => x.IndexOf("/LogName:", StringComparison.OrdinalIgnoreCase) > -1)) {
                Console.WriteLine("Must specify /LogName: with /ComputerFqdn:.");
                PrintUsage();
                return;
            }
            #endregion

            var filePath = string.Empty;
            var computerFqdn = string.Empty;
            var logName = string.Empty;
            var format = ReportFormat.CSV;

            #region Get file path
            var filePathArg = args
                .Where(x => x.StartsWith("/FilePath:", StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(filePathArg)) {

                if (filePathArg.Length > "/FilePath:".Length) {
                    filePath = filePathArg.Substring("/FilePath:".Length);
                }

                if (string.IsNullOrWhiteSpace(filePath)) {
                    Console.WriteLine("FilePath not specified.");
                    return;
                }

                if (!File.Exists(filePath)) {
                    Console.WriteLine($"FilePath does not exist: {filePath}");
                    return;
                }
            }
            #endregion

            #region Get ComputerFqdn and LogName
            else {
                var computerFqdnArg = args
                    .Where(x => x.StartsWith("/ComputerFqdn:", StringComparison.OrdinalIgnoreCase))
                    .FirstOrDefault();
                if (!string.IsNullOrWhiteSpace(computerFqdnArg)) {
                    if (computerFqdnArg.Length > "/ComputerFqdn:".Length) {
                        computerFqdn = computerFqdnArg.Substring("/ComputerFqdn:".Length).Trim().ToUpperInvariant();
                    }

                    int dotCount = computerFqdn.Count(x => x == '.');
                    if (dotCount < 2) {
                        Console.WriteLine($"ComputerFqdn: {computerFqdn} must be in fqdn format: computername.company.com");
                        return;
                    }
                }

                if (string.IsNullOrWhiteSpace(computerFqdn)) {
                    Console.WriteLine("/ComputerFqdn: not specified.");
                    return;
                }

                var logNameArg = args
                    .Where(x => x.StartsWith("/LogName:", StringComparison.OrdinalIgnoreCase))
                    .FirstOrDefault();
                if (!string.IsNullOrWhiteSpace(logNameArg)) {
                    if (logNameArg.Length > "/LogName:".Length) {
                        logName = logNameArg.Substring("/LogName:".Length);
                    }
                }

                if (string.IsNullOrWhiteSpace(logName)) {
                    Console.WriteLine("/LogName: not specified.");
                    return;
                }
            }
            #endregion

            #region Get event Ids
            var eventIds = new List<int>();
            var eventsArg = args
                .Where(x => x.StartsWith("/Events:", StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(eventsArg)) {
                if (eventsArg.Length > "/Events:".Length) {
                    var eventIdsText = eventsArg.Substring("/Events:".Length).Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                    if (eventIdsText.Length > 0) {
                        foreach (var eventIdText in eventIdsText) {
                            if (int.TryParse(eventIdText, out int eventId)) {
                                eventIds.Add(eventId);
                            }
                        }
                    }
                }

                if (eventIds.Count == 0) {
                    Console.WriteLine("/Events: specified without required event id's.");
                    return;
                }
            }
            #endregion

            #region Start date
            var startDate = new DateTime(2010, 1, 1);
            var startDateArg = args
                .Where(x => x.StartsWith("/StartDate:", StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(startDateArg)) {
                if (startDateArg.Length > "/StartDate:".Length) {
                    if (!DateTime.TryParseExact(startDateArg.Substring("/StartDate:".Length), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out startDate)) {
                        Console.WriteLine($"Invalid start date: {startDateArg.Substring("/StartDate:".Length)}");
                        return;
                    }
                }
                else {
                    Console.WriteLine("/StartDate: specified without required value.");
                    return;
                }
            }
            #endregion

            #region End date
            var endDate = DateTime.UtcNow;
            var endDateArg = args
                .Where(x => x.StartsWith("/EndDate:", StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(endDateArg)) {
                if (endDateArg.Length > "/EndDate:".Length) {
                    if (!DateTime.TryParseExact(endDateArg.Substring("/EndDate:".Length), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out endDate)) {
                        Console.WriteLine($"Invalid end date: {endDateArg.Substring("/StartDate:".Length)}");
                        return;
                    }
                }
                else {
                    Console.WriteLine("/EndDate: specified without required value.");
                    return;
                }
            }

            if (endDate < startDate) {
                Console.WriteLine($"Start date: {startDate.YMDFriendly()} must be less than end date: {endDate.YMDFriendly()}");
                return;
            }
            #endregion

            #region Report Format
            var formatArg = args
                .Where(x => x.StartsWith("/Format:", StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(formatArg)) {
                if (formatArg.Length > "/Format:".Length) {
                    if (Enum.TryParse<ReportFormat>(formatArg.Substring("/Format:".Length), ignoreCase: true, result: out format)) {
                        if (!Enum.IsDefined(typeof(ReportFormat), value: format)) {
                            Console.WriteLine($"Invalid report format specified: {formatArg.Substring("/Format:".Length)}");
                            return;
                        }
                    }
                    else {
                        Console.WriteLine($"Invalid report format specified: {formatArg.Substring("/Format:".Length)}");
                        return;
                    }
                }
                else {
                    Console.WriteLine("/Format: specified without required value.");
                    return;
                }
            }
            #endregion

            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
            TaskProcessor.DoWork(filePath, computerFqdn, logName, eventIds, startDate, endDate, format);
        }

        private static void PrintUsage() {
            Console.WriteLine("Usage: EventLogParser.exe /FilePath:<pathToEvtx> | /ComputerFqdn:<computerName.company.com> /LogName:<logName> [/Events:EventId,EventId,EventId] [/Startdate:yyyy-MM-dd] [/Enddate:yyyy-MM-dd] [/Format:CSV | XML]");
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
