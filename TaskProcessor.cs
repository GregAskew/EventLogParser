namespace EventLogParser {

    #region Usings
    using Extensions;
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Diagnostics;
    using System.Diagnostics.Eventing.Reader;
    using System.IO;
    using System.Linq;
    using System.Runtime;
    using System.Security;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Xml;
    using System.Xml.Linq;
    #endregion

    public class TaskProcessor {

        #region Members

        private static bool EventLogQueryReverseDirection { get; set; }

        /// <summary>
        /// The format for the report.
        /// </summary>
        private static ReportFormat ReportFormat { get; set; }

        #region Collections
        /// <summary>
        /// Friendly description/category for an Event Id. Example: 4264 Logon, 4625 Logon Failure
        /// </summary>
        private static Dictionary<int, string> EventCategories { get; set; }

        /// <summary>
        /// The Event Id - Column Name cross-reference
        /// Key: Event Id Value: Dictionary of column names
        /// </summary>
        private static Dictionary<int, Dictionary<string, int>> EventColumns { get; set; }

        /// <summary>
        /// The Event Id - Event cross-reference
        /// Key: Event Id Value: Events
        /// </summary>
        private static Dictionary<int, List<EventBase>> EventData { get; set; }
        #endregion
        #endregion

        #region Constructor
        static TaskProcessor() {
            EventCategories = new Dictionary<int, string>();
            EventColumns = new Dictionary<int, Dictionary<string, int>>();
            EventData = new Dictionary<int, List<EventBase>>();
        }
        #endregion

        #region Methods

        /// <summary>
        /// Entry for report creation.  Will call CreateCSVReport or CreateXMLReport.
        /// </summary>
        /// <param name="reportFileBasePath">The report file base path.  
        /// If specified, the directory for the report files is obtained from this path.
        /// If not specified, the Desktop directory is used.
        /// </param>
        private static void CreateReport(string reportFileBasePath = "") {
            Console.WriteLine($"{DateTime.Now.YMDHMSFriendly()} - {ObjectExtensions.CurrentMethodName()}");

            if (EventData.Count == 0) {
                Console.WriteLine($"{DateTime.Now.YMDHMSFriendly()} - {ObjectExtensions.CurrentMethodName()} Report not created due to no events collected.");
                return;
            }

            if (!string.IsNullOrWhiteSpace(reportFileBasePath)) {
                reportFileBasePath = Path.GetDirectoryName(reportFileBasePath);
            }

            if (string.IsNullOrWhiteSpace(reportFileBasePath)) {
                reportFileBasePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory, Environment.SpecialFolderOption.DoNotVerify), "EventLogParser");
                if (!Directory.Exists(reportFileBasePath)) {
                    Directory.CreateDirectory(reportFileBasePath);
                }
            }

            if (ReportFormat == ReportFormat.CSV) {
                CreateCSVReport(reportFileBasePath);
            }
            else if (ReportFormat == ReportFormat.XML) {
                CreateXMLReport(reportFileBasePath);
            }
        }

        private static void CreateCSVReport(string reportFileBasePath) {
            Console.WriteLine($"{DateTime.Now.YMDHMSFriendly()} - {ObjectExtensions.CurrentMethodName()}");

            #region Validation
            if (string.IsNullOrWhiteSpace(reportFileBasePath)) {
                throw new ArgumentNullException("reportFileBasePath");
            }
            #endregion

            foreach (var kvpEventData in EventData) {
                Console.WriteLine($" - EventId: {kvpEventData.Key} Events: {kvpEventData.Value.Count}");

                var lines = new List<string>();
                var headerLine = new StringBuilder();

                #region Column headers for all events
                headerLine.AppendFormat("EventId,");
                headerLine.AppendFormat("EventRecordId,");
                headerLine.AppendFormat("EventSourceMachine,");
                headerLine.AppendFormat("DateTimeUTC,");
                headerLine.AppendFormat("Channel,");
                headerLine.AppendFormat("Level,");
                #endregion

                foreach (var kvpColumnName in EventColumns[kvpEventData.Key]) {
                    headerLine.AppendFormat("{0},", kvpColumnName.Key);
                }

                lines.Add(headerLine.ToString());

                var firstEvent = DateTime.MaxValue;
                var lastEvent = DateTime.MinValue;

                foreach (var eventBase in kvpEventData.Value) {
                    if (eventBase.DateTimeUTC < firstEvent) firstEvent = eventBase.DateTimeUTC;
                    if (eventBase.DateTimeUTC > lastEvent) lastEvent = eventBase.DateTimeUTC;

                    var line = new StringBuilder();
                    #region Column values for all events/lines
                    line.AppendFormat("{0},", eventBase.EventId);
                    line.AppendFormat("{0},", eventBase.EventRecordId);
                    line.AppendFormat("{0},", eventBase.EventSourceMachine ?? "NULL");
                    line.AppendFormat("{0},", eventBase.DateTimeUTC.YMDHMSFriendly());
                    line.AppendFormat("{0},", eventBase.Channel);
                    line.AppendFormat("{0},", eventBase.Level);
                    #endregion

                    #region Get column values for the line
                    foreach (var columnName in EventColumns[eventBase.EventId].Keys) {
                        var columnValue = "N/A,";
                        if ((eventBase.EventDataNameElements != null) && (eventBase.EventDataNameElements.Count > 0)) {
                            var columnElement = eventBase.EventDataNameElements
                                .Where(x => x.Attributes()
                                    .Any(y =>
                                        (y != null) && (y.Name != null) && !string.IsNullOrWhiteSpace(y.Name.LocalName)
                                        && string.Equals(y.Name.LocalName, "Name", StringComparison.OrdinalIgnoreCase)
                                        && !string.IsNullOrWhiteSpace(y.Value)
                                        && string.Equals(y.Value, columnName, StringComparison.OrdinalIgnoreCase)))
                                .FirstOrDefault();

                            if ((columnElement != null) && (columnElement.Value != null)) {
                                columnValue = columnElement.Value.Trim();
                            }
                        }

                        line.AppendFormat("\"{0}\",", columnValue);
                    }

                    #endregion

                    lines.Add(line.ToString());
                } // foreach (var eventBase in kvpEventData.Value) {

                #region Create report for the specific event id
                var eventCategory = string.Empty;
                if (EventCategories.ContainsKey(kvpEventData.Key)) {
                    eventCategory = $"-{EventCategories[kvpEventData.Key]}";
                }

                var reportFileName = string.Format("EventId-{0}{1}-{2}-{3}.csv",
                    kvpEventData.Key,
                    eventCategory,
                    firstEvent.YMDHMFriendly().Replace(":", "-").Replace(" ", "-"),
                    lastEvent.YMDHMFriendly().Replace(":", "-").Replace(" ", "-"));
                var reportFilePath = Path.Combine(reportFileBasePath, reportFileName);

                try {
                    Console.WriteLine($"{DateTime.Now.YMDHMSFriendly()} - {ObjectExtensions.CurrentMethodName()} Writing: {lines.Count} lines to file: {reportFilePath}");
                    File.WriteAllLines(reportFilePath, lines, Encoding.UTF8);
                }
                catch (Exception e) {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error parsing events");
                    Console.WriteLine(e.VerboseExceptionString());
                    Console.ResetColor();
                }
                #endregion

            } // foreach (var kvpEventData in EventData) {
        }

        private static void CreateXMLReport(string reportFileBasePath) {
            Console.WriteLine($"{DateTime.Now.YMDHMSFriendly()} - {ObjectExtensions.CurrentMethodName()}");

            #region Validation
            if (string.IsNullOrWhiteSpace(reportFileBasePath)) {
                throw new ArgumentNullException("reportFileBasePath");
            }
            #endregion

            foreach (var kvpEventData in EventData) {
                Console.WriteLine($" - EventId: {kvpEventData.Key} Events: {kvpEventData.Value.Count}");

                var rootElement = new XElement($"ArrayOfEventId{kvpEventData.Key}");

                var firstEvent = DateTime.MaxValue;
                var lastEvent = DateTime.MinValue;

                foreach (var eventBase in kvpEventData.Value) {

                    if (eventBase.DateTimeUTC < firstEvent) firstEvent = eventBase.DateTimeUTC;
                    if (eventBase.DateTimeUTC > lastEvent) lastEvent = eventBase.DateTimeUTC;

                    #region Element values for all events/lines
                    var eventChildElement = new XElement("Event");
                    eventChildElement.Add(new XElement("EventId", eventBase.EventId));
                    eventChildElement.Add(new XElement("EventRecordId", eventBase.EventRecordId));
                    eventChildElement.Add(new XElement("EventSourceMachine", eventBase.EventSourceMachine ?? "NULL"));
                    eventChildElement.Add(new XElement("DateTimeUTC", eventBase.DateTimeUTC.YMDHMSFriendly()));
                    eventChildElement.Add(new XElement("Channel", eventBase.Channel));
                    eventChildElement.Add(new XElement("Level", eventBase.Level));
                    #endregion

                    #region Get element values unique for the event
                    foreach (var elementName in EventColumns[eventBase.EventId].Keys) {
                        var elementValue = "N/A";

                        if ((eventBase.EventDataNameElements != null) && (eventBase.EventDataNameElements.Count > 0)) {
                            var columnElement = eventBase.EventDataNameElements
                                .Where(x => x.Attributes()
                                    .Any(y =>
                                        (y != null) && (y.Name != null) && !string.IsNullOrWhiteSpace(y.Name.LocalName)
                                        && string.Equals(y.Name.LocalName, "Name", StringComparison.OrdinalIgnoreCase)
                                        && !string.IsNullOrWhiteSpace(y.Value)
                                        && string.Equals(y.Value, elementName, StringComparison.OrdinalIgnoreCase)))
                                .FirstOrDefault();

                            if ((columnElement != null) && (columnElement.Value != null)) {
                                elementValue = columnElement.Value.Trim();
                                if (string.IsNullOrWhiteSpace(elementValue)) {
                                    elementValue = "N/A";
                                }
                            }
                        }

                        eventChildElement.Add(new XElement(elementName, elementValue));
                    }
                    #endregion

                    rootElement.Add(eventChildElement);

                } // foreach (var eventBase in kvpEventData.Value) {

                #region Create report for the specific event id
                var eventCategory = string.Empty;
                if (EventCategories.ContainsKey(kvpEventData.Key)) {
                    eventCategory = $"-{EventCategories[kvpEventData.Key]}";
                }

                var reportFileName = string.Format("EventId-{0}{1}-{2}-{3}.xml",
                    kvpEventData.Key,
                    eventCategory,
                    firstEvent.YMDHMFriendly().Replace(":", "-").Replace(" ", "-"),
                    lastEvent.YMDHMFriendly().Replace(":", "-").Replace(" ", "-"));
                var reportFilePath = Path.Combine(reportFileBasePath, reportFileName);

                var xDocument = new XDocument(new XDeclaration("1.0", "UTF-8", string.Empty), rootElement);

                try {
                    Console.WriteLine($"{DateTime.Now.YMDHMSFriendly()} - {ObjectExtensions.CurrentMethodName()} Writing: {rootElement.Elements().Count()} events to file: {reportFilePath}");
                    xDocument.Save(reportFilePath);
                }
                catch (Exception e) {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error parsing events");
                    Console.WriteLine(e.VerboseExceptionString());
                    Console.ResetColor();
                }
                #endregion

            } // foreach (var kvpEventData in EventData) {
        }

        /// <summary>
        /// Primary entry for application execution.
        /// </summary>
        /// <param name="file">The event log evtx file to parse.</param>
        /// <param name="computerFqdn">The computer to process event logs (if not processing a file).</param>
        /// <param name="eventLogName">The event log name (Required if computerFqdn is specified). Example: Security</param>
        /// <param name="eventIds">Optional.  List of event Ids for the query filter.</param>
        /// <param name="startDate">The start date for the query filter.</param>
        /// <param name="endDate">The end date for the query filter.</param>
        /// <param name="format">The report format (CSV or XML).</param>
        public static void DoWork(string file, string computerFqdn, string eventLogName, List<int> eventIds, DateTime startDate, DateTime endDate, ReportFormat reportFormat = ReportFormat.CSV) {

            #region Validation
            if (string.IsNullOrWhiteSpace(file) && string.IsNullOrWhiteSpace(computerFqdn)) {
                throw new ArgumentException("Must specify either file or computerFqdn.");
            }
            if (!string.IsNullOrWhiteSpace(file) && !string.IsNullOrWhiteSpace(computerFqdn)) {
                throw new ArgumentException("Must specify only file or computerFqdn.");
            }
            if (!string.IsNullOrWhiteSpace(computerFqdn) && string.IsNullOrWhiteSpace(eventLogName)) {
                throw new ArgumentException("Must specify eventLogName with computerFqdn.");
            }
            if (eventIds == null) {
                throw new ArgumentNullException("eventIds");
            }
            if (startDate > endDate) {
                throw new ArgumentOutOfRangeException("startDate must be less than endDate.");
            }

            if (!Enum.IsDefined(typeof(ReportFormat), reportFormat)) {
                throw new ArgumentOutOfRangeException($"Invalid report format: {reportFormat}");
            }
            else {
                ReportFormat = reportFormat;
            }
            #endregion

            Console.WriteLine($"{DateTime.Now.YMDHMSFriendly()} - {ObjectExtensions.CurrentMethodName()}");

            var stopwatch = Stopwatch.StartNew();

            try {
                Initialize();
                if (!string.IsNullOrWhiteSpace(file) && file.EndsWith(".XML", StringComparison.OrdinalIgnoreCase)) {
                    GetEventsFromXml(file);
                }
                else {
                    GetEvents(file, computerFqdn, eventLogName, eventIds, startDate, endDate);
                }
                CreateReport(file);
            }
            catch (Exception e) {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error parsing events");
                Console.WriteLine(e.VerboseExceptionString());
                Console.ResetColor();
            }
            finally {
                Console.WriteLine($"{DateTime.Now.YMDHMSFriendly()} - {ObjectExtensions.CurrentMethodName()} Finished.  Time required: {stopwatch.Elapsed} Memory used: {Process.GetCurrentProcess().PeakWorkingSet64.ToString("N0")} Processor time: {Process.GetCurrentProcess().TotalProcessorTime}");

                // 755,000 events and XML report statistics:
                // Time required: 00:14:05.4639642 Memory used: 6,572,908,544 Processor time: 00:06:15.7812500

                // 755,000 events and CSV report statistics:
                // Time required: 00:07:15.7816481 Memory used: 6,426,861,568 Processor time: 00:05:21.7187500
            }
        }

        private static void GetEvents(string file, string computerFqdn, string logName, List<int> eventIds, DateTime startDate, DateTime endDate) {
            Console.WriteLine("{0} - {1} File: {2} Computer: {3} LogName: {4} EventIds: {5} Start Date: {6} End Date: {7}",
                DateTime.Now.YMDHMSFriendly(), ObjectExtensions.CurrentMethodName(),
                !string.IsNullOrWhiteSpace(file) ? file : "N/A",
                !string.IsNullOrWhiteSpace(computerFqdn) ? computerFqdn : "N/A",
                !string.IsNullOrWhiteSpace(logName) ? logName : "N/A",
                (eventIds.Count > 0) ? eventIds.ToDelimitedString() : "<All Events>",
                startDate.YMDFriendly(), endDate.YMDFriendly());

            long eventsProcessed = 0;
            var stopwatch = Stopwatch.StartNew();
            var eventReadTimeout = TimeSpan.FromMinutes(5);

            if (!string.IsNullOrWhiteSpace(file)) {
                if (!file.EndsWith(".EVTX", StringComparison.OrdinalIgnoreCase)) {
                    Console.WriteLine($"{DateTime.Now.YMDHMSFriendly()} - {ObjectExtensions.CurrentMethodName()} File: {file} if not .XML, file must be an .EVTX file type. Exiting.");
                    return;
                }
            }

            try {

                #region Construct event query filter
                var eventIdsFilter = new StringBuilder();
                if (eventIds.Count > 0) {
                    eventIdsFilter.Append("(");
                    for (int index = 0; index < eventIds.Count; index++) {
                        eventIdsFilter.AppendFormat("EventID={0}", eventIds[index]);
                        if (index < eventIds.Count - 1) {
                            eventIdsFilter.Append(" or ");
                        }
                    }
                    eventIdsFilter.Append(") and ");
                }

                var query = $"*[System[{eventIdsFilter.ToString()}TimeCreated[@SystemTime>='{startDate.YMDFriendly()}T00:00:00.000Z' and @SystemTime<'{endDate.Add(TimeSpan.FromDays(1)).YMDFriendly()}T00:00:00.000Z']]]";

                Console.WriteLine($"{DateTime.Now.YMDHMSFriendly()} - {ObjectExtensions.CurrentMethodName()} Event query: {query}");
                #endregion

                EventLogQuery eventLogQuery = null;
                EventLogSession session = null;

                try {

                    if (!string.IsNullOrWhiteSpace(computerFqdn)) {
                        eventLogQuery = new EventLogQuery(logName, PathType.LogName, query);
                        session = new EventLogSession(computerFqdn);
                        eventLogQuery.Session = session;
                    }
                    else {
                        eventLogQuery = new EventLogQuery(file, PathType.FilePath, query);
                    }

                    eventLogQuery.ReverseDirection = EventLogQueryReverseDirection;

                    using (var eventLogReader = new EventLogReader(eventLogQuery)) {

                        EventRecord eventRecord = null;
                        string eventRecordXml = string.Empty;

                        do {

                            eventRecord = null;
                            eventRecordXml = string.Empty;

                            #region Read event from event log
                            try {
                                eventRecord = eventLogReader.ReadEvent(eventReadTimeout);
                                if (eventRecord == null) break;
                                eventsProcessed++;
                                eventRecordXml = eventRecord.ToXml();
                                if (string.IsNullOrWhiteSpace(eventRecordXml)) continue;
                                if (GetEventFromXml(eventRecordXml) != 0) continue;
                            }
                            catch (XmlException e) {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine($"Error parsing event record Xml:{eventRecordXml ?? "NULL"}");
                                Console.WriteLine(e.VerboseExceptionString());
                                Console.ResetColor();
                                continue;
                            }
                            catch (EventLogException e) {
                                if (Regex.IsMatch(e.Message, "The array bounds are invalid", RegexOptions.IgnoreCase)) continue;
                                if (Regex.IsMatch(e.Message, "The data area passed to a system call is too small", RegexOptions.IgnoreCase)) continue;
                                throw;
                            }
                            #endregion

                            #region Log statistics
                            if (eventsProcessed % 5000 == 0) {
                                Console.WriteLine($"{DateTime.Now.YMDHMSFriendly()} Events processed: {eventsProcessed}");
                            }
                            #endregion

                        } while (eventRecord != null);
                    }
                }
                finally {
                    if (session != null) {
                        try {
                            session.Dispose();
                        }
                        catch { }
                    }
                }
            }
            catch (Exception e) {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error parsing events");
                Console.WriteLine(e.VerboseExceptionString());
                Console.ResetColor();
                throw;
            }
        }

        private static int GetEventFromXml(string eventRecordXml) {

            if (string.IsNullOrWhiteSpace(eventRecordXml)) {
                throw new ArgumentNullException("eventRecordXml");
            }

            // 0 Success, 1 Error
            int returnCode = 0;

            var eventBase = new EventBase(eventRecordXml);
            var validationResults = eventBase.Validate().ToList();

            if (validationResults.Count > 0) {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Event is not valid: {eventRecordXml}");
                foreach (var validationResult in validationResults) {
                    Console.WriteLine($"Validation result: Member names: {validationResult.MemberNames.ToDelimitedString()} Message: {validationResult.ErrorMessage}");
                }
                Console.ResetColor();
                return 1;
            }

            if (!EventColumns.ContainsKey(eventBase.EventId)) {
                EventColumns.Add(eventBase.EventId, new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase));
            }

            foreach (var element in eventBase.EventDataNameElements) {

                var columnValue = string.Empty;
                var nameAttributeElement = element.Attributes()
                        .Where(x =>
                            (x != null) && (x.Name != null) && !string.IsNullOrWhiteSpace(x.Name.LocalName)
                            && string.Equals(x.Name.LocalName, "Name", StringComparison.OrdinalIgnoreCase))
                    .FirstOrDefault();

                if ((nameAttributeElement != null) && (nameAttributeElement.Value != null)) {
                    columnValue = nameAttributeElement.Value.Trim();
                }

                if (!string.IsNullOrWhiteSpace(columnValue)) {
                    if (!EventColumns[eventBase.EventId].ContainsKey(columnValue)) {
                        EventColumns[eventBase.EventId].Add(columnValue, 0);
                    }
                }
            }

            if (!EventData.ContainsKey(eventBase.EventId)) {
                EventData.Add(eventBase.EventId, new List<EventBase>());
            }
            EventData[eventBase.EventId].Add(eventBase);

            return returnCode;
        }

        private static void GetEventsFromXml(string file) {
            Console.WriteLine($"{DateTime.Now.YMDHMSFriendly()} - {ObjectExtensions.CurrentMethodName()} File: {file ?? "NULL"}");

            if (string.IsNullOrWhiteSpace(file)) {
                throw new ArgumentNullException("file");
            }
            if (!File.Exists(file)) {
                throw new FileNotFoundException($"File not found: {file}");
            }

            try {
                var rootElement = XElement.Parse(File.ReadAllText(file));
                var eventElements = rootElement
                    .Elements()
                    .Where(x => x.Name.LocalName == "Event")
                    .ToList();

                Console.WriteLine($"{DateTime.Now.YMDHMSFriendly()} - {ObjectExtensions.CurrentMethodName()} File: {file} Xml Events: {eventElements.Count}");

                foreach (var eventElement in eventElements) {
                    var eventRecordXml = eventElement.ToString(SaveOptions.DisableFormatting);
                    GetEventFromXml(eventRecordXml);
                }
            }
            catch (Exception e) {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error parsing events from file: {file}");
                Console.WriteLine(e.VerboseExceptionString());
                Console.ResetColor();
                throw;
            }
        }

        private static void Initialize() {

            #region EventCategories
            if (ConfigurationManager.AppSettings["EventCategories"] != null) {
                var eventCategories = ConfigurationManager.AppSettings["EventCategories"].Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var eventIdAndCategory in eventCategories) {
                    var eventCategorySplit = eventIdAndCategory.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                    if (eventCategorySplit.Length == 2) {
                        var eventId = -1;
                        if (int.TryParse(eventCategorySplit[0], out eventId)) {
                            if (eventId > 0) {
                                var eventCategory = eventCategorySplit[1];
                                if (!EventCategories.ContainsKey(eventId)) {
                                    EventCategories.Add(eventId, eventCategory);
                                }
                            }
                        }
                    }
                }
            }

            Console.WriteLine("EventCategories:");
            foreach (var item in EventCategories) {
                Console.WriteLine($" - Id: {item.Key} Description: {item.Value}");
            }
            #endregion

            #region EventLogQueryReverseDirection
            if (ConfigurationManager.AppSettings["EventLogQueryReverseDirection"] != null) {
                EventLogQueryReverseDirection = Convert.ToBoolean(ConfigurationManager.AppSettings["EventLogQueryReverseDirection"]);
            }
            else {
                EventLogQueryReverseDirection = true;
            }
            Console.WriteLine($"EventLogQueryReverseDirection: {EventLogQueryReverseDirection}");
            #endregion

        }
        #endregion
    }
}
