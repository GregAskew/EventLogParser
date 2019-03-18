# EventLogParser
Utility for parsing events from an event log evtx file or a remote computer.
Usage: EventLogParser.exe
	/FilePath:<pathToEvtx> | /ComputerFqdn:<computerName.company.com> /LogName:<logName>
	[/Events:EventId,EventId,EventId]
	[/Startdate:yyyy-MM-dd]
	[/Enddate:yyyy-MM-dd]
	[/Format:CSV | XML]
