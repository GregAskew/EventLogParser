# EventLogParser
Utility for parsing events from an event log evtx file or a remote computer.  
Usage: EventLogParser.exe  
	/FilePath:<pathToEvtx> | /ComputerFqdn:<computerName.company.com> /LogName:<logName>  
	[/Events:EventId,EventId,EventId]  
	[/StartDate:yyyy-MM-dd]  
	[/EndDate:yyyy-MM-dd]  
	[/Format:CSV | XML]  
