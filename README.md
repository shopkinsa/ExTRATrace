# ExTRATrace
This script starts and stops logman trace collection on one or many Exchange servers simultaneously. After collection is stopped logs are collected to local server for review.Engineers can generate ExTRA configuration strings to provide to end users to collect specific datafrom Exchange.

Exchange 2010SP3, 2013, and 2016 supported as long as compatible tags are provided.

# Usage Examples

  - Start ExTRA trace after prompting for configuration
 
    *.\ExTRAtrace.ps1 -Start*

  - Start ExTRA trace on local server and consolidate logs into D:\logs\extra\
 
    *.\ExTRAtrace.ps1 -Start -LogPath "D:\logs\extra\"*

  - Start ExTRA trace on multiple servers and consolidate all logs into D:\logs\extra\ on the local server
 
    *.\ExTRAtrace.ps1 -Start -Servers NA-EXCH01,NA-EXCH02,NA-EXCH03 -LogPath "D:\logs\extra\"*

  - Interactive Configuration generator
 
    *.\ExTRAtrace.ps1 -Generate*

# Parameters

Parameter | Description
--------- | -----------
-Generate | Interactive session to create an ExTRA Trace configuration 
-FreeBusy | Use preconfigured tags for Free/Busy diagnostic
-LogPath | Sets the directory location of where the ETL files will be consolidated. Default Location is C:\extra on the local server.
-Manual | (WIP) Use prebuilt EnabledTraces.Config. Must be located in the same path as ExTRATrace.ps1
-Servers | Specify remote server(s) for tracing.  If no server(s) is specified, the local server will be used.
-Start | Starts an ExTRA trace

# Credits

Thanks to Matthew Huynh for his initial log collection script.
https://blogs.technet.microsoft.com/mahuynh/2016/08/05/script-enable-and-collect-extra-tracing-across-all-exchange-servers/

