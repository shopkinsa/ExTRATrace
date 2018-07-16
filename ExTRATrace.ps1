<#
.NOTES
	Name: ExTRAtrace.ps1
	Author: Shaun Hopkins
	Original Author: Matthew Huynh
	Requires: Exchange Management Shell and administrator rights on the target Exchange
	server as well as the local machine.
	Version History:
	06/28/2017 - Initial Public Release.
	07/12/2018 - Initial Public Release of version 2. - rewritten by Shaun Hopkins.
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
	Collects ExTRA debug logs from Exchange for torubleshooting issues.
.DESCRIPTION
	This script starts and stops logman trace collection on one or many Exchange servers simultaneously.
    After collection is stopped logs are collected to local server for review.
    Engineers can generate ExTRA configuration strings to provide to end users to collect specific data
    from Exchange.
    
    Exchange 2010SP3, 2013, and 2016 supported as long as compatible tags are provided. 
.PARAMETER Servers
	This optional parameter allows multiple target Exchange servers to be specified. If it is not the 		
	local server is assumed.
.PARAMETER Start
	Starts log trace after prompting for configuration data
.PARAMETER FreeBusy
	Use prebuilt configuration for Free Busy tracing
.PARAMETER Stop
	Stops log tracing, cleans up collection, and consolidates all logs to a central folder.
.PARAMETER Generate
	Used to generate Base64 configuration file for debuging tags
.PARAMETER LogPath
	Specify local log consolidation path. Only used with -Stop
.EXAMPLE
	.\ExTRAtrace.ps1 -Generate
	Interactive Configuration generator
.EXAMPLE
	.\ExTRAtrace.ps1 -Start
	Start ExTRA log generation after prompting for configuration
.EXAMPLE
	.\ExTRAtrace.ps1 -Stop -LogPath "D:\logs\extra\"
	Stop ExTRA tracing and consolidate logs into D:\logs\extra\
.LINK
    https://blogs.technet.microsoft.com/mahuynh/2016/08/05/script-enable-and-collect-extra-tracing-across-all-exchange-servers/
#>

[CmdletBinding()]
Param(
 [string]$Servers,
 [string]$logpath, 
 [switch]$Start,
 [switch]$FreeBusy,
 [switch]$Stop,
 [switch]$Generate
 
)

# network path to save the resulting traces (default is local c:\temp\extra)
$script:nl = "`r`n"
# check that user ran with either Start or Stop switch params
if (($Start -and $Generate) -or ($Stop -and $Generate) -or ($Start -and $Stop) -or (-not $Start -and -not $Stop -and -not $Generate)) {
	Write-Error "Please specify only 1 parameter: -Start -Stop or -Generate."
	exit
}

function Set-CB() {
    Param(
      [Parameter(ValueFromPipeline=$true)]
      [string] $text
    )
    Add-Type -AssemblyName System.Windows.Forms
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true
    $tb.Text = $text
    $tb.SelectAll()
    $tb.Copy()
}

function CreateExtraTraceConfig
{
	$string = "TraceLevels:Debug,Warning,Error,Fatal,Info,Pfd`nSystemLogging:SystemNet`n"
	if ($FreeBusy) {$string += [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes('QURQcm92aWRlcjogVG9wb2xvZ3lQcm92aWRlciwgQURUb3BvbG9neSwgQ29ubmVjdGlvbiwgQ29ubmVjdGlvbkRldGFpbHMsIEdldENvbm5lY3Rpb24sIEFERmluZCwgQURSZWFkLCBBRFJlYWREZXRhaWxzLCBBRFNhdmUsIEFEU2F2ZURldGFpbHMsIEFERGVsZXRlLCBWYWxpZGF0aW9uLCBBRE5vdGlmaWNhdGlvbnMsIERpcmVjdG9yeUV4Y2VwdGlvbiwgTGRhcEZpbHRlckJ1aWxkZXIsIEFEUHJvcGVydHlSZXF1ZXN0LCBBRE9iamVjdCwgQ29udGVudFR5cGVNYXBwaW5nLCBMY2lkTWFwcGVyLCBSZWNpcGllbnRVcGRhdGVTZXJ2aWNlLCBVTUF1dG9BdHRlbmRhbnQsIEV4Y2hhbmdlVG9wb2xvZ3ksIFBlcmZDb3VudGVycywgQ2xpZW50VGhyb3R0bGluZywgU2VydmVyU2V0dGluZ3NQcm92aWRlciwgUmV0cnlNYW5hZ2VyLCBTeXN0ZW1Db25maWd1cmF0aW9uQ2FjaGUsIEZlZGVyYXRlZElkZW50aXR5LCBGYXVsdEluamVjdGlvbiwgQWRkcmVzc0xpc3QsIE5zcGlScGNDbGllbnRDb25uZWN0aW9uLCBTY29wZVZlcmlmaWNhdGlvbiwgU2NoZW1hSW5pdGlhbGl6YXRpb24sIElzTWVtYmVyT2ZSZXNvbHZlciwgT3dhU2VnbWVudGF0aW9uLCBBRFBlcmZvcm1hbmNlLCBSZXNvdXJjZUhlYWx0aE1hbmFnZXIsIEJ1ZGdldERlbGF5LCBHTFMsIE1TZXJ2LCBUZW5hbnRSZWxvY2F0aW9uLCBTdGF0ZU1hbmFnZW1lbnQsIFNlcnZlckNvbXBvbmVudFN0YXRlTWFuYWdlciwgU2Vzc2lvblNldHRpbmdzLCBBRENvbmZpZ0xvYWRlciwgU2xpbVRlbmFudCwgVGVuYW50VXBncmFkZVNlcnZpY2VsZXQsIERpcmVjdG9yeVRhc2tzLCBDb21wbGlhbmNlLCBMaW5rZWRSb2xlR3JvdXAsIE1TZXJ2RGF0YSwgQWNjb3VudFZhbGlkYXRpb24sIEFER2xvYmFsQ29uZmlnLCBBZ2dyZWdhdGVSZWNpcGllbnQsIFRlbmFudEh5ZHJhdGlvblNlcnZpY2VsZXQsIFRlbmFudFJlcGFpclNjYW5uZXJTZXJ2aWNlbGV0LCBUZW5hbnRSZXBhaXJQcm9jZXNzb3JTZXJ2aWNlbGV0LCBUZW5hbnRBbGxvd0Jsb2NrTGlzdEFsbG93QmxvY2tMaXN0LCBEZWRpY2F0ZWRNYWlsYm94UGxhbnNDdXN0b21BdHRyaWJ1dGUKSW5mb1dvcmtlci5BdmFpbGFiaWxpdHk6IEluaXRpYWxpemUsIFNlY3VyaXR5LCBDYWxlbmRhclZpZXcsIENvbmZpZ3VyYXRpb24sIFB1YmxpY0ZvbGRlclJlcXVlc3QsIEludHJhU2l0ZUNhbGVuZGFyUmVxdWVzdCwgTWVldGluZ1N1Z2dlc3Rpb25zLCBBdXRvRGlzY292ZXIsIE1haWxib3hDb25uZWN0aW9uQ2FjaGUsIFBGRCwgRG5zUmVhZGVyLCBNZXNzYWdlLCBGYXVsdEluamVjdGlvbkluZm9Xb3JrZXIuUmVxdWVzdERpc3BhdGNoOiBSZXF1ZXN0Um91dGluZywgRGlzdHJpYnV0aW9uTGlzdEhhbmRsaW5nLCBQcm94eVdlYlJlcXVlc3QsIEZhdWx0SW5qZWN0aW9uLCBHZXRGb2xkZXJSZXF1ZXN0TVNFeGNoYW5nZUF1dG9kaXNjb3ZlcjogRnJhbWV3b3JrLCBPdXRsb29rUHJvdmlkZXIsIE1vYmlsZVN5bmNQcm92aWRlciwgRmF1bHRJbmplY3Rpb24sIEF1dGhNZXRhZGF0YQpNU0V4Y2hhbmdlV2ViU2VydmljZXM6IENhbGVuZGFyQWxnb3JpdGhtLCBDYWxlbmRhckRhdGEsIENhbGVuZGFyQ2FsbCwgQ29tbW9uQWxnb3JpdGhtLCBGb2xkZXJBbGdvcml0aG0sIEZvbGRlckRhdGEsIEZvbGRlckNhbGwsIEl0ZW1BbGdvcml0aG0sIEl0ZW1EYXRhLCBJdGVtQ2FsbCwgRXhjZXB0aW9uLCBTZXNzaW9uQ2FjaGUsIEV4Y2hhbmdlUHJpbmNpcGFsQ2FjaGUsIFNlYXJjaCwgVXRpbEFsZ29yaXRobSwgVXRpbERhdGEsIFV0aWxDYWxsLCBTZXJ2ZXJUb1NlcnZlckF1dGhaLCBTZXJ2aWNlQ29tbWFuZEJhc2VDYWxsLCBTZXJ2aWNlQ29tbWFuZEJhc2VEYXRhLCBGYWNhZGVCYXNlQ2FsbCwgQ3JlYXRlSXRlbUNhbGwsIEdldEl0ZW1DYWxsLCBVcGRhdGVJdGVtQ2FsbCwgRGVsZXRlSXRlbUNhbGwsIFNlbmRJdGVtQ2FsbCwgTW92ZUNvcHlDb21tYW5kQmFzZUNhbGwsIE1vdmVDb3B5SXRlbUNvbW1hbmRCYXNlQ2FsbCwgQ29weUl0ZW1DYWxsLCBNb3ZlSXRlbUNhbGwsIENyZWF0ZUZvbGRlckNhbGwsIEdldEZvbGRlckNhbGwsIFVwZGF0ZUZvbGRlckNhbGwsIERlbGV0ZUZvbGRlckNhbGwsIE1vdmVDb3B5Rm9sZGVyQ29tbWFuZEJhc2VDYWxsLCBDb3B5Rm9sZGVyQ2FsbCwgTW92ZUZvbGRlckNhbGwsIEZpbmRDb21tYW5kQmFzZUNhbGwsIEZpbmRJdGVtQ2FsbCwgRmluZEZvbGRlckNhbGwsIFV0aWxDb21tYW5kQmFzZUNhbGwsIEV4cGFuZERMQ2FsbCwgUmVzb2x2ZU5hbWVzQ2FsbCwgU3Vic2NyaWJlQ2FsbCwgVW5zdWJzY3JpYmVDYWxsLCBHZXRFdmVudHNDYWxsLCBTdWJzY3JpcHRpb25zLCBTdWJzY3JpcHRpb25CYXNlLCBQdXNoU3Vic2NyaXB0aW9uLCBTeW5jRm9sZGVySGllcmFyY2h5Q2FsbCwgU3luY0ZvbGRlckl0ZW1zQ2FsbCwgU3luY2hyb25pemF0aW9uLCBQZXJmb3JtYW5jZU1vbml0b3IsIENvbnZlcnRJZENhbGwsIEdldERlbGVnYXRlQ2FsbCwgQWRkRGVsZWdhdGVDYWxsLCBSZW1vdmVEZWxlZ2F0ZUNhbGwsIFVwZGF0ZURlbGVnYXRlQ2FsbCwgUHJveHlFdmFsdWF0b3IsIEdldE1haWxUaXBzQ2FsbCwgQWxsUmVxdWVzdHMsIEF1dGhlbnRpY2F0aW9uLCBXQ0YsIEdldFVzZXJDb25maWd1cmF0aW9uQ2FsbCwgQ3JlYXRlVXNlckNvbmZpZ3VyYXRpb25DYWxsLCBEZWxldGVVc2VyQ29uZmlndXJhdGlvbkNhbGwsIFVwZGF0ZVVzZXJDb25maWd1cmF0aW9uQ2FsbCwgVGhyb3R0bGluZywgRXh0ZXJuYWxVc2VyLCBHZXRPcmdhbml6YXRpb25Db25maWd1cmF0aW9uQ2FsbCwgR2V0Um9vbXNDYWxsLCBHZXRGZWRlcmF0aW9uSW5mb3JtYXRpb24sIFBhcnRpY2lwYW50TG9va3VwQmF0Y2hpbmcsIEFsbFJlc3BvbnNlcywgRmF1bHRJbmplY3Rpb24sIEdldEluYm94UnVsZXNDYWxsLCBVcGRhdGVJbmJveFJ1bGVzQ2FsbCwgR2V0Q0FTTWFpbGJveCwgRmFzdFRyYW5zZmVyLCBTeW5jQ29udmVyc2F0aW9uQ2FsbCwgRUxDLCBBY3Rpdml0eUNvbnZlcnRlciwgU3luY1Blb3BsZUNhbGwsIEdldENhbGVuZGFyRm9sZGVyc0NhbGwsIEdldFJlbWluZGVyc0NhbGwsIFN5bmNDYWxlbmRhckNhbGwsIFBlcmZvcm1SZW1pbmRlckFjdGlvbkNhbGwsIFByb3Zpc2lvbkNhbGwsIFJlbmFtZUNhbGVuZGFyR3JvdXBDYWxsLCBEZWxldGVDYWxlbmRhckdyb3VwQ2FsbCwgQ3JlYXRlQ2FsZW5kYXJDYWxsLCBSZW5hbWVDYWxlbmRhckNhbGwsIERlbGV0ZUNhbGVuZGFyQ2FsbCwgU2V0Q2FsZW5kYXJDb2xvckNhbGwsIFNldENhbGVuZGFyR3JvdXBPcmRlckNhbGwsIENyZWF0ZUNhbGVuZGFyR3JvdXBDYWxsLCBNb3ZlQ2FsZW5kYXJDYWxsLCBHZXRGYXZvcml0ZXNDYWxsLCBVcGRhdGVGYXZvcml0ZUZvbGRlckNhbGwsIEdldFRpbWVab25lT2Zmc2V0c0NhbGwsIEF1dGhvcml6YXRpb24sIFNlbmRDYWxlbmRhclNoYXJpbmdJbnZpdGVDYWxsLCBHZXRDYWxlbmRhclNoYXJpbmdSZWNpcGllbnRJbmZvQ2FsbCwgQWRkU2hhcmVkQ2FsZW5kYXJDYWxsLCBGaW5kUGVvcGxlQ2FsbCwgRmluZFBsYWNlc0NhbGwsIFVzZXJQaG90b3MsIEdldFBlcnNvbmFDYWxsLCBHZXRFeHRlbnNpYmlsaXR5Q29udGV4dENhbGwsIFN1YnNjcmliZUludGVybmFsQ2FsZW5kYXJDYWxsLCBTdWJzY3JpYmVJbnRlcm5ldENhbGVuZGFyQ2FsbCwgR2V0VXNlckF2YWlsYWJpbGl0eUludGVybmFsQ2FsbCwgQXBwbHlDb252ZXJzYXRpb25BY3Rpb25DYWxsLCBHZXRDYWxlbmRhclNoYXJpbmdQZXJtaXNzaW9uc0NhbGwsIFNldENhbGVuZGFyU2hhcmluZ1Blcm1pc3Npb25zQ2FsbCwgU2V0Q2FsZW5kYXJQdWJsaXNoaW5nQ2FsbCwgVUNTLCBHZXRUYXNrRm9sZGVyc0NhbGwsIENyZWF0ZVRhc2tGb2xkZXJDYWxsLCBSZW5hbWVUYXNrRm9sZGVyQ2FsbCwgRGVsZXRlVGFza0ZvbGRlckNhbGwsIE1hc3RlckNhdGVnb3J5TGlzdENhbGwsIEdldENhbGVuZGFyRm9sZGVyQ29uZmlndXJhdGlvbkNhbGwsIE9ubGluZU1lZXRpbmcsIE1vZGVybkdyb3VwcywgQ3JlYXRlVW5pZmllZE1haWxib3gsIEFkZEFnZ3JlZ2F0ZWRBY2NvdW50LCBSZW1pbmRlcnMsIEdldEFnZ3JlZ2F0ZWRBY2NvdW50LCBSZW1vdmVBZ2dyZWdhdGVkQWNjb3VudCwgU2V0QWdncmVnYXRlZEFjY291bnQsIFdlYXRoZXIsIEdldFBlb3BsZUlLbm93R3JhcGhDYWxsLCBBZGRFdmVudFRvTXlDYWxlbmRhciwgQ29udmVyc2F0aW9uQWdncmVnYXRpb24sIElzT2ZmaWNlMzY1RG9tYWluLCBSZWZyZXNoR0FMQ29udGFjdHNGb2xkZXIsIE9wdGlvbnMsIE9wZW5UZW5hbnRNYW5hZ2VyLCBNYXJrQWxsSXRlbXNBc1JlYWQsIEdldENvbnZlcnNhdGlvbkl0ZW1zLCBHZXRMaWtlcnMsIEdldFVzZXJVbmlmaWVkR3JvdXBzLCBQZW9wbGVJQ29tbXVuaWNhdGVXaXRoLCBTeW5jUGVyc29uYUNvbnRhY3RzQmFzZSwgU3luY0F1dG9Db21wbGV0ZVJlY2lwaWVudHMsIFNldFVuaWZpZWRHcm91cEZhdm9yaXRlU3RhdGUsIEdldFVuaWZpZWRHcm91cERldGFpbHMsIEdldFVuaWZpZWRHcm91cE1lbWJlcnMsIFNldFVuaWZpZWRHcm91cFVzZXJTdWJzY3JpYmVTdGF0ZSwgSm9pblByaXZhdGVVbmlmaWVkR3JvdXAsIEFwcGx5QnVsa0l0ZW1BY3Rpb25DYWxsLCBDcmVhdGVTd2VlcFJ1bGVGb3JTZW5kZXJDYWxsLCBDcmVhdGVVbmlmaWVkR3JvdXAsIFZhbGlkYXRlVW5pZmllZEdyb3VwQWxpYXMsIEltcG9ydENhbGVuZGFyRXZlbnQsIFJlbW92ZVVuaWZpZWRHcm91cCwgU2V0VW5pZmllZEdyb3VwTWVtYmVyc2hpcFN0YXRlLCBHZXRVbmlmaWVkR3JvdXBVbnNlZW5EYXRhLCBTZXRVbmlmaWVkR3JvdXBVbnNlZW5EYXRhLCBVcGRhdGVVbmlmaWVkR3JvdXAsIEdldEF2YWlsYWJsZUN1bHR1cmVzLCBVc2VyU29jaWFsQWN0aXZpdHlOb3RpZmljYXRpb24sIEdldFNvY2lhbEFjdGl2aXR5Tm90aWZpY2F0aW9ucywgQ2hhbm5lbEV2ZW50LCBHZXRVbmlmaWVkR3JvdXBzU2V0dGluZ3MsIEdldFBlb3BsZUluc2lnaHRzLCBHZXRVbmlmaWVkR3JvdXBVbnNlZW5Db3VudCwgU2V0VW5pZmllZEdyb3VwTGFzdFZpc2l0ZWRUaW1lLCBDb250YWN0UHJvcGVydHlTdWdnZXN0aW9uLCBHZXRVbmlmaWVkR3JvdXBTZW5kZXJSZXN0cmljdGlvbnMsIFNldFVuaWZpZWRHcm91cFNlbmRlclJlc3RyaWN0aW9ucywgQ29udmVydEljc1RvQ2FsZW5kYXJJdGVtLCBNZXNzYWdlTGF0ZW5jeSwgR2V0UGVvcGxlSW5zaWdodHNUb2tlbnMsIFNldFBlb3BsZUluc2lnaHRzVG9rZW5zLCBEZWxldGVQZW9wbGVJbnNpZ2h0c1Rva2VucywgRXhlY3V0ZVNlYXJjaCwgUHJvY2Vzc0NvbXBsaWFuY2VPcGVyYXRpb25DYWxsLCBEZWxlZ2F0ZUNvbW1hbmRCYXNlLCBHZXRTdWdnZXN0ZWRVbmlmaWVkR3JvdXBzLCBPRGF0YUNvbW1vbiwgT0RhdGFQdXNoTm90aWZpY2F0aW9uTmV0d29ya2luZ0xheWVyOiBETlMsIE5ldHdvcmssIEF1dGhlbnRpY2F0aW9uLCBDZXJ0aWZpY2F0ZSwgRGlyZWN0b3J5U2VydmljZXMsIFByb2Nlc3NNYW5hZ2VyLCBIdHRwQ2xpZW50LCBQcm90b2NvbExvZywgUmlnaHRzTWFuYWdlbWVudCwgTGl2ZUlEQXV0aGVudGljYXRpb25DbGllbnQsIERlbHRhU3luY0NsaWVudCwgRGVsdGFTeW5jUmVzcG9uc2VIYW5kbGVyLCBMYW5ndWFnZVBhY2tJbmZvLCBXU1RydXN0LCBFd3NDbGllbnQsIENvbmZpZ3VyYXRpb24sIFNtdHBDbGllbnQsIFhyb3BTZXJ2aWNlQ2xpZW50LCBYcm9wU2VydmljZVNlcnZlciwgQ2xhaW0sIEZhY2Vib29rLCBMaW5rZWRJbiwgTW9uaXRvcmluZ1dlYkNsaWVudCwgUnVsZXNCYXNlZEh0dHBNb2R1bGUsIEFBRENsaWVudCwgQXBwU2V0dGluZ3MsIENvbW1vbgo='))} 
	elseif ($Transport) {$string += [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes('STRINGHERE'))}
	elseif ($Manual) {$string += [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes('STRINGHERE'))}
	else {$string += [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($(Read-Host -Prompt 'Please enter ExTRA configuration')))}
	$string += "`nTransportRuleAgent:FaultInjection`nFilteredTracing:No`nInMemoryTracing:No`n"
	new-item -path "C:\EnabledTraces.Config" -type file -force | Out-Null
	$string | Out-File -filepath "C:\EnabledTraces.Config" -Encoding ASCII -Append | Out-Null
}

function GetExchServers
{
    # if no server is specified to the script, use the local computer name
	if(!$Servers)
	{
		$Servers = ${env:computername}
        Write-Debug "No Server list specified. Using local Server..."
	}
	foreach($serv in $Servers)
	{$return += (Get-ExchangeServer $serv)}
	return $return
}

Function StartTrace 
{
	$servlist = GetExchServers
	CreateExtraTraceConfig
	$filepath = "c:\tracing\"
	$ts = get-date -f HHmmssddMMyy
	foreach ($s in $servlist)
	{
		Write-Host "Enabling ExTRA tracing on" ($s) -ForegroundColor green $nl
		If (Test-Connection -BufferSize 32 -Count 1 -ComputerName $s -Quiet)
		{
			# Check if ExTRA Trace already exists
			$CheckExTRA = @(logman query -s $s) -match "ExchangeDebugTraces"
			if (!$CheckExTRA)
			{
				Write-Host "Creating Trace... " -NoNewline
				$ExTRAcmd = "logman create trace ExchangeDebugTraces -p '{79bb49e6-2a2c-46e4-9167-fa122525d540}' -o $filepath$s-ExTRA-$ts.etl -s $s -ow -f bin -max 1024"
				# Create ExTRA Trace
				Write-Debug $ExTRAcmd
				Invoke-Expression -Command $ExTRAcmd
				while (!($CheckifCreated = @(logman query -s $s) -match "ExchangeDebugTraces"))
				{
					Write-Host " Traced failed to create. Would you like to try creating it again? " -NoNewline
					$answer = ConfirmAnswer
					if ($answer -eq "yes"){Invoke-Expression -Command $ExTRAcmd}
					if ($answer -eq "no"){Continue}
				}
				Write-Host "COMPLETED" -ForegroundColor green
				Write-Host "Starting Trace... " -NoNewline
				$ExTRAcmd = "logman start ExchangeDebugTraces -s $s"
				# Create ExTRA Trace
				Write-Debug $ExTRAcmd
				Invoke-Expression -Command $ExTRAcmd
				$CheckExTRA = @(logman query -s $ServerName) -match "ExchangeDebugTraces"
				$CheckifRunning = select-string -InputObject $CheckExTRA -pattern "Running" -quiet
				if ($CheckifRunning)
				{
					Write-Host "COMPLETED" -ForegroundColor green
					$cmd = "logman stop -n ExchangeDebugTraces -s $Servername"
					$StopExmon = Invoke-Expression -Command $cmd
					Write-Host ""
				}
			}
			else
			{
				Write-Host "ExchangeDebugTraces already exists. Checking if already running"
				$CheckifRunning = select-string -InputObject $CheckExTRA -pattern "Running" -quiet
				if ($CheckifRunning)
				{
					$cmd = "logman stop ExchangeDebugTraces -s $s"
					$StopExTRA = Invoke-Expression -Command $Cmd
					Start-Sleep 2
				}
				#Delete and recreate ExTRA tracing
				Write-Host "Deleting and recreating ExchangeDebugTraces"
				$cmd = "logman delete ExchangeDebugTraces -s $s"
				$DeleteExTRA = Invoke-Expression -Command $Cmd 
				# Create ExTRA Trace
				$ExTRAcmd = "logman create trace ExchangeDebugTraces -p '{79bb49e6-2a2c-46e4-9167-fa122525d540}' -o $filepath$s-ExTRA-$ts.etl -s $s -ow -f bin -max 1024"
				Write-Debug $ExTRAcmd
				Invoke-Expression -Command $ExTRAcmd
				while (!($CheckifCreated = @(logman query -s $s) -match "ExchangeDebugTraces"))
				{
					Write-Host "ExTRA Traced failed to create. Would you like to try creating it again? " -NoNewline
					$answer = ConfirmAnswer
					if ($answer -eq "yes"){Invoke-Expression -Command $ExTRAcmd}
					if ($answer -eq "no"){Continue}
				}
			}
		}
		Else
		{
			Write-Host "Server $s cannot be contacted. Skipping..."  -foregroundcolor Red $nl
			Continue
		}
	}
}

Function StopTrace
{
	if ($logpath -eq "") {$logpath = "C:\extra\"} elseif ($logpath.EndsWith("\")) {$logpath = $logpath} else {$logpath = $logpath + "\"}
	# Convert logpath to UNC adminshare path
	$TRACES_FILEPATH = "\\" + (hostname) + "\"+ $logpath.replace(':','$') + $(get-date -f HHmmssddMMyy)
	# create target path if it does not exist yet
	if (-not (Test-Path $TRACES_FILEPATH)) {
		New-Item $TRACES_FILEPATH -ItemType Directory | Out-Null
		Write-Host "Created $TRACES_FILEPATH as it did not exist yet" $nl
	}
	$servlist = GetExchServers
	foreach ($s in $servlist)
	{
		$CheckExmon = @(logman query -s $s) -match "ExchangeDebugTraces"
		$CheckifRunning = select-string -InputObject $CheckExmon -pattern "Running" -quiet
		if (!$CheckifRunning)
		{
			Write-Host "Stopping trace on $s... " -NoNewline
			$Error.Clear()
			$cmd = "logman stop -n ExchangeDebugTraces -s $s"
			$StopTrace = Invoke-Expression -Command $cmd
			if ($Error){Write-host "Error encountered" -ForegroundColor Red; Continue}
			else {Write-Host "COMPLETED`n" -ForegroundColor Green}
			
			Write-Host "Removing trace on $s... " -NoNewline
			$Error.Clear()
			$cmd = "logman delete ExchangeDebugTraces -s $s"
			$StopTrace = Invoke-Expression -Command $cmd
			if ($Error){Write-host "Error encountered" -ForegroundColor Red}
			else {Write-Host "COMPLETED`n" -ForegroundColor Green}
		}
		Write-Host "Transfering trace logs from $s... " -NoNewline
		$fileToMovePath = "\\" + $s + "\c$\tracing\*.etl"
	    try { Move-Item $fileToMovePath $TRACES_FILEPATH -Force}
        Catch { Write-Host "FAILED "-ForegroundColor red $nl; Continue }
		Write-Host "COMPLETED`n" -ForegroundColor Green
	}
	Write-Host "Logs can be found in" $logpath
}

Function Generate
{
    Write-Host "Input trace lines. Empty line to finish" -ForegroundColor Green $nl																					 
    #prompt for trace definations
    while ($true) {
        $input = Read-Host -Prompt ' '
        if ($input -eq '') {break}
        Else {$string += $input + "`n"}
    }																					
    $EncodedText =[Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($string))
    Write-Host "Send the following line to customer for ExTRA Trace Configuration.`nConfig has been copied to clipboard" -ForegroundColor Green $nl
	Start-Sleep 1
	Set-CB $EncodedText
    Write-Host $EncodedText + $nl

}
		
if ($generate) {Generate; exit;}
if ($start) {StartTrace; exit;}
if ($stop) {StopTrace; exit;}
