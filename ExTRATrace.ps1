<#
.NOTES
	Name: ExTRAtrace.ps1
	Version: 0.9.80
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
	Specify local log consolidation path.
.EXAMPLE
	.\ExTRAtrace.ps1 -Generate
	Interactive Configuration generator
.EXAMPLE
	.\ExTRAtrace.ps1 -Start
	Start ExTRA log generation after prompting for configuration
.EXAMPLE
	.\ExTRAtrace.ps1 -Start -LogPath "D:\logs\extra\"
	Start ExTRA log generation after prompting for configuration
.EXAMPLE
	.\ExTRAtrace.ps1 -Start -Servers NA-EXCH01,NA-EXCH02,NA-EXCH04
	Start ExTRA log generation on multiple servers
.LINK
    https://blogs.technet.microsoft.com/mahuynh/2016/08/05/script-enable-and-collect-extra-tracing-across-all-exchange-servers/
#>

[CmdletBinding()]
Param(
 [Array]$Servers,
 [string]$LogPath, 
 [switch]$Start,
 [switch]$FreeBusy,
 [switch]$Generate,
 [switch]$Manual
 
)

$script:nl = "`r`n"
# check that user ran with either Start or Stop switch params
if (($Start -and $Generate)) {
	Write-Error "Please specify only 1 parameter: -Start or -Generate."
	exit
}

# Set clipboard
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

#Function to test if you are an admin on the server 
Function Is-Admin {
	$currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
	If( $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
		return $true
	}
	else {
		return $false
	}
}

function CreateExtraTraceConfig
{
	$string = "TraceLevels:Debug,Warning,Error,Fatal,Info,Pfd`n"
	if ($FreeBusy) {$string += [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String('U3lzdGVtTG9nZ2luZzogU3lzdGVtTmV0Ck1TRXhjaGFuZ2VXZWJTZXJ2aWNlczogQ2FsZW5kYXJBbGdvcml0aG0sIENhbGVuZGFyRGF0YSwgQ2FsZW5kYXJDYWxsLCBDb21tb25BbGdvcml0aG0sIEZvbGRlckFsZ29yaXRobSwgRm9sZGVyRGF0YSwgRm9sZGVyQ2FsbCwgSXRlbUFsZ29yaXRobSwgSXRlbURhdGEsIEl0ZW1DYWxsLCBFeGNlcHRpb24sIFNlc3Npb25DYWNoZSwgRXhjaGFuZ2VQcmluY2lwYWxDYWNoZSwgU2VhcmNoLCBVdGlsQWxnb3JpdGhtLCBVdGlsRGF0YSwgVXRpbENhbGwsIFNlcnZlclRvU2VydmVyQXV0aFosIFNlcnZpY2VDb21tYW5kQmFzZUNhbGwsIFNlcnZpY2VDb21tYW5kQmFzZURhdGEsIEZhY2FkZUJhc2VDYWxsLCBDcmVhdGVJdGVtQ2FsbCwgR2V0SXRlbUNhbGwsIFVwZGF0ZUl0ZW1DYWxsLCBEZWxldGVJdGVtQ2FsbCwgU2VuZEl0ZW1DYWxsLCBNb3ZlQ29weUNvbW1hbmRCYXNlQ2FsbCwgTW92ZUNvcHlJdGVtQ29tbWFuZEJhc2VDYWxsLCBDb3B5SXRlbUNhbGwsIE1vdmVJdGVtQ2FsbCwgQ3JlYXRlRm9sZGVyQ2FsbCwgR2V0Rm9sZGVyQ2FsbCwgVXBkYXRlRm9sZGVyQ2FsbCwgRGVsZXRlRm9sZGVyQ2FsbCwgTW92ZUNvcHlGb2xkZXJDb21tYW5kQmFzZUNhbGwsIENvcHlGb2xkZXJDYWxsLCBNb3ZlRm9sZGVyQ2FsbCwgRmluZENvbW1hbmRCYXNlQ2FsbCwgRmluZEl0ZW1DYWxsLCBGaW5kRm9sZGVyQ2FsbCwgVXRpbENvbW1hbmRCYXNlQ2FsbCwgRXhwYW5kRExDYWxsLCBSZXNvbHZlTmFtZXNDYWxsLCBTdWJzY3JpYmVDYWxsLCBVbnN1YnNjcmliZUNhbGwsIEdldEV2ZW50c0NhbGwsIFN1YnNjcmlwdGlvbnMsIFN1YnNjcmlwdGlvbkJhc2UsIFB1c2hTdWJzY3JpcHRpb24sIFN5bmNGb2xkZXJIaWVyYXJjaHlDYWxsLCBTeW5jRm9sZGVySXRlbXNDYWxsLCBTeW5jaHJvbml6YXRpb24sIFBlcmZvcm1hbmNlTW9uaXRvciwgQ29udmVydElkQ2FsbCwgR2V0RGVsZWdhdGVDYWxsLCBBZGREZWxlZ2F0ZUNhbGwsIFJlbW92ZURlbGVnYXRlQ2FsbCwgVXBkYXRlRGVsZWdhdGVDYWxsLCBQcm94eUV2YWx1YXRvciwgR2V0TWFpbFRpcHNDYWxsLCBBbGxSZXF1ZXN0cywgQXV0aGVudGljYXRpb24sIFdDRiwgR2V0VXNlckNvbmZpZ3VyYXRpb25DYWxsLCBDcmVhdGVVc2VyQ29uZmlndXJhdGlvbkNhbGwsIERlbGV0ZVVzZXJDb25maWd1cmF0aW9uQ2FsbCwgVXBkYXRlVXNlckNvbmZpZ3VyYXRpb25DYWxsLCBUaHJvdHRsaW5nLCBFeHRlcm5hbFVzZXIsIEdldE9yZ2FuaXphdGlvbkNvbmZpZ3VyYXRpb25DYWxsLCBHZXRSb29tc0NhbGwsIEdldEZlZGVyYXRpb25JbmZvcm1hdGlvbiwgUGFydGljaXBhbnRMb29rdXBCYXRjaGluZywgQWxsUmVzcG9uc2VzLCBGYXVsdEluamVjdGlvbiwgR2V0SW5ib3hSdWxlc0NhbGwsIFVwZGF0ZUluYm94UnVsZXNDYWxsLCBHZXRDQVNNYWlsYm94LCBGYXN0VHJhbnNmZXIsIFN5bmNDb252ZXJzYXRpb25DYWxsLCBFTEMsIEFjdGl2aXR5Q29udmVydGVyLCBTeW5jUGVvcGxlQ2FsbCwgR2V0Q2FsZW5kYXJGb2xkZXJzQ2FsbCwgR2V0UmVtaW5kZXJzQ2FsbCwgU3luY0NhbGVuZGFyQ2FsbCwgUGVyZm9ybVJlbWluZGVyQWN0aW9uQ2FsbCwgUHJvdmlzaW9uQ2FsbCwgUmVuYW1lQ2FsZW5kYXJHcm91cENhbGwsIERlbGV0ZUNhbGVuZGFyR3JvdXBDYWxsLCBDcmVhdGVDYWxlbmRhckNhbGwsIFJlbmFtZUNhbGVuZGFyQ2FsbCwgRGVsZXRlQ2FsZW5kYXJDYWxsLCBTZXRDYWxlbmRhckNvbG9yQ2FsbCwgU2V0Q2FsZW5kYXJHcm91cE9yZGVyQ2FsbCwgQ3JlYXRlQ2FsZW5kYXJHcm91cENhbGwsIE1vdmVDYWxlbmRhckNhbGwsIEdldEZhdm9yaXRlc0NhbGwsIFVwZGF0ZUZhdm9yaXRlRm9sZGVyQ2FsbCwgR2V0VGltZVpvbmVPZmZzZXRzQ2FsbCwgQXV0aG9yaXphdGlvbiwgU2VuZENhbGVuZGFyU2hhcmluZ0ludml0ZUNhbGwsIEdldENhbGVuZGFyU2hhcmluZ1JlY2lwaWVudEluZm9DYWxsLCBBZGRTaGFyZWRDYWxlbmRhckNhbGwsIEZpbmRQZW9wbGVDYWxsLCBGaW5kUGxhY2VzQ2FsbCwgVXNlclBob3RvcywgR2V0UGVyc29uYUNhbGwsIEdldEV4dGVuc2liaWxpdHlDb250ZXh0Q2FsbCwgU3Vic2NyaWJlSW50ZXJuYWxDYWxlbmRhckNhbGwsIFN1YnNjcmliZUludGVybmV0Q2FsZW5kYXJDYWxsLCBHZXRVc2VyQXZhaWxhYmlsaXR5SW50ZXJuYWxDYWxsLCBBcHBseUNvbnZlcnNhdGlvbkFjdGlvbkNhbGwsIEdldENhbGVuZGFyU2hhcmluZ1Blcm1pc3Npb25zQ2FsbCwgU2V0Q2FsZW5kYXJTaGFyaW5nUGVybWlzc2lvbnNDYWxsLCBTZXRDYWxlbmRhclB1Ymxpc2hpbmdDYWxsLCBVQ1MsIEdldFRhc2tGb2xkZXJzQ2FsbCwgQ3JlYXRlVGFza0ZvbGRlckNhbGwsIFJlbmFtZVRhc2tGb2xkZXJDYWxsLCBEZWxldGVUYXNrRm9sZGVyQ2FsbCwgTWFzdGVyQ2F0ZWdvcnlMaXN0Q2FsbCwgR2V0Q2FsZW5kYXJGb2xkZXJDb25maWd1cmF0aW9uQ2FsbCwgT25saW5lTWVldGluZywgTW9kZXJuR3JvdXBzLCBDcmVhdGVVbmlmaWVkTWFpbGJveCwgQWRkQWdncmVnYXRlZEFjY291bnQsIFJlbWluZGVycywgR2V0QWdncmVnYXRlZEFjY291bnQsIFJlbW92ZUFnZ3JlZ2F0ZWRBY2NvdW50LCBTZXRBZ2dyZWdhdGVkQWNjb3VudCwgV2VhdGhlciwgR2V0UGVvcGxlSUtub3dHcmFwaENhbGwsIEFkZEV2ZW50VG9NeUNhbGVuZGFyLCBDb252ZXJzYXRpb25BZ2dyZWdhdGlvbiwgSXNPZmZpY2UzNjVEb21haW4sIFJlZnJlc2hHQUxDb250YWN0c0ZvbGRlciwgT3B0aW9ucywgT3BlblRlbmFudE1hbmFnZXIsIE1hcmtBbGxJdGVtc0FzUmVhZCwgR2V0Q29udmVyc2F0aW9uSXRlbXMsIEdldExpa2VycywgR2V0VXNlclVuaWZpZWRHcm91cHMsIFBlb3BsZUlDb21tdW5pY2F0ZVdpdGgsIFN5bmNQZXJzb25hQ29udGFjdHNCYXNlLCBTeW5jQXV0b0NvbXBsZXRlUmVjaXBpZW50cywgU2V0VW5pZmllZEdyb3VwRmF2b3JpdGVTdGF0ZSwgR2V0VW5pZmllZEdyb3VwRGV0YWlscywgR2V0VW5pZmllZEdyb3VwTWVtYmVycywgU2V0VW5pZmllZEdyb3VwVXNlclN1YnNjcmliZVN0YXRlLCBKb2luUHJpdmF0ZVVuaWZpZWRHcm91cCwgQXBwbHlCdWxrSXRlbUFjdGlvbkNhbGwsIENyZWF0ZVN3ZWVwUnVsZUZvclNlbmRlckNhbGwsIENyZWF0ZVVuaWZpZWRHcm91cCwgVmFsaWRhdGVVbmlmaWVkR3JvdXBBbGlhcywgSW1wb3J0Q2FsZW5kYXJFdmVudCwgUmVtb3ZlVW5pZmllZEdyb3VwLCBTZXRVbmlmaWVkR3JvdXBNZW1iZXJzaGlwU3RhdGUsIEdldFVuaWZpZWRHcm91cFVuc2VlbkRhdGEsIFNldFVuaWZpZWRHcm91cFVuc2VlbkRhdGEsIFVwZGF0ZVVuaWZpZWRHcm91cCwgR2V0QXZhaWxhYmxlQ3VsdHVyZXMsIFVzZXJTb2NpYWxBY3Rpdml0eU5vdGlmaWNhdGlvbiwgR2V0U29jaWFsQWN0aXZpdHlOb3RpZmljYXRpb25zLCBDaGFubmVsRXZlbnQsIEdldFVuaWZpZWRHcm91cHNTZXR0aW5ncywgR2V0UGVvcGxlSW5zaWdodHMsIEdldFVuaWZpZWRHcm91cFVuc2VlbkNvdW50LCBTZXRVbmlmaWVkR3JvdXBMYXN0VmlzaXRlZFRpbWUsIENvbnRhY3RQcm9wZXJ0eVN1Z2dlc3Rpb24sIEdldFVuaWZpZWRHcm91cFNlbmRlclJlc3RyaWN0aW9ucywgU2V0VW5pZmllZEdyb3VwU2VuZGVyUmVzdHJpY3Rpb25zLCBDb252ZXJ0SWNzVG9DYWxlbmRhckl0ZW0sIE1lc3NhZ2VMYXRlbmN5LCBHZXRQZW9wbGVJbnNpZ2h0c1Rva2VucywgU2V0UGVvcGxlSW5zaWdodHNUb2tlbnMsIERlbGV0ZVBlb3BsZUluc2lnaHRzVG9rZW5zLCBFeGVjdXRlU2VhcmNoLCBQcm9jZXNzQ29tcGxpYW5jZU9wZXJhdGlvbkNhbGwsIERlbGVnYXRlQ29tbWFuZEJhc2UsIEdldFN1Z2dlc3RlZFVuaWZpZWRHcm91cHMsIE9EYXRhQ29tbW9uLCBPRGF0YVB1c2hOb3RpZmljYXRpb24KSW5mb1dvcmtlci5SZXF1ZXN0RGlzcGF0Y2g6IFJlcXVlc3RSb3V0aW5nLCBEaXN0cmlidXRpb25MaXN0SGFuZGxpbmcsIFByb3h5V2ViUmVxdWVzdCwgRmF1bHRJbmplY3Rpb24sIEdldEZvbGRlclJlcXVlc3QKSW5mb1dvcmtlci5BdmFpbGFiaWxpdHk6IEluaXRpYWxpemUsIFNlY3VyaXR5LCBDYWxlbmRhclZpZXcsIENvbmZpZ3VyYXRpb24sIFB1YmxpY0ZvbGRlclJlcXVlc3QsIEludHJhU2l0ZUNhbGVuZGFyUmVxdWVzdCwgTWVldGluZ1N1Z2dlc3Rpb25zLCBBdXRvRGlzY292ZXIsIE1haWxib3hDb25uZWN0aW9uQ2FjaGUsIFBGRCwgRG5zUmVhZGVyLCBNZXNzYWdlLCBGYXVsdEluamVjdGlvbgo='))} 
	elseif ($Transport) {$string += [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String('STRINGHERE'))}
	else 
	{
		#Replaced Base63 with Base64+GZip
		$data = [System.Convert]::FromBase64String([regex]::matches($(Read-Host -Prompt 'Please enter ExTRA configuration'),'(?<=@).*(?=\^)').value)
		$ms = New-Object System.IO.MemoryStream
		$ms.Write($data, 0, $data.Length)
		$ms.Seek(0,0) | Out-Null
		$string += $(New-Object System.IO.StreamReader(New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Decompress))).readtoend()
	}	
	$string += "TransportRuleAgent:FaultInjection`nFilteredTracing:No`nInMemoryTracing:No"
	new-item -path "C:\EnabledTraces.Config" -type file -force | Out-Null
	$string | Out-File -filepath "C:\EnabledTraces.Config" -Encoding ASCII -Append | Out-Null
}

function GetExchServers
{
	$return = @()
	# if no server is specified to the script, use the local computer name
	if(!$Servers)
	{
		$Servers = ${env:computername}
        Write-Debug "No Server list specified. Using local Server..."
	}
	foreach($serv in $Servers) {If (Test-Connection -BufferSize 32 -Count 1 -ComputerName $serv -Quiet) {$return += (Get-ExchangeServer $serv)}}
	if($return.Count -eq 0)
	{
		Write-Host "No Exchnage servers found using the specified names"  -foregroundcolor Red $nl
		Exit
	}
	return $return
}

Function ConfirmAnswer
{
	$Confirm = "" 
	while ($Confirm -eq "") 
	{ 
		switch (Read-Host "(Y/N)") 
		{ 
			"yes" {$Confirm = "yes"} 
			"no" {$Confirm = "No"} 
			"y" {$Confirm = "yes"} 
			"n" {$Confirm = "No"} 
			default {Write-Host "Invalid entry, please answer question again " -NoNewline} 
		} 
	} 
	return $Confirm 
}


Function CreateTrace($s)
{

	if($Servers)
	{
        Write-Host "Moving EnabledTraces.Config... " -NoNewline
	    try { Copy-Item "C:\EnabledTraces.Config" ("\\" + $s + "\c$\EnabledTraces.config") -Force }
	    Catch { Write-Host "FAILED."-ForegroundColor red $nl; return $false}
	}
	Write-Host "Creating Trace... " -NoNewline
	$ver = Invoke-Command -ComputerName $s.Name -ScriptBlock {$(Get-Command Exsetup.exe).version.ToString()}
	$ExTRAcmd = "logman create trace ExchangeDebugTraces -p '{79bb49e6-2a2c-46e4-9167-fa122525d540}' -o $filepath$s-$ver-$ts.etl -s $s -ow -f bin -max 1024"
	# Create ExTRA Trace
	Write-Debug $ExTRAcmd
	Invoke-Expression -Command $ExTRAcmd | Out-Null
	while (!($CheckifCreated = @(logman query -s $s) -match "ExchangeDebugTraces"))
	{
		Write-Host " Traced failed to create. Would you like to try creating it again? " -NoNewline
		$answer = ConfirmAnswer
		if ($answer -eq "yes"){Invoke-Expression -Command $ExTRAcmd | Out-Null}
		if ($answer -eq "no"){End}
	}
	return $true
}

Function InitTrace($s)
{
	Write-Host "Starting Trace... " -NoNewline
	$ExTRAcmd = "logman start ExchangeDebugTraces -s $s"
	Write-Debug $ExTRAcmd
	Invoke-Expression -Command $ExTRAcmd | Out-Null
	$CheckExTRA = @(logman query -s $s) -match "ExchangeDebugTraces"
	$CheckifRunning = select-string -InputObject $CheckExTRA -pattern "Running" -quiet
	if ($CheckifRunning)
	{
		Write-Host "COMPLETED" -ForegroundColor green
	}
}

Function StartTrace
{
	if(-not (Is-Admin))
	{
        Write-Warning "The script needs to be executed in elevated mode. Start the Exchange Mangement Shell as an Administrator."
        exit 
	}
	$servlist = GetExchServers
	#Write-Host "Creating Trace... " -NoNewline
	if ($manual -AND [System.IO.File]::Exists("$(Split-Path -parent $PSCommandPath)\EnabledTraces.Config"))
	{
		#Code for running trace with existing EnabledTraces.Config
		Write-Host "Using existing EnabledTraces.Config... " -NoNewline
		try { Move-Item "$(Split-Path -parent $PSCommandPath)\EnabledTraces.Config" "C:\EnabledTraces.Config" -Force}
        Catch { Write-Host "FAILED."-ForegroundColor red $nl; Write-Host "Prompting for configuration... "; CreateExtraTraceConfig; Continue }
		Write-Host "COMPLETED`n" -ForegroundColor Green
	}
	else
	{
	CreateExtraTraceConfig
	}
	$ts = get-date -f HHmmssddMMyy
	$filepath = "c:\tracing\$ts\"
	foreach ($s in $servlist)
	{
		Write-Host "`nEnabling ExTRA tracing on" ($s) -ForegroundColor green
		If (Test-Connection -BufferSize 32 -Count 1 -ComputerName $s -Quiet)
		{
			# Check if ExTRA Trace already exists
			$CheckExTRA = @(logman query -s $s) -match "ExchangeDebugTraces"
			if (!$CheckExTRA)
			{
				if (createtrace($s)){
					inittrace($s)
				}
			}
			else
			{
				Write-Host "ExchangeDebugTraces already exists. Checking if already running"
				$CheckifRunning = select-string -InputObject $CheckExTRA -pattern "Running" -quiet
				if ($CheckifRunning)
				{
					Write-Host "Trace is running. Would you like stop it and recreate it? " -NoNewline
					$answer = ConfirmAnswer
					if ($answer -eq "yes"){
						$cmd = "logman stop ExchangeDebugTraces -s $s"
						$StopExTRA = Invoke-Expression -Command $Cmd
						Start-Sleep 2
					}
					if ($answer -eq "no"){End}
				}
				#Delete and recreate ExTRA tracing
				Write-Host "Deleting and recreating ExchangeDebugTraces"
				$cmd = "logman delete ExchangeDebugTraces -s $s"
				$DeleteExTRA = Invoke-Expression -Command $Cmd 
				createtrace($s)
				inittrace($s)
			}
		}
		Else {Write-Host "Server $s cannot be contacted. Skipping..."  -foregroundcolor Red $nl; Continue}
	}
}

Function StopTrace
{
	if ($LogPath -eq "") {$LogPath = "C:\extra\" + $ts} elseif ($LogPath.EndsWith("\")) {$LogPath = $LogPath + $(get-date -f HHmmssddMMyy) + "\"} else {$LogPath = $LogPath + "\" +  + $(get-date -f HHmmssddMMyy) + "\"}
	# Convert logpath to UNC adminshare path
	$TRACES_FILEPATH = "\\" + (hostname) + "\"+ $LogPath.replace(':','$')
	# create target path if it does not exist yet
	if (-not (Test-Path $TRACES_FILEPATH)) {
		New-Item $TRACES_FILEPATH -ItemType Directory | Out-Null
		#Write-Host "Created $LogPath as it did not exist yet" $nl
	}
	$servlist = GetExchServers
	foreach ($s in $servlist)
	{
		If (Test-Connection -BufferSize 32 -Count 1 -ComputerName $s -Quiet)
		{
			$CheckExmon = @(logman query -s $s) -match "ExchangeDebugTraces"
			$CheckifRunning = select-string -InputObject $CheckExmon -pattern "Running" -quiet
			if ($CheckifRunning)
			{
				Write-Host "Stopping trace on $s... " -NoNewline
				$Error.Clear()
				$cmd = "logman stop -n ExchangeDebugTraces -s $s"
				$StopTrace = Invoke-Expression -Command $cmd
				if ($Error){Write-host "Error encountered" -ForegroundColor Red; Continue}
				else {Write-Host "COMPLETED" -ForegroundColor Green}
				Write-Host "Removing trace on $s... " -NoNewline
				$Error.Clear()
				$cmd = "logman delete ExchangeDebugTraces -s $s"
				$StopTrace = Invoke-Expression -Command $cmd
				if ($Error){Write-host "Error encountered" -ForegroundColor Red}
				else {Write-Host "COMPLETED" -ForegroundColor Green}
			}
			Write-Host "Transfering trace logs from $s... " -NoNewline
			$fileToMovePath = "\\" + $s + "\c$\tracing\*.etl"
			try { Move-Item $fileToMovePath $TRACES_FILEPATH -Force}
			Catch { Write-Host "FAILED "-ForegroundColor red $nl; Continue }
			Write-Host "COMPLETED`n" -ForegroundColor Green
		}
		else {Write-Host "Server $s cannot be contacted. Skipping..."  -foregroundcolor Red $nl; Continue}
	}
	Write-Host "Logs can be found in" $LogPath $nl
}

Function Generate
{
	$comment = $nul
	Write-Host "Input trace lines. Empty line to finish" -ForegroundColor Green $nl
	#prompt for trace definations
	while ($true) {
        $input = Read-Host -Prompt ' '
        if ($input -eq '') {break}
        Else {$string += $input + "`n"}
	}																					
	#Replaced Base63 with Base64+GZip
	# $Encodedstring =[Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($string))
	$ms = New-Object System.IO.MemoryStream
	$cs = New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Compress)
	$sw = New-Object System.IO.StreamWriter($cs)
	$sw.Write($string)
	$sw.Close();
	$Encodedstring = [Convert]::ToBase64String($ms.ToArray())
	
	# Add padding and comment
	Write-Host "Input comment." -ForegroundColor Green $nl	
	$comment = Read-Host -Prompt ' '
	$EncodedText = "#*ExTRACFG-*$comment-@$Encodedstring^end#"
	Write-Host "Config has been copied to clipboard" -ForegroundColor Green $nl
	Write-Host $EncodedText
	Set-CB $EncodedText
}

if ($generate) {Generate; exit;}
if ($start) {StartTrace; [void](Read-Host 'Trace in progress. Press ENTER to stop'); StopTrace; exit;}
if ($stop) {StopTrace; exit;}
