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
.PARAMETER Server
	This optional parameter allows multiple target Exchange servers to be specified. If it is not the 		
	local server is assumed.
.PARAMETER Start
	Starts log trace after prompting for configuration data
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
 [string]$Server,
 [string]$logpath, 
 [switch]$Start,
 [switch]$Stop,
 [switch]$Generate
)

# network path to save the resulting traces (default is local c:\temp\extra)
$script:nl = "`r`n"
$servers = $null
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
	$string = Read-Host -Prompt 'Please enter ExTRA configuration'
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
    $string = "TraceLevels:Debug,Warning,Error,Fatal,Info,Pfd`nSystemLogging:SystemNet`n"
    #prompt for trace definations
    while ($true) {
        $input = Read-Host -Prompt ' '
        if ($input -eq '') {break}
        Else {$string += $input + '`n'}
    }
    $string += "TransportRuleAgent:FaultInjection`nFilteredTracing:No`nInMemoryTracing:No`n"
    $EncodedText =[Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($string))
    Write-Host "Send the following line to customer for ExTRA Trace Configuration.`nConfig has been copied to clipboard" -ForegroundColor Green $nl
	Start-Sleep 1
	Set-CB $EncodedText
    Write-Host $EncodedText + $nl

}
		
if ($generate) {Generate; exit;}
if ($start) {StartTrace; exit;}
if ($stop) {StopTrace; exit;}
