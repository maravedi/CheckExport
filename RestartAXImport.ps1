# RestartAXImport.ps1
# Monitors the folder in which referrals are placed for pickup and transmission to AppXtender
# "Normal" cause is commonly the AxImport service gets hung-up and needs to be restarted
# to get documents flowing to AppXtender again.
#
# David Frazer 10/27/2015
#
# Change History
# - 01-15-2016: Re-designed this script to be object oriented.  Though it wasn't necessary, I think it makes it easier to understand.
# - 04-06-2016: Renamed the script to be more descriptive.  Also allowed parameters to be passed to this script so a single email is sent to the admin with all pertinent details.

Param([int]$count, [string]$noun, [string]$verb, [string[]]$stuckFiles)

$global:output = '';

# Class definition
$serviceclass = new-object psobject -Property @{
	status = $null
}

# Class constructor
function ServiceClass {
    param(
          [Parameter(Mandatory=$true)]
          [String]$status
    )
    $service = $ServiceClass.psobject.copy()
    $service.status = $status
    $service
}

# Class method - Stop
$ServiceClass | Add-Member -MemberType ScriptMethod -Name "Stop" -Value {
	$global:output += '(' + (Get-Date).ToString() + '):   AX Import Service: STOPPING<br>';
	write-host 'AX Import Service is: STOPPING' -foreground 'red' `n
	C:\windows\system32\sc.exe \\server stop 'AX Import Service' | out-null
}

# Class method - Start
$ServiceClass | Add-Member -MemberType ScriptMethod -Name "Start" -Value {
	$global:output += '(' + (Get-Date).ToString() + '):   AX Import Service: STARTING<br>';
	write-host 'AX Import Service is: STARTING' -foreground 'green' `n
	C:\windows\system32\sc.exe \\server start 'AX Import Service' | out-null
}

# Class method - Kill
$ServiceClass | Add-Member -MemberType ScriptMethod -Name "Kill" -Value {
	if ($process_id = Get-Process aximport -computername server | select -expand id) {
		$global:output += '(' + (Get-Date).ToString() + '):   Process ID is: ' + $process_id +  '<br>';
		$global:output += '(' + (Get-Date).ToString() + '):   Kill process!<br>';
		write-host 'Process ID is: ' $process_id;
		Write-Host 'Kill process!';
		Invoke-Command -computername server {Stop-Process $process_id -Force}
		}
		else {
			$global:output += '(' + (Get-Date).ToString() + '):   Process not found.<br>';
			Write-Host 'Process not found'`n`n;
		}
}

# Class method - State
$ServiceClass | Add-Member -MemberType ScriptMethod -Name "State" -Value {
	$state = C:\windows\system32\sc.exe \\server query `
			'AX Import Service' ^| findstr "STATE"
		if ($state -like '*STOPPED*')
		{
			$status = 'STOPPED';
			return $status;
		} elseif ($state -like '*STOP_PENDING*') {
			$status = 'STOP_PENDING';
			return $status;
		} elseif ($state -like '*RUNNING*')	{
			$status = 'RUNNING';
			return $status;
		} elseif ($state -like '*START_PENDING*') {
			$status = 'START_PENDING';
			return $status;
		} else {
			$global:output += 'Unable to check service status.<br>';
			write-host "Unable to check service status."`n;
			return $state;
		}
}

# Class method - Restart
$ServiceClass | Add-Member -MemberType ScriptMethod -Name "Restart" -Value {
	$global:output += '(' + (Get-Date).ToString() + '):   Service status is: ' + $service.State() + '<br>';
	Write-host 'Service status is: ' $service.State()

	if ($service.State() -eq 'STOPPED') {
		$global:output += '(' + (Get-Date).ToString() + '):   Not stopping the service because it is already stopped.<br>';
		write-host 'Not stopping the service because it is already stopped.';
		$start = $True;
	} else {
		$service.Stop()
		Start-Sleep -s 10
		if ($service.State() -ne 'STOPPED') {
			$global:output += '(' + (Get-Date).ToString() + '):   Failed to stop AX Import Service. Going to see if process needs to be killed.<br>';
			Write-Host 'Failed to stop AX Import Service. Going to see if process needs to be killed.';
			# Wait 60 seconds to see if the service will clear-up on its own
			Start-Sleep -s 60
			$global:output += '(' + (Get-Date).ToString() + '):   Waiting 60 seconds to make sure the service fully stops.<br>';
			write-host 'Waiting 60 seconds to make sure the service fully stops.';
			if ($service.State() -ne 'STOPPED') {
				$service.Kill()
				$service.Start()
			}
		}
		$start = $True;
	}

	if ($start -eq $True) {
		$global:output += '(' + (Get-Date).ToString() + '):   Status before start is: ' + $service.State() + '<br>';
		write-host 'Status before start is: ' $service.State()`n`n;
		if ($status -eq 'RUNNING') {
			$global:output += '(' + (Get-Date).ToString() + '):   Not starting the process because it is already started.<br>';
			write-host 'Not starting the process because it is already started.';
		} else {
			$service.Start()
			Start-Sleep -s 10
		}
		if ($service.State() -ne 'RUNNING') {
			$global:output += '(' + (Get-Date).ToString() + '):   Failed to start AX Import Service. Going to see if process needs to be killed.<br>';
			Write-Host 'Failed to start AX Import Service. Going to see if process needs to be killed.';
			# Wait 60 seconds to see if the service will clear-up on its own
			Start-Sleep -s 60
			$global:output += '(' + (Get-Date).ToString() + '):   Waiting 60 seconds to make sure the service fully starts.<br>';
			write-host 'Waiting 60 seconds to make sure the service fully starts.';
			if ($service.State() -ne 'RUNNING') {
				$service.Kill()
				$service.Start()
			}
		}
	}

	$global:output += '(' + (Get-Date).ToString() + '):   Most recent status is: ' + $service.State() + '<br>';
	write-host 'Most recent status is: ' $service.State()`n`n;
}

# Assigning a value to status
$service = ServiceClass -status 0

try {
	$service.Restart()
	style = "<style>table, td{border-width:1px;border-style:solid;}a{text-decoration:none;color:#000000;}</style>"
	& 'C:\Program Files (x86)\BlatEmailClient\blat.exe' -server 1.1.1.1 -f admin@admin.org -html -q -to anyone@anyone.com -subject 'FYI: AX Import Service Restarted' -body "Found $count $noun that $verb older than 30 minutes.  If files are showing from multiple users, then the service probably needs to be restarted.<br/><br/>Otherwise, it could be that a user has the files open and they cannot be processed because they are locked.<br/><br/>$style<table>$stuckFiles</table></style><br><br>AX Import Service was restarted.  See below for details.<br><br>$global:output<br><br>";

} catch {
    Write-Error $_.Exception.ToString()
    Read-Host -Prompt 'The above error occurred. Press Enter to exit.';
}
