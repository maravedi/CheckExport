# CheckReferralExport.ps1
# Monitors the folder in which referrals are placed for pickup and transmission 
# to AppXtender
# 
# "Normal" cause is commonly the AxImport service gets hung-up and needs to be 
# restarted to get documents flowing to AppXtender again.
#
# David Frazer 10/27/2015
#
# Change History
# - Moved the service restarting aspects to a completely different script since it didn't make sense to have them combined

function Service_status () {
	$state = C:\windows\system32\sc.exe \\server query `
		'AX Import Service' ^| findstr "STATE"
	if ($state -like '*STOPPED*')
	{
		$status = 'STOPPED'
		$color = "-f 'Black' -b 'Red'"
		return $status, $color;
	} elseif ($state -like '*STOP_PENDING*') {
		$status = 'STOP_PENDING'
		$color = "-f 'Black' -b 'Red'"
		return $status, $color;
	} elseif ($state -like '*RUNNING*')	{
		$status = 'RUNNING'
		$color = "-f 'Black' -b 'Green'"
		return $status, $color;
	} elseif ($state -like '*START_PENDING*') {
		$status = 'START_PENDING'
		$color = "-f 'Black' -b 'Red'"
		return $status, $color;
	} else {
		write-host "Unable to check service status."`n;
		return $state;
	}
}


try {
	$now = get-date;
	$exportStuck = $False; # default should be false -- set to true for dev
	$start = $False;
	$dir = '\\server\dir\';

	# Get all the files in the directory. Should be empty except at the time
	# the export is taking place
	$forms = @(get-childitem -path $dir -recurse -include *.pdf,*.tif | 
		sort-object -Property LastWriteTime)
	
	# Iterate through the files and check to see if any of them are older than
	# they should be...
	# meaning there is an unexpected processing delay. All it takes is one to 
	# indicate a problem.
	# 
	# Start a counter to count total number of files in the directories
	$i = 0;
	$stuckFiles += ,('<b><tr><td>','File Path','</td><td>','File Owner',
		'</td><td>','Last Modified','</td></tr></b>');
	foreach ($file in $forms) {
		$fileWriteTime = $file.LastWriteTime;
		$owner = (get-acl $file).Owner
		$file_dir = $file.Directory.Name;
		$timediff = ($now - $fileWriteTime).TotalHours;
		# Check to see if the file is older than 30 minutes
		if ($timediff -ge 0.5) {
			# Increment the counter for each file found
			$i++;
			$exportStuck = $True;
			$stuckFiles += ,('<tr><td>',"$file",'</td><td>',$owner,
			'</td><td>',$fileWriteTime,'</td></tr>');
		}
	}
	$stuckFiles = $stuckFiles | %{"$_"};
	if (!$exportStuck) {
		Write-Host ' + ' -f Black -b Green -NoNewline;
		Write-Host ' No problems detected with process to export HHC forms to AppXtender.' `
		-f Yellow -b Black;
	} else {
		Write-Host ' - ' -f Black -b Red -NoNewline;
		write-host ' A problem was detected with process to export HHC forms to AppXtender.' `
		-f Yellow -b Black;
		if ($i -gt 1 -or $i -eq 0) {
			$noun = 'files';
			$verb = 'are';
		} else {
			$noun = 'file';
			$verb = 'is';
		}
		write-host 'Found' $i $noun 'that' $verb 'older than 30 minutes.';
		#
		# If old files are found send an email alert and resolve the problem by 
		# restarting the Inscrybe Connector service
		& 'C:\Users\maravedi\RestartAXImport.ps1' -count $i -noun $noun -verb $verb -stuckFiles $StuckFiles 
	}
}
 
catch {
    Write-Error $_.Exception.ToString()
    Read-Host -Prompt 'The above error occurred. Press Enter to exit.';
}

$status = Service_status;
# Even though the files may be fine at this point, checking to see if the 
# service is stuck and needs to be restarted
Invoke-Expression("Write-host 'Serivce status is: ' -f 'Yellow' -b 'Black' -NoNewline; Write-Host $status$color;")
