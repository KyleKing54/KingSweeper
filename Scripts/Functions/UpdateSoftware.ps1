<# 

    Update Software - Function to scan Software on a computer
    Published 8/28/2020
    Created By Kyle King

#>

function UpdateSoftware($pc)
{
	# Set home location
	$homeDir = get-location
	$path = "$homeDir\CSV\Software\${pc}.csv"

	# array that stores list of scanned software
	$array = @()

	# used for call center software installed into chrome and not saved in the same registry
	$ver10 = $false
	$ver11 = $false
	$hangouts = $false

	# Open registry on system
    $UninstallKey = "SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
    try
    {
		$reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$pc)
		$regkey = $reg.OpenSubKey($UninstallKey)
		$chrome = $reg.OpenSubKey('SOFTWARE\\WOW6432Node\\Policies\\Google\\Chrome\\ExtensionInstallForcelist')
    }
    catch
    {
        write-host "Error pulling Software data from: $pc"
        continue
    }

	# loop through registry keys and add each to the array
    $subkeys = $regkey.GetSubKeyNames() 
    foreach($key in $subkeys)
    {
		$displayName = $null
		$displayVersion = $null
		$installLocation = $null
		$publisher = $null
		$time = $null

		$thisKey = $UninstallKey+"\\"+$key 
        $thisSubKey = $reg.OpenSubKey($thisKey)

		$displayName = $thisSubKey.GetValue('DisplayName')
		$displayVersion = $thisSubKey.GetValue('DisplayVersion')
		$installLocation = $thisSubKey.GetValue('InstallLocation')
		$publisher = $thisSubKey.GetValue('Publisher')
        $time = $thisSubKey.GetValue('InstallDate')
			
        if($time -ne $null)
        {
			try {$installTime = [datetime]::ParseExact($time, 'yyyyMMdd', $null)}
			catch {$installTime = $time}			
		}

        $obj = New-Object PSObject
        $obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $pc
        if($displayName -ne $null)	   {$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $displayName}
        if($displayVersion -ne $null)  {$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $displayVersion}
        if($installLocation -ne $null) {$obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $installLocation}
        if($publisher -ne $null)       {$obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $publisher}
        if($time -ne $null)            {$obj | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value $installTime}

        $array += $obj         
	}

	# loop through the chrome regisrty and search for call center related software
	for($j = 0; $j -lt 53; $j++)
	{
		try{$keyValue = $chrome.GetValue($j)}
		catch{break}			

		if($keyValue -eq 'offljimabadkdnhijlbopjlkaolpdgip;https://clients2.google.com/service/update2/crx')
		{$ver10 = $true}
		if($keyValue -eq 'ngfbofmmhadhjnlodckfafckdhmlcpeh;https://clients2.google.com/service/update2/crx')
		{$ver11 = $true}
		if($keyValue -eq 'nckgahadagoaajjgafhacjanaoiihapd;https://clients2.google.com/service/update2/crx')
		{$hangouts = $true}
	}

	# if found, add to the software array
	if($ver10 -eq $true)
	{
		$obj = New-Object PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $pc
		$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "Five9 Software Service (Chrome Extension)"
		$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value "10"
		$obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value "Five9"
		$array += $obj
	}
	if($ver11 -eq $true)
	{
		$obj = New-Object PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $pc
		$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "Five9 Software Service (Chrome Extension)"
		$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value "11"
		$obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value "Five9"
		$array += $obj
	}
	if($hangouts -eq $true)
	{
		$obj = New-Object PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $pc
		$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "Google Hangouts (Chrome Extension)"
		$obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value "Google"
		$array += $obj
	}
	
	# save array into CSV file
    $array | Where-Object { $_.DisplayName } | select ComputerName, DisplayName, DisplayVersion, Publisher, InstallDate | sort ComputerName, DisplayName | export-csv -path $path -nti
    write-host "Successfully pulled Software data from: $computer" -f green
}