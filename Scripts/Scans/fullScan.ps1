<# 

    Full Scan - Scans AD for all computers and calls hardware and software scanners for each computer that is online.
    Version 2.0 - Updated 12/9/2021
    Published 8/28/2020
    Created By Kyle King

#>

import-module activedirectory

# Set home location and import required function and CSV files
cd ../../
$homeDir = get-location

. "$homeDir\Scripts\Functions\UpdateSoftware.ps1"
. "$homeDir\Scripts\Functions\UpdateHardware.ps1"
. "$homeDir\Scripts\Functions\UploadFiles.ps1"

$csvPath = "$homeDir\CSV\Hardware\HCPF_PCs.csv"
$serverCSV = "$homeDir\CSV\ServerFiles\Hardware.csv"

# Import all computers in AD that have been logged into in the 30 days, running non server Windows OS
$logonTimeLimit = $(((Get-Date).AddDays(-30)).ToFileTime())
$computers = Get-ADComputer -Filter {lastLogonTimeStamp -gt $logonTimeLimit -and OperatingSystem -like 'Windows*' -and OperatingSystem -notlike "*server*"} | Select-Object -ExpandProperty Name

# Attempt to import existing CSV file, if it does not exist create a new one
Try {
    $csvFile = Import-Csv -path $csvPath -ErrorAction Stop
}
Catch [Exception] {
    Write-Host "CSV file doesn't currently exist; Creating new file"
    Add-Content -path $csvPath -Value '"ComputerName", "Enabled", "IpAddress", "OU", "Type", "SerialNumber", "Manufacturer", "Model", "RAM", "ProcessorName", "OperatingSystem", "LastLogon", "FullName", "LastLogonDate", "LastScanned", "BitLockerRecoveryKey"'
    $csvFile = Import-Csv -path $csvPath
}

# Set count for progress bar
$count = 1

# Loop through all computers imported from AD, if it is online scan hardware then software
Foreach ($computer in $computers) {
    Write-Progress -Id 0 -Activity "Updating Database..." -Status "$count of $($computers.Count)" -PercentComplete (($count / $computers.Count) * 100)
    $count++

    if (Test-Connection -Cn $computer -Count 1 -quiet) {
        ScanHardware $computer $csvFile
        UpdateSoftware $computer
    }
    
    else {
        Write-Host "$computer is offline" -f white
    }
}

# Start file upload
Upload