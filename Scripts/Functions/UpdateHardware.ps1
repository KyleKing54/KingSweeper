<# 

    Update Hardware - Function to scan Hardware on a computer
    Published 8/28/2020
    Created By Kyle King

#>

function ScanHardware($computer, $csv) {
    Try {
        # collect hardware infomation
        $serialnumber = Get-WmiObject win32_bios -ComputerName $computer -ErrorAction Stop
        $processorname = Get-WmiObject -Class Win32_processor -ComputerName $computer -ErrorAction Stop
        $adProperties = get-adcomputer $computer -properties Name, Lastlogondate, operatingsystem, ipv4Address, enabled, DistinguishedName -ErrorAction Stop
        $ram = "{0} GB" -f ((Get-WmiObject Win32_PhysicalMemory -ComputerName $computer | Measure-Object Capacity  -Sum).Sum / 1GB) 
        $model = Get-WmiObject win32_computersystem -ComputerName $computer | select @{Name = "Type"; Expression = { if (($_.pcsystemtype -eq '2')  ) { 'Laptop' } Else { 'Desktop/VM' } } },
        Manufacturer, @{Name = "Model"; Expression = { if (($_.model -eq "$null")  ) { 'Virtual' } Else { $_.model } } }, username -ErrorAction Stop
        $time = Get-Date -Format g
        $FullName = Get-ADUser ($model.username).split('\')[1]
        $blComputer = Get-ADComputer $computer
        $BitLockerPassword = Get-ADObject -Filter "objectClass -eq 'msFVE-RecoveryInformation'" -SearchBase $blComputer.distinguishedName -Properties msFVE-RecoveryPassword, whenCreated | Sort whenCreated -Descending | Select -First 1 | Select-Object -ExpandProperty msFVE-RecoveryPassword    
        $OSVersion = Get-WmiObject -Class win32_OperatingSystem -ComputerName $computer | select BuildNumber -ErrorAction Stop
        switch ($OSVersion.BuildNumber) {
            10240 { $OSVersion.BuildNumber = "1507" }
            10586 { $OSVersion.BuildNumber = "1511" }
            14393 { $OSVersion.BuildNumber = "1607" }
            15063 { $OSVersion.BuildNumber = "1703" }
            16299 { $OSVersion.BuildNumber = "1709" }
            17134 { $OSVersion.BuildNumber = "1803" }
            17763 { $OSVersion.BuildNumber = "1809" }
            18362 { $OSVersion.BuildNumber = "1903" }
            18363 { $OSVersion.BuildNumber = "1909" }
            19041 { $OSVersion.BuildNumber = "2004" }
            19042 { $OSVersion.BuildNumber = "20H2" }
            19043 { $OSVersion.BuildNumber = "21H1" }
            19044 { $OSVersion.BuildNumber = "21H2" }
            default { $OSVersion.BuildNumber = "" }
        }

        # Save scanned data to object
        $computer_object = New-Object PSObject -Property @{
            ComputerName         = $adProperties.name
            Enabled              = $adProperties.Enabled
            IpAddress            = $adProperties.ipv4Address
            OU                   = $adProperties.DistinguishedName.split(',')[1].split('=')[1]
            Type                 = $model.type
            SerialNumber         = $serialnumber.serialnumber
            Manufacturer         = $model.Manufacturer
            Model                = $model.Model
            RAM                  = $ram
            ProcessorName        = ($processorname.name | Out-String).Trim()
            OperatingSystem      = $adProperties.operatingsystem + " " + $OSVersion.BuildNumber
            LastLogon            = $model.username
            FullName             = $fullName.name
            LastLogonDate        = $adProperties.lastlogondate
            LastScanned          = $time
            BitLockerRecoveryKey = $BitLockerPassword
        }
    }
    # If error occured during scan, log error and quit
    Catch {
        write-host "Error pulling Hardware data from: $computer" -f yellow
        return
    }
    # Loop through CSV and see if computer already exists
    for ($i = 0; $i -lt $csv.count; $i++) {
        # If it does, update the entry with newly scanned info
        if ($csv[$i].computername -eq $computer) {
            
            $Csv[$i].ComputerName = $computer_object.ComputerName
            $Csv[$i].OperatingSystem = $computer_object.OperatingSystem
            $Csv[$i].LastLogonDate = $computer_object.LastLogonDate
            if ($computer_object.LastLogon -ne $null) { $Csv[$i].LastLogon = $computer_object.LastLogon }
            $Csv[$i].IpAddress = $computer_object.IpAddress
            $Csv[$i].Enabled = $computer_object.Enabled
            $Csv[$i].OU = $computer_object.OU
            $Csv[$i].LastScanned = $computer_object.LastScanned
            if ($computer_object.BitLockerRecoveryKey -ne $null) { $Csv[$i].BitLockerRecoveryKey = $computer_object.BitLockerRecoveryKey }
            else { $Csv[$i].BitLockerRecoveryKey = "Bitlocker not Enabled" }
            $Csv[$i].FullName = $computer_object.FullName
            
            $Csv | export-csv -path $path -nti
            break
            
        }
        # If it is not, add system to the end of the CSV
        elseif ($i -eq ($csv.count - 1) -or $csv.count -eq 0) {
            $computer_object | select ComputerName, Enabled, IpAddress, OU, Type, SerialNumber, Manufacturer, Model, RAM, ProcessorName, OperatingSystem, LastLogon, FullName, LastLogonDate, LastScanned, BitLockerRecoveryKey | export-csv -Append -path $path -nti
        }
    }
    write-host "Successfully pulled Hardware data from: $computer" -f Green
}