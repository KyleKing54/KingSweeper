<# 

    KINGSweeper - Replacement for LANSweeper for HCPF
    Published 8/2/19 
    Created By Kyle King

    UPDATES:
    4/23/20 - Added bitlocker recovery key text box
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '880,650'
$Form.text                       = "KINGSweeper"
$Form.BackColor                  = "#252525"
$Form.TopMost                    = $false

$SearchCombo                     = New-Object system.Windows.Forms.ComboBox
$SearchCombo.text                = "Search..."
$SearchCombo.BackColor           = "#5e5e5e"
$SearchCombo.width               = 345
$SearchCombo.height              = 20
$SearchCombo.location            = New-Object System.Drawing.Point(10,10)
$SearchCombo.Font                = 'Microsoft Sans Serif,10'
$SearchCombo.ForeColor           = "#ffffff"

$SearchButton                    = New-Object system.Windows.Forms.Button
$SearchButton.BackColor          = "#5e5e5e"
$SearchButton.text               = "Search"
$SearchButton.width              = 65
$SearchButton.height             = 25
$SearchButton.location           = New-Object System.Drawing.Point(364,10)
$SearchButton.Font               = 'Microsoft Sans Serif,10'
$SearchButton.ForeColor          = "#ffffff"

$OnlineLabel                     = New-Object system.Windows.Forms.Label
$OnlineLabel.text                = "Online:"
$OnlineLabel.AutoSize            = $true
$OnlineLabel.width               = 25
$OnlineLabel.height              = 10
$OnlineLabel.location            = New-Object System.Drawing.Point(805,14)
$OnlineLabel.Font                = 'Microsoft Sans Serif,10'
$OnlineLabel.ForeColor           = "#ffffff"

$OnlineBox                       = New-Object system.Windows.Forms.Label
$OnlineBox.text                  = " "
$OnlineBox.BackColor             = "#d0021b"
$OnlineBox.AutoSize              = $false
$OnlineBox.width                 = 15
$OnlineBox.height                = 12
$OnlineBox.location              = New-Object System.Drawing.Point(851,17)
$OnlineBox.Font                  = 'Microsoft Sans Serif,10'

$UpdateButton                    = New-Object system.Windows.Forms.Button
$UpdateButton.BackColor          = "#5e5e5e"
$UpdateButton.text               = "Update"
$UpdateButton.width              = 65
$UpdateButton.height             = 25
$UpdateButton.location           = New-Object System.Drawing.Point(730,10)
$UpdateButton.Font               = 'Microsoft Sans Serif,10'
$UpdateButton.ForeColor          = "#ffffff"

$AddButton                       = New-Object system.Windows.Forms.Button
$AddButton.BackColor             = "#5e5e5e"
$AddButton.text                  = "Add"
$AddButton.width                 = 65
$AddButton.height                = 25
$AddButton.location              = New-Object System.Drawing.Point(660,10)
$AddButton.Font                  = 'Microsoft Sans Serif,10'
$AddButton.ForeColor             = "#ffffff"


$SoftwareButton                  = New-Object system.Windows.Forms.Button
$SoftwareButton.BackColor        = "#000000"
$SoftwareButton.text             = "Software"
$SoftwareButton.width            = 75
$SoftwareButton.height           = 25
$SoftwareButton.location         = New-Object System.Drawing.Point(768,552)
$SoftwareButton.Font             = 'Microsoft Sans Serif,10'
$SoftwareButton.ForeColor        = "#ffffff"

$NotFound                        = New-Object system.Windows.Forms.Label
$NotFound.text                   = "Computer Not Found"
$NotFound.AutoSize               = $true
$NotFound.width                  = 25
$NotFound.height                 = 10
$NotFound.location               = New-Object System.Drawing.Point(290,301)
$NotFound.Font                   = 'Microsoft Sans Serif,24'
$NotFound.ForeColor              = "#ffffff"

$King                            = New-Object system.Windows.Forms.Label
$King.text                       = "KING"
$King.AutoSize                   = $true
$King.width                      = 25
$King.height                     = 10
$King.location                   = New-Object System.Drawing.Point(444,10)
$King.Font                       = 'Microsoft Sans Serif,18,style=Bold'
$King.ForeColor                  = "#606060"

$Sweeper                         = New-Object system.Windows.Forms.Label
$Sweeper.text                    = "Sweeper"
$Sweeper.AutoSize                = $true
$Sweeper.width                   = 25
$Sweeper.height                  = 10
$Sweeper.location                = New-Object System.Drawing.Point(510,10)
$Sweeper.Font                    = 'Microsoft Sans Serif,18,style=Bold'
$Sweeper.ForeColor               = "#e08b00"

$ComputerName                    = New-Object system.Windows.Forms.TextBox
$ComputerName.multiline          = $false
$ComputerName.BackColor          = "#000000"
$ComputerName.width              = 200
$ComputerName.height             = 20
$ComputerName.location           = New-Object System.Drawing.Point(200,100)
$ComputerName.Font               = 'Microsoft Sans Serif,10'
$ComputerName.ForeColor          = "#ffffff"

$LastLogon                       = New-Object system.Windows.Forms.TextBox
$LastLogon.multiline             = $false
$LastLogon.BackColor             = "#000000"
$LastLogon.width                 = 200
$LastLogon.height                = 20
$LastLogon.location              = New-Object System.Drawing.Point(590,150)
$LastLogon.Font                  = 'Microsoft Sans Serif,10'
$LastLogon.ForeColor             = "#ffffff"

$Model                           = New-Object system.Windows.Forms.TextBox
$Model.multiline                 = $false
$Model.BackColor                 = "#000000"
$Model.width                     = 200
$Model.height                    = 20
$Model.location                  = New-Object System.Drawing.Point(200,200)
$Model.Font                      = 'Microsoft Sans Serif,10'
$Model.ForeColor                 = "#ffffff"

$OperatingSystem                 = New-Object system.Windows.Forms.TextBox
$OperatingSystem.multiline       = $false
$OperatingSystem.BackColor       = "#000000"
$OperatingSystem.width           = 200
$OperatingSystem.height          = 20
$OperatingSystem.location        = New-Object System.Drawing.Point(200,150)
$OperatingSystem.Font            = 'Microsoft Sans Serif,10'
$OperatingSystem.ForeColor       = "#ffffff"

$IPAddress                       = New-Object system.Windows.Forms.TextBox
$IPAddress.multiline             = $false
$IPAddress.BackColor             = "#000000"
$IPAddress.width                 = 200
$IPAddress.height                = 20
$IPAddress.location              = New-Object System.Drawing.Point(590,200)
$IPAddress.Font                  = 'Microsoft Sans Serif,10'
$IPAddress.ForeColor             = "#ffffff"

$UserBox                         = New-Object system.Windows.Forms.TextBox
$UserBox.multiline               = $false
$UserBox.BackColor               = "#000000"
$UserBox.width                   = 200
$UserBox.height                  = 20
$UserBox.location                = New-Object System.Drawing.Point(590,100)
$UserBox.Font                    = 'Microsoft Sans Serif,10'
$UserBox.ForeColor               = "#ffffff"

$ComputerNameLabel               = New-Object system.Windows.Forms.Label
$ComputerNameLabel.text          = "Computer Name: "
$ComputerNameLabel.AutoSize      = $true
$ComputerNameLabel.width         = 25
$ComputerNameLabel.height        = 10
$ComputerNameLabel.location      = New-Object System.Drawing.Point(65,100)
$ComputerNameLabel.Font          = 'Microsoft Sans Serif,12'
$ComputerNameLabel.ForeColor     = "#ffffff"

$ModelLabel                      = New-Object system.Windows.Forms.Label
$ModelLabel.text                 = "Model:"
$ModelLabel.AutoSize             = $true
$ModelLabel.width                = 25
$ModelLabel.height               = 10
$ModelLabel.location             = New-Object System.Drawing.Point(138,200)
$ModelLabel.Font                 = 'Microsoft Sans Serif,12'
$ModelLabel.ForeColor            = "#ffffff"

$TypeLabel                       = New-Object system.Windows.Forms.Label
$TypeLabel.text                  = "Operating System:"
$TypeLabel.AutoSize              = $true
$TypeLabel.width                 = 25
$TypeLabel.height                = 10
$TypeLabel.location              = New-Object System.Drawing.Point(55,150)
$TypeLabel.Font                  = 'Microsoft Sans Serif,12'
$TypeLabel.ForeColor             = "#ffffff"

$IPAddressLabel                  = New-Object system.Windows.Forms.Label
$IPAddressLabel.text             = "IP Address:"
$IPAddressLabel.AutoSize         = $true
$IPAddressLabel.width            = 25
$IPAddressLabel.height           = 10
$IPAddressLabel.location         = New-Object System.Drawing.Point(490,200)
$IPAddressLabel.Font             = 'Microsoft Sans Serif,12'
$IPAddressLabel.ForeColor        = "#ffffff"

$LastLogonLabel                  = New-Object system.Windows.Forms.Label
$LastLogonLabel.text             = "Last Logon:"
$LastLogonLabel.AutoSize         = $true
$LastLogonLabel.width            = 25
$LastLogonLabel.height           = 10
$LastLogonLabel.location         = New-Object System.Drawing.Point(491,150)
$LastLogonLabel.Font             = 'Microsoft Sans Serif,12'
$LastLogonLabel.ForeColor        = "#ffffff"

$UserLabel                       = New-Object system.Windows.Forms.Label
$UserLabel.text                  = "User:"
$UserLabel.AutoSize              = $true
$UserLabel.width                 = 25
$UserLabel.height                = 10
$UserLabel.location              = New-Object System.Drawing.Point(535,100)
$UserLabel.Font                  = 'Microsoft Sans Serif,12'
$UserLabel.ForeColor             = "#ffffff"

$ManufacturerLabel               = New-Object system.Windows.Forms.Label
$ManufacturerLabel.text          = "Manufacturer:"
$ManufacturerLabel.AutoSize      = $true
$ManufacturerLabel.width         = 25
$ManufacturerLabel.height        = 10
$ManufacturerLabel.location      = New-Object System.Drawing.Point(89,250)
$ManufacturerLabel.Font          = 'Microsoft Sans Serif,12'
$ManufacturerLabel.ForeColor     = "#ffffff"

$ProcessorLabel                  = New-Object system.Windows.Forms.Label
$ProcessorLabel.text             = "Processor:"
$ProcessorLabel.AutoSize         = $true
$ProcessorLabel.width            = 25
$ProcessorLabel.height           = 10
$ProcessorLabel.location         = New-Object System.Drawing.Point(110,300)
$ProcessorLabel.Font             = 'Microsoft Sans Serif,12'
$ProcessorLabel.ForeColor        = "#ffffff"

$RAMLabel                        = New-Object system.Windows.Forms.Label
$RAMLabel.text                   = "RAM:"
$RAMLabel.AutoSize               = $true
$RAMLabel.width                  = 25
$RAMLabel.height                 = 10
$RAMLabel.location               = New-Object System.Drawing.Point(144,350)
$RAMLabel.Font                   = 'Microsoft Sans Serif,12'
$RAMLabel.ForeColor              = "#ffffff"

$OULabel                         = New-Object system.Windows.Forms.Label
$OULabel.text                    = "OU:"
$OULabel.AutoSize                = $true
$OULabel.width                   = 25
$OULabel.height                  = 10
$OULabel.location                = New-Object System.Drawing.Point(542,250)
$OULabel.Font                    = 'Microsoft Sans Serif,12'
$OULabel.ForeColor               = "#ffffff"

$EnabledLabel                    = New-Object system.Windows.Forms.Label
$EnabledLabel.text               = "Enabled:"
$EnabledLabel.AutoSize           = $true
$EnabledLabel.width              = 25
$EnabledLabel.height             = 10
$EnabledLabel.location           = New-Object System.Drawing.Point(507,300)
$EnabledLabel.Font               = 'Microsoft Sans Serif,12'
$EnabledLabel.ForeColor          = "#ffffff"

$LastLogonDateLabel              = New-Object system.Windows.Forms.Label
$LastLogonDateLabel.text         = "Last Logon Date:"
$LastLogonDateLabel.AutoSize     = $true
$LastLogonDateLabel.width        = 25
$LastLogonDateLabel.height       = 10
$LastLogonDateLabel.location     = New-Object System.Drawing.Point(454,350)
$LastLogonDateLabel.Font         = 'Microsoft Sans Serif,12'
$LastLogonDateLabel.ForeColor    = "#ffffff"

$OU                              = New-Object system.Windows.Forms.TextBox
$OU.multiline                    = $false
$OU.BackColor                    = "#000000"
$OU.width                        = 200
$OU.height                       = 20
$OU.location                     = New-Object System.Drawing.Point(590,250)
$OU.Font                         = 'Microsoft Sans Serif,10'
$OU.ForeColor                    = "#ffffff"

$Enabled                         = New-Object system.Windows.Forms.TextBox
$Enabled.multiline               = $false
$Enabled.BackColor               = "#000000"
$Enabled.width                   = 200
$Enabled.height                  = 20
$Enabled.location                = New-Object System.Drawing.Point(590,300)
$Enabled.Font                    = 'Microsoft Sans Serif,10'
$Enabled.ForeColor               = "#ffffff"

$LastLogonDate                   = New-Object system.Windows.Forms.TextBox
$LastLogonDate.multiline         = $false
$LastLogonDate.BackColor         = "#000000"
$LastLogonDate.width             = 200
$LastLogonDate.height            = 20
$LastLogonDate.location          = New-Object System.Drawing.Point(590,350)
$LastLogonDate.Font              = 'Microsoft Sans Serif,10'
$LastLogonDate.ForeColor         = "#ffffff"

$SerialNumberLabel               = New-Object system.Windows.Forms.Label
$SerialNumberLabel.text          = "Serial Number:"
$SerialNumberLabel.AutoSize      = $true
$SerialNumberLabel.width         = 25
$SerialNumberLabel.height        = 10
$SerialNumberLabel.location      = New-Object System.Drawing.Point(80,400)
$SerialNumberLabel.Font          = 'Microsoft Sans Serif,12'
$SerialNumberLabel.ForeColor     = "#ffffff"

$bitLockerLabel                  = New-Object system.Windows.Forms.Label
$bitLockerLabel.text             = "BitLocker Recovery Key:"
$bitLockerLabel.AutoSize         = $true
$bitLockerLabel.width            = 25
$bitLockerLabel.height           = 10
$bitLockerLabel.location         = New-Object System.Drawing.Point(12,450)
$bitLockerLabel.Font             = 'Microsoft Sans Serif,12'
$bitLockerLabel.ForeColor        = "#ffffff"

$LastScannedLabel                = New-Object system.Windows.Forms.Label
$LastScannedLabel.text           = "Last Scanned: "
$LastScannedLabel.AutoSize       = $true
$LastScannedLabel.width          = 25
$LastScannedLabel.height         = 10
$LastScannedLabel.location       = New-Object System.Drawing.Point(469,400)
$LastScannedLabel.Font           = 'Microsoft Sans Serif,12'
$LastScannedLabel.ForeColor      = "#ffffff"

$LastScanned                     = New-Object system.Windows.Forms.TextBox
$LastScanned.multiline           = $false
$LastScanned.BackColor           = "#000000"
$LastScanned.width               = 200
$LastScanned.height              = 20
$LastScanned.location            = New-Object System.Drawing.Point(590,400)
$LastScanned.Font                = 'Microsoft Sans Serif,10'
$LastScanned.ForeColor           = "#ffffff"

$Manufacturer                    = New-Object system.Windows.Forms.TextBox
$Manufacturer.multiline          = $false
$Manufacturer.BackColor          = "#000000"
$Manufacturer.width              = 200
$Manufacturer.height             = 20
$Manufacturer.location           = New-Object System.Drawing.Point(200,250)
$Manufacturer.Font               = 'Microsoft Sans Serif,10'
$Manufacturer.ForeColor          = "#ffffff"

$Processor                       = New-Object system.Windows.Forms.TextBox
$Processor.multiline             = $false
$Processor.BackColor             = "#000000"
$Processor.width                 = 200
$Processor.height                = 20
$Processor.location              = New-Object System.Drawing.Point(200,300)
$Processor.Font                  = 'Microsoft Sans Serif,10'
$Processor.ForeColor             = "#ffffff"

$RAM                             = New-Object system.Windows.Forms.TextBox
$RAM.multiline                   = $false
$RAM.BackColor                   = "#000000"
$RAM.width                       = 200
$RAM.height                      = 20
$RAM.location                    = New-Object System.Drawing.Point(200,350)
$RAM.Font                        = 'Microsoft Sans Serif,10'
$RAM.ForeColor                   = "#ffffff"

$SerialNumber                    = New-Object system.Windows.Forms.TextBox
$SerialNumber.multiline          = $false
$SerialNumber.BackColor          = "#000000"
$SerialNumber.width              = 200
$SerialNumber.height             = 20
$SerialNumber.location           = New-Object System.Drawing.Point(200,400)
$SerialNumber.Font               = 'Microsoft Sans Serif,10'
$SerialNumber.ForeColor          = "#ffffff"

$bitLocker                       = New-Object system.Windows.Forms.TextBox
$bitLocker.multiline             = $false
$bitLocker.BackColor             = "#000000"
$bitLocker.width                 = 200
$bitLocker.height                = 20
$bitLocker.location              = New-Object System.Drawing.Point(200,450)
$bitLocker.Font                  = 'Microsoft Sans Serif,10'
$bitLocker.ForeColor             = "#ffffff"

$MultipleFound                   = New-Object system.Windows.Forms.Label
$MultipleFound.text              = "Multiple Computers were Found"
$MultipleFound.AutoSize          = $true
$MultipleFound.width             = 25
$MultipleFound.height            = 10
$MultipleFound.location          = New-Object System.Drawing.Point(253,159)
$MultipleFound.Font              = 'Microsoft Sans Serif,19'
$MultipleFound.ForeColor         = "#ffffff"

$MultipleCombo                   = New-Object system.Windows.Forms.ComboBox
$MultipleCombo.text              = "Select..."
$MultipleCombo.BackColor         = "#5e5e5e"
$MultipleCombo.width             = 286
$MultipleCombo.height            = 20
$MultipleCombo.location          = New-Object System.Drawing.Point(253,221)
$MultipleCombo.Font              = 'Microsoft Sans Serif,10'
$MultipleCombo.ForeColor         = "#ffffff"

$MultipleButton                  = New-Object system.Windows.Forms.Button
$MultipleButton.BackColor        = "#5e5e5e"
$MultipleButton.text             = "Select"
$MultipleButton.width            = 65
$MultipleButton.height           = 25
$MultipleButton.location         = New-Object System.Drawing.Point(548,221)
$MultipleButton.Font             = 'Microsoft Sans Serif,10'
$MultipleButton.ForeColor        = "#ffffff"

$Panel1                          = New-Object system.Windows.Forms.Panel
$Panel1.height                   = 588
$Panel1.width                    = 856
$Panel1.BackColor                = "#4a4a4a"
$Panel1.location                 = New-Object System.Drawing.Point(10,49)

$Form.controls.AddRange(@($SearchCombo,$SearchButton,$OnlineLabel,$OnlineBox,$UpdateButton,$AddButton,$NotFound,$King,$Sweeper,$MultipleFound,$MultipleCombo,$MultipleButton,$Panel1))
$Panel1.controls.AddRange(@($SoftwareButton,$ComputerName,$LastLogon,$Model,$OperatingSystem,$IPAddress,$UserBox,$ComputerNameLabel,$ModelLabel,$TypeLabel,$IPAddressLabel,$LastLogonLabel,$UserLabel,$ManufacturerLabel,$ProcessorLabel,$RAMLabel,$OULabel,$EnabledLabel,$LastLogonDateLabel,$OU,$Enabled,$LastLogonDate,$SerialNumberLabel,$LastScannedLabel,$LastScanned,$Manufacturer,$Processor,$RAM,$SerialNumber,$bitLocker,$bitLockerLabel))




#Write your logic code here
$homeDir = get-location

. "$homeDir\Scripts\Functions\UpdateSoftware.ps1"
$Csv = Import-Csv -path "$homeDir\CSV\Hardware\HCPF_PCs.csv"

import-module activedirectory

$NotFound.visible = $False
$MultipleFound.visible = $False
$MultipleCombo.visible = $False
$MultipleButton.visible = $False
$Panel1.visible = $False
$UpdateButton.visible = $False
$AddButton.visible = $False
$OnlineLabel.visible = $False
$OnlineBox.visible = $False
$SoftwareButton.visible = $False

$ADUsers = get-aduser -filter * -Properties Name, SamAccountName
$ADComputers = Get-ADComputer -filter * -Properties SamAccountName

Foreach ($User in $ADUsers)
{
    $SearchCombo.Items.Add($User.Name)
    $SearchCombo.Items.Add($User.SamAccountName)
}
Foreach ($Computer in $ADComputers)
{
    $SearchCombo.Items.Add($Computer.SamAccountName)
}
$SearchCombo.Sorted = $True;

$SearchButton.Add_Click(
    {
        #Search by User ID
        $found = @()
        $selected = $SearchCombo.selectedItem
        $testname = Get-ADUser -filter "Name -eq '$selected'" | Select-Object -ExpandProperty SamAccountName
        for($i = 0; $i -lt $Csv.count; $i++)
        {
            if($Csv[$i].LastLogon -eq "HCPF\"+$SearchCombo.selectedItem)
            {
                $found += $Csv[$i]
            }
            if($Csv[$i].computername -match $SearchCombo.selectedItem)
            {
                $found += $Csv[$i]
            }
            if($Csv[$i].LastLogon -eq "HCPF\"+$testname)
            {
                $found += $Csv[$i]
            }
        }
        #Search by Fullname
        
        
        #Handle no computers found
        if($found.count -eq 0)
        {
            $NotFound.visible = $True
            $MultipleFound.visible = $False
            $MultipleCombo.visible = $False
            $MultipleButton.visible = $False
            $Panel1.visible = $False
            $UpdateButton.visible = $False
            $AddButton.visible = $True
            $OnlineLabel.visible = $False
            $OnlineBox.visible = $False
        }
        #Handle 1 computer found
        if($found.count -eq 1)
        {
            $NotFound.visible = $False
            $MultipleFound.visible = $False
            $MultipleCombo.visible = $False
            $MultipleButton.visible = $False
            $Panel1.visible = $True
            $UpdateButton.visible = $False
            $AddButton.visible = $False
            $OnlineLabel.visible = $False
            $OnlineBox.visible = $False
            
            if(Test-Connection -Cn $found[0].ComputerName -Count 1 -quiet)
            {
                $UpdateButton.visible = $True
                $OnlineLabel.visible = $True
                $OnlineBox.visible = $True
                $OnlineBox.BackColor = "#7ed321"
            }
            else
            {
                $UpdateButton.visible = $False
                $OnlineLabel.visible = $True
                $OnlineBox.visible = $True
                $OnlineBox.BackColor = "#d0021b"
            }
            
            $ComputerName.Text = $found[0].ComputerName
            $Enabled.Text = $found[0].Enabled
            $IpAddress.Text = $found[0].IpAddress
            $OU.Text = $found[0].OU
            $SerialNumber.Text = $found[0].SerialNumber
            $Manufacturer.Text = $found[0].Manufacturer
            $Model.Text = $found[0].Model
            $RAM.Text = $found[0].RAM
            $Processor.Text = $found[0].ProcessorName
            $OperatingSystem.Text = $found[0].OperatingSystem
            $LastLogon.Text = $found[0].LastLogon
            $LastScanned.Text = $found[0].LastScanned
            $LastLogonDate.Text = $found[0].LastLogonDate
            $name = $found[0].ComputerName
            
            $tempName = $LastLogon.Text.Split('\')[1]
            $newUser = Get-ADUser $tempName | Select-Object -ExpandProperty Name
            $UserBox.text = $newUser

            if ($found[0].BitLockerRecoveryKey -ne $null) {$bitLocker.text = $found[0].BitLockerRecoveryKey}
            else {$bitLocker.text = "Key not on file"}
            
            $path
            Set-Variable -Name "path" -Value "$homeDir\CSV\Software\${name}.csv" -Scope Global
            try
            {
                $Software = Import-Csv -path $path 
                $SoftwareButton.visible = $True
            }
            catch [Exception]
            {
                $SoftwareButton.visible = $False
            }
        }
        #Handle multiple computers found
        if($found.count -gt 1)
        {
            $NotFound.visible = $False
            $MultipleFound.visible = $True
            $MultipleCombo.visible = $True
            $MultipleButton.visible = $True
            $Panel1.visible = $False
            $UpdateButton.visible = $False
            $AddButton.visible = $False
            $OnlineLabel.visible = $False
            $OnlineBox.visible = $False
            
            $MultipleCombo.Items.Clear();
			$MultipleCombo.text = "Select..."
			            
            for($i = 0; $i -lt $found.count; $i++)
            {
                $MultipleCombo.Items.Add($found[$i].ComputerName + "    - " + $found[$i].LastLogonDate)
            }
        }
    }    
)

$SoftwareButton.Add_Click(
    {
        $Software = Import-Csv -path $path | Out-GridView
    }    
)

$MultipleButton.Add_Click(
    {
        #Load 1 computer found after multple were found for selected computer
        $selectedComputer = $MultipleCombo.SelectedItem.Split(' ')[0]
        write-host $selectedComputer
        for($i = 0; $i -lt $Csv.count; $i++)
        {
            if($Csv[$i].computername -match $selectedComputer)
            {
                $found += $Csv[$i]
                break
            }
        }
        $NotFound.visible = $False
        $MultipleFound.visible = $False
        $MultipleCombo.visible = $False
        $MultipleButton.visible = $False
        $Panel1.visible = $True
        $UpdateButton.visible = $False
        $AddButton.visible = $False
        $OnlineLabel.visible = $False
        $OnlineBox.visible = $False
        
        if(Test-Connection -Cn $found[0].ComputerName -Count 1 -quiet)
        {
            $UpdateButton.visible = $True
            $OnlineLabel.visible = $True
            $OnlineBox.visible = $True
            $OnlineBox.BackColor = "#7ed321"
        }
        else
        {
            $UpdateButton.visible = $False
            $OnlineLabel.visible = $True
            $OnlineBox.visible = $True
            $OnlineBox.BackColor = "#d0021b"
        }
        
        $ComputerName.Text = $found[0].ComputerName
        $Enabled.Text = $found[0].Enabled
        $IpAddress.Text = $found[0].IpAddress
        $OU.Text = $found[0].OU
        $SerialNumber.Text = $found[0].SerialNumber
        $Manufacturer.Text = $found[0].Manufacturer
        $Model.Text = $found[0].Model
        $RAM.Text = $found[0].RAM
        $Processor.Text = $found[0].ProcessorName
        $OperatingSystem.Text = $found[0].OperatingSystem
        $LastLogon.Text = $found[0].LastLogon
        $LastScanned.Text = $found[0].LastScanned
        $LastLogonDate.Text = $found[0].LastLogonDate
        $name = $found[0].ComputerName
        $path
        
        $tempName = $LastLogon.Text.Split('\')[1]
        $newUser = Get-ADUser $tempName | Select-Object -ExpandProperty Name
        $UserBox.text = $newUser

        if ($found[0].BitLockerRecoveryKey -ne $null) {$bitLocker.text = $found[0].BitLockerRecoveryKey}
        else {$bitLocker.text = "Key not on file"}
        
        Set-Variable -Name "path" -Value "$homeDir\CSV\Software\${name}.csv" -Scope Global
        try
        {
            $Software = Import-Csv -path $path 
            $SoftwareButton.visible = $True
        }
        catch [Exception]
        {
            $SoftwareButton.visible = $False
        }
    }   
)

$UpdateButton.Add_Click(
    {
        UpdateSoftware $ComputerName.Text
    }    
)

$AddButton.Add_Click(
    {
        UpdateSoftware $SearchCombo.Text
    }
)
[void]$Form.ShowDialog()