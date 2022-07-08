<# 

    Upload Files - Function upload saved CSV to a webserver for web portal version
    Published 12/9/2021
    Created By Kyle King

#>

function Upload()
{
	# Set file locations
	$homeDir = get-location
    $hardwarePath = "$homeDir\CSV\ServerFiles\hardware.csv"
    $AD_ObjPath = "$homeDir\CSV\ServerFiles\ADOjbects.csv"
    $softwarePath = "$homeDir\CSV\ServerFiles\allHCPF_Software.csv"

    ### Generate Server Files ###

    # Copy updated Hardware file to ServerFiles
    Get-Content "$homeDir\CSV\Hardware\HCPF_PCs.csv" | Set-Content $hardwarePath

    # Create AD objects CSV
    Write-Host "Generating AD Objects files..."
    Get-ADUser -filter * | select Name,SamAccountName | Export-Csv -Path $AD_ObjPath

    # Create compiled software list
    Write-Host "Generating Software file..."
    Get-ChildItem "$homeDir\CSV\Software\*.csv" | ForEach-Object {Import-Csv $_} | Export-Csv $softwarePath -Force


    ### Export file to KS Server ###

    # Define clear text string for username and password
    [string]$userName = '' <# removed for privacy #>
    [string]$userPassword = '' <# removed for privacy #>
    
    # Convert to SecureString
    [securestring]$secStringPassword = ConvertTo-SecureString $userPassword -AsPlainText -Force
    [pscredential]$credObject = New-Object System.Management.Automation.PSCredential ($userName, $secStringPassword)
    
    # Export files to webserver
    Set-SCPFile -ComputerName '' <# removed for privacy #> -Credential $credObject -RemotePath  '/var/www/kingsweeper' -LocalFile $hardwarePath
    Set-SCPFile -ComputerName '' <# removed for privacy #> -Credential $credObject -RemotePath  '/var/www/kingsweeper' -LocalFile $AD_ObjPath
    Set-SCPFile -ComputerName '' <# removed for privacy #> -Credential $credObject -RemotePath  '/var/www/kingsweeper' -LocalFile $softwarePath
}