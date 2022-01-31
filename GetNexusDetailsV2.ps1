#just install of copy Posh-SSH module to %userprofile%\Documents\WindowsPowerShell\Modules\Posh-SSH

#Specify variables
$User = "admin"
$Pass = "VMwar3!!"
$HostMachine = "10.253.4.211"
$LocalDir = "C:\Users\Rao2S\Desktop\RCM Switch Automation"
$PythonScript = "GetNexusInfoV2.py"
$JsonFile = $HostMachine.Replace('.','_') + '.json'

#Create credential object
[securestring]$secStringPassword = ConvertTo-SecureString $Pass -AsPlainText -Force
$Creds = New-Object System.Management.Automation.PSCredential -ArgumentList ($User, $secStringPassword)

#Create ssh session
$SSHSession = New-SSHSession -ComputerName $HostMachine -Credential $Creds

#Transfer python script using scp
$PythonScriptPath = Join-Path $LocalDir $PythonScript
Set-SCPFile -ComputerName $HostMachine -Credential $Creds -LocalFile $PythonScriptPath -RemotePath '/'
Write-Host "Python Script Copied"

#Execute python script
if ($SSHSession.Connected -eq "True") {
    $RunPythonScriptCmd = "python bootflash:/" + $PythonScript
    $JsonOutput = Invoke-SSHCommandStream -Command $RunPythonScriptCmd -SSHSession $SSHSession
}
Write-Host "Python Script Executed"

#Print Json captured from switch - testing only
Write-Output $JsonOutput | ConvertFrom-Json| ConvertTo-Json -depth 100

## Exploring PS custom Objects
$JsonOutput.GetType() 
($JsonOutput | ConvertFrom-Json).GetType()
($JsonOutput | ConvertFrom-String).GetType()
$custom_obj = $JsonOutput | ConvertFrom-Json
$custom_obj.sh_ver.chassis_id
$custom_obj.sh_inv.TABLE_inv.ROW_inv

$custom_obj.sh_inv.TABLE_inv.ROW_inv | ForEach-Object {
Write-Host "$($_.serialnum, $_.productid)"
}

$custom_obj.sh_mod.TABLE_modwwninfo.ROW_modwwninfo
$custom_obj.sh_mod.TABLE_modinfo.ROW_modinfo
$custom_obj.sh_mod.TABLE_modinfo.ROW_modinfo.model

$custom_obj.sh_mod.TABLE_modmacinfo.ROW_modmacinfo

$custom_obj.sh_env.powersup.TABLE_psinfo.ROW_psinfo
$custom_obj.sh_env.powersup.TABLE_psinfo.power_summary
$custom_obj.sh_env.powersup.TABLE_psinfo.voltage_level
$custom_obj.sh_env.powersup.TABLE_mod_pow_info.ROW_mod_pow_info
$custom_obj.sh_env.powersup.TABLE_mod_pow_info.ROW_mod_pow_info
$custom_obj.sh_env.TABLE_tempinfo.ROW_tempinfo
$custom_obj.sh_env.TABLE_tempinfo.fandetails.TABLE_faninfo.ROW_faninfo





# Write Json to File
$jsonFilePath = Join-Path $LocalDir  $JsonFile
$JsonOutput | ConvertFrom-Json| ConvertTo-Json -depth 100 |Out-File $jsonFilePath
Write-Host "JSON file saved at " $jsonFilePath


Remove-SSHSession -SSHSession $SSHSession