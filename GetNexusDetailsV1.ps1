#just install of copy Posh-SSH module to %userprofile%\Documents\WindowsPowerShell\Modules\Posh-SSH

#Specify variables
$User = "admin"
$Pass = "VMwar3!!"
$HostMachine = "10.253.4.211"
$LocalDir = "C:\Users\Rao2S\Desktop\RCM Switch Automation"
$PythonScript = "get_nexus_info.py"

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
    $JsonFile = Invoke-SSHCommandStream -Command $RunPythonScriptCmd -SSHSession $SSHSession
    Write-Host $JsonFile
}

Write-Host "Python Script Executed"


# Below section is not working yet
<#
#Download json output file from switch to Element Manager VM using SCP
$json_remote_file_path = '/' +  $json_file
$json_local_file_path = Join-Path $local_dir $json_file

Write-Host $json_remote_file_path $json_local_file_path
get-SCPFile -ComputerName $HostMachine -Credential $Creds -LocalFile 'C:\Users\Rao2S\Desktop\RCM Switch Automation\10_253_4_211_op_2022_01_29_18_07_07.json' -RemoteFile '/10_253_4_211_op_2022_01_29_18_07_07.json'
#>

#Delete the json file from boot flash post download to local directory
if ($SSHSession.Connected -eq "True") {
    $DeleteJsonFileCmd = "delete bootflash:/" + $JsonFile
    Invoke-SSHCommandStream -Command $DeleteJsonFileCmd -SSHSession $SSHSession
    Write-Host "Json File Deleted"
}


Remove-SSHSession -SSHSession $SSHSession