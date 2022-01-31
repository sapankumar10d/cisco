#just install of copy Posh-SSH module to %userprofile%\Documents\WindowsPowerShell\Modules\Posh-SSH

#Specify variables
$User = "admin"
$Pass = "VMwar3!!"
$HostMachine = "10.253.4.211"

#Create credential object
[securestring]$secStringPassword = ConvertTo-SecureString $Pass -AsPlainText -Force

$Creds = New-Object System.Management.Automation.PSCredential -ArgumentList ($User, $secStringPassword)

#Create ssh session

$SSHSession = New-SSHSession -ComputerName $HostMachine -Credential $Creds

if ($SSHSession.Connected -eq "True") {

#Run commands
$Out = Invoke-SSHCommandStream -Command "show version" -SSHSession $SSHSession

#Print output
#$Out = $Out.split("`n")

$Out | Select-String -Pattern ":"

<#
#Transfer files using scp
Set-SCPFile -ComputerName $HostMachine -Credential $Creds -LocalFile 'C:\Dell\pan_gp_hrpt.xml' -RemotePath '/deepak'

#Download files using scp
get-SCPFile -ComputerName $HostMachine -Credential $Creds -LocalFile 'C:\Dell\pan_gp_hrpt.xml1' -RemoteFile '/deepak/pan_gp_hrpt.xml'

#Set-SCPFolder -ComputerName $HostMachine -Credential $Creds -LocalFolder 'C:\Dell\PrintAnywhere\' -RemoteFolder '/deepak/PrintAnywhere'

#>

}

Remove-SSHSession -SSHSession $SSHSession