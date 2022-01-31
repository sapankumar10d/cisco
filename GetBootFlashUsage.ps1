#Specify variables
$User = "admin"
$Pass = "VMwar3!!"
$HostMachine = "10.253.4.211"


#Create credential object
[securestring]$secStringPassword = ConvertTo-SecureString $Pass -AsPlainText -Force
$Creds = New-Object System.Management.Automation.PSCredential -ArgumentList ($User, $secStringPassword)

#Create ssh session
$SSHSession = New-SSHSession -ComputerName $HostMachine -Credential $Creds


#Execute command 
$ExecuteCmd = "dir bootflash:"

if ($SSHSession.Connected -eq "True") {
    $CmdOutput = Invoke-SSHCommandStream -Command $ExecuteCmd -SSHSession $SSHSession
}


#Print cmd output
#Write-Output $CmdOutput 
<#  Manually Parsing the text and building the JSON
$BootFlashTotal = -split ($CmdOutput| Select-String -Pattern "bytes total") |Select-Object -Index 0
$BootFlashUsed = -split ($CmdOutput| Select-String -Pattern "bytes used") |Select-Object -Index 0
$BootFlashFree = -split ($CmdOutput| Select-String -Pattern "bytes free") |Select-Object -Index 0

$BootFlashTotalMB = [math]::Round(($BootFlashTotal/1024/1024),2)
$BootFlashUsedMB = [math]::Round(($BootFlashUsed/1024/1024),2)
$BootFlashFreeMB = [math]::Round(($BootFlashFree/1024/1024),2)
$BootFlashUsedPercent = [math]::Round(($BootFlashUsed/$BootFlashTotal)*100,2)

$BootFlashUsage = @"
{ 
"BootFlashTotalMB": $BootFlashTotalMB,
"BootFlashUsedMB": $BootFlashUsedMB,
"BootFlashFreeMB": $BootFlashFreeMB,
"BootFlashUsedPercent": $BootFlashUsedPercent
}
"@
#>

#$BootFlashUsage| ConvertFrom-Json | ConvertTo-Json 
## Writing to Console
#Write-Host "`nBootFlash Usage:: Total MB:" $BootFlashTotalMB ,"  Used MB:" $BootFlashUsedMB ,"  Free MB:" $BootFlashFreeMB ,"  Used %:" $BootFlashUsedPercent

## Writing to Console
Write-Host "`Here are top 5 large Files/Dirs`n"
$CmdOutput | Select-Object -SkipLast 5 | Sort-Object –Descending | Select-Object -First 5


## Testing Converting to JSON
#$BootFlashFiles = $CmdOutput | Select-Object -SkipLast 5 | ConvertFrom-String -PropertyNames Ignore,Bytes,Month,Date,Time,Year,Name|Select-Object Month,Date,Year,Name,Bytes | ConvertTo-Json
#$BootFlashFiles = $CmdOutput | Select-Object -SkipLast 5 | ConvertFrom-String -PropertyNames Ignore,Bytes,Month,Date,Time,Year,Name|Select-Object Month,Date,Year,Name,Bytes | Sort-Object -Descending -Property Bytes | ConvertTo-Json
#$BootFlashUsage = $CmdOutput | Select-Object -Last 3 | ConvertFrom-String -PropertyNames Ignore1, Bytes, Ignore2, Type|Select-Object Type, Bytes | ConvertTo-Json

 

## Testing Writing to CSV
#$CmdOutput | Select-Object -SkipLast 5 | ConvertFrom-String -PropertyNames Ignore,Bytes,Month,Date,Time,Year,Name|Select-Object Month,Date,Year,Name,Bytes | Sort-Object -Descending -Property Bytes | Export-Csv -Path "C:\Users\Rao2S\Desktop\RCM Switch Automation\BootFlashFiles.csv" -NoTypeInformation


## Testing Reading from CSV
#$BootFlashFilesCsv = Import-Csv -Path "C:\Users\Rao2S\Desktop\RCM Switch Automation\BootFlashFiles.csv"


<# Testing iteration through CSV
$BootFlashFiles | ForEach-Object {
Write-Host "$($_.Name)"
}
#>


## Collecting PS Objects and using them as CSV to write to Excel
$BootFlashFiles = $CmdOutput | Select-Object -SkipLast 5 | ConvertFrom-String -PropertyNames Ignore,Bytes,Month,Date,Time,Year,Name|Select-Object Month,Date,Year,Name,Bytes | Sort-Object -Descending -Property Bytes
$BootFlashUsage = $CmdOutput | Select-Object -Last 3 | ConvertFrom-String -PropertyNames Ignore1, Bytes, Ignore2, Type|Select-Object Type, Bytes

Remove-SSHSession -SSHSession $SSHSession

### Excel 
$excel = New-Object -ComObject excel.application
$excel.visible = $False
$workbook = $excel.Workbooks.Add()
$bootflashfileswkst= $workbook.Worksheets.Item(1)
$bootflashfileswkst.Name = "Bootflash Files"

$bootflashusagewkst= $workbook.Worksheets.Item(2)
$bootflashusagewkst.Name = "Bootflash Usage"
 
## writing Headers for Bootflash Files
$bootflashfileswkst.Cells.Item(1,1) = 'Bootflash Files'
$bootflashfileswkst.Cells.Item(2,1) = 'Month'
$bootflashfileswkst.Cells.Item(2,2) = 'Date'
$bootflashfileswkst.Cells.Item(2,3) = 'Year'
$bootflashfileswkst.Cells.Item(2,4) = 'Name'
$bootflashfileswkst.Cells.Item(2,5) = 'Bytes'

## Formatting the header cell for Bootflash Files
$bootflashfileswkst.Cells.Item(1,1).Font.Size = 12
$bootflashfileswkst.Cells.Item(1,1).Font.Bold=$True
$bootflashfileswkst.Cells.Item(1,1).Font.Name = "Calibri"
$bootflashfileswkst.Cells.Item(1,1).Font.ThemeFont = 1
$bootflashfileswkst.Cells.Item(1,1).Font.ThemeColor = 4
$bootflashfileswkst.Cells.Item(1,1).Font.ColorIndex = 55
$bootflashfileswkst.Cells.Item(1,1).Font.Color = 8210719



## Populating data to Bootflash Files worksheet
$i = 3
foreach($filename in $BootFlashFiles) 
{ 
 $bootflashfileswkst.cells.item($i,1) = $filename.Month
 $bootflashfileswkst.cells.item($i,2) = $filename.Date
 $bootflashfileswkst.cells.item($i,3) = $filename.Year
 $bootflashfileswkst.cells.item($i,4) = $filename.Name
 $bootflashfileswkst.cells.item($i,5) = $filename.Bytes
 $i++ 
} 

## writing Headers for Bootflash Usage
$bootflashusagewkst.Cells.Item(1,1) = 'Bootflash Usage'
$bootflashusagewkst.Cells.Item(2,1) = 'Type'
$bootflashusagewkst.Cells.Item(2,2) = 'Bytes'


## Formatting the header cell for Bootflash Usage
$bootflashusagewkst.Cells.Item(1,1).Font.Size = 12
$bootflashusagewkst.Cells.Item(1,1).Font.Bold=$True
$bootflashusagewkst.Cells.Item(1,1).Font.Name = "Calibri"
$bootflashusagewkst.Cells.Item(1,1).Font.ThemeFont = 1
$bootflashusagewkst.Cells.Item(1,1).Font.ThemeColor = 4
$bootflashusagewkst.Cells.Item(1,1).Font.ColorIndex = 55
$bootflashusagewkst.Cells.Item(1,1).Font.Color = 8210719

## Populating data to Bootflash Usage worksheet
$i = 3
foreach($item in $BootFlashUsage) 
{ 
 $bootflashusagewkst.cells.item($i,1) = $item.Type
 $bootflashusagewkst.cells.item($i,2) = $item.Bytes

 $i++ 
} 

#Save Excel
$excel.DisplayAlerts = 'False'
$ext=".xlsx"
$path="C:\Users\Rao2S\Desktop\RCM Switch Automation\NexusReport$ext"
$workbook.SaveAs($path) 
$workbook.Close
$excel.DisplayAlerts = 'False'
$excel.Quit()