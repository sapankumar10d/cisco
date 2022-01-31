##### This Script will read the JSON output file created by python script + local run commands 


### Define Variables

## The file names later will have time stamp and switch_ips

$LocalDir = "C:\Users\kumars92\OneDrive - Dell Technologies\PowerShell\"
$JsonFile = "10_253_4_211.json"
$ExcelFile = "Nexus_Report.xlsx"


$JsonFileLocalPath = Join-Path $LocalDir $JsonFile
$ExcelFileLocalPath = Join-Path $LocalDir $ExcelFile

Write-Host $JsonFileLocalPath
Write-Host $ExcelFileLocalPath

### Read JSON File
$json = Get-Content $JsonFileLocalPath | ConvertFrom-Json

### Making yourself familier with JSON Object/ PS custom Object 
<#
$json.GetType()
$json.sh_ver
$json.sh_ver.chassis_id
$json.sh_ver.bootflash_size
$json.sh_ver.kick_file_name
$json.sh_ver.kickstart_ver_str
$json.sh_ver.sys_ver_str
$json.sh_ver.kick_cmpl_time
$json.sh_ver.host_name
#>

#$json.sh_inv.TABLE_inv.ROW_inv

### Excel 
<#
The Code to decide - to replace the file or to add time stamp 
#>
$excel = New-Object -ComObject excel.application
$excel.visible = $False
$workbook = $excel.Workbooks.Add()

$sh_ver_wkst= $workbook.Worksheets.Item(1)
$sh_ver_wkst.Name = "Show Version"

$sh_inv_wkst= $workbook.Worksheets.Item(2)
$sh_inv_wkst.Name = "Show Inventory"

 
## writing Headers for Show Version
$sh_ver_wkst.Cells.Item(1,1) = 'Show Version'
$sh_ver_wkst.Cells.Item(2,1) = 'chassis_id'
$sh_ver_wkst.Cells.Item(3,1) = 'bootflash_size'
$sh_ver_wkst.Cells.Item(4,1) = 'kick_file_name'
$sh_ver_wkst.Cells.Item(5,1) = 'sys_ver_str'
$sh_ver_wkst.Cells.Item(6,1) = 'host_name'

## writing Headers for Show Environment
$sh_inv_wkst.Cells.Item(1,1) = 'Show Inventory'
$sh_inv_wkst.Cells.Item(2,1) = 'vendorid'
$sh_inv_wkst.Cells.Item(2,2) = 'serialnum'
$sh_inv_wkst.Cells.Item(2,3) = 'productid'
$sh_inv_wkst.Cells.Item(2,4) = 'name'
$sh_inv_wkst.Cells.Item(2,5) = 'desc'


<# Need to Adjust the Column width 
#>

## Formatting the header cell for Show Version
$sh_ver_wkst.Cells.Item(1,1).Font.Size = 12
$sh_ver_wkst.Cells.Item(1,1).Font.Bold=$True
$sh_ver_wkst.Cells.Item(1,1).Font.Name = "Calibri"
$sh_ver_wkst.Cells.Item(1,1).Font.ThemeFont = 1
$sh_ver_wkst.Cells.Item(1,1).Font.ThemeColor = 4
$sh_ver_wkst.Cells.Item(1,1).Font.ColorIndex = 55
$sh_ver_wkst.Cells.Item(1,1).Font.Color = 8210719

## Formatting the header cell for Show Inventory
$sh_inv_wkst.Cells.Item(1,1).Font.Size = 12
$sh_inv_wkst.Cells.Item(1,1).Font.Bold=$True
$sh_inv_wkst.Cells.Item(1,1).Font.Name = "Calibri"
$sh_inv_wkst.Cells.Item(1,1).Font.ThemeFont = 1
$sh_inv_wkst.Cells.Item(1,1).Font.ThemeColor = 4
$sh_inv_wkst.Cells.Item(1,1).Font.ColorIndex = 55
$sh_inv_wkst.Cells.Item(1,1).Font.Color = 8210719


## Populating data to Show Version worksheet

foreach($item in $json.sh_ver) 
{ 
 $sh_ver_wkst.cells.item(2,2) = $item.chassis_id
 $sh_ver_wkst.cells.item(3,2) = $item.bootflash_size
 $sh_ver_wkst.cells.item(4,2) = $item.kick_file_name
 $sh_ver_wkst.cells.item(5,2) = $item.sys_ver_str
 $sh_ver_wkst.cells.item(6,2) = $item.host_name
} 


## Populating data to Show Inventory worksheet

$i = 3
foreach($item in $json.sh_inv.TABLE_inv.ROW_inv) 
{ 
 $sh_inv_wkst.cells.item($i,1) = $item.vendorid
 $sh_inv_wkst.cells.item($i,2) = $item.serialnum
 $sh_inv_wkst.cells.item($i,3) = $item.productid
 $sh_inv_wkst.cells.item($i,4) = $item.name
 $sh_inv_wkst.cells.item($i,5) = $item.desc
 $i++ 
} 

#Save Excel
$excel.DisplayAlerts = 'False'
$workbook.SaveAs($ExcelFileLocalPath) 
$workbook.Close
$excel.DisplayAlerts = 'False'
$excel.Quit()







###################
### Break Down of some addtional JSON Objects 

<#

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

#>