Import-Module Inventorys

$Script:DirDir = "D:\Inventory"

workflow start-parallel{
param($ADList)
$InventoryList=@()
foreach -parallel ($Comp in $AdList)
{
$WORKFLOW:InventoryList+=do-inventory $Comp
}
return $InventoryList
}

function start-Inventory{

if((Test-Path "$DirDir\Data") -eq $false){get-NewInventory;break} else
{
start-Archive 1000
#$ADList=get-OldList
$ADList=get-UpdateList
$ADList=start-parallel $ADList
}

$ADList=$ADlist|select name,operatingsystem,OS_Architecture,OS_SerialNumber,
HDD_model,HDD_size,HDD_SerialNumber,RAM,RAM_SerialNumber,
CPU,MB_name,MB_Manufacturer,MB_SerialNumber,
PCType,printer,IP_Address,MAC_Address,Status|sort-object name

$ADlist|sort-object name|Export-Clixml $DirDir\Data\adcomp.xml

$ADList=$ADlist|select name,operatingsystem,OS_Architecture,#OS_SerialNumber,
HDD_model,HDD_size,HDD_SerialNumber,RAM,#RAM_SerialNumber,
CPU,MB_name,MB_Manufacturer,#MB_SerialNumber,
PCType,printer,IP_Address,MAC_Address|sort-object name
$ADList|format-table

$ADlist|ConvertTo-Html > $DirDir\fullcomp.html
}

start-Inventory
