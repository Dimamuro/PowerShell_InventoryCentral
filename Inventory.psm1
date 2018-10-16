
$Script:DirDir = "D:\Inventory"
#find directory yourself
#$Script:PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

function get-ADList{
	import-module activedirectory
	$Object=Get-ADComputer -filter * -Properties name, DistinguishedName, LastLogon, operatingsystem|where{$_.DistinguishedName.ToString() -like "*OU=_KMZ*"}|Select-Object name, operatingsystem|sort-object name
	$Object|Add-Member -MemberType NoteProperty -Name OS_Architecture -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name OS_SerialNumber -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name HDD_model -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name HDD_size -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name HDD_SerialNumber -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name RAM -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name RAM_SerialNumber -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name CPU -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name MB_name -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name MB_Manufacturer -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name MB_SerialNumber -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name PCType -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name printer -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name IP_Address -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name MAC_Address -Value '' -Force
	$Object|Add-Member -MemberType NoteProperty -Name Status -Value 'NONE' -Force
	#вариант для get-UpdateList
	#$Object|Add-Member -MemberType NoteProperty -Name Status -Value 'Off' -Force

	return $Object
}

function get-TypePC{
    param($PCname)
	$PCmodel=""
	$WMIinfo=gwmi win32_systemenclosure -ComputerName $PCname
	switch($WMIinfo.chassistypes){
		1 {$PCmodel="Другое";break}
		2 {$PCmodel="Unknown";break}
		3 {$PCmodel="Настольный ПК";break}
		4 {$PCmodel="Low Profile Desktop";break}
		5 {$PCmodel="Pizza Box";break}
		6 {$PCmodel="Mini Tower";break}
		7 {$PCmodel="Tower";break}
		8 {$PCmodel="Portable";break}
		9 {$PCmodel="Laptop";break}
		10 {$PCmodel="Ноутбук";break}
		11 {$PCmodel="Handheld";break}
		12 {$PCmodel="Docking Station";break}
		13 {$PCmodel="All-in-One";break}
		14 {$PCmodel="Sub-Notebook";break}
		15 {$PCmodel="Space Saving";break}
		16 {$PCmodel="Lunch Box";break}
		17 {$PCmodel="Main System Chassis";break}
		18 {$PCmodel="Expansion Chassis";break}
		19 {$PCmodel="Sub-Chassis";break}
		20 {$PCmodel="Bus Expansion Chassis";break}
		21 {$PCmodel="Peripheral Chassis";break}
		22 {$PCmodel="Storage Chassis";break}
		23 {$PCmodel="Rack Mount Chassis";break}
		24 {$PCmodel="Sealed-Case PC";break}
	}
	return $PCmodel
}

function get-HHDRightSize{
    param([int32]$size)
	switch($size){
		75 {$size=80;break}
		149 {$size=160;break}
		233 {$size=240;break}
		234 {$size=240;break}
		298 {$size=320;break}
		466 {$size=500;break}
		596 {$size=650;break}
		default {$size=$size;break}
	}
	return $size
}

function start-Archive{
    param([int32]$CountArchive)
	if((Test-Path "$DirDir\Archive") -eq $false){new-item "$DirDir\Archive" -Type directory}   
	#archive data directory

	 $date=[string](Get-Date -Format "%d_%M_%y-%h_%m")
	 cd ${env:ProgramFiles}
	 .\7-Zip\7z.exe a -tzip -mx9 "$DirDir\Archive\$date.zip" "$DirDir\Data"

	 #delete old archive
	dir $DirDir\Archive\*.zip|Sort-Object LastWriteTime -Descending|Select-Object -Skip $CountArchive|foreach{del $_.fullname}
	return
}

function get-OldList{
	$OldList=Import-Clixml $DirDir\Data\adcomp.xml
	return $OldList
}

function get-UpdateList{
	$OldList=get-OldList
	$NewList=get-ADList

	for($i=0;$i -lt $NewList.length;$i++)
	{
		$NewList[$i].Status="Ok"
		for($j=0;$j -lt $OldList.length;$j++)
		{
			if($NewList[$i].name -like $OldList[$j].name)
			{
				if($NewList[$i].operatingsystem -like $OldList[$j].operatingsystem)
				{
					$OldList[$j].Status="Ok"
					$NewList[$i]=$OldList[$j]
					if($j -eq 0){$OldList=$OldList[($j+1)..($OldList.length-1)]} else
					{$OldList=$OldList[0..($j-1)+($j+1)..($OldList.length-1)]}
				} else
				{$OldList[$j].Status = "NONE"}
			} else
			{$OldList[$j].Status = "NONE"}     
		}
	}
	$ADList=$NewList+$OldList
	$ADList|where{$_.Status -like "NONE"}|foreach{
	if($_.name -like "OLD_*"){} else
		{$_.name = "OLD_"+$_.name}
	}
	return $ADList
}


function get-NewInventory{

	if((Test-Path "$DirDir\Data") -eq $false){new-item "$DirDir\Data" -Type directory}

	$ADFile=get-ADList

	$ADFile|Export-Clixml $DirDir\Data\adcomp.xml

	start-Inventory
}


#Поиск изменений в составе компа
function get-terrorist($NewList){
	$OldList=get-OldList
	$ADList=$NewList
	for($i=0;$i -lt $OldList.length;$i++)
	{
		for($j=0;$j -lt $NewList.length;$j++)
		{
			if($NewList[$i].name -like "OLD_*"){}else{
				if($NewList[$i].CPU -like ""){}else{
					if($OldList[$i] -like $NewList[$j]){}else{
						$OldList[$i].name="Terrorist_"+$OldList[$i].name
						$ADList=$ADList+$OldList[$i]
					}
				}
			}
		}
	}
	return $ADList
}

function do-inventory(){
    [CmdletBinding()]
	param($Comp)


	if($Comp.status -like "Ok")
	{
		if((Test-Connection $Comp.name -Count 1 -Quiet) -eq $false)
		{
			#Write-Host $Comp.name "Выключен..." -ForegroundColor Red
			Write-Verbose $Comp.name "Выключен `t x"
			$Comp.Status="Off"
		} else
		{
			Write-Verbose $Comp.name " Включен `t :)"
			#Write-Host $Comp.name "Включен..." -ForegroundColor Green
			$Buff=Get-WmiObject -ComputerName $Comp.name -Class win32_operatingsystem
			if($? -eq $True)
			{
				$hdd=get-wmiobject -ComputerName $Comp.name -Class win32_diskdrive|where{$_.DeviceID -like "*PHYSICALDRIVE0"}|Select-Object SerialNumber, @{n="size"; e={[int32]($_.size/1048576/1024)}}, Model
				$os=Get-WmiObject -ComputerName $Comp.name -Class win32_operatingsystem|Select-Object OSArchitecture, serialnumber
				$mb=Get-WmiObject -ComputerName $Comp.name -Class win32_baseboard|Select-Object product, serialnumber, Manufacturer
				$cpu=Get-WmiObject -ComputerName $Comp.name -Class win32_processor|Select-Object name
				$printer=Get-WmiObject -ComputerName $Comp.name -Class Win32_Printer|where{$_.name -notlike "*pdf*" -and $_.name -notlike "fax"-and $_.name -notlike "microsoft xps*" -and $_.name -notlike "*OneNote*"}|Select-Object name
				$ram=Get-WmiObject -ComputerName $Comp.name -Class win32_physicalmemory|Select-Object serialnumber, @{n="size"; e={($_.capacity/1024/1024/1024)}}
				$net=Get-WmiObject -ComputerName $Comp.name -Class Win32_NetworkAdapterConfiguration|where{$_.DNSDomain -like "netgate.kmz" }
				$mac=Get-WmiObject -ComputerName $Comp.name -Class Win32_NetworkAdapter|where{$_.ServiceName -like $net.ServiceName}

				$Comp.RAM='';$Comp.RAM+=$ram|foreach{[string]$_.size+"GB"}
				$Comp.RAM_SerialNumber='';$Comp.RAM_SerialNumber+=$ram|foreach{$_.serialnumber+";"}
				$Comp.OS_Architecture=$os.OSArchitecture
				$Comp.OS_SerialNumber=$os.serialnumber
				$Comp.Status="Ok"
				$Comp.HDD_model= $hdd.model
				$Comp.HDD_size= [string](get-HHDRightSize $hdd.size)+"GB"
				$Comp.HDD_SerialNumber=$hdd.SerialNumber
				if($Comp.HDD_SerialNumber -notlike $null -and $Comp.HDD_SerialNumber.Length -ge 30)
				{
					$hdd_dex=$Comp.HDD_SerialNumber -split '(.{2})' |%{ if ($_ -ne ""){[CHAR]([CONVERT]::toint16("$_",16))}}
					$Comp.HDD_SerialNumber=''
					$Comp.HDD_SerialNumber+=$hdd_dex|foreach{$_}
					$Comp.HDD_SerialNumber=$Comp.HDD_SerialNumber -replace " "
				}
				$Comp.MB_name=$mb.product
				$Comp.MB_Manufacturer=$mb.Manufacturer
				$Comp.MB_SerialNumber=$mb.serialnumber
				$Comp.printer='';$Comp.printer+=$printer|foreach{$_.name+";"}
				$Comp.CPU=$cpu.name
				$Comp.PCType=get-TypePC $Comp.name
				$Comp.IP_Address=$net.IPAddress[0]
				$Comp.MAC_Address=($mac|where{$_.Speed -eq 100000000}).MACAddress
			}
		}
	}
	return $Comp
}
