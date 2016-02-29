
## Uses Wmi queries to return information about the computer; this can be modified to return additional information as needed.
## MAKE SURE TO UPDATE THE PATH VARIABLE AT THE BOTTOM (LINE 27).  This is where the reported .CSV will be saved.  It should used a FQDN or IP, not a mapped network drive.

Function Compstats {
$computer = $env:COMPUTERNAME

$mobo = get-wmiobject -computer $computer -class win32_baseboard 
$processor = get-wmiobject -computer $computer -class win32_processor
$drives = $drives = (get-wmiobject -ComputerName $computer -class win32_diskDrive | select-object @{name="Drives"; expression={$_.Model + " " + $_.Size / 1GB}}) | ft -HideTableHeaders | Out-String
$bios = get-wmiobject -computer $computer -class win32_bios
$networkAdapter = ((get-wmiobject -ComputerName $computer -class win32_networkadapter -filter "netconnectionstatus = 2" | select-object -property description, macaddress )) | ft -HideTableHeaders | out-string
$antivirus = get-wmiobject -Namespace root/SecurityCenter2 -class AntiVirusProduct -computer $computer | select-object -Property displayname -ExpandProperty displayname | out-string
$name = (Get-WmiObject -ComputerName $computer win32_computersystem).name

new-object psobject -property @{
Name = ($name)
Motherboard = ($mobo.manufacturer + " " + $mobo.product)
Processor = ($processor.Name + " " + $processor.Socketdesignation)
Drives = $drives
BIOS = ($bios.Manufacturer + $bios.name)
NetworkAdapter = $networkAdapter
Antivirus = $antivirus
}
}

$path = ("INSERT PATH TO SAVE DIR HERE"+$env:COMPUTERNAME+".csv")

Compstats | export-csv -force $path