Clear
$Servers = Get-MailboxServer *|sort name
Out-File d:\Scripts\NeedToCleanLogDisk.txt
$result = Foreach ($Server in $Servers) {
	$Volumes = Get-WmiObject -Class "Win32_Volume" -ComputerName $Server | where{$_.label -like "*log*"} | Where { $_.Freespace -lt "64424509440"}
	foreach ($Volume in $Volumes) {$Volume.label >> d:\Scripts\NeedToCleanLogDisk.txt}
}
$Disks = get-content d:\Scripts\NeedToCleanLogDisk.txt
If ($Disks -gt 0){
(get-content d:\Scripts\NeedToCleanLogDisk.txt) | foreach {$_ -replace "Logs",""} | set-content d:\Scripts\NeedToCleanLogDisk.txt
Write-Host -ForegroundColor Green "=====================================DBs which log disk Free space < 60GB====================================="
Foreach ($Server in $Servers) {
Foreach($Disk in $Disks) {
	$Volumes = Get-WmiObject -Class "Win32_Volume" -ComputerName $Server |Where {$_.Label -eq $Disk}
	$Volumes|select systemname,label,name,@{Name="Capacity(GB)";Expression={[Math]::Truncate($_.capacity/1073741824)}},@{Name="freespace(GB)";Expression={[Math]::Truncate($_.freespace/1073741824)}},@{Name="% Free";Expression={[Math]::Truncate(($_.freespace/$_.capacity)*100)}} |sort "% Free"|ft -autosize
}
}
Write-Host -ForegroundColor Green "=====================================Enable Circular Logging for the DBs====================================="
(get-content d:\Scripts\NeedToCleanLogDisk.txt) | foreach {Set-mailboxdatabase $_ -CircularLoggingEnabled:$true}
(get-content d:\Scripts\NeedToCleanLogDisk.txt) | foreach {get-mailboxdatabase $_ |ft name, server, CircularLoggingEnabled}
for ($i = 0; $i -lt 100; $i++) { Write-Progress -Activity "Wait for 15 minutes to clean logs, then will disable Circular Logging...." -Status "Status: $i" -PercentComplete $i; Start-Sleep -Milliseconds 9000}
Write-Host -ForegroundColor Green "=====================================Disable Circular Logging for the DBs====================================="
(get-content d:\Scripts\NeedToCleanLogDisk.txt) | foreach {Set-mailboxdatabase $_ -CircularLoggingEnabled:$false}
(get-content d:\Scripts\NeedToCleanLogDisk.txt) | foreach {get-mailboxdatabase $_ |ft name, server, CircularLoggingEnabled}
Foreach ($Server in $Servers) {
Foreach($Disk in $Disks) {
	$Volumes = Get-WmiObject -Class "Win32_Volume" -ComputerName $Server |Where {$_.Label -eq $Disk}
	$Volumes|select systemname,label,name,@{Name="Capacity(GB)";Expression={[Math]::Truncate($_.capacity/1073741824)}},@{Name="freespace(GB)";Expression={[Math]::Truncate($_.freespace/1073741824)}},@{Name="% Free";Expression={[Math]::Truncate(($_.freespace/$_.capacity)*100)}} |sort "% Free"|ft -autosize
}
}
}
Else {Write-Host -ForegroundColor Green "ALL DB log disk free space are above 30GB now"}