Clear
$Servers = Get-MailboxServer nnexmrdk*|sort name
Write "========================Disk space check=======================" | Out-File d:\LogDiskSpaceCheck.txt
Write "`n" >> d:\qjw\LogDiskSpaceCheck.txt
$result = Foreach ($Server in $Servers) {
	$Volumes = Get-WmiObject -Class "Win32_Volume" -ComputerName $Server | where{$_.label -like "*log*"}
	Write "----------------------Disk space check on $Server-----------------------" >>d:\qjw\LogDiskSpaceCheck.txt
	$Volumes|sort freespace|ft systemname,label,name,@{Name="Capacity(GB)";Expression={[Math]::Truncate($_.capacity/1073741824)}},@{Name="freespace(GB)";Expression={[Math]::Truncate($_.freespace/1073741824)}},@{Name="% Free";Expression={[Math]::Truncate(($_.freespace/$_.capacity)*100)}} -autosize >>d:\qjw\LogDiskSpaceCheck.txt
}
