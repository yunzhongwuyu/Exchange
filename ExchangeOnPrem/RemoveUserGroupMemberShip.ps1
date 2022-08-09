$user = read-host "Write the user who need to be removed from all groups who is member of"
$Groups=Get-ADUser $User -Properties Memberof | select -ExpandProperty:Memberof | foreach {Get-ADGroup $_ | select -ExpandProperty:Name}
$groups | foreach {Remove-DistributionGroupMember -Identity $_ -Member $User -BypassSecurityGroupManagerCheck}