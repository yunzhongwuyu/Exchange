<#

This script is used to grant group owner edit group member permission in ADUC.

Assign-GroupMemberEditpermission.ps1

#>

Function Assign-GroupMemberEditpermission {

    param(
        [String]$Group,
        [string]$Owner
    
    )

    Add-ADPermission -Identity $group -User $Owner -AccessRights WriteProperty -Properties "Member"

}

$Groups = Get-DistributionGroup -Filter {ManagedBy -ne $null} -ResultSize unlimited | select -ExpandProperty name

Foreach($Group in $Groups){

$GroupName = $Group

$Owners = Get-DistributionGroup -Identity $Groupname | select -ExpandProperty ManagedBy
Foreach($owner in $Owners){

    Assign-GroupMemberEditpermission -Group $Group -Owner $owner

}

}