
######################################################################################################
#                                                                                                    #
# Name:        Tickle-MailRecipients.ps1                                                             #
#                                                                                                    #
# Version:     1.0                                                                                   #
#                                                                                                    #
# Description: Address Lists in Exchange Online do not automatically populate during provisioning    #
#              and there is no "Update-AddressList" cmdlet.  This script "tickles" mailboxes, mail   #
#              users and distribution groups so the Address List populates.                          #
#                                                                                                    #
# Author:      Joseph Palarchio                                                                      #
#                                                                                                    #
# Usage:       Additional information on the usage of this script can found at the following         #
#              blog post:  http://blogs.perficient.com/microsoft/?p=25536                            #
#                                                                                                    #
# Disclaimer:  This script is provided AS IS without any support. Please test in a lab environment   #
#              prior to production use.                                                              #
#                                                                                                    #
######################################################################################################

$mailboxes = Get-Mailbox -Resultsize Unlimited
$count = $mailboxes.count
$i=0

Write-Host
Write-Host "Mailboxes Found:" $count

foreach($mailbox in $mailboxes){
  $i++
  Set-Mailbox $mailbox.alias -SimpleDisplayName $mailbox.SimpleDisplayName -WarningAction silentlyContinue
  Write-Progress -Activity "Tickling Mailboxes [$count]..." -Status $i
}

$mailusers = Get-MailUser -Resultsize Unlimited
$count = $mailusers.count
$i=0

Write-Host
Write-Host "Mail Users Found:" $count

foreach($mailuser in $mailusers){
  $i++
  Set-MailUser $mailuser.alias -SimpleDisplayName $mailuser.SimpleDisplayName -WarningAction silentlyContinue
  Write-Progress -Activity "Tickling Mail Users [$count]..." -Status $i
}

$distgroups = Get-DistributionGroup -Resultsize Unlimited
$count = $distgroups.count
$i=0

Write-Host
Write-Host "Distribution Groups Found:" $count

foreach($distgroup in $distgroups){
  $i++
  Set-DistributionGroup $distgroup.alias -SimpleDisplayName $distgroup.SimpleDisplayName -WarningAction silentlyContinue
  Write-Progress -Activity "Tickling Distribution Groups. [$count].." -Status $i
}

Write-Host
Write-Host "Tickling Complete"