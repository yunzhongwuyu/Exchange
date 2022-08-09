[void][System.reflection.assembly]::LoadWithPartialName('microsoft.visualbasic')

$recipient=[Microsoft.visualbasic.interaction]::inputbox("Enter the Recipient Name to delete the Meeting requests from all the room mailoxes","name","Sankar_M")


$mailid=[Microsoft.visualbasic.interaction]::inputbox(" Enter your mail id to create a foler called REPORT in your outlook in which you can find the deleted contents","name","Sankar@Domain.com")

if(get-recipient $recipient -warningaction:silentlycontinue)
{
 [Microsoft.visualbasic.interaction]::msgbox("Recipient is Existing in Production. Ensure that You are running this program for Terminated User")

Write-host "Recipient is Existing in Production. Ensure that You are running this program for Terminated User"

break 
}
else
{

[Microsoft.visualbasic.interaction]::msgbox("Hello, Please Wait Until the Program completes")

Write-Progress -Activity "Preparing" -Status "Retrieving Roommailbox list" -PercentComplete 0
$rooms=get-mailbox -recipienttypedetails roommailbox -resultsize unlimited -warningaction:silentlycontinue| where {$_.name -notlike "*test*"}
 
$count=$rooms.count




foreach($room in $rooms)

{
  
  $i=$i+1
  $percentage=$i/$count*100
  
  
  Write-Progress -Activity "Collecting mailbox details" -Status "Processing mailbox $i of $Count - $room" -PercentComplete $percentage
  

 $room | search-mailbox -searchquery "kind:meetings from:$recipient" -targetmailbox $mailid -targetfolder "REPORT" -deletecontent -force
  }}
 
