$MailboxName = "user@domain.com"
$StartDate = new-object System.DateTime(2009, 08, 01)
$EndDate = new-object System.DateTime(2009, 11, 01)

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)
$CalendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($service,$folderid)
$cvCalendarview = new-object Microsoft.Exchange.WebServices.Data.CalendarView($StartDate,$EndDate,2000)
$cvCalendarview.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$frCalendarResult = $CalendarFolder.FindAppointments($cvCalendarview)
foreach ($apApointment in $frCalendarResult.Items){
     $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
     $apApointment.load($psPropset)
    "Appointment : " + $apApointment.Subject.ToString() 
    "Start : " + $apApointment.Start.ToString()
    "End : " + $apApointment.End.ToString()
    "Organizer : " + $apApointment.Organizer.ToString()
    "Required Attendees :"
    foreach($attendee in $apApointment.RequiredAttendees){
		"	" + $attendee.Address
	}
    "Optional Attendees :"
     foreach($attendee in $apApointment.OptionalAttendees){
		"	" + $attendee.Address
     }
    "Resources :"
     foreach($attendee in $apApointment.Resources){
		"	" + $attendee.Address
     }
    " "
}