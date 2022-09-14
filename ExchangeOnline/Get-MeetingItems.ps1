
<#
.SYNOPSIS
Gather statistics regarding meeting room usage

.DESCRIPTION
This script uses Exchange Web Services to connect to one or more meeting rooms and gather statistics regarding their usage between to specific dates

IMPORTANT:
  - You must use the room's SMTP address;
  - You must have at least Reviewer rights to the meeting room's calendar (FullAccess to the mailbox will also work);
  - Maximum range to search is two years;
  - Maximum of 1000 meetings are returned;
  - Exchange AutoDiscover needs to be working.


.EXAMPLE
C:\PS> .\Get-MeetingItems.ps1 -MailboxListSMTP "room.1@domain.com, room.2@domain.com" -From "01/01/2017" -To "01/02/2017" -Verbose

Description
-----------
This command will:
   1. Process room.1@domain.com and room.2@domain.com meeting rooms;
   2. Gather statistics for both room between 1st of Jan and 1st of Feb (please be aware of your date format: day/month vs month/day);
   3. Write progress information as it goes along because of the -Verbose switch


.EXAMPLE
C:\PS> Get-Help .\Get-MeetingRoomStats_EWS.ps1 -Full

Description
-----------
Shows this help manual.
#>



[CmdletBinding()]
Param (
	[Parameter(Position = 0, Mandatory = $True)]
	[String] $MailboxListSMTP,

	[Parameter(Position = 1, Mandatory = $False)]
	[DateTime] $From = (Get-Date "01/01/2018" -Day 1 -Hour 0 -Minute -0 -Second 0),
	
	[Parameter(Position = 2, Mandatory = $False)]
	[DateTime] $To = ((Get-Date -Day 1 -Hour 0 -Minute -0 -Second 0).AddMonths(1).AddSeconds(-1))
)


Function Load-EWS {
	Write-Verbose "Loading EWS Managed API"
	$EWSdll = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services' | Sort Name -Descending | Select -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")

	If (Test-Path $EWSdll) {
		Try {
			Import-Module $EWSdll -ErrorAction Stop
		} Catch {
			Write-Verbose -Message "Unable to load EWS Managed API: $($_.Exception.Message). Exiting Script."
			Exit
		}
	} Else {
		Write-Verbose "EWS Managed API not installed. Please download and install the current version of the EWS Managed API from http://go.microsoft.com/fwlink/?LinkId=255472. Exiting Script."
		Exit
	}
}


Function Connect-Exchange {
	Param ([String]$Mailboxes)
	
	# Load EWS Managed API dll
	Load-EWS

	# Create Exchange Service Object and set Exchange version
	Write-Verbose "Creating Exchange Service Object using AutoDiscover"
	$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2)
	$service.AutodiscoverUrl($Mailboxes)
    #$service.Url = new-object Uri("https://exchangewebservice.domain.com/ews/exchange.asmx")

	If (!$service.URL) {
		Write-Verbose -Message "Error conneting to Exchange Web Services (no AutoDiscover URL). Exiting Script."
		Exit
	} Else {
		Return $service
	}
}



#################################################################
# Script Start
#################################################################

# Initialize some variables that will be used later in the script
[Array] $MailboxsCol = @()

# Connect to Exchange Server
$service = Connect-Exchange -Mailboxes ($MailboxListSMTP.Split(",")[0])

## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

         # create an instance of the TrustAll and attach it to the ServicePointManager
        $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
        [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

ForEach ($Mailbox in $MailboxListSMTP.Split(",") -replace (" ", "")) {
	$topOrganizers = @{}
	$topAttendees = @{}

	# Bind to the room's Calendar folder
	Try {

		Write-Verbose -Message "Binding to the $Mailbox Calendar folder."
		$folderID = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $Mailbox) -ErrorAction Stop
		$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderID)
	} Catch {
		Write-Verbose "Unable to connect to $Mailbox. Please check permissions: $($_.Exception.Message). Skipping $Mailbox."
		Continue
	}

	#Define the calendar view and properties to load (required to get attendees)
	Try {
		$psPropset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
		$CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($From, $To, 1000)    
		$fiItems = $service.FindAppointments($Calendar.Id,$CalendarView)    
		If ($fiItems.Items.Count -gt 0) {[Void] $service.LoadPropertiesForItems($fiItems, $psPropset)}
	} Catch {
		Write-Verbose "Unable to retrieve data from $Mailbox calendar. Please check permissions: $($_.Exception.Message). Skipping $Mailbox."
		Continue
	}

	[Int] $totalMeetings = $totalDuration = $totalAttendees = $totalReqAttendees = $totalOptAttendees = $totalAM = $totalPM = $totalRecurring = 0
	ForEach($meeting in $fiItems.Items) {
		
        
        # Top Organizers
		If ($meeting.Organizer -and $topOrganizers.ContainsKey($meeting.Organizer.Address)) {
			$topOrganizers.Set_Item($meeting.Organizer.Address, $topOrganizers.Get_Item($meeting.Organizer.Address) + 1)
		} Else {
			$topOrganizers.Add($meeting.Organizer.Address, 1)
		}
		
		# Top Attendees
		ForEach ($attendant in $meeting.RequiredAttendees) {
			If (!$attendant.Address) {Continue}
			If ($topAttendees.ContainsKey($attendant.Address)) {
				$topAttendees.Set_Item($attendant.Address, $topAttendees.Get_Item($attendant.Address) + 1)
			} Else {
				$topAttendees.Add($attendant.Address, 1)
			}
		}

		ForEach ($attendant in $meeting.OptionalAttendees) {
			If (!$attendant.Address) {Continue}
			If ($topAttendees.ContainsKey($attendant.Address)) {
				$topAttendees.Set_Item($attendant.Address, $topAttendees.Get_Item($attendant.Address) + 1)
			} Else {
				$topAttendees.Add($attendant.Address, 1)
			}
		}

		$totalMeetings++
		$totalDuration += $meeting.Duration.TotalMinutes
		$totalAttendees += $meeting.RequiredAttendees.Count + $meeting.OptionalAttendees.Count
		$totalReqAttendees += $meeting.RequiredAttendees.Count
		$totalOptAttendees += $meeting.OptionalAttendees.Count
		If ((Get-Date $meeting.Start -UFormat %p) -eq "AM") {$totalAM++} Else {$totalPM++}
		If ($meeting.IsRecurring) {$totalRecurring++}
	}
    

    
    $StartTime = $Meeting.Start
    $EndTime = $Meeting.End
    $AllDayEvent = $meeting.IsAllDayEvent
    $Organizer = $meeting.Organizer
    [String]$RequiredAtt = $Meeting.RequiredAttendees.Address -join ";"
    [String]$OptionalAtt = $meeting.OptionalAttendees.Address -join ";"
    $Location = $Meeting.Location
    $Priority = $Meeting.Importance
    if($Meeting.Sensitivity -eq "Private"){$Private = "Y"}Else{$Private = "N"}
    $CalendarOwner = $Mailbox
    $InviteDate = $Meeting.DateTimeCreated

	# Save the information gathered into an object and add it to our object collection
	$romObj = New-Object PSObject -Property @{
		StartTime		= $StartTime
		EndTime			= $EndTime
		AllDayEvent		= $AllDayEvent
		Organizer		= $Organizer
		RequiredAtt		= $RequiredAtt
		OptionalAtt		= $OptionalAtt
		Location		= $Location
		Priority		= $Priority
		Private			= $Private
		Owner			= $CalendarOwner
		InviteDate		= $InviteDate
	}
	
	$MailboxsCol += $romObj
}



$StartTime = $from.ToString('yyyyMMdd')

# Print and export the results
$MailboxsCol | Select StartTime, EndTime, AllDayEvent, Organizer, RequiredAtt, OptionalAtt, Location, Priority, Private, Owner, InviteDate | Sort InviteDate | Export-Csv  "D:\FGLH\O365\Calendar-$Mailbox-$StrartTime.csv" -NoTypeInformation
}
#$MailboxsCol | Select From, To, RoomEmail, RoomName, Meetings, Duration, AvgDuration, TotAttendees, AvgAttendees, RecAttendees, OptAttendees, AMtotal, AMperc, PMtotal, PMperc, RecTotal, RecPerc, TopOrg, TopAtt | Sort Date
#$MailboxsCol | Select From, To, RoomEmail, RoomName, Meetings, Duration, AvgDuration, TotAttendees, AvgAttendees, RecAttendees, OptAttendees, AMtotal, AMperc, PMtotal, PMperc, RecTotal, RecPerc, TopOrg, TopAtt | Sort Date | Export-Csv "MeetingRoomStats_$((Get-Date).ToString('yyyyMMdd')).csv" -NoTypeInformation
