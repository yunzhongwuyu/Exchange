#
# Remove-DuplicateAppointments.ps1
#
# By David Barrett, Microsoft Ltd. 2015. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

param (
	[Parameter(Position=0,Mandatory=$False,HelpMessage="Specifies the mailbox to be accessed")]
	[ValidateNotNullOrEmpty()]
	[string]$Mailbox,

	[Parameter(Mandatory=$False,HelpMessage="Folder to search - if omitted, the mailbox calendar folder is assumed")]
	[string]$FolderPath,

	[Parameter(Mandatory=$False,HelpMessage="Folder to which any duplicates will be moved.  If not specified, duplicate items are soft deleted (will go to Deleted Items folder)")]
	[string]$DuplicatesTargetFolder,

	[Parameter(Mandatory=$False,HelpMessage="If this switch is present, folder path is required and the path points to a public folder")]
	[switch]$PublicFolders,

	[Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")]
    [System.Management.Automation.PSCredential]$Credentials,
				
	[Parameter(Mandatory=$False,HelpMessage="Username used to authenticate with EWS")]
	[string]$Username,
	
	[Parameter(Mandatory=$False,HelpMessage="Password used to authenticate with EWS")]
	[string]$Password,
	
	[Parameter(Mandatory=$False,HelpMessage="Domain used to authenticate with EWS")]
	[string]$Domain,
	
	[Parameter(Mandatory=$False,HelpMessage="Whether we are using impersonation to access the mailbox")]
	[switch]$Impersonate,
	
	[Parameter(Mandatory=$False,HelpMessage="EWS Url (if omitted, then autodiscover is used)")]	
	[string]$EwsUrl,
	
	[Parameter(Mandatory=$False,HelpMessage="Path to managed API (if omitted, a search of standard paths is performed)")]	
	[string]$EWSManagedApiPath = "",
	
	[Parameter(Mandatory=$False,HelpMessage="Whether to ignore any SSL errors (e.g. invalid certificate)")]	
	[switch]$IgnoreSSLCertificate,
	
	[Parameter(Mandatory=$False,HelpMessage="Whether to allow insecure redirects when performing autodiscover")]	
	[switch]$AllowInsecureRedirection,
	
	[Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified")]	
	[string]$LogFile = "",
	
	[Parameter(Mandatory=$False,HelpMessage="Do not apply any changes, just report what would be updated")]	
	[switch]$WhatIf	
)


# Define our functions

Function Log([string]$Details, [ConsoleColor]$Colour)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
	Write-Host $Details -ForegroundColor $Colour
	if ( $LogFile -eq "" ) { return	}
	$Details | Out-File $LogFile -Append
}

Function LoadEWSManagedAPI()
{
	# Find and load the managed API
	
	if ( ![string]::IsNullOrEmpty($EWSManagedApiPath) )
	{
		if ( Test-Path $EWSManagedApiPath )
		{
			Add-Type -Path $EWSManagedApiPath
			return $true
		}
		Write-Host ( [string]::Format("Managed API not found at specified location: {0}", $EWSManagedApiPath) ) -ForegroundColor Yellow
	}
	
	$a = Get-ChildItem -Recurse "C:\Program Files (x86)\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) }
	if (!$a)
	{
		$a = Get-ChildItem -Recurse "C:\Program Files\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) }
	}
	
	if ($a)	
	{
		# Load EWS Managed API
		Write-Host ([string]::Format("Using managed API {0} found at: {1}", $a.VersionInfo.FileVersion, $a.VersionInfo.FileName)) -ForegroundColor Gray
		Add-Type -Path $a.VersionInfo.FileName
		return $true
	}
	return $false
}

Function CurrentUserPrimarySmtpAddress()
{
    # Attempt to retrieve the current user's primary SMTP address
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
    $result = $searcher.FindOne()

    if ($result -ne $null)
    {
        $mail = $result.Properties["mail"]
        return $mail
    }
    return $null
}

Function TrustAllCerts() {
    <#
    .SYNOPSIS
    Set certificate trust policy to trust self-signed certificates (for test servers).
    #>

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
        namespace Local.ToolkitExtensions.Net.CertificatePolicy {
        public class TrustAll : System.Net.ICertificatePolicy {
            public TrustAll()
            { 
            }
            public bool CheckValidationResult(System.Net.ServicePoint sp,
                                                System.Security.Cryptography.X509Certificates.X509Certificate cert, 
                                                System.Net.WebRequest req, int problem)
            {
                return true;
            }
        }
        }
'@ 
    $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
    $TAAssembly=$TAResults.CompiledAssembly

    ## We now create an instance of the TrustAll and attach it to the ServicePointManager
    $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
    [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
}

function CreateService($targetMailbox)
{
    # Creates and returns an ExchangeService object to be used to access mailboxes
    $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)

    # Set credentials if specified, or use logged on user.
    if ($Credentials -ne $Null)
    {
        Write-Verbose "Applying given credentials"
        $exchangeService.Credentials = $Credentials.GetNetworkCredential()
    }
    elseif ($Username -and $Password)
    {
	    Write-Verbose "Applying given credentials for $Username"
	    if ($Domain)
	    {
		    $exchangeService.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password,$Domain)
	    } else {
		    $exchangeService.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password)
	    }
    }
    else
    {
	    Write-Verbose "Using default credentials"
        $exchangeService.UseDefaultCredentials = $true
    }

    # Set EWS URL if specified, or use autodiscover if no URL specified.
    if ($EwsUrl)
    {
    	$exchangeService.URL = New-Object Uri($EwsUrl)
    }
    else
    {
    	try
    	{
		    Write-Verbose "Performing autodiscover for $targetMailbox"
		    if ( $AllowInsecureRedirection )
		    {
			    $exchangeService.AutodiscoverUrl($targetMailbox, {$True})
		    }
		    else
		    {
			    $exchangeService.AutodiscoverUrl($targetMailbox)
		    }
		    if ([string]::IsNullOrEmpty($exchangeService.Url))
		    {
			    Log "$targetMailbox : autodiscover failed" Red
			    return $Null
		    }
		    Write-Verbose "EWS Url found: $($exchangeService.Url)"
    	}
    	catch
    	{
    	}
    }
 
    if ($Impersonate)
    {
		$exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $targetMailbox)
	}

    return $exchangeService
}

function SearchForDuplicates($folder)
{
    # Search the folder for duplicate appointments
    # We read all the items in the folder, and build a list of all the duplicates

    $subject = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
    $icaluids = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
    $duplicateItems = @()
    $itemCount = $folder.TotalCount
    $processedCount = 0


    $offset = 0
    $moreItems = $true
    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(500, 0)
    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
    $propset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start)
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End)
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::ICalUid)
    $view.PropertySet = $propset

    Write-Progress -Activity "Reading all folder items" -Status "0 of $itemCount items read" -PercentComplete 0

    while ($moreItems)
    {
        $results = $Folder.FindItems($view)
        $moreItems = $results.MoreAvailable
        $view.Offset = $results.NextPageOffset
        foreach ($item in $results)
        {
            Write-Verbose "Processing: $($item.Subject)"
            $isDupe = $False
            if ($icaluids.ContainsKey($item.ICalUid))
            {
                # Duplicate ICalUid exists
                $duplicateItems += $item
                $isDupe = $True
            }
            else
            {
                $icaluids.Add($item.ICalUid, $item.Id.UniqueId)

                $subject_cmp = $item.Subject
                if ([String]::IsNullOrEmpty($subject_cmp))
                {
                    $subject_cmp = "[No Subject]" # If the subject is blank, we need to give it an arbitrary value to prevent checks failing
                }
                if ($subject.ContainsKey($subject_cmp))
                {
                    # Duplicate subject exists, so we now check the start and end date to confirm if this is a duplicate
                    $dupSubjectList = $subject[$subject_cmp]
                    foreach ($dupSubject in $dupSubjectList)
                    {
                        if (($dupSubject.Start -eq $item.Start) -and ($dupSubject.End -eq $item.End))
                        {
                            # Same subject, start, and end date, so this is a duplicate
                            $duplicateItems += $item
                            $isDupe = $True
                        }
                    }
                    if (!$isDupe)
                    {
                        # Add this item to the list of items with the same subject (as it is not a duplicate)
                        $subject[$subject_cmp] += $item
                    }
                }
                else
                {
                    # Add this to our subject list
                    $subject.Add($subject_cmp, @($item))
                }
            }

            $processedCount++
            Write-Progress -Activity "Reading all folder items" -Status "$processedCount of $itemCount items read" -PercentComplete ($processedCount/$itemCount)
        }
    }
    Write-Progress -Activity "Reading all folder items" -Completed

    if ($duplicateItems.Count -eq 0)
    {
        Log "No duplicate items found!" Green
        return
    }
    Log ([string]::Format("{0} duplicates found", $duplicateItems.Count)) Green

    # We now have a list of duplicate items, so we can process them

    $action = "delet"
    if ($script:targetFolder -ne $null)
    {
        $action = "mov"
    }

    if (!$WhatIf)
    {
	    # Delete (or move) the items (we will do this in batches of 500)
	    $itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
	    $itemIdType = [Type] $itemId.GetType()
	    $baseList = [System.Collections.Generic.List``1]
	    $genericItemIdList = $baseList.MakeGenericType(@($itemIdType))
	    $deleteIds = [Activator]::CreateInstance($genericItemIdList)
        $pauseForThrottling = $False
	    ForEach ($dupe in $duplicateItems)
	    {



            Log ([string]::Format("{1}ing: {0}", $dupe.Subject, $action)) Gray
		    $deleteIds.Add($dupe.Id)
		    if ($deleteIds.Count -ge 500)
		    {
			    # Send the delete request
                try
                {
                    if ($script:targetFolder -ne $null)
                    {
                        [void]$script:service.MoveItems( $deleteIds, $script:targetFolder.Id )
                    }
                    else
                    {
			            [void]$script:service.DeleteItems( $deleteIds, [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, $null )
                    }
                    if ($pauseForThrottling)
                    {
                        Start-Sleep -s 1
                    }
                }
                catch
                {
                    if ($Error[0].Contains("Try again later"))
                    {
                        # Most likely we've been throttled, so we'll wait before sending any more requests
                        Write-Host "Pausing for ten seconds to allow for throttling limits" -ForegroundColor DarkYellow
                        Start-Sleep -s 10
                        $pauseForThrottling = $True
                    }
                }
                Write-Verbose ([string]::Format("{0} items {1}ed", $deleteIds.Count, $action))
			    $deleteIds = [Activator]::CreateInstance($genericItemIdList)
		    }
	    }
	    if ($deleteIds.Count -gt 0)
	    {
            if ($script:targetFolder -ne $null)
            {
                [void]$script:service.MoveItems( $deleteIds, $script:targetFolder.Id )
            }
            else
            {
			    [void]$script:service.DeleteItems( $deleteIds, [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, $null )
            }
	    }
	    Write-Verbose ([string]::Format("{0} items {1}ed", $deleteIds.Count, $action))
    }
    else
    {
        # We aren't actually deleting, so just report what we would delete
	    ForEach ($dupe in $duplicateItems)
	    {
            if ([String]::IsNullOrEmpty($dupe.Subject))
            {
                Log "Would $($action)e: [No Subject]" Gray
            }
            else
            {
                Log ([string]::Format("Would {1}e: {0}", $dupe.Subject, $action)) Gray
            }
        }
    }
}

Function GetFolder()
{
	# Return a reference to a folder specified by path

    $FolderPath, $Create = $args[0]
	
    if ($PublicFolders)
    {
        $mbx = ""
        $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)
    }
    else
    {
		$mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
		$folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mbx )
	    $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId)
    }

	if ($FolderPath -ne '\')
	{
		$PathElements = $FolderPath -split '\\'
		For ($i=0; $i -lt $PathElements.Count; $i++)
		{
			if ($PathElements[$i])
			{
				$View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0)
				$View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
						
				$SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i])
				
				$FolderResults = $Folder.FindFolders($SearchFilter, $View)
				if ($FolderResults.TotalCount -gt 1)
				{
					# We have more than one folder returned (this should never actually happen)
					$Folder = $null
					Write-Host "Failed to find $($PathElements[$i]), path requested was $FolderPath" -ForegroundColor Red
					break
				}
                elseif ($FolderResults.TotalCount -eq 0)
                {
                    if (!$Create)
                    {
					    $Folder = $null
					    Write-Host "Folder $($PathElements[$i]) doesn't exist, path requested was $FolderPath" -ForegroundColor Red
					    break
                    }
                    # Attempt to create the folder
					$subfolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($Folder.Service)
					$subfolder.DisplayName = $PathElements[$i]
                    $subfolder.FolderClass = "IPF.Appointment"
                    try
                    {
					    $subfolder.Save($Folder.Id)
                        Write-Host "Created folder $($PathElements[$i])"
                    }
                    catch
                    {
					    # Failed to create the subfolder
					    $subfolder = $null
					    Log "Failed to create folder $($PathElements[$i]) in path $FolderPath" Red
					    break
                    }
                    $Folder = $subfolder
                }
                else
				{
				    $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $FolderResults.Folders[0].Id)
                }
			}
		}
	}
	
	return $Folder
}

function ProcessMailbox()
{
    # Process the mailbox
    Write-Host ([string]::Format("Processing mailbox {0}", $Mailbox)) -ForegroundColor Gray
	$script:service = CreateService($Mailbox)
	if ($script:service -eq $Null)
	{
		Write-Host "Failed to create ExchangeService" -ForegroundColor Red
	}
	
    $Folder = $Null
	if ($FolderPath)
	{
		$Folder = GetFolder($FolderPath, $False)
		if (!$Folder)
		{
			Write-Host "Failed to find folder $FolderPath" -ForegroundColor Red
			return
		}
	}
    else
    {
        if ($PublicFolders)
        {
            Write-Host "You must specify folder path when searching public folders" -ForegroundColor Red
            return
        }
        else
        {
		    $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
		    $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $mbx )
	        $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId)
        }
    }

    if (![string]::IsNullOrEmpty($DuplicatesTargetFolder))
    {
        # We are moving any duplicates to a specific folder (instead of deleting), so find that folder
        $script:targetFolder = GetFolder($DuplicatesTargetFolder, $True)
		if (!$script:targetFolder)
		{
			Write-Host "Failed to find target folder for duplicates $($script:targetFolder)" -ForegroundColor Red
			return
		}
        Write-Host "Will move duplicate items to folder $($script:targetFolder.DisplayName)" -ForegroundColor Yellow
    }

	SearchForDuplicates $Folder
}


# The following is the main script

if ( [string]::IsNullOrEmpty($Mailbox) )
{
    $Mailbox = CurrentUserPrimarySmtpAddress
    if ( [string]::IsNullOrEmpty($Mailbox) )
    {
	    Write-Host "Mailbox not specified.  Failed to determine current user's SMTP address." -ForegroundColor Red
	    Exit
    }
    else
    {
        Write-Host ([string]::Format("Current user's SMTP address is {0}", $Mailbox)) -ForegroundColor Green
    }
}

# Check if we need to ignore any certificate errors
# This needs to be done *before* the managed API is loaded, otherwise it doesn't work consistently (i.e. usually doesn't!)
if ($IgnoreSSLCertificate)
{
	Write-Host "WARNING: Ignoring any SSL certificate errors" -foregroundColor Yellow
    TrustAllCerts
}
 
# Load EWS Managed API
if (!(LoadEWSManagedAPI))
{
	Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Red
	Exit
}

# Check we have valid credentials
if ($Credentials -ne $Null)
{
    If ($Username -or $Password)
    {
        Write-Host "Please specify *either* -Credentials *or* -Username and -Password" Red
        Exit
    }
}

  

Write-Host ""

# Check whether we have a CSV file as input...
$FileExists = Test-Path $Mailbox
If ( $FileExists )
{
	# We have a CSV to process
	$csv = Import-CSV $Mailbox
	foreach ($entry in $csv)
	{
		$Mailbox = $entry.PrimarySmtpAddress
		if ( [string]::IsNullOrEmpty($Mailbox) -eq $False )
		{
			ProcessMailbox
		}
	}
}
Else
{
	# Process as single mailbox
	ProcessMailbox
}