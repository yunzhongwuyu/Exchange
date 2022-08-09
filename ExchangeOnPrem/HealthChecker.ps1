﻿<#
.NOTES
	Name: HealthChecker.ps1
	Original Author: Marc Nivens
    Author: David Paulson
    Contributor: Jason Shinbaum 
	Contributor: Michael Schatte
	Requires: Exchange Management Shell and administrator rights on the target Exchange
	server as well as the local machine.
	Version History:
	1.31 - 9/21/2016
	3/30/2015 - Initial Public Release.
    1/18/2017 - Initial Public Release of version 2. - rewritten by David Paulson.
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
	Checks the target Exchange server for various configuration recommendations from the Exchange product group.
.DESCRIPTION
	This script checks the Exchange server for various configuration recommendations outlined in the 
	"Exchange 2013 Performance Recommendations" section on TechNet, found here:

	https://technet.microsoft.com/en-us/library/dn879075(v=exchg.150).aspx

	Informational items are reported in Grey.  Settings found to match the recommendations are
	reported in Green.  Warnings are reported in yellow.  Settings that can cause performance
	problems are reported in red.  Please note that most of these recommendations only apply to Exchange
	2013/2016.  The script will run against Exchange 2010/2007 but the output is more limited.
.PARAMETER Server
	This optional parameter allows the target Exchange server to be specified.  If it is not the 		
	local server is assumed.
.PARAMETER OutputFilePath
	This optional parameter allows an output directory to be specified.  If it is not the local 		
	directory is assumed.  This parameter must not end in a \.  To specify the folder "logs" on 		
	the root of the E: drive you would use "-OutputFilePath E:\logs", not "-OutputFilePath E:\logs\".
.PARAMETER MailboxReport
	This optional parameter gives a report of the number of active and passive databases and
	mailboxes on the server.
.PARAMETER LoadBalancingReport
    This optional parameter will check the connection count of the Default Web Site for every server
    running Exchange 2013/2016 with the Client Access role in the org.  It then breaks down servers by percentage to 
    give you an idea of how well the load is being balanced.
.PARAMETER CasServerList
    Used with -LoadBalancingReport.  A comma separated list of CAS servers to operate against.  Without 
    this switch the report will use all 2013/2016 Client Access servers in the organization.
.PARAMETER SiteName
	Used with -LoadBalancingReport.  Specifies a site to pull CAS servers from instead of querying every server
    in the organization.
.PARAMETER XMLDirectoryPath
    Used in combination with BuildHtmlServersReport switch for the location of the HealthChecker XML files for servers 
    which you want to be included in the report. Default location is the current directory.
.PARAMETER BuildHtmlServersReport 
    Switch to enable the script to build the HTML report for all the servers XML results in the XMLDirectoryPath location.
.PARAMETER HtmlReportFile 
    Name of the HTML output file from the BuildHtmlServersReport. Default is ExchangeAllServersReport.html
.PARAMETER DCCoreRatio 
    Gathers the Exchange to DC/GC Core ratio and displays the results in the current site that the script is running in.
.PARAMETER Verbose	
	This optional parameter enables verbose logging.
.EXAMPLE
	.\HealthChecker.ps1 -Server SERVERNAME
	Run against a single remote Exchange server
.EXAMPLE
	.\HealthChecker.ps1 -Server SERVERNAME -MailboxReport -Verbose
	Run against a single remote Exchange server with verbose logging and mailbox report enabled.
.EXAMPLE
    Get-ExchangeServer | ?{$_.AdminDisplayVersion -Match "^Version 15"} | %{.\HealthChecker.ps1 -Server $_.Name}
    Run against all Exchange 2013/2016 servers in the Organization.
.EXAMPLE
    .\HealthChecker.ps1 -LoadBalancingReport
    Run a load balancing report comparing all Exchange 2013/2016 CAS servers in the Organization.
.EXAMPLE
    .\HealthChecker.ps1 -LoadBalancingReport -CasServerList CAS01,CAS02,CAS03
    Run a load balancing report comparing servers named CAS01, CAS02, and CAS03.
.LINK
    https://technet.microsoft.com/en-us/library/dn879075(v=exchg.150).aspx
    https://technet.microsoft.com/en-us/library/36184b2f-4cd9-48f8-b100-867fe4c6b579(v=exchg.150)#BKMK_Prereq
#>
[CmdletBinding()]
param(
    #Default to use the local computer 
    [string]$Server=($env:COMPUTERNAME),
    [ValidateScript({-not $_.ToString().EndsWith('\')})]$OutputFilePath = ".",
    [switch]$MailboxReport,
    [switch]$LoadBalancingReport,
    $CasServerList = $null,
    $SiteName = $null,
    [ValidateScript({-not $_.ToString().EndsWith('\')})]$XMLDirectoryPath = ".",
    [switch]$BuildHtmlServersReport,
    [string]$HtmlReportFile="ExchangeAllServersReport.html",
    [switch]$DCCoreRatio
)

<#
Note to self. "New Release Update" are functions that i need to update when a new release of Exchange is published
#>

$healthCheckerVersion = "2.25"
$VirtualizationWarning = @"
Virtual Machine detected.  Certain settings about the host hardware cannot be detected from the virtual machine.  Verify on the VM Host that: 

    - There is no more than a 1:1 Physical Core to Virtual CPU ratio (no oversubscribing)
    - If Hyper-Threading is enabled do NOT count Hyper-Threaded cores as physical cores
    - Do not oversubscribe memory or use dynamic memory allocation
    
Although Exchange technically supports up to a 2:1 physical core to vCPU ratio, a 1:1 ratio is strongly recommended for performance reasons.  Certain third party Hyper-Visors such as VMWare have their own guidance.  VMWare recommends a 1:1 ratio.  Their guidance can be found at https://www.vmware.com/files/pdf/Exchange_2013_on_VMware_Best_Practices_Guide.pdf.  For further details, please review the virtualization recommendations on TechNet at https://technet.microsoft.com/en-us/library/36184b2f-4cd9-48f8-b100-867fe4c6b579(v=exchg.150)#BKMK_Prereq.  Related specifically to VMWare, if you notice you are experiencing packet loss on your VMXNET3 adapter, you may want to review the following article from VMWare:  http://kb.vmware.com/selfservice/microsites/search.do?language=en_US&cmd=displayKC&externalId=2039495. 

"@

#this is to set the verbose information to a different color 
if($PSBoundParameters["Verbose"]){
    #Write verose output in cyan since we already use yellow for warnings 
    $Script:VerboseEnabled = $true
    $VerboseForeground = $Host.PrivateData.VerboseForegroundColor #ToDo add a way to add the default setings back 
    $Host.PrivateData.VerboseForegroundColor = "Cyan"
}

$oldErrorAction = $ErrorActionPreference
$ErrorActionPreference = "Stop"

try{

#Enums and custom data types 
Add-Type -TypeDefinition @"
    namespace HealthChecker
    {
        public class HealthExchangeServerObject
        {
            public string ServerName;        //String of the server that we are working with 
            public HardwareObject HardwareInfo;  // Hardware Object Information 
            public OperatingSystemObject  OSVersion; // OS Version Object Information 
            public NetVersionObject NetVersionInfo; //.net Framework object information 
            public ExchangeInformationObject ExchangeInformation; //Detailed Exchange Information 

        }

        public class ExchangeInformationObject 
        {
            public ServerRole ExServerRole;          // Roles that are currently installed - Exchange 2013 makes a note if both roles aren't installed 
            public ExchangeVersion ExchangeVersion;  //Exchange Version (Exchange 2010/2013/2016)
            public string ExchangeFriendlyName;       // Friendly Name is provided 
            public string ExchangeBuildNumber;       //Exchange Build number 
            public string BuildReleaseDate;           //Provides the release date for which the CU they are currently on 
            public object ExchangeServerObject;      //Stores the Get-ExchangeServer Object 
            public bool SupportedExchangeBuild;      //Deteremines if we are within the correct build of Exchange 
            public bool InbetweenCUs;                //bool to provide if we are between main releases of CUs. Hotfixes/IUs. 
            public bool RecommendedNetVersion; //RecommendNetVersion Info includes all the factors. Windows Version & CU. 
            public ExchangeBuildObject ExchangeBuildObject; //Store the build object
            public System.Array KBsInstalled;         //Stored object for IU or Security KB fixes 
            public bool MapiHttpEnabled; //Stored from ogranzation config 
            public string MapiFEAppGCEnabled; //to determine if we were able to get information regarding GC mode being enabled or not
            public string ExchangeServicesNotRunning; //Contains the Exchange services not running by Test-ServiceHealth 
           
        }

        public class ExchangeInformationTempObject 
        {
            public string FriendlyName;    //String of the friendly name of the Exchange version 
            public bool Error;             //To report back an error and address how to handle it
            public string ExchangeBuildNumber;  //Exchange full build number 
            public string ReleaseDate;        // The release date of that version of Exchange 
            public bool SupportedCU;          //bool to determine if we are on a supported build of Exchange 
            public bool InbetweenCUs;         //Bool to determine if we are inbetween CUs. FIU/Hotfixes 
            public ExchangeBuildObject ExchangeBuildObject; //Holds the Exchange Build Object for debugging and function use reasons 
        }

        public class ExchangeBuildObject
        {
            public ExchangeVersion ExchangeVersion;  //enum for Exchange 2010/2013/2016 
            public ExchangeCULevel CU;               //enum for the CU value 
            public bool InbetweenCUs;                //bool for if we are between CUs 
        }

        //enum for CU levels of Exchange
        //New Release Update 
        public enum ExchangeCULevel
        {
            Unknown,
            Preview,
            RTM,
            CU1,
            CU2,
            CU3,
            CU4,
            CU5,
            CU6,
            CU7,
            CU8,
            CU9,
            CU10,
            CU11,
            CU12,
            CU13,
            CU14,
            CU15,
            CU16,
            CU17,
            CU18,
            CU19,
            CU20,
            CU21

        }

        //enum for the server roles that the computer is 
        public enum ServerRole
        {
            MultiRole,
            Mailbox,
            ClientAccess,
            Hub,
            Edge,
            None
        }
        
        public class NetVersionObject 
        {
            public NetVersion NetVersion; //NetVersion value 
            public string FriendlyName;  //string of the friendly name 
            public bool SupportedVersion; //bool to determine if the .net framework is on a supported build for the version of Exchange that we are running 
            public string DisplayWording; //used to display what is going on
            public int NetRegValue; //store the registry value 
        }

        public class NetVersionCheckObject
        {
            public bool Error;         //bool for error handling 
            public bool Supported;     //to provide if we are supported or not. This should throw a red warning if false 
            public bool RecommendedNetVersion;  //Bool to determine if there is a recommended version that we should be on instead of the supported version 
            public string DisplayWording;   //string value to display what is wrong with the .NET version that we are on. 
        }

        //enum for the dword value of the .NET frame 4 that we are on 
        public enum NetVersion 
        {

            Unknown = 0,
            Net4d5 = 378389,
			Net4d5d1 = 378675,
			Net4d5d2 = 379893,
			Net4d5d2wFix = 380035,
			Net4d6 = 393297,
			Net4d6d1 = 394271,
            Net4d6d1wFix = 394294,
			Net4d6d2 = 394806,
            Net4d7 = 460805,
            Net4d7d1 = 461310,
            Net4d7d2 = 461814
        }

        public class HardwareObject
        {
            public string Manufacturer; //String to display the hardware information 
            public ServerType ServerType; //Enum to determine if the hardware is VMware, HyperV, Physical, or Unknown 
            public double TotalMemory; //Stores the total memory available 
            public object System;   //objec to store the system information that we have collected 
            public ProcessorInformationObject Processor;   //Detailed processor Information 
            public bool AutoPageFile; //True/False if we are using a page file that is being automatically set 
            public string Model; //string to display Model 
            
        }

        //enum for the type of computer that we are
        public enum ServerType
        {
            VMWare,
            HyperV,
            Physical,
            Unknown
        }

        public class ProcessorInformationObject 
        {
            public int NumberOfPhysicalCores;    //Number of Physical cores that we have 
            public int NumberOfLogicalProcessors;  //Number of Logical cores that we have presented to the os 
            public int NumberOfProcessors; //Total number of processors that we have in the system 
            public int MaxMegacyclesPerCore; //Max speed that we can get out of the cores 
            public int CurrentMegacyclesPerCore; //Current speed that we are using the cores at 
            public bool ProcessorIsThrottled;  //True/False if we are throttling our processor 
            public string ProcessorName;    //String of the processor name 
            public object Processor;        // object to store the processor information 
            public bool DifferentProcessorsDetected; //true/false to detect if we have different processor types detected 
			public int EnvProcessorCount; //[system.environment]::processorcount 
            
        }

        public class OperatingSystemObject 
        {
            public OSVersionName  OSVersion; //enum for the version name 
            public string OSVersionBuild;    //string to hold the build number 
            public string OperatingSystemName; //string for the OS version friendly name
            public object OperatingSystem;   //object to store the OS information that we pulled 
            public bool HighPerformanceSet;  //True/False for the power plan setting being set correctly 
            public string PowerPlanSetting; //string value for the power plan setting being set correctly 
            public object PowerPlan;       // object to store the power plan information 
            public System.Array NetworkAdapters; //array to keep all the nics on the servers 
            public double TCPKeepAlive;       //value used for the TCP/IP keep alive setting 
            public System.Array HotFixes; //array to keep all the hotfixes of the server
            public System.Array HotFixInfo;     //objec to store hotfix information
			public string HttpProxy;
            public PageFileObject PageFile;
            public ServerLmCompatibilityLevel LmCompat;
            public bool ServerPendingReboot; //bool to determine if a server is pending a reboot to properly apply fixes

        }

        public class HotfixObject
        {
            public string KBName; //KB that we are using to check against 
            public System.Array FileInformation; //store FileVersion information
            public bool ValidFileLevelCheck;  
        }

        public class FileVersionCheckObject 
        {
            public string FriendlyFileName;
            public string FullPath; 
            public string BuildVersion;
        }

        public class NICInformationObject 
        {
            public string Description;  //Friendly name of the adapter 
            public string LinkSpeed;    //speed of the adapter 
            public string DriverDate;   // date of the driver that is currently installed on the server 
            public string DriverVersion; // version of the driver that we are on 
            public string RSSEnabled;  //bool to determine if RSS is enabled 
            public string Name;        //name of the adapter 
            public object NICObject; //objec to store the adapter info 
             
        }

        //enum for the Exchange version 
        public enum ExchangeVersion
        {
            Unknown,
            Exchange2010,
            Exchange2013,
            Exchange2016,
            Exchange2019
        }

        //enum for the OSVersion that we are
        public enum OSVersionName
        {
            Unknown,
            Windows2008, 
            Windows2008R2,
            Windows2012,
            Windows2012R2,
            Windows2016,
            Windows2019
        }

        public class PageFileObject 
        {
            public object PageFile;  //object to store the information that we got for the page file 
            public double MaxPageSize; //value to hold the information of what our page file is set to 
        }

        public class ServerLmCompatibilityLevel
        {
            public int LmCompatibilityLevel;  //The LmCompatibilityLevel for the server (INT 1 - 5)
            public string LmCompatibilityLevelDescription; //The description of the lmcompat that the server is set too
            public string LmCompatibilityLevelRef; //The URL for the LmCompatibilityLevel technet (https://technet.microsoft.com/en-us/library/cc960646.aspx)
        }
    }

"@

}

catch {
    Write-Warning "There was an error trying to add custom classes to the current PowerShell session. You need to close this session and open a new one to have the script properly work."
    sleep 5
    exit 
}

finally {
    $ErrorActionPreference = $oldErrorAction
}

##################
#Helper Functions#
##################

#Output functions
function Write-Red($message)
{
    Write-Host $message -ForegroundColor Red
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Yellow($message)
{
    Write-Host $message -ForegroundColor Yellow
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Green($message)
{
    Write-Host $message -ForegroundColor Green
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Grey($message)
{
    Write-Host $message
    $message | Out-File ($OutputFullPath) -Append
}

function Write-VerboseOutput($message)
{
    Write-Verbose $message
    if($Script:VerboseEnabled)
    {
        $message | Out-File ($OutputFullPath) -Append
    }
}

Function Write-Break {
    Write-Host ""
}


############################################################
############################################################

Function Load-ExShell {
	#Verify that we are on Exchange 2010 or newer 
	if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup') -or (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup'))
	{
		#If we are on Exchange Server, we need to make sure that Exchange Management Snapin is loaded 
		try
		{
			Get-ExchangeServer | Out-Null
		}
		catch
		{
			Write-Host "Loading Exchange PowerShell Module..."
			Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
		}
	}
	else
	{
		Write-Host "Not on Exchange 2010 or newer. Going to exit."
		sleep 2
		exit
	}
}

Function Is-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
    If( $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
        return $true
    }
    else {
        return $false
    }
}

Function Get-OperatingSystemVersion {
param(
[Parameter(Mandatory=$true)][string]$OS_Version
)

    Write-VerboseOutput("Calling: Get-OperatingSystemVersion")
    Write-VerboseOutput("Passed: $OS_Version")
    
    switch($OS_Version)
    {
        "6.0.6000" {Write-VerboseOutput("Returned: Windows2008"); return [HealthChecker.OSVersionName]::Windows2008}
        "6.1.7600" {Write-VerboseOutput("Returned: Windows2008R2"); return [HealthChecker.OSVersionName]::Windows2008R2}
        "6.1.7601" {Write-VerboseOutput("Returned: Windows2008R2"); return [HealthChecker.OSVersionName]::Windows2008R2}
        "6.2.9200" {Write-VerboseOutput("Returned: Windows2012"); return [HealthChecker.OSVersionName]::Windows2012}
        "6.3.9600" {Write-VerboseOutput("Returned: Windows2012R2"); return [HealthChecker.OSVersionName]::Windows2012R2}
        "10.0.14393" {Write-VerboseOutput("Returned: Windows2016"); return [HealthChecker.OSVersionName]::Windows2016}
        "10.0.17713" {Write-VerboseOutput("Returned: Windows2019"); return [HealthChecker.OSVersionName]::Windows2019}
        default{Write-VerboseOutput("Returned: Unknown"); return [HealthChecker.OSVersionName]::Unknown}
    }

}

Function Get-PageFileObject {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name
)
    Write-VerboseOutput("Calling: Get-PageFileObject")
    Write-Verbose("Passed: $Machine_Name")
    [HealthChecker.PageFileObject]$page_obj = New-Object HealthChecker.PageFileObject
    $pagefile = Get-WmiObject -ComputerName $Machine_Name -Class Win32_PageFileSetting
    if($pagefile -ne $null) 
    { 
        if($pagefile.GetType().Name -eq "ManagementObject")
        {
            $page_obj.MaxPageSize = $pagefile.MaximumSize
        }
        $page_obj.PageFile = $pagefile
    }
    else
    {
        Write-VerboseOutput("Return Null value")
    }

    return $page_obj
}


Function Build-NICInformationObject {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name,
[Parameter(Mandatory=$true)][HealthChecker.OSVersionName]$OSVersion
)

    Write-VerboseOutput("Calling: Build-NICInformationObject")
    Write-VerboseOutput("Passed: $Machine_Name")
    Write-VerboseOutput("Passed: $OSVersion")

    [array]$aNICObjects = @() 
    if($OSVersion -ge [HealthChecker.OSVersionName]::Windows2012R2)
    {
        Write-VerboseOutput("Detected OS Version greater than or equal to Windows 2012R2")
        $cimSession = New-CimSession -ComputerName $Machine_Name
        $NetworkCards = Get-NetAdapter -CimSession $cimSession | ?{$_.MediaConnectionState -eq "Connected"}
        foreach($adapter in $NetworkCards)
        {
            Write-VerboseOutput("Working on getting netAdapeterRSS information for adapter: " + $adapter.InterfaceDescription)
            [HealthChecker.NICInformationObject]$nicObject = New-Object -TypeName HealthChecker.NICInformationObject 
            try
            {
                $RSS_Settings = $adapter | Get-netAdapterRss -ErrorAction Stop
                $nicObject.RSSEnabled = $RSS_Settings.Enabled
            }
            catch 
            {
                $Script:iErrorExcluded++
                Write-Yellow("Warning: Unable to get the netAdapterRSS Information for adapter: {0}" -f $adapter.InterfaceDescription)
                $nicObject.RSSEnabled = "NoRSS"
            }
            $nicObject.Description = $adapter.InterfaceDescription
            $nicObject.DriverDate = $adapter.DriverDate
            $nicObject.DriverVersion = $adapter.DriverVersionString
            $nicObject.LinkSpeed = (($adapter.Speed)/1000000).ToString() + " Mbps"
            $nicObject.Name = $adapter.Name
            $nicObject.NICObject = $adapter 
            $aNICObjects += $nicObject
        }

    }
    
    #Else we don't have all the correct powershell options to get more detailed information remotely 
    else
    {
        Write-VerboseOutput("Detected OS Version less than Windows 2012R2")
        $NetworkCards2008 = Get-WmiObject -ComputerName $Machine_Name -Class Win32_NetworkAdapter | ?{$_.NetConnectionStatus -eq 2}
        foreach($adapter in $NetworkCards2008)
        {
            [HealthChecker.NICInformationObject]$nicObject = New-Object -TypeName HealthChecker.NICInformationObject 
            $nicObject.Description = $adapter.Description
            $nicObject.LinkSpeed = $adapter.Speed
            $nicObject.NICObject = $adapter 
            $aNICObjects += $nicObject
        }

    }

    return $aNICObjects 

}

Function Get-HttpProxySetting {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name
)
	$httpProxy32 = [String]::Empty
	$httpProxy64 = [String]::Empty
	Write-VerboseOutput("Calling  Get-HttpProxySetting")
	Write-VerboseOutput("Passed: {0}" -f $Machine_Name)
	$orgErrorPref = $ErrorActionPreference
    $ErrorActionPreference = "Stop"
    
    Function Get-WinHttpSettings {
    param(
        [Parameter(Mandatory=$true)][string]$RegistryLocation
    )
        $connections = Get-ItemProperty -Path $RegistryLocation
        $Proxy = [string]::Empty
        if(($connections -ne $null) -and ($Connections | gm).Name -contains "WinHttpSettings")
        {
            foreach($Byte in $Connections.WinHttpSettings)
            {
                if($Byte -ge 48)
                {
                    $Proxy += [CHAR]$Byte
                }
            }
        }
        return $(if($Proxy -eq [string]::Empty){"<None>"} else {$Proxy})
    }

	try
	{
        $httpProxyPath32 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"
        $httpProxyPath64 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"
        
        if($Machine_Name -ne $env:COMPUTERNAME) 
        {
            Write-VerboseOutput("Calling Get-WinHttpSettings via Invoke-Command")
            $httpProxy32 = Invoke-Command -ComputerName $Machine_Name -ScriptBlock ${Function:Get-WinHttpSettings} -ArgumentList $httpProxyPath32
            $httpProxy64 = Invoke-Command -ComputerName $Machine_Name -ScriptBlock ${Function:Get-WinHttpSettings} -ArgumentList $httpProxyPath64
        }
        else 
        {
            Write-VerboseOutput("Calling Get-WinHttpSettings via local session")
            $httpProxy32 = Get-WinHttpSettings -RegistryLocation $httpProxyPath32
            $httpProxy64 = Get-WinHttpSettings -RegistryLocation $httpProxyPath64
        }
		
		
        Write-VerboseOutput("Http Proxy 32: {0}" -f $httpProxy32)
		Write-VerboseOutput("Http Proxy 64: {0}" -f $httpProxy64)
	}

	catch
	{
        $Script:iErrorExcluded++
		Write-Yellow("Warning: Unable to get the Http Proxy Settings for server {0}" -f $Machine_Name)
	}
	finally
	{
		$ErrorActionPreference = $orgErrorPref
	}

	if($httpProxy32 -ne "<None>")
	{
		return $httpProxy32
	}
	else
	{
		return $httpProxy64
	}

}

Function New-FileLevelHotfixObject {
param(
[parameter(Mandatory=$true)][string]$FriendlyName,
[parameter(Mandatory=$true)][string]$FullFilePath, 
[Parameter(Mandatory=$true)][string]$BuildVersion
)
    #Write-VerboseOutput("Calling Function: New-FileLevelHotfixObject")
    #Write-VerboseOutput("Passed - FriendlyName: {0} FullFilePath: {1} BuldVersion: {2}" -f $FriendlyName, $FullFilePath, $BuildVersion)
    [HealthChecker.FileVersionCheckObject]$FileVersion_obj = New-Object HealthChecker.FileVersionCheckObject
    $FileVersion_obj.FriendlyFileName = $FriendlyName
    $FileVersion_obj.FullPath = $FullFilePath
    $FileVersion_obj.BuildVersion = $BuildVersion
    return $FileVersion_obj
}

Function Get-HotFixListInfo{
param(
[Parameter(Mandatory=$true)][HealthChecker.OSVersionName]$OS_Version
)
    $hotfix_objs = @()
    switch ($OS_Version)
    {
        ([HealthChecker.OSVersionName]::Windows2008R2)
        {
            [HealthChecker.HotfixObject]$hotfix_obj = New-Object HealthChecker.HotfixObject
            $hotfix_obj.KBName = "KB3004383"
            $hotfix_obj.ValidFileLevelCheck = $true
            $hotfix_obj.FileInformation += (New-FileLevelHotfixObject -FriendlyName "Appidapi.dll" -FullFilePath "C:\Windows\SysWOW64\Appidapi.dll" -BuildVersion "6.1.7601.22823")
            #For this check, we are only going to check for one file, becuase there are a ridiculous amount in this KB. Hopefullly we don't see many false positives 
            $hotfix_objs += $hotfix_obj
            return $hotfix_objs
        }
        ([HealthChecker.OSVersionName]::Windows2012R2)
        {
            [HealthChecker.HotfixObject]$hotfix_obj = New-Object HealthChecker.HotfixObject
            $hotfix_obj.KBName = "KB3041832"
            $hotfix_obj.ValidFileLevelCheck = $true
            $hotfix_obj.FileInformation += (New-FileLevelHotfixObject -FriendlyName "Hwebcore.dll" -FullFilePath "C:\Windows\SysWOW64\inetsrv\Hwebcore.dll" -BuildVersion "8.5.9600.17708")
            $hotfix_obj.FileInformation += (New-FileLevelHotfixObject -FriendlyName "Iiscore.dll" -FullFilePath "C:\Windows\SysWOW64\inetsrv\Iiscore.dll" -BuildVersion "8.5.9600.17708")
            $hotfix_obj.FileInformation += (New-FileLevelHotfixObject -FriendlyName "W3dt.dll" -FullFilePath "C:\Windows\SysWOW64\inetsrv\W3dt.dll" -BuildVersion "8.5.9600.17708")
            $hotfix_objs += $hotfix_obj
            
            return $hotfix_objs
        }
        ([HealthChecker.OSVersionName]::Windows2016)
        {
            [HealthChecker.HotfixObject]$hotfix_obj = New-Object HealthChecker.HotfixObject
            $hotfix_obj.KBName = "KB3206632"
            $hotfix_obj.ValidFileLevelCheck = $false
            $hotfix_obj.FileInformation += (New-FileLevelHotfixObject -FriendlyName "clusport.sys" -FullFilePath "C:\Windows\System32\drivers\clusport.sys" -BuildVersion "10.0.14393.576")
            $hotfix_objs += $hotfix_obj
            return $hotfix_objs
        }
    }

    return $null
}

Function Remote-GetFileVersionInfo {
param(
[Parameter(Mandatory=$true)][object]$PassedObject 
)
    $KBsInfo = $PassedObject.KBCheckList
    $ReturnList = @()
    foreach($KBInfo in $KBsInfo)
    {
        $main_obj = New-Object PSCustomObject
        $main_obj | Add-Member -MemberType NoteProperty -Name KBName -Value $KBInfo.KBName 
        $kb_info_List = @()
        foreach($FilePath in $KBInfo.KBInfo)
        {
            $obj = New-Object PSCustomObject
            $obj | Add-Member -MemberType NoteProperty -Name FriendlyName -Value $FilePath.FriendlyName
            $obj | Add-Member -MemberType NoteProperty -Name FilePath -Value $FilePath.FilePath
            $obj | Add-Member -MemberType NoteProperty -Name Error -Value $false
            if(Test-Path -Path $FilePath.FilePath)
            {
            $info = Get-childItem $FilePath.FilePath
            $obj | Add-Member -MemberType NoteProperty -Name ChildItemInfo -Value $info 
            $buildVersion = "{0}.{1}.{2}.{3}" -f $info.VersionInfo.FileMajorPart, $info.VersionInfo.FileMinorPart, $info.VersionInfo.FileBuildPart, $info.VersionInfo.FilePrivatePart
            $obj | Add-Member -MemberType NoteProperty -Name BuildVersion -Value $buildVersion
            
            }
            else 
            {
                $obj.Error = $true
            }
            $kb_info_List += $obj
        }
        $main_obj | Add-Member -MemberType NoteProperty -Name KBInfo -Value $kb_info_List
        $ReturnList += $main_obj
    }

    return $ReturnList
}

Function Get-RemoteHotFixInforamtion {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name,
[Parameter(Mandatory=$true)][HealthChecker.OSVersionName]$OS_Version
)
    $HotfixListObjs = Get-HotFixListInfo -OS_Version $OS_Version
    if($HotfixListObjs -ne $null)    
    {
        $oldErrorAction = $ErrorActionPreference
        $ErrorActionPreference = "stop"
        try 
        {
            $kbList = @() 
            $results = @()
            foreach($HotfixListObj in $HotfixListObjs)
            {
                #HotfixListObj contains all files that we should check for that particluar KB to make sure we are on the correct build 
                $kb_obj = New-Object PSCustomObject
                $kb_obj | Add-Member -MemberType NoteProperty -Name KBName -Value $HotfixListObj.KBName
                $list = @()
                foreach($FileCheck in $HotfixListObj.FileInformation)
                {
                    $obj = New-Object PSCustomObject
                    $obj | Add-Member -MemberType NoteProperty -Name FilePath -Value $FileCheck.FullPath
                    $obj | Add-Member -MemberType NoteProperty -Name FriendlyName -Value $FileCheck.FriendlyFileName
                    $list += $obj
                    #$results += Invoke-Command -ComputerName $Machine_Name -ScriptBlock $script_block -ArgumentList $FileCheck.FullPath
                }
                $kb_obj | Add-Member -MemberType NoteProperty -Name KBInfo -Value $list   
                $kbList += $kb_obj             
            }
            $argList = New-Object PSCustomObject
            $argList | Add-Member -MemberType NoteProperty -Name "KBCheckList" -Value $kbList
            
            if($Machine_Name -ne $env:COMPUTERNAME)
            {
                Write-VerboseOutput("Calling Remote-GetFileVersionInfo via Invoke-Command")
                $results = Invoke-Command -ComputerName $Machine_Name -ScriptBlock ${Function:Remote-GetFileVersionInfo} -ArgumentList $argList
            }
            else 
            {
                Write-VerboseOutput("Calling Remote-GetFileVersionInfo via local session")
                $results = Remote-GetFileVersionInfo -PassedObject $argList 
            }
            
            
            return $results
        }
        catch 
        {
            $Script:iErrorExcluded++ 
        }
        finally
        {
            $ErrorActionPreference = $oldErrorAction
        }
        
    }
}

Function Get-ServerRebootPending {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name
)
    Write-VerboseOutput("Calling: Get-ServerRebootPending")
    Write-VerboseOutput("Passed: {0}" -f $Machine_Name)

    $PendingFileReboot = $false
    $PendingAutoUpdateReboot = $false
    $PendingCBSReboot = $false #Component-Based Servicing Reboot 
    $PendingSCCMReboot = $false
    $ServerPendingReboot = $false

    #Pending File Rename operations 
    Function Get-PendingFileReboot {

        $PendingFileKeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\"
        $file = Get-ItemProperty -Path $PendingFileKeyPath -Name PendingFileRenameOperations
        if($file)
        {
            return $true
        }
        return $false
    }

    Function Get-PendingAutoUpdateReboot {

        if(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired")
        {
            return $true
        }
        return $false
    }

    Function Get-PendingCBSReboot {

        if(Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending")
        {
            return $true
        }
        return $false
    }

    Function Get-PendingSCCMReboot {

        $SCCMReboot = Invoke-CimMethod -Namespace 'Root\ccm\clientSDK' -ClassName 'CCM_ClientUtilities' -Name 'DetermineIfRebootPending'

        if($SCCMReboot)
        {
            If($SCCMReboot.RebootPending -or $SCCMReboot.IsHardRebootPending)
            {
                return $true
            }
        }
        return $false
    }

    Function Execute-ScriptBlock{
    param(
    [Parameter(Mandatory=$true)][string]$Machine_Name,
    [Parameter(Mandatory=$true)][scriptblock]$Script_Block,
    [Parameter(Mandatory=$true)][string]$Script_Block_Name
    )
        Write-VerboseOutput("Calling Script Block {0} for server {1}." -f $Script_Block_Name, $Machine_Name)
        $oldErrorAction = $ErrorActionPreference
        $ErrorActionPreference = "Stop"
        $returnValue = $false
        try 
        {
            $returnValue = Invoke-Command -ComputerName $Machine_Name -ScriptBlock $Script_Block
        }
        catch 
        {
            Write-VerboseOutput("Failed to run Invoke-Command for Script Block {0} on Server {1} --- Note: This could be normal" -f $Script_Block_Name, $Machine_Name)
            $Script:iErrorExcluded++
        }
        finally 
        {
            $ErrorActionPreference = $oldErrorAction
        }
        return $returnValue
    }

    Function Execute-LocalMethods {
    param(
    [Parameter(Mandatory=$true)][string]$Machine_Name,
    [Parameter(Mandatory=$true)][ScriptBlock]$Script_Block,
    [Parameter(Mandatory=$true)][string]$Script_Block_Name
    )
        $oldErrorAction = $ErrorActionPreference
        $ErrorActionPreference = "Stop"
        $returnValue = $false
        Write-VerboseOutput("Calling Local Script Block {0} for server {1}." -f $Script_Block_Name, $Machine_Name)
        try 
        {
            $returnValue = & $Script_Block
        }
        catch 
        {
            Write-VerboseOutput("Failed to run local for Script Block {0} on Server {1} --- Note: This could be normal" -f $Script_Block_Name, $Machine_Name)
            $Script:iErrorExcluded++
        }
        finally 
        {
            $ErrorActionPreference = $oldErrorAction
        }
        return $returnValue
    }

    if($Machine_Name -eq $env:COMPUTERNAME)
    {
        Write-VerboseOutput("Calling Server Reboot Pending options via local session")
        $PendingFileReboot = Execute-LocalMethods -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingFileReboot} -Script_Block_Name "Get-PendingFileReboot"
        $PendingAutoUpdateReboot = Execute-LocalMethods -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingAutoUpdateReboot} -Script_Block_Name "Get-PendingAutoUpdateReboot"
        $PendingCBSReboot = Execute-LocalMethods -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingCBSReboot} -Script_Block_Name "Get-PendingCBSReboot"
        $PendingSCCMReboot = Execute-LocalMethods -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingSCCMReboot} -Script_Block_Name "Get-PendingSCCMReboot"
    }
    else 
    {
        Write-VerboseOutput("Calling Server Reboot Pending options via Invoke-Command")
        $PendingFileReboot = Execute-ScriptBlock -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingFileReboot} -Script_Block_Name "Get-PendingFileReboot"
        $PendingAutoUpdateReboot = Execute-ScriptBlock -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingAutoUpdateReboot} -Script_Block_Name "Get-PendingAutoUpdateReboot"
        $PendingCBSReboot = Execute-ScriptBlock -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingCBSReboot} -Script_Block_Name "Get-PendingCBSReboot"
        $PendingSCCMReboot = Execute-ScriptBlock -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingSCCMReboot} -Script_Block_Name "Get-PendingSCCMReboot"
    }

    Write-VerboseOutput("Results - PendingFileReboot: {0} PendingAutoUpdateReboot: {1} PendingCBSReboot: {2} PendingSCCMReboot: {3}" -f $PendingFileReboot, $PendingAutoUpdateReboot, $PendingCBSReboot, $PendingSCCMReboot)
    if($PendingFileReboot -or $PendingAutoUpdateReboot -or $PendingCBSReboot -or $PendingSCCMReboot)
    {
        $ServerPendingReboot = $true
    }

    Write-VerboseOutput("Exit: Get-ServerRebootPending")
    return $ServerPendingReboot
}

Function Build-OperatingSystemObject {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name
)
 
    Write-VerboseOutput("Calling: Build-OperatingSystemObject")
    Write-VerboseOutput("Passed: $Machine_Name")

    [HealthChecker.OperatingSystemObject]$os_obj = New-Object HealthChecker.OperatingSystemObject
    $os = Get-WmiObject -ComputerName $Machine_Name -Class Win32_OperatingSystem
    try
    {
        $plan = Get-WmiObject -ComputerName $Machine_Name -Class Win32_PowerPlan -Namespace root\cimv2\power -Filter "isActive='true'"
    }
    catch
    {
        Write-VerboseOutput("Unable to get power plan from the server")
        $Script:iErrorExcluded++
        $plan = $null
    }
    $os_obj.OSVersionBuild = $os.Version
    $os_obj.OSVersion = (Get-OperatingSystemVersion -OS_Version $os_obj.OSVersionBuild)
    $os_obj.OperatingSystemName = $os.Caption
    $os_obj.OperatingSystem = $os
    
    if($plan -ne $null)
    {
        if($plan.InstanceID -eq "Microsoft:PowerPlan\{8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c}")
        {
            Write-VerboseOutput("High Performance Power Plan is set to true")
            $os_obj.HighPerformanceSet = $true
        }
        $os_obj.PowerPlanSetting = $plan.ElementName
        
    }
    else
    {
        Write-VerboseOutput("Power Plan Information could not be read")
        $os_obj.HighPerformanceSet = $false
        $os_obj.PowerPlanSetting = "N/A"
    }
    $os_obj.PowerPlan = $plan 
    $os_obj.PageFile = (Get-PageFileObject -Machine_Name $Machine_Name)
    $os_obj.NetworkAdapters = (Build-NICInformationObject -Machine_Name $Machine_Name -OSVersion $os_obj.OSVersion) 

    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Machine_Name)
    $RegKey= $Reg.OpenSubKey("SYSTEM\CurrentControlSet\Services\Tcpip\Parameters")
    $os_obj.TCPKeepAlive = $RegKey.GetValue("KeepAliveTime")
	$os_obj.HttpProxy = Get-HttpProxySetting -Machine_Name $Machine_Name
    $os_obj.HotFixes = (Get-HotFix -ComputerName $Machine_Name -ErrorAction SilentlyContinue) #old school check still valid and faster and a failsafe 
    $os_obj.HotFixInfo = Get-RemoteHotFixInforamtion -Machine_Name $Machine_Name -OS_Version $os_obj.OSVersion 
    $os_obj.LmCompat = (Build-LmCompatibilityLevel -Machine_Name $Machine_Name)
    $os_obj.ServerPendingReboot = (Get-ServerRebootPending -Machine_Name $Machine_Name)

    return $os_obj
}

Function Get-ServerType {
param(
[Parameter(Mandatory=$true)][string]$ServerType
)
    Write-VerboseOutput("Calling: Get-ServerType")
    Write-VerboseOutput("Passed: $serverType")



    if($ServerType -like "VMware*"){Write-VerboseOutput("Returned: VMware"); return [HealthChecker.ServerType]::VMWare}
    elseif($ServerType -like "*Microsoft Corporation*"){Write-VerboseOutput("Returned: HyperV"); return [HealthChecker.ServerType]::HyperV}
    elseif($ServerType.Length -gt 0) {Write-VerboseOutput("Returned: Physical"); return [HealthChecker.ServerType]::Physical}
    else{Write-VerboseOutput("Returned: unknown") ;return [HealthChecker.ServerType]::Unknown}
    
}


Function Get-ProcessorInformationObject {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name
)
    Write-VerboseOutput("Calling: Get-ProcessorInformationObject")
    Write-VerboseOutput("Passed: $Machine_Name")
    [HealthChecker.ProcessorInformationObject]$processor_info_object = New-Object HealthChecker.ProcessorInformationObject
    $wmi_obj_processor = Get-WmiObject -ComputerName $Machine_Name -Class Win32_Processor
    $object_Type = $wmi_obj_processor.Gettype().Name 
    Write-VerboseOutput("Processor object type: $object_Type")
    
    #if it is a single processor 
    if($object_Type -eq "ManagementObject") {
        Write-VerboseOutput("single processor detected")
        $processor_info_object.ProcessorName = $wmi_obj_processor.Name
        $processor_info_object.MaxMegacyclesPerCore = $wmi_obj_processor.MaxClockSpeed
    }
    else{
        Write-VerboseOutput("multiple processor detected")
        $processor_info_object.ProcessorName = $wmi_obj_processor[0].Name
        $processor_info_object.MaxMegacyclesPerCore = $wmi_obj_processor[0].MaxClockSpeed
    }

    #Get the total number of cores in the processors 
    Write-VerboseOutput("getting the total number of cores in the processor(s)")
    foreach($processor in $wmi_obj_processor) 
    {
        $processor_info_object.NumberOfPhysicalCores += $processor.NumberOfCores 
        $processor_info_object.NumberOfLogicalProcessors += $processor.NumberOfLogicalProcessors
        $processor_info_object.NumberOfProcessors += 1 #may want to call Win32_ComputerSystem and use NumberOfProcessors for this instead.. but this should get the same results. 

        #Test to see if we are throttling the processor 
        if($processor.CurrentClockSpeed -lt $processor.MaxClockSpeed) 
        {
            Write-VerboseOutput("We see the processor being throttled")
            $processor_info_object.CurrentMegacyclesPerCore = $processor.CurrentClockSpeed
            $processor_info_object.ProcessorIsThrottled = $true 
        }

        if($processor.Name -ne $processor_info_object.ProcessorName -or $processor.MaxClockSpeed -ne $processor_info_object.MaxMegacyclesPerCore){$processor_info_object.DifferentProcessorsDetected = $true; Write-VerboseOutput("Different Processors are detected"); Write-Yellow("Warning: Different Processors are detected. This shouldn't occur")}
    }

	Write-VerboseOutput("Trying to get the System.Environment ProcessorCount")
	$oldError = $ErrorActionPreference
    $ErrorActionPreference = "Stop"
    Function Get-ProcessorCount {
        [System.Environment]::ProcessorCount
    }
	try
	{
        if($Machine_Name -ne $env:COMPUTERNAME)
        {
            Write-VerboseOutput("Getting System.Environment ProcessorCount from Invoke-Command")
            $processor_info_object.EnvProcessorCount = (
                Invoke-Command -ComputerName $Machine_Name -ScriptBlock ${Function:Get-ProcessorCount}
            )
        }
        else 
        {
            Write-VerboseOutput("Getting System.Environment ProcessorCount from local session")
            $processor_info_object.EnvProcessorCount = Get-ProcessorCount
        }

	}
	catch
	{
        $Script:iErrorExcluded++
		Write-Red("Error: Unable to get Environment Processor Count on server {0}" -f $Machine_Name)
		$processor_info_object.EnvProcessorCount = -1 
	}
	finally
	{
		$ErrorActionPreference = $oldError
	}

    $processor_info_object.Processor = $wmi_obj_processor
    return $processor_info_object

}

Function Build-HardwareObject {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name
)
    Write-VerboseOutput("Calling: Build-HardwareObject")
    Write-VerboseOutput("Passed: $Machine_Name")
    [HealthChecker.HardwareObject]$hardware_obj = New-Object HealthChecker.HardwareObject
    $system = Get-WmiObject -ComputerName $Machine_Name -Class Win32_ComputerSystem
    $hardware_obj.Manufacturer = $system.Manufacturer
    $hardware_obj.System = $system
    $hardware_obj.AutoPageFile = $system.AutomaticManagedPagefile
    $hardware_obj.TotalMemory = $system.TotalPhysicalMemory
    $hardware_obj.ServerType = (Get-ServerType -ServerType $system.Manufacturer)
    $hardware_obj.Processor = Get-ProcessorInformationObject -Machine_Name $Machine_Name 
    $hardware_obj.Model = $system.Model 

    return $hardware_obj
}


Function Get-NetFrameworkVersionFriendlyInfo{
param(
[Parameter(Mandatory=$true)][int]$NetVersionKey,
[Parameter(Mandatory=$true)][HealthChecker.OSVersionName]$OSVersionName 
)
    Write-VerboseOutput("Calling: Get-NetFrameworkVersionFriendlyInfo")
    Write-VerboseOutput("Passed: " + $NetVersionKey.ToString())
    Write-VerboseOutput("Passed: " + $OSVersionName.ToString())
    [HealthChecker.NetVersionObject]$versionObject = New-Object -TypeName HealthChecker.NetVersionObject
        if(($NetVersionKey -ge [HealthChecker.NetVersion]::Net4d5) -and ($NetVersionKey -lt [HealthChecker.NetVersion]::Net4d5d1))
    {
        $versionObject.FriendlyName = "4.5"
        $versionObject.NetVersion = [HealthChecker.NetVersion]::Net4d5
    }
    elseif(($NetVersionKey -ge [HealthChecker.NetVersion]::Net4d5d1) -and ($NetVersionKey -lt [HealthChecker.NetVersion]::Net4d5d2))
    {
        $versionObject.FriendlyName = "4.5.1"
        $versionObject.NetVersion = [HealthChecker.NetVersion]::Net4d5d1
    }
    elseif(($NetVersionKey -ge [HealthChecker.NetVersion]::Net4d5d2) -and ($NetVersionKey -lt [HealthChecker.NetVersion]::Net4d5d2wFix))
    {
        $versionObject.FriendlyName = "4.5.2"
        $versionObject.NetVersion = [HealthChecker.NetVersion]::Net4d5d2
    }
    elseif(($NetVersionKey -ge [HealthChecker.NetVersion]::Net4d5d2wFix) -and ($NetVersionKey -lt [HealthChecker.NetVersion]::Net4d6))
    {
        $versionObject.FriendlyName = "4.5.2 with Hotfix 3146718"
        $versionObject.NetVersion = [HealthChecker.NetVersion]::Net4d5d2wFix
    }
    elseif(($NetVersionKey -ge [HealthChecker.NetVersion]::Net4d6) -and ($NetVersionKey -lt [HealthChecker.NetVersion]::Net4d6d1))
    {
        $versionObject.FriendlyName = "4.6"
        $versionObject.NetVersion = [HealthChecker.NetVersion]::Net4d6
    }
    elseif(($NetVersionKey -ge [HealthChecker.NetVersion]::Net4d6d1) -and ($NetVersionKey -lt [HealthChecker.NetVersion]::Net4d6d1wFix))
    {
        $versionObject.FriendlyName = "4.6.1"
        $versionObject.NetVersion = [HealthChecker.NetVersion]::Net4d6d1
    }
    elseif($NetVersionKey -eq 394802 -and $OSVersionName -eq [HealthChecker.OSVersionName]::Windows2016)
    {
        $versionObject.FriendlyName = "Windows Server 2016 .NET 4.6.2"
        $versionObject.NetVersion = [HealthChecker.NetVersion]::Net4d6d2
    }
    elseif(($NetVersionKey -ge [HealthChecker.NetVersion]::Net4d6d1wFix) -and ($NetVersionKey -lt [HealthChecker.NetVersion]::Net4d6d2))
    {
        $versionObject.FriendlyName = "4.6.1 with Hotfix 3146716/3146714/3146715"
        $versionObject.NetVersion = [HealthChecker.NetVersion]::Net4d6d1wFix
    }
    elseif(($NetVersionKey -ge [HealthChecker.NetVersion]::Net4d6d2) -and ($NetVersionKey -lt [HealthChecker.NetVersion]::Net4d7))
    {
        $versionObject.FriendlyName = "4.6.2"
        $versionObject.NetVersion = [HealthChecker.NetVersion]::Net4d6d2
    }
	elseif($NetVersionKey -ge [HealthChecker.NetVersion]::Net4d7 -and ($NetVersionKey -lt [HealthChecker.NetVersion]::Net4d7d1))
	{
		$versionObject.FriendlyName = "4.7"
		$versionObject.NetVersion = [HealthChecker.NetVersion]::Net4d7
    }
    elseif($NetVersionKey -ge [HealthChecker.NetVersion]::Net4d7d1 -and ($NetVersionKey -lt [HealthChecker.NetVersion]::Net4d7d2))
    {
        $versionObject.FriendlyName = "4.7.1"
        $versionObject.NetVersion = [HealthChecker.NetVersion]::Net4d7d1
    }
    elseif($NetVersionKey -ge [HealthChecker.NetVersion]::Net4d7d2)
    {
        $versionObject.FriendlyName = "4.7.2"
        $versionObject.NetVersion = [HealthChecker.NetVersion]::Net4d7d2
    }
    else
    {
        $versionObject.FriendlyName = "Unknown" 
        $versionObject.NetVersion = [HealthChecker.NetVersion]::Unknown
    }
    $versionObject.NetRegValue = $NetVersionKey


    Write-VerboseOutput("Returned: " + $versionObject.FriendlyName)
    return $versionObject
    
}


#Uses registry build numbers from https://msdn.microsoft.com/en-us/library/hh925568(v=vs.110).aspx
Function Build-NetFrameWorkVersionObject {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name,
[Parameter(Mandatory=$true)][HealthChecker.OSVersionName]$OSVersionName
)
    Write-VerboseOutput("Calling: Build-NetFrameWorkVersionObject")
    Write-VerboseOutput("Passed: $Machine_Name")
    Write-VerboseOutput("Passed: $OSVersionName")

    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Machine_Name)
    $RegKey = $Reg.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full")
    [int]$NetVersionKey = $RegKey.GetValue("Release")
    $sNetVersionKey = $NetVersionKey.ToString()
    Write-VerboseOutput("Got $sNetVersionKey from the registry")

    [HealthChecker.NetVersionObject]$versionObject = Get-NetFrameworkVersionFriendlyInfo -NetVersionKey $NetVersionKey -OSVersionName $OSVersionName
    return $versionObject

}

Function Get-ExchangeVersion {
param(
[Parameter(Mandatory=$true)][object]$AdminDisplayVersion
)
    Write-VerboseOutput("Calling: Get-ExchangeVersion")
    Write-VerboseOutput("Passed: " + $AdminDisplayVersion.ToString())
    $iBuild = $AdminDisplayVersion.Major + ($AdminDisplayVersion.Minor / 10)
    Write-VerboseOutput("Determing build based of of: " + $iBuild) 
    switch($iBuild)
    {
        14.3 {Write-VerboseOutput("Returned: Exchange2010"); return [HealthChecker.ExchangeVersion]::Exchange2010}
        15 {Write-VerboseOutput("Returned: Exchange2013"); return [HealthChecker.ExchangeVersion]::Exchange2013}
        15.1{Write-VerboseOutput("Returned: Exchange2016"); return [HealthChecker.ExchangeVersion]::Exchange2016}
        15.2{Write-VerboseOutput("Returned: Exchange2019"); return [HealthChecker.ExchangeVersion]::Exchange2019}
        default {Write-VerboseOutput("Returned: Unknown"); return [HealthChecker.ExchangeVersion]::Unknown}
    }

}

Function Get-BuildNumberToString {
param(
[Parameter(Mandatory=$true)][object]$AdminDisplayVersion
)
    $sAdminDisplayVersion = $AdminDisplayVersion.Major.ToString() + "." + $AdminDisplayVersion.Minor.ToString() + "."  + $AdminDisplayVersion.Build.ToString() + "."  + $AdminDisplayVersion.Revision.ToString()
    Write-VerboseOutput("Called: Get-BuildNumberToString")
    Write-VerboseOutput("Returned: " + $sAdminDisplayVersion)
    return $sAdminDisplayVersion
}

<#
New Release Update 
#>
Function Get-ExchangeBuildObject {
param(
[Parameter(Mandatory=$true)][object]$AdminDisplayVersion
)
    Write-VerboseOutput("Calling: Get-ExchangeBuildObject")
    Write-VerboseOutput("Passed: " + $AdminDisplayVersion.ToString())
    [HealthChecker.ExchangeBuildObject]$exBuildObj = New-Object -TypeName HealthChecker.ExchangeBuildObject
    $iRevision = if($AdminDisplayVersion.Revision -lt 10) {$AdminDisplayVersion.Revision /10} else{$AdminDisplayVersion.Revision /100}
    $buildRevision = $AdminDisplayVersion.Build + $iRevision
    Write-VerboseOutput("Revision Value: " + $iRevision)
    Write-VerboseOutput("Build Plus Revision Value: " + $buildRevision)
    #https://technet.microsoft.com/en-us/library/hh135098(v=exchg.150).aspx

    if($AdminDisplayVersion.Major -eq 15 -and $AdminDisplayVersion.Minor -eq 2)
    {
        Write-VerboseOutput("Determined that we are on Exchange 2019")
        $exBuildObj.ExchangeVersion = [HealthChecker.ExchangeVersion]::Exchange2019
        if($buildRevision -ge 196.0){$exBuildObj.CU = [HealthChecker.ExchangeCULevel]::Preview}
    }
    elseif($AdminDisplayVersion.Major -eq 15 -and $AdminDisplayVersion.Minor -eq 1)
    {
        Write-VerboseOutput("Determined that we are on Exchange 2016")
        $exBuildObj.ExchangeVersion = [HealthChecker.ExchangeVersion]::Exchange2016
        if($buildRevision -ge 225.16 -and $buildRevision -lt 225.42) {if($buildRevision -gt 225.16){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::Preview}
        elseif($buildRevision -lt 396.30) {if($buildRevision -gt 225.42){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::RTM}
        elseif($buildRevision -lt 466.34) {if($buildRevision -gt 396.30){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU1}
        elseif($buildRevision -lt 544.27) {if($buildRevision -gt 466.34){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU2}
        elseif($buildRevision -lt 669.32) {if($buildRevision -gt 544.27){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU3}
        elseif($buildRevision -lt 845.34) {if($buildRevision -gt 669.32){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU4}
        elseif($buildRevision -lt 1034.26) {if($buildRevision -gt 845.34){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU5}
        elseif($buildRevision -lt 1261.35) {if($buildRevision -gt 1034.26){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU6}
        elseif($buildRevision -lt 1415.2) {if($buildRevision -gt 1261.35){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU7}
        elseif($buildRevision -lt 1466.3) {if($buildRevision -gt 1415.2){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU8}
        elseif($buildRevision -lt 1531.3) {if($buildRevision -gt 1466.3){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU9}
        elseif($buildRevision -ge 1531.3) {if($buildRevision -gt 1531.3){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU10}

    }
    elseif($AdminDisplayVersion.Major -eq 15 -and $AdminDisplayVersion.Minor -eq 0)
    {
        Write-VerboseOutput("Determined that we are on Exchange 2013")
        $exBuildObj.ExchangeVersion = [HealthChecker.ExchangeVersion]::Exchange2013
        if($buildRevision -ge 516.32 -and $buildRevision -lt 620.29) {if($buildRevision -gt 516.32){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::RTM}
        elseif($buildRevision -lt 712.24) {if($buildRevision -gt 620.29){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU1}
        elseif($buildRevision -lt 775.38) {if($buildRevision -gt 712.24){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU2}
        elseif($buildRevision -lt 847.32) {if($buildRevision -gt 775.38){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU3}
        elseif($buildRevision -lt 913.22) {if($buildRevision -gt 847.32){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU4}
        elseif($buildRevision -lt 995.29) {if($buildRevision -gt 913.22){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU5}
        elseif($buildRevision -lt 1044.25) {if($buildRevision -gt 995.29){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU6}
        elseif($buildRevision -lt 1076.9) {if($buildRevision -gt 1044.25){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU7}
        elseif($buildRevision -lt 1104.5) {if($buildRevision -gt 1076.9){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU8}
        elseif($buildRevision -lt 1130.7) {if($buildRevision -gt 1104.5){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU9}
        elseif($buildRevision -lt 1156.6) {if($buildRevision -gt 1130.7){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU10}
        elseif($buildRevision -lt 1178.4) {if($buildRevision -gt 1156.6){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU11}
        elseif($buildRevision -lt 1210.3) {if($buildRevision -gt 1178.4){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU12}
        elseif($buildRevision -lt 1236.3) {if($buildRevision -gt 1210.3){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU13}
        elseif($buildRevision -lt 1263.5) {if($buildRevision -gt 1236.3){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU14}
        elseif($buildRevision -lt 1293.2) {if($buildRevision -gt 1263.5){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU15}
        elseif($buildRevision -lt 1320.4) {if($buildRevision -gt 1293.2){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU16}
        elseif($buildRevision -lt 1347.2) {if($buildRevision -gt 1320.4){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU17}
        elseif($buildRevision -lt 1365.1) {if($buildRevision -gt 1347.2){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU18}
        elseif($buildRevision -lt 1367.3) {if($buildRevision -gt 1365.1){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU19}
        elseif($buildRevision -lt 1395.4) {if($buildRevision -gt 1367.3){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU20}
        elseif($buildRevision -ge 1395.4) {if($buildRevision -gt 1395.4){$exBuildObj.InbetweenCUs = $true} $exBuildObj.CU = [HealthChecker.ExchangeCULevel]::CU21}
    }
    else
    {
        Write-Red "Error: Didn't know how to process the Admin Display Version Provided"
        
    }

    return $exBuildObj

}

#New Release Update 
Function Get-ExchangeBuildInformation {
param(
[Parameter(Mandatory=$true)][object]$AdminDisplayVersion
)
    Write-VerboseOutput("Calling: Get-ExchangeBuildInformation")
    Write-VerboseOutput("Passed: " + $AdminDisplayVersion.ToString())
    [HealthChecker.ExchangeInformationTempObject]$tempObject = New-Object -TypeName HealthChecker.ExchangeInformationTempObject
    
    #going to remove the minor checks. Not sure I see a value in keeping them. 
    if($AdminDisplayVersion.Major -eq 15)
    {
       Write-VerboseOutput("Determined that we are working with Exchange 2013 or greater")
       [HealthChecker.ExchangeBuildObject]$exBuildObj = Get-ExchangeBuildObject -AdminDisplayVersion $AdminDisplayVersion 
       Write-VerboseOutput("Got the exBuildObj")
       Write-VerboseOutput("Exchange Version is set to: " + $exBuildObj.ExchangeVersion.ToString())
       Write-VerboseOutput("CU is set to: " + $exBuildObj.CU.ToString())
       Write-VerboseOutput("Inbetween CUs: " + $exBuildObj.InbetweenCUs.ToString())
       switch($exBuildObj.ExchangeVersion)
       {
        ([HealthChecker.ExchangeVersion]::Exchange2019)
            {
                Write-VerboseOutput("Working with Exchange 2019")
                switch($exBuildObj.CU)
                {
                    ([HealthChecker.ExchangeCULevel]::Preview) {$tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.FriendlyName = "Exchange 2019 Preview"; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.ReleaseDate = "07/24/2018"; break}
                    default {Write-Red("Error: Unknown Exchange 2019 Build was detected"); $tempObject.Error = $true; break;}
                }
            }

        ([HealthChecker.ExchangeVersion]::Exchange2016)
            {
                Write-VerboseOutput("Working with Exchange 2016")
                switch($exBuildObj.CU)
                {
                    ([HealthChecker.ExchangeCULevel]::Preview) {$tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.FriendlyName = "Exchange 2016 Preview"; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.ReleaseDate = "07/22/2015"; break}
                    ([HealthChecker.ExchangeCULevel]::RTM) {$tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.FriendlyName = "Exchange 2016 RTM"; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.ReleaseDate = "10/01/2015"; break}
                    ([HealthChecker.ExchangeCULevel]::CU1) {$tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.FriendlyName = "Exchange 2016 CU1"; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.ReleaseDate = "03/15/2016"; break}
                    ([HealthChecker.ExchangeCULevel]::CU2) {$tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.FriendlyName = "Exchange 2016 CU2"; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.ReleaseDate = "06/21/2016"; break}
                    ([HealthChecker.ExchangeCULevel]::CU3) {$tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.FriendlyName = "Exchange 2016 CU3"; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.ReleaseDate = "09/20/2016"; break}
                    ([HealthChecker.ExchangeCULevel]::CU4) {$tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.FriendlyName = "Exchange 2016 CU4"; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.ReleaseDate = "12/13/2016"; break}
                    ([HealthChecker.ExchangeCULevel]::CU5) {$tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.FriendlyName = "Exchange 2016 CU5"; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.ReleaseDate = "03/21/2017"; break}
                    ([HealthChecker.ExchangeCULevel]::CU6) {$tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.FriendlyName = "Exchange 2016 CU6"; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.ReleaseDate = "06/24/2017"; break}
                    ([HealthChecker.ExchangeCULevel]::CU7) {$tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.FriendlyName = "Exchange 2016 CU7"; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.ReleaseDate = "09/16/2017"; break}
                    ([HealthChecker.ExchangeCULevel]::CU8) {$tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.FriendlyName = "Exchange 2016 CU8"; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.ReleaseDate = "12/19/2017"; break}
                    ([HealthChecker.ExchangeCULevel]::CU9) {$tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.FriendlyName = "Exchange 2016 CU9"; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.ReleaseDate = "03/20/2018"; $tempObject.SupportedCU = $true; break}
                    ([HealthChecker.ExchangeCULevel]::CU10) {$tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.FriendlyName = "Exchange 2016 CU10"; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.ReleaseDate = "06/19/2018"; $tempObject.SupportedCU = $true; break}
                    default {Write-Red "Error: Unknown Exchange 2016 build was detected"; $tempObject.Error = $true; break;}
                }
                break;
            }
        ([HealthChecker.ExchangeVersion]::Exchange2013)
            {
                Write-VerboseOutput("Working with Exchange 2013")
                switch($exBuildObj.CU)
                {
                    ([HealthChecker.ExchangeCULevel]::RTM) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 RTM"; $tempObject.ReleaseDate = "12/03/2012"; break}
                    ([HealthChecker.ExchangeCULevel]::CU1) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU1"; $tempObject.ReleaseDate = "04/02/2013"; break}
                    ([HealthChecker.ExchangeCULevel]::CU2) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU2"; $tempObject.ReleaseDate = "07/09/2013"; break}
                    ([HealthChecker.ExchangeCULevel]::CU3) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU3"; $tempObject.ReleaseDate = "11/25/2013"; break}
                    ([HealthChecker.ExchangeCULevel]::CU4) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU4"; $tempObject.ReleaseDate = "02/25/2014"; break}
                    ([HealthChecker.ExchangeCULevel]::CU5) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU5"; $tempObject.ReleaseDate = "05/27/2014"; break}
                    ([HealthChecker.ExchangeCULevel]::CU6) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU6"; $tempObject.ReleaseDate = "08/26/2014"; break}
                    ([HealthChecker.ExchangeCULevel]::CU7) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU7"; $tempObject.ReleaseDate = "12/09/2014"; break}
                    ([HealthChecker.ExchangeCULevel]::CU8) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU8"; $tempObject.ReleaseDate = "03/17/2015"; break}
                    ([HealthChecker.ExchangeCULevel]::CU9) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU9"; $tempObject.ReleaseDate = "06/17/2015"; break}
                    ([HealthChecker.ExchangeCULevel]::CU10) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU10"; $tempObject.ReleaseDate = "09/15/2015"; break}
                    ([HealthChecker.ExchangeCULevel]::CU11) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU11"; $tempObject.ReleaseDate = "12/15/2015"; break}
                    ([HealthChecker.ExchangeCULevel]::CU12) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU12"; $tempObject.ReleaseDate = "03/15/2016"; break}
                    ([HealthChecker.ExchangeCULevel]::CU13) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU13"; $tempObject.ReleaseDate = "06/21/2016"; break}
                    ([HealthChecker.ExchangeCULevel]::CU14) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU14"; $tempObject.ReleaseDate = "09/20/2016"; break}
                    ([HealthChecker.ExchangeCULevel]::CU15) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU15"; $tempObject.ReleaseDate = "12/13/2016"; break}
                    ([HealthChecker.ExchangeCULevel]::CU16) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU16"; $tempObject.ReleaseDate = "03/21/2017"; break}
                    ([HealthChecker.ExchangeCULevel]::CU17) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU17"; $tempObject.ReleaseDate = "06/24/2017"; break}
                    ([HealthChecker.ExchangeCULevel]::CU18) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU18"; $tempObject.ReleaseDate = "09/16/2017"; break}
                    ([HealthChecker.ExchangeCULevel]::CU19) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU19"; $tempObject.ReleaseDate = "12/19/2017"; $tempObject.SupportedCU = $true; break}
                    ([HealthChecker.ExchangeCULevel]::CU20) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU20"; $tempObject.ReleaseDate = "03/20/2018"; $tempObject.SupportedCU = $true; break}
                    ([HealthChecker.ExchangeCULevel]::CU21) {$tempObject.ExchangeBuildObject = $exBuildObj; $tempObject.InbetweenCUs = $exBuildObj.InbetweenCUs; $tempObject.ExchangeBuildNumber = (Get-BuildNumberToString $AdminDisplayVersion); $tempObject.FriendlyName = "Exchange 2013 CU21"; $tempObject.ReleaseDate = "06/19/2018"; $tempObject.SupportedCU = $true; break}
                    default {Write-Red "Error: Unknown Exchange 2013 build was detected"; $tempObject.Error = $TRUE; break;}
                }
                break;
            }
            
        default {$tempObject.Error = $true; Write-Red "Error: Unknown error in Get-ExchangeBuildInformation"}   
       }
    }

    else
    {
        Write-VerboseOutput("Error occur because we weren't on Exchange 2013 or greater")
        $tempObject.Error = $true
    }

    return $tempObject
}

<#

Exchange 2013 Support 
https://technet.microsoft.com/en-us/library/aa996719(v=exchg.150).aspx

Exchange 2016 Support 
https://technet.microsoft.com/en-us/library/aa996719(v=exchg.160).aspx

Team Blog Articles 

.NET Framework 4.7 and Exchange Server
https://blogs.technet.microsoft.com/exchange/2017/06/13/net-framework-4-7-and-exchange-server/

Released: December 2016 Quarterly Exchange Updates
https://blogs.technet.microsoft.com/exchange/2016/12/13/released-december-2016-quarterly-exchange-updates/

Released: September 2016 Quarterly Exchange Updates
https://blogs.technet.microsoft.com/exchange/2016/09/20/released-september-2016-quarterly-exchange-updates/

Released: June 2016 Quarterly Exchange Updates
https://blogs.technet.microsoft.com/exchange/2016/06/21/released-june-2016-quarterly-exchange-updates/

Released: December 2017 Quarterly Exchange Updates
https://blogs.technet.microsoft.com/exchange/2017/12/19/released-december-2017-quarterly-exchange-updates/

Summary:
Exchange 2013 CU19 & 2016 CU8 .NET Framework 4.7.1 Supported on all OSs 
Exchange 2013 CU15 & 2016 CU4 .Net Framework 4.6.2 Supported on All OSs
Exchange 2016 CU3 .NET Framework 4.6.2 Supported on Windows 2016 OS - however, stuff is broke on this OS. 

Exchange 2013 CU13 & Exchange 2016 CU2 .NET Framework 4.6.1 Supported on all OSs


Exchange 2013 CU12 & Exchange 2016 CU1 Supported on .NET Framework 4.5.2 

The upgrade to .Net 4.6.2, while strongly encouraged, is optional with these releases. As previously disclosed, the cumulative updates released in our March 2017 quarterly updates will require .Net 4.6.2.

#>
Function Check-DotNetFrameworkSupportedLevel {
param(
[Parameter(Mandatory=$true)][HealthChecker.ExchangeBuildObject]$exBuildObj,
[Parameter(Mandatory=$true)][HealthChecker.OSVersionName]$OSVersionName,
[Parameter(Mandatory=$true)][HealthChecker.NetVersion]$NetVersion
)
    Write-VerboseOutput("Calling: Check-DotNetFrameworkSupportedLevel")


    Function Check-NetVersionToExchangeVersion {
    param(
    [Parameter(Mandatory=$true)][HealthChecker.NetVersion]$CurrentNetVersion,
    [Parameter(Mandatory=$true)][HealthChecker.NetVersion]$MinSupportNetVersion,
    [Parameter(Mandatory=$true)][HealthChecker.NetVersion]$RecommendedNetVersion
    
    )
        [HealthChecker.NetVersionCheckObject]$NetCheckObj = New-Object -TypeName HealthChecker.NetVersionCheckObject
        $NetCheckObj.RecommendedNetVersion = $true 
        Write-VerboseOutput("Calling: Check-NetVersionToExchangeVersion")
        Write-VerboseOutput("Passed: Current Net Version: " + $CurrentNetVersion.ToString())
        Write-VerboseOutput("Passed: Min Support Net Version: " + $MinSupportNetVersion.ToString())
        Write-VerboseOutput("Passed: Recommnded/Max Net Version: " + $RecommendedNetVersion.ToString())

        #If we are on the recommended/supported version of .net then we should be okay 
        if($CurrentNetVersion -eq $RecommendedNetVersion)
        {
            Write-VerboseOutput("Current Version of .NET equals the Recommended Version of .NET")
            $NetCheckObj.Supported = $true    
        }
        elseif($CurrentNetVersion -eq [HealthChecker.NetVersion]::Net4d6 -and $RecommendedNetVersion -ge [HealthChecker.NetVersion]::Net4d6d1wFix)
        {
            Write-VerboseOutput("Current version of .NET equals 4.6 while the recommended version of .NET is equal to or greater than 4.6.1 with hotfix. This means that we are on an unsupported version because we never supported just 4.6")
            $NetCheckObj.Supported = $false
            $NetCheckObj.RecommendedNetVersion = $false
            [HealthChecker.NetVersionObject]$RecommendedNetVersionObject = Get-NetFrameworkVersionFriendlyInfo -NetVersionKey $RecommendedNetVersion.value__ -OSVersionName $OSVersionName
            $NetCheckObj.DisplayWording = "On .NET 4.6 and this is an unsupported build of .NET for Exchange. Only .NET 4.6.1 with the hotfix and greater are supported. Please upgrade to " + $RecommendedNetVersionObject.FriendlyName + " as soon as possible to get into a supported state."
        }
		elseif($CurrentNetVersion -eq [HealthChecker.NetVersion]::Net4d6d1 -and $RecommendedNetVersion -ge [HealthChecker.NetVersion]::Net4d6d1wFix)
		{
			Write-VerboseOutput("Current version of .NET equals 4.6.1 while the recommended version of .NET is equal to or greater than 4.6.1 with hotfix. This means that we are on an unsupported version because we never supported just 4.6.1 without the hotfix")
			$NetCheckObj.Supported = $false
            $NetCheckObj.RecommendedNetVersion = $false
			[HealthChecker.NetVersionObject]$RecommendedNetVersionObject = Get-NetFrameworkVersionFriendlyInfo -NetVersionKey $RecommendedNetVersion.value__ -OSVersionName $OSVersionName
			$NetCheckObj.DisplayWording = "On .NET 4.6.1 and this is an unsupported build of .NET for Exchange. Only .NET 4.6.1 with the hotfix and greater are supported. Please upgrade to " + $RecommendedNetVersionObject.FriendlyName + " as soon as possible to get into a supported state."
		}

        #this catch is for when you are on a version of exchange where we can be on let's say 4.5.2 without fix, but there isn't a better option available.
        elseif($CurrentNetVersion -lt $MinSupportNetVersion -and $MinSupportNetVersion -eq $RecommendedNetVersion)
        {
            Write-VerboseOutput("Current version of .NET is less than Min Supported Version. Need to upgrade to this version as soon as possible")
            $NetCheckObj.Supported = $false
            $NetCheckObj.RecommendedNetVersion = $false 
            [HealthChecker.NetVersionObject]$currentNetVersionObject = Get-NetFrameworkVersionFriendlyInfo -NetVersionKey $CurrentNetVersion.value__ -OSVersionName $OSVersionName
            [HealthChecker.NetVersionObject]$MinSupportNetVersionObject = Get-NetFrameworkVersionFriendlyInfo -NetVersionKey $MinSupportNetVersion.value__ -OSVersionName $OSVersionName
            $NetCheckObj.DisplayWording = "On .NET " + $currentNetVersionObject.FriendlyName + " and the minimum supported version is " + $MinSupportNetVersionObject.FriendlyName + ". Upgrade to this version as soon as possible."
        }
        #here we are assuming that we are able to get to a much better version of .NET then the min 
        elseif($CurrentNetVersion -lt $MinSupportNetVersion)
        {
            Write-VerboseOutput("Current Version of .NET is less than Min Supported Version. However, the recommended version is the one we want to upgrade to")
            $NetCheckObj.Supported = $false
            $NetCheckObj.RecommendedNetVersion = $false
            [HealthChecker.NetVersionObject]$currentNetVersionObject = Get-NetFrameworkVersionFriendlyInfo -NetVersionKey $CurrentNetVersion.value__ -OSVersionName $OSVersionName
            [HealthChecker.NetVersionObject]$MinSupportNetVersionObject = Get-NetFrameworkVersionFriendlyInfo -NetVersionKey $MinSupportNetVersion.value__ -OSVersionName $OSVersionName
            [HealthChecker.NetVersionObject]$RecommendedNetVersionObject = Get-NetFrameworkVersionFriendlyInfo -NetVersionKey $RecommendedNetVersion.value__ -OSVersionName $OSVersionName
            $NetCheckObj.DisplayWording = "On .NET " + $currentNetVersionObject.FriendlyName + " and the minimum supported version is " + $MinSupportNetVersionObject.FriendlyName + ", but the recommended version is " + $RecommendedNetVersionObject.FriendlyName + ". upgrade to this version as soon as possible." 
        }
        elseif($CurrentNetVersion -lt $RecommendedNetVersion)
        {
            Write-VerboseOutput("Current version is less than the recommended version, but we are at or higher than the Min Supported level. Should upgrade to the recommended version as soon as possible.")
            $NetCheckObj.Supported = $true
            $NetCheckObj.RecommendedNetVersion = $false 
            [HealthChecker.NetVersionObject]$currentNetVersionObject = Get-NetFrameworkVersionFriendlyInfo -NetVersionKey $CurrentNetVersion.value__ -OSVersionName $OSVersionName
            [HealthChecker.NetVersionObject]$RecommendedNetVersionObject = Get-NetFrameworkVersionFriendlyInfo -NetVersionKey $RecommendedNetVersion.value__ -OSVersionName $OSVersionName
            $NetCheckObj.DisplayWording = "On .NET " + $currentNetVersionObject.FriendlyName + " and the recommended version of .NET for this build of Exchange is " + $RecommendedNetVersionObject.FriendlyName + ". Upgrade to this version as soon as possible." 
        }
        elseif($CurrentNetVersion -gt $RecommendedNetVersion)
        {
            Write-VerboseOutput("Current version is greater than the recommended version. This is an unsupported state.")
            $NetCheckObj.Supported = $false
            $NetCheckObj.RecommendedNetVersion = $false 
            [HealthChecker.NetVersionObject]$currentNetVersionObject = Get-NetFrameworkVersionFriendlyInfo -NetVersionKey $CurrentNetVersion.value__ -OSVersionName $OSVersionName
            [HealthChecker.NetVersionObject]$RecommendedNetVersionObject = Get-NetFrameworkVersionFriendlyInfo -NetVersionKey $RecommendedNetVersion.value__ -OSVersionName $OSVersionName
            $NetCheckObj.DisplayWording = "On .NET " + $currentNetVersionObject.FriendlyName + " and the max recommnded version of .NET for this build of Exchange is " + $RecommendedNetVersionObject.FriendlyName + ". Correctly remove the .NET version that you are on and reinstall the recommended max value. Generic catch message for current .NET version being greater than Max .NET version, so ask or lookup on the correct steps to address this issue."
        }
        else
        {
            $NetCheckObj.Error = $true
            Write-VerboseOutput("unknown version of .net detected or combination with Exchange build")
        }

        Return $NetCheckObj
    }

    switch($exBuildObj.ExchangeVersion)
    {
        ([HealthChecker.ExchangeVersion]::Exchange2013)
            {
                Write-VerboseOutput("Exchange 2013 Detected...checking .NET version")
				#change -lt to -le as we don't support CU12 with 4.6.1 
                if($exBuildObj.CU -le ([HealthChecker.ExchangeCULevel]::CU12))
                {
                    $NetCheckObj = Check-NetVersionToExchangeVersion -CurrentNetVersion $NetVersion -MinSupportNetVersion Net4d5d2wFix -RecommendedNetVersion Net4d5d2wFix
                }
                elseif($exBuildObj.CU -lt ([HealthChecker.ExchangeCULevel]::CU15))
                {
                    $NetCheckObj = Check-NetVersionToExchangeVersion -CurrentNetVersion $NetVersion -MinSupportNetVersion Net4d5d2wFix -RecommendedNetVersion Net4d6d1wFix
                }
                elseif($exBuildObj.CU -eq ([HealthChecker.ExchangeCULevel]::CU15))
                {
                    $NetCheckObj = Check-NetVersionToExchangeVersion -CurrentNetVersion $NetVersion -MinSupportNetVersion Net4d5d2wFix -RecommendedNetVersion Net4d6d2
                    $NetCheckObj.DisplayWording = $NetCheckObj.DisplayWording + " NOTE: Starting with CU16 we will require .NET 4.6.2 before you can install this version of Exchange." 
                }
                elseif($exBuildObj.CU -lt ([HealthChecker.ExchangeCULevel]::CU19))
                {
                    $NetCheckObj = Check-NetVersionToExchangeVersion -CurrentNetVersion $NetVersion -MinSupportNetVersion Net4d6d2 -RecommendedNetVersion Net4d6d2
                }
                else
                {
                    $NetCheckObj = Check-NetVersionToExchangeVersion -CurrentNetVersion $NetVersion -MinSupportNetVersion Net4d6d2 -RecommendedNetVersion Net4d7d1
                }


                break;
                
            }
        ([HealthChecker.ExchangeVersion]::Exchange2016)
            {
                Write-VerboseOutput("Exchange 2016 detected...checking .NET version")

                if($exBuildObj.CU -lt [HealthChecker.ExchangeCULevel]::CU2)
                {
                    $NetCheckObj = Check-NetVersionToExchangeVersion -CurrentNetVersion $NetVersion -MinSupportNetVersion Net4d5d2wFix -RecommendedNetVersion Net4d5d2wFix
                }
                elseif($exBuildObj.CU -eq [HealthChecker.ExchangeCULevel]::CU2)
                {
                    $NetCheckObj = Check-NetVersionToExchangeVersion -CurrentNetVersion $NetVersion -MinSupportNetVersion Net4d5d2wFix -RecommendedNetVersion Net4d6d1wFix 
                }
                elseif($exBuildObj.CU -eq [HealthChecker.ExchangeCULevel]::CU3)
                {
                    if($OSVersionName -eq [HealthChecker.OSVersionName]::Windows2016)
                    {
                        $NetCheckObj = Check-NetVersionToExchangeVersion -CurrentNetVersion $NetVersion -MinSupportNetVersion Net4d5d2wFix -RecommendedNetVersion Net4d6d2
                        $NetCheckObj.DisplayWording = $NetCheckObj.DisplayWording + " NOTE: Starting with CU16 we will require .NET 4.6.2 before you can install this version of Exchange."
                    }
                    else
                    {
                        $NetCheckObj = Check-NetVersionToExchangeVersion -CurrentNetVersion $NetVersion -MinSupportNetVersion Net4d5d2wFix -RecommendedNetVersion Net4d6d1wFix
                    }
                }
                elseif($exBuildObj.CU -eq [HealthChecker.ExchangeCULevel]::CU4)
                {
                    $NetCheckObj = Check-NetVersionToExchangeVersion -CurrentNetVersion $NetVersion -MinSupportNetVersion Net4d5d2wFix -RecommendedNetVersion Net4d6d2 
                    $NetCheckObj.DisplayWording = $NetCheckObj.DisplayWording + " NOTE: Starting with CU5 we will require .NET 4.6.2 before you can install this version of Exchange."
                }
                elseif($exBuildObj.CU -lt [HealthChecker.ExchangeCULevel]::CU8)
                {
                    $NetCheckObj = Check-NetVersionToExchangeVersion -CurrentNetVersion $NetVersion -MinSupportNetVersion Net4d6d2 -RecommendedNetVersion Net4d6d2 
                }
                else
                {
                    $NetCheckObj = Check-NetVersionToExchangeVersion -CurrentNetVersion $NetVersion -MinSupportNetVersion Net4d6d2 -RecommendedNetVersion Net4d7d1
                }
                

                break;
            }
        ([HealthChecker.ExchangeVersion]::Exchange2019)
            {
                Write-VerboseOutput("Exchange 2019 detected...checking .NET version")
                if($exBuildObj.CU -lt [HealthChecker.ExchangeCULevel]::CU2)
                {
                    $NetCheckObj = Check-NetVersionToExchangeVersion -CurrentNetVersion $NetVersion -MinSupportNetVersion Net4d7d1 -RecommendedNetVersion Net4d7d2
                }

            }
        default {$NetCheckObj.Error = $true; Write-VerboseOutput("Error trying to determine major version of Exchange for .NET fix level")}
    }

    return $NetCheckObj

}

Function Get-MapiFEAppPoolGCMode{
param(
[Parameter(Mandatory=$true)][string]$Machine_Name
)
    Write-VerboseOutput("Calling: Get-MapiFEAppPoolGCMode")
    Write-VerboseOutput("Passed: {0}" -f $Machine_Name)
    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Machine_Name)
    $RegLocation = "SOFTWARE\Microsoft\ExchangeServer\v15\Setup\"
    $RegKey = $Reg.OpenSubKey($RegLocation)
    $MapiConfig = ("{0}bin\MSExchangeMapiFrontEndAppPool_CLRConfig.config" -f $RegKey.GetValue("MsiInstallPath"))
    Write-VerboseOutput("Mapi FE App Pool Config Location: {0}" -f $MapiConfig)
    $mapiGCMode = "Unknown"

    Function Get-MapiConfigGCSetting {
    param(
        [Parameter(Mandatory=$true)][string]$ConfigPath
    )
        if(Test-Path $ConfigPath)
        {
            $xml = [xml](Get-Content $ConfigPath)
            $rString =  $xml.configuration.runtime.gcServer.enabled
            return $rString
        }
        else 
        {
            Return "Unknown"    
        }
    }

    try 
    {
        if($Machine_Name -ne $env:COMPUTERNAME)
        {
            Write-VerboseOutput("Calling Get-MapiConfigGCSetting via Invoke-Command")
            $mapiGCMode = Invoke-Command -ComputerName $Machine_Name -ScriptBlock ${Function:Get-MapiConfigGCSetting} -ArgumentList $MapiConfig
        }
        else 
        {
            Write-VerboseOutput("Calling Get-MapiConfigGCSetting via local session")
            $mapiGCMode = Get-MapiConfigGCSetting -ConfigPath $MapiConfig    
        }
        
    }
    catch
    {
        #don't need to do anything here
    }

    Write-VerboseOutput("Returning GC Mode: {0}" -f $mapiGCMode)
    return $mapiGCMode
}

Function Get-ExchangeUpdates {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name,
[Parameter(Mandatory=$true)][HealthChecker.ExchangeVersion]$ExchangeVersion
)
    Write-VerboseOutput("Calling: Get-ExchangeUpdates")
    Write-VerboseOutput("Passed: " + $Machine_Name)
    Write-VerboseOutput("Passed: {0}" -f $ExchangeVersion.ToString())
    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Machine_Name)
    $RegLocation = $null 
    if([HealthChecker.ExchangeVersion]::Exchange2013 -eq $ExchangeVersion)
    {
        $RegLocation = "SOFTWARE\Microsoft\Updates\Exchange 2013"
    }
    else 
    {
        $RegLocation = "SOFTWARE\Microsoft\Updates\Exchange 2016"
    }
    $RegKey= $Reg.OpenSubKey($RegLocation)
    if($RegKey -ne $null)
    {
        $IU = $RegKey.GetSubKeyNames()
        if($IU -ne $null)
        {
            Write-VerboseOutput("Detected fixes installed on the server")
            $fixes = @()
            foreach($key in $IU)
            {
                $IUKey = $Reg.OpenSubKey($RegLocation + "\" + $key)
                $IUName = $IUKey.GetValue("PackageName")
                Write-VerboseOutput("Found: " + $IUName)
                $fixes += $IUName
            }
            return $fixes
        }
    }
    return $null
}

Function Get-ServerRole {
param(
[Parameter(Mandatory=$true)][object]$ExchangeServerObj
)
    Write-VerboseOutput("Calling: Get-ServerRole")
    $roles = $ExchangeServerObj.ServerRole.ToString()
    Write-VerboseOutput("Roll: " + $roles)
    #Need to change this to like because of Exchange 2010 with AIO with the hub role.
    if($roles -like "Mailbox, ClientAccess*")
    {
        return [HealthChecker.ServerRole]::MultiRole
    }
    elseif($roles -eq "Mailbox")
    {
        return [HealthChecker.ServerRole]::Mailbox
    }
    elseif($roles -eq "Edge")
    {
        return [HealthChecker.ServerRole]::Edge
    }
    elseif($roles -like "*ClientAccess*")
    {
        return [HealthChecker.ServerRole]::ClientAccess
    }
    else
    {
        return [HealthChecker.ServerRole]::None
    }
}

Function Build-ExchangeInformationObject {
param(
[Parameter(Mandatory=$true)][HealthChecker.HealthExchangeServerObject]$HealthExSvrObj
)
    $Machine_Name = $HealthExSvrObj.ServerName
    $OSVersionName = $HealthExSvrObj.OSVersion.OSVersion
    Write-VerboseOutput("Calling: Build-ExchangeInformationObject")
    Write-VerboseOutput("Passed: $Machine_Name")

    [HealthChecker.ExchangeInformationObject]$exchInfoObject = New-Object -TypeName HealthChecker.ExchangeInformationObject
    $exchInfoObject.ExchangeServerObject = (Get-ExchangeServer -Identity $Machine_Name)
    $exchInfoObject.ExchangeVersion = (Get-ExchangeVersion -AdminDisplayVersion $exchInfoObject.ExchangeServerObject.AdminDisplayVersion) 
    $exchInfoObject.ExServerRole = (Get-ServerRole -ExchangeServerObj $exchInfoObject.ExchangeServerObject)

    #Exchange 2013 and 2016 things to check 
    if($exchInfoObject.ExchangeVersion -ge [HealthChecker.ExchangeVersion]::Exchange2013) 
    {
        Write-VerboseOutput("Exchange 2013 or greater detected")
        $HealthExSvrObj.NetVersionInfo = Build-NetFrameWorkVersionObject -Machine_Name $Machine_Name -OSVersionName $OSVersionName
        $versionObject =  $HealthExSvrObj.NetVersionInfo 
        [HealthChecker.ExchangeInformationTempObject]$tempObject = Get-ExchangeBuildInformation -AdminDisplayVersion $exchInfoObject.ExchangeServerObject.AdminDisplayVersion
        if($tempObject.Error -ne $true) 
        {
            Write-VerboseOutput("No error detected when getting temp information")
            $exchInfoObject.BuildReleaseDate = $tempObject.ReleaseDate
            $exchInfoObject.ExchangeBuildNumber = $tempObject.ExchangeBuildNumber
            $exchInfoObject.ExchangeFriendlyName = $tempObject.FriendlyName
            $exchInfoObject.InbetweenCUs = $tempObject.InbetweenCUs
            $exchInfoObject.SupportedExchangeBuild = $tempObject.SupportedCU
            $exchInfoObject.ExchangeBuildObject = $tempObject.ExchangeBuildObject 
            [HealthChecker.NetVersionCheckObject]$NetCheckObj = Check-DotNetFrameworkSupportedLevel -exBuildObj $exchInfoObject.ExchangeBuildObject -OSVersionName $OSVersionName -NetVersion $versionObject.NetVersion
            if($NetCheckObj.Error)
            {
                Write-Yellow "Warnign: Unable to determine if .NET is supported"
            }
            else
            {
                $versionObject.SupportedVersion = $NetCheckObj.Supported
                $versionObject.DisplayWording = $NetCheckObj.DisplayWording
                $exchInfoObject.RecommendedNetVersion = $NetCheckObj.RecommendedNetVersion

            }
            
        }
        else
        {
            Write-Yellow "Warning: Couldn't get acturate information on server: $Machine_Name"
        }

        
        $exchInfoObject.MapiHttpEnabled = (Get-OrganizationConfig).MapiHttpEnabled
        if($exchInfoObject.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2013 -and $exchInfoObject.MapiHttpEnabled)
        {
            $exchInfoObject.MapiFEAppGCEnabled = Get-MapiFEAppPoolGCMode -Machine_Name $Machine_Name
        }

        $exchInfoObject.KBsInstalled = Get-ExchangeUpdates -Machine_Name $Machine_Name -ExchangeVersion $exchInfoObject.ExchangeVersion
    }
    elseif($exchInfoObject.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2010)
    {
        Write-VerboseOutput("Exchange 2010 detected")
        $exchInfoObject.ExchangeFriendlyName = "Exchange 2010"
        $exchInfoObject.ExchangeBuildNumber = $exchInfoObject.ExchangeServerObject.AdminDisplayVersion
    }
    else
    {
        Write-Red "Error: Unknown version of Exchange detected for server: $Machine_Name"
    }

    if($exchInfoObject.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2013 -and $exchInfoObject.ExServerRole -eq [HealthChecker.ServerRole]::ClientAccess)
    {
        Write-VerboseOutput("Exchange 2013 CAS only detected. Not going to run Test-ServiceHealth against this server.")
    }
    else 
    {
        Write-VerboseOutput("Exchange 2013 CAS only not detected. Going to run Test-ServiceHealth against this server.")
        $exchInfoObject.ExchangeServicesNotRunning = Test-ServiceHealth -Server $Machine_Name | %{$_.ServicesNotRunning}
    }
	
    $HealthExSvrObj.ExchangeInformation = $exchInfoObject
    return $HealthExSvrObj

}


Function Build-HealthExchangeServerObject {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name
)

    Write-VerboseOutput("Calling: Build-HealthExchangeServerObject")
    Write-VerboseOutput("Passed: $Machine_Name")

    [HealthChecker.HealthExchangeServerObject]$HealthExSvrObj = New-Object -TypeName HealthChecker.HealthExchangeServerObject 
    $HealthExSvrObj.ServerName = $Machine_Name 
    $HealthExSvrObj.HardwareInfo = Build-HardwareObject -Machine_Name $Machine_Name 
    $HealthExSvrObj.OSVersion = Build-OperatingSystemObject -Machine_Name $Machine_Name  
    $HealthExSvrObj = Build-ExchangeInformationObject -HealthExSvrObj $HealthExSvrObj
    Write-VerboseOutput("Finished building health Exchange Server Object for server: " + $Machine_Name)
    return $HealthExSvrObj
}


Function Get-MailboxDatabaseAndMailboxStatistics {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name
)
    Write-VerboseOutput("Calling: Get-MailboxDatabaseAndMailboxStatistics")
    Write-VerboseOutput("Passed: " + $Machine_Name)

    $AllDBs = Get-MailboxDatabaseCopyStatus -server $Machine_Name -ErrorAction SilentlyContinue 
    $MountedDBs = $AllDBs | ?{$_.Status -eq 'Healthy'}
    if($MountedDBs.Count -gt 0)
    {
        Write-Grey("`tActive Database:")
        foreach($db in $MountedDBs)
        {
            Write-Grey("`t`t" + $db.Name)
        }
        $MountedDBs.DatabaseName | %{Write-VerboseOutput("Calculating User Mailbox Total for Active Database: $_"); $TotalActiveUserMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited).Count}
        Write-Grey("`tTotal Active User Mailboxes on server: " + $TotalActiveUserMailboxCount)
        $MountedDBs.DatabaseName | %{Write-VerboseOutput("Calculating Public Mailbox Total for Active Database: $_"); $TotalActivePublicFolderMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited -PublicFolder).Count}
        Write-Grey("`tTotal Active Public Folder Mailboxes on server: " + $TotalActivePublicFolderMailboxCount)
        Write-Grey("`tTotal Active Mailboxes on server " + $Machine_Name + ": " + ($TotalActiveUserMailboxCount + $TotalActivePublicFolderMailboxCount).ToString())
    }
    else
    {
        Write-Grey("`tNo Active Mailbox Databases found on server " + $Machine_Name + ".")
    }
    $HealthyDbs = $AllDBs | ?{$_.Status -eq 'Healthy'}
    if($HealthyDbs.count -gt 0)
    {
        Write-Grey("`r`n`tPassive Databases:")
        foreach($db in $HealthyDbs)
        {
            Write-Grey("`t`t" + $db.Name)
        }
        $HealthyDbs.DatabaseName | %{Write-VerboseOutput("`tCalculating User Mailbox Total for Passive Healthy Databases: $_"); $TotalPassiveUserMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited).Count}
        Write-Grey("`tTotal Passive user Mailboxes on Server: " + $TotalPassiveUserMailboxCount)
        $HealthyDbs.DatabaseName | %{Write-VerboseOutput("`tCalculating Passive Mailbox Total for Passive Healthy Databases: $_"); $TotalPassivePublicFolderMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited -PublicFolder).Count}
        Write-Grey("`tTotal Passive Public Mailboxes on server: " + $TotalPassivePublicFolderMailboxCount)
        Write-Grey("`tTotal Passive Mailboxes on server: " + ($TotalPassiveUserMailboxCount + $TotalPassivePublicFolderMailboxCount).ToString()) 
    }
    else
    {
        Write-Grey("`tNo Passive Mailboxes found on server " + $Machine_Name + ".")
    }

}

#This function will return a true if the version level is the same or greater than the CheckVersionObject - keeping it simple so it can be done remotely as well 
Function Get-BuildLevelVersionCheck {
param(
[Parameter(Mandatory=$true)][object]$ActualVersionObject,
[Parameter(Mandatory=$true)][object]$CheckVersionObject,
[Parameter(Mandatory=$false)][bool]$DebugFunction = $false
)
Add-Type -TypeDefinition @"
public enum VersionDetection 
{
    Unknown,
    Lower,
    Equal,
    Greater
}
"@
    #unsure of how we do build numbers for all types of DLLs on the OS, but we are going to try to cover all bases here and it is up to the caller to make sure that we are passing the correct values to be checking 
    #FileMajorPart
    if($ActualVersionObject.FileMajorPart -lt $CheckVersionObject.FileMajorPart){$FileMajorPart = [VersionDetection]::Lower}
    elseif($ActualVersionObject.FileMajorPart -eq $CheckVersionObject.FileMajorPart){$FileMajorPart = [VersionDetection]::Equal}
    elseif($ActualVersionObject.FileMajorPart -gt $CheckVersionObject.FileMajorPart){$FileMajorPart = [VersionDetection]::Greater}
    else{$FileMajorPart =  [VersionDetection]::Unknown}

    if($ActualVersionObject.FileMinorPart -lt $CheckVersionObject.FileMinorPart){$FileMinorPart = [VersionDetection]::Lower}
    elseif($ActualVersionObject.FileMinorPart -eq $CheckVersionObject.FileMinorPart){$FileMinorPart = [VersionDetection]::Equal}
    elseif($ActualVersionObject.FileMinorPart -gt $CheckVersionObject.FileMinorPart){$FileMinorPart = [VersionDetection]::Greater}
    else{$FileMinorPart = [VersionDetection]::Unknown}

    if($ActualVersionObject.FileBuildPart -lt $CheckVersionObject.FileBuildPart){$FileBuildPart = [VersionDetection]::Lower}
    elseif($ActualVersionObject.FileBuildPart -eq $CheckVersionObject.FileBuildPart){$FileBuildPart = [VersionDetection]::Equal}
    elseif($ActualVersionObject.FileBuildPart -gt $CheckVersionObject.FileBuildPart){$FileBuildPart = [VersionDetection]::Greater}
    else{$FileBuildPart = [VersionDetection]::Unknown}

    
    if($ActualVersionObject.FilePrivatePart -lt $CheckVersionObject.FilePrivatePart){$FilePrivatePart = [VersionDetection]::Lower}
    elseif($ActualVersionObject.FilePrivatePart -eq $CheckVersionObject.FilePrivatePart){$FilePrivatePart = [VersionDetection]::Equal}
    elseif($ActualVersionObject.FilePrivatePart -gt $CheckVersionObject.FilePrivatePart){$FilePrivatePart = [VersionDetection]::Greater}
    else{$FilePrivatePart = [VersionDetection]::Unknown}

    if($DebugFunction)
    {
        Write-VerboseOutput("ActualVersionObject - FileMajorPart: {0} FileMinorPart: {1} FileBuildPart: {2} FilePrivatePart: {3}" -f $ActualVersionObject.FileMajorPart, 
        $ActualVersionObject.FileMinorPart, $ActualVersionObject.FileBuildPart, $ActualVersionObject.FilePrivatePart)
        Write-VerboseOutput("CheckVersionObject - FileMajorPart: {0} FileMinorPart: {1} FileBuildPart: {2} FilePrivatePart: {3}" -f $CheckVersionObject.FileMajorPart,
        $CheckVersionObject.FileMinorPart, $CheckVersionObject.FileBuildPart, $CheckVersionObject.FilePrivatePart)
        Write-VerboseOutput("Switch Detection - FileMajorPart: {0} FileMinorPart: {1} FileBuildPart: {2} FilePrivatePart: {3}" -f $FileMajorPart, $FileMinorPart, $FileBuildPart, $FilePrivatePart)
    }

    if($FileMajorPart -eq [VersionDetection]::Greater){return $true}
    if($FileMinorPart -eq [VersionDetection]::Greater){return $true}
    if($FileBuildPart -eq [VersionDetection]::Greater){return $true}
    if($FilePrivatePart -ge [VersionDetection]::Equal){return $true}

    return $false
}

Function Get-CASLoadBalancingReport {

    Write-VerboseOutput("Calling: Get-CASLoadBalancingReport")
    Write-Yellow("Note: CAS Load Balancing Report has known issues with attempting to get counter from servers. If you see errors regarding 'Get-Counter path not valid', please ignore for the time being. This is going to be addressed in later versions")
    #Connection and requests per server and client type values
    $CASConnectionStats = @{}
    $TotalCASConnectionCount = 0
    $AutoDStats = @{}
    $TotalAutoDRequests = 0
    $EWSStats = @{}
    $TotalEWSRequests = 0
    $MapiHttpStats = @{}
    $TotalMapiHttpRequests = 0
    $EASStats = @{}
    $TotalEASRequests = 0
    $OWAStats = @{}
    $TotalOWARequests = 0
    $RpcHttpStats = @{}
    $TotalRpcHttpRequests = 0
    $CASServers = @()

    if($CasServerList -ne $null)
    {
		Write-Grey("Custom CAS server list is being used.  Only servers specified after the -CasServerList parameter will be used in the report.")
        foreach($cas in $CasServerList)
        {
            $CASServers += (Get-ExchangeServer $cas)
        }
    }
	elseif($SiteName -ne $null)
	{
		Write-Grey("Site filtering ON.  Only Exchange 2013/2016 CAS servers in " + $SiteName + " will be used in the report.")
		$CASServers = Get-ExchangeServer | ?{($_.IsClientAccessServer -eq $true) -and ($_.AdminDisplayVersion -Match "^Version 15") -and ($_.Site.Name -eq $SiteName)}
	}
    else
    {
		Write-Grey("Site filtering OFF.  All Exchange 2013/2016 CAS servers will be used in the report.")
        $CASServers = Get-ExchangeServer | ?{($_.IsClientAccessServer -eq $true) -and ($_.AdminDisplayVersion -Match "^Version 15")}
    }

	if($CASServers.Count -eq 0)
	{
		Write-Red("Error: No CAS servers found using the specified search criteria.")
		Exit
	}

    #Pull connection and request stats from perfmon for each CAS
    foreach($cas in $CASServers)
    {
        #Total connections
        $TotalConnectionCount = (Get-Counter ("\\" + $cas.Name + "\Web Service(Default Web Site)\Current Connections")).CounterSamples.CookedValue
        $CASConnectionStats.Add($cas.Name, $TotalConnectionCount)
        $TotalCASConnectionCount += $TotalConnectionCount

        #AutoD requests
        $AutoDRequestCount = (Get-Counter ("\\" + $cas.Name + "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_Autodiscover)\Requests Executing")).CounterSamples.CookedValue
        $AutoDStats.Add($cas.Name, $AutoDRequestCount)
        $TotalAutoDRequests += $AutoDRequestCount

        #EWS requests
        $EWSRequestCount = (Get-Counter ("\\" + $cas.Name + "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_EWS)\Requests Executing")).CounterSamples.CookedValue
        $EWSStats.Add($cas.Name, $EWSRequestCount)
        $TotalEWSRequests += $EWSRequestCount

        #MapiHttp requests
        $MapiHttpRequestCount = (Get-Counter ("\\" + $cas.Name + "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_mapi)\Requests Executing")).CounterSamples.CookedValue
        $MapiHttpStats.Add($cas.Name, $MapiHttpRequestCount)
        $TotalMapiHttpRequests += $MapiHttpRequestCount

        #EAS requests
        $EASRequestCount = (Get-Counter ("\\" + $cas.Name + "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_Microsoft-Server-ActiveSync)\Requests Executing")).CounterSamples.CookedValue
        $EASStats.Add($cas.Name, $EASRequestCount)
        $TotalEASRequests += $EASRequestCount

        #OWA requests
        $OWARequestCount = (Get-Counter ("\\" + $cas.Name + "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_owa)\Requests Executing")).CounterSamples.CookedValue
        $OWAStats.Add($cas.Name, $OWARequestCount)
        $TotalOWARequests += $OWARequestCount

        #RPCHTTP requests
        $RpcHttpRequestCount = (Get-Counter ("\\" + $cas.Name + "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_Rpc)\Requests Executing")).CounterSamples.CookedValue
        $RpcHttpStats.Add($cas.Name, $RpcHttpRequestCount)
        $TotalRpcHttpRequests += $RpcHttpRequestCount
    }

    #Report the results for connection count
    Write-Grey("")
    Write-Grey("Connection Load Distribution Per Server")
    Write-Grey("Total Connections: " + $TotalCASConnectionCount)
    #Calculate percentage of connection load
    $CASConnectionStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
    Write-Grey($_.Key + ": " + $_.Value + " Connections = " + [math]::Round((([int]$_.Value/$TotalCASConnectionCount)*100)) + "% Distribution")
    }

    #Same for each client type.  These are request numbers not connection numbers.
    #AutoD
    if($TotalAutoDRequests -gt 0)
    {
        Write-Grey("")
        Write-Grey("Current AutoDiscover Requests Per Server")
        Write-Grey("Total Requests: " + $TotalAutoDRequests)
        $AutoDStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
        Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value/$TotalAutoDRequests)*100)) + "% Distribution")
        }
    }

    #EWS
    if($TotalEWSRequests -gt 0)
    {
        Write-Grey("")
        Write-Grey("Current EWS Requests Per Server")
        Write-Grey("Total Requests: " + $TotalEWSRequests)
        $EWSStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
        Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value/$TotalEWSRequests)*100)) + "% Distribution")
        }
    }

    #MapiHttp
    if($TotalMapiHttpRequests -gt 0)
    {
        Write-Grey("")
        Write-Grey("Current MapiHttp Requests Per Server")
        Write-Grey("Total Requests: " + $TotalMapiHttpRequests)
        $MapiHttpStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
        Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value/$TotalMapiHttpRequests)*100)) + "% Distribution")
        }
    }

    #EAS
    if($TotalEASRequests -gt 0)
    {
        Write-Grey("")
        Write-Grey("Current EAS Requests Per Server")
        Write-Grey("Total Requests: " + $TotalEASRequests)
        $EASStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
        Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value/$TotalEASRequests)*100)) + "% Distribution")
        }
    }

    #OWA
    if($TotalOWARequests -gt 0)
    {
        Write-Grey("")
        Write-Grey("Current OWA Requests Per Server")
        Write-Grey("Total Requests: " + $TotalOWARequests)
        $OWAStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
        Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value/$TotalOWARequests)*100)) + "% Distribution")
        }
    }

    #RpcHttp
    if($TotalRpcHttpRequests -gt 0)
    {
        Write-Grey("")
        Write-Grey("Current RpcHttp Requests Per Server")
        Write-Grey("Total Requests: " + $TotalRpcHttpRequests)
        $RpcHttpStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
        Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value/$TotalRpcHttpRequests)*100)) + "% Distribution")
        }
    }

    Write-Grey("")

}


Function Verify-PagefileEqualMemoryPlus10{
param(
[Parameter(Mandatory=$true)][HealthChecker.PageFileObject]$page_obj,
[Parameter(Mandatory=$true)][HealthChecker.HardwareObject]$hardware_obj
)
    Write-VerboseOutput("Calling: Verify-PagefileEqualMemoryPlus10")
    Write-VerboseOutput("Passed: total memory: " + $hardware_obj.TotalMemory)
    Write-VerboseOutput("Passed: max page file size: " + $page_obj.MaxPageSize)
    $sReturnString = "Good"
    $iMemory = [System.Math]::Round(($hardware_obj.TotalMemory / 1048576) + 10)
    Write-VerboseOutput("Server Memory Plus 10 MB: " + $iMemory) 
    
    if($page_obj.MaxPageSize -lt $iMemory)
    {
        $sReturnString = "Page file is set to (" + $page_obj.MaxPageSize + ") which appears to be less than the Total System Memory plus 10 MB which is (" + $iMemory + ") this appears to be set incorrectly."
    }
    elseif($page_obj.MaxPageSize -gt $iMemory)
    {
        $sReturnString = "Page file is set to (" + $page_obj.MaxPageSize + ") which appears to be More than the Total System Memory plus 10 MB which is (" + $iMemory + ") this appears to be set incorrectly." 
    }

    return $sReturnString

}

Function Get-LmCompatibilityLevel {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name
)
    #LSA Reg Location "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa"
    #Check if valuename LmCompatibilityLevel exists, if not, then value is 3
    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Machine_Name)
    $RegKey = $reg.OpenSubKey("SYSTEM\CurrentControlSet\Control\Lsa")
    $RegValue = $RegKey.GetValue("LmCompatibilityLevel")
    If ($RegValue)
    {
        Return $RegValue
    }
    Else
    {
        Return 3
    }

}

Function Build-LmCompatibilityLevel {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name
)

    Write-VerboseOutput("Calling: Build-LmCompatibilityLevel")
    Write-VerboseOutput("Passed: $Machine_Name")

    [HealthChecker.ServerLmCompatibilityLevel]$ServerLmCompatObject = New-Object -TypeName HealthChecker.ServerLmCompatibilityLevel
    
    $ServerLmCompatObject.LmCompatibilityLevelRef = "https://technet.microsoft.com/en-us/library/cc960646.aspx"
    $ServerLmCompatObject.LmCompatibilityLevel    = Get-LmCompatibilityLevel $Machine_Name
    Switch ($ServerLmCompatObject.LmCompatibilityLevel)
    {
        0 {$ServerLmCompatObject.LmCompatibilityLevelDescription = "Clients use LM and NTLM authentication, but they never use NTLMv2 session security. Domain controllers accept LM, NTLM, and NTLMv2 authentication." }
        1 {$ServerLmCompatObject.LmCompatibilityLevelDescription = "Clients use LM and NTLM authentication, and they use NTLMv2 session security if the server supports it. Domain controllers accept LM, NTLM, and NTLMv2 authentication." }
        2 {$ServerLmCompatObject.LmCompatibilityLevelDescription = "Clients use only NTLM authentication, and they use NTLMv2 session security if the server supports it. Domain controller accepts LM, NTLM, and NTLMv2 authentication." }
        3 {$ServerLmCompatObject.LmCompatibilityLevelDescription = "Clients use only NTLMv2 authentication, and they use NTLMv2 session security if the server supports it. Domain controllers accept LM, NTLM, and NTLMv2 authentication." }
        4 {$ServerLmCompatObject.LmCompatibilityLevelDescription = "Clients use only NTLMv2 authentication, and they use NTLMv2 session security if the server supports it. Domain controller refuses LM authentication responses, but it accepts NTLM and NTLMv2." }
        5 {$ServerLmCompatObject.LmCompatibilityLevelDescription = "Clients use only NTLMv2 authentication, and they use NTLMv2 session security if the server supports it. Domain controller refuses LM and NTLM authentication responses, but it accepts NTLMv2." }
    }

    Return $ServerLmCompatObject
}

Function Display-KBHotfixCheckFailSafe {
param(
[Parameter(Mandatory=$true)][HealthChecker.HealthExchangeServerObject]$HealthExSvrObj
)

    Write-Grey("`r`nHotfix Check:")
    $2008HotfixList = $null
  $2008R2HotfixList = @("KB3004383")
  $2012HotfixList = $null
  $2012R2HotfixList = @("KB3041832")
  $2016HotfixList = @("KB3206632")
  
  
  Function Check-Hotfix 
  {
      param(
      [Parameter(Mandatory=$true)][Array]$Hotfixes,
      [Parameter(Mandatory=$true)][Array]$CheckListHotFixes
      )
      $hotfixesneeded = $false
      foreach($check in $CheckListHotFixes)
      {
          if($Hotfixes.Contains($check) -eq $false)
          {
              $hotfixesneeded = $true
              Write-Yellow("Warning: Hotfix " + $check + " is recommended for this OS and was not detected.  Please consider installing it to prevent performance issues. --- Note that this KB update may be superseded by another KB update. To verify, check the file versions in the KB against your machine. This is a temporary workaround till the script gets properly updated for all KB checks.")
          }
      }
      if($hotfixesneeded -eq $false)
      {
          Write-Grey("Hotfix check complete.  No action required.")
      }
  }

  switch($HealthExSvrObj.OSVersion.OSVersion) 
  {
      ([HealthChecker.OSVersionName]::Windows2008)
      {
          if($2008HotfixList -ne $null) {Check-Hotfix -Hotfixes $HealthExSvrObj.OSVersion.HotFixes.Hotfixid -CheckListHotFixes $2008HotfixList}
      }
      ([HealthChecker.OSVersionName]::Windows2008R2)
      {
          if($2008R2HotfixList -ne $null) {Check-Hotfix -Hotfixes $HealthExSvrObj.OSVersion.HotFixes.Hotfixid -CheckListHotFixes $2008R2HotfixList}
      }
      ([HealthChecker.OSVersionName]::Windows2012)
      {
          if($2012HotfixList -ne $null) {Check-Hotfix -Hotfixes $HealthExSvrObj.OSVersion.HotFixes.Hotfixid -CheckListHotFixes $2012HotfixList}
      }
      ([HealthChecker.OSVersionName]::Windows2012R2)
      {
          if($2012R2HotfixList -ne $null) {Check-Hotfix -Hotfixes $HealthExSvrObj.OSVersion.HotFixes.Hotfixid -CheckListHotFixes $2012R2HotfixList}
      }
      ([HealthChecker.OSVersionName]::Windows2016)
      {
          if($2016HotfixList -ne $null) {Check-Hotfix -Hotfixes $HealthExSvrObj.OSVersion.HotFixes.Hotfixid -CheckListHotFixes $2016HotfixList}
      }

      default {}
  }
}

Function Get-BuildVersionObjectFromString {
param(
[Parameter(Mandatory=$true)][string]$BuildString 
)
    $aBuild = $BuildString.Split(".")
    if($aBuild.Count -ge 4)
    {
        $obj = New-Object PSCustomObject 
        $obj | Add-Member -MemberType NoteProperty -Name FileMajorPart -Value ([System.Convert]::ToInt32($aBuild[0]))
        $obj | Add-Member -MemberType NoteProperty -Name FileMinorPart -Value ([System.Convert]::ToInt32($aBuild[1]))
        $obj | Add-Member -MemberType NoteProperty -Name FileBuildPart -Value ([System.Convert]::ToInt32($aBuild[2]))
        $obj | Add-Member -MemberType NoteProperty -Name FilePrivatePart -Value ([System.Convert]::ToInt32($aBuild[3]))
        return $obj 
    }
    else 
    {
        Return "Error"    
    }
}

#Addressed issue 69
Function Display-KBHotFixCompareIssues{
param(
[Parameter(Mandatory=$true)][HealthChecker.HealthExchangeServerObject]$HealthExSvrObj
)
    Write-VerboseOutput("Calling: Display-KBHotFixCompareIssues")
    Write-VerboseOutput("For Server: {0}" -f $HealthExSvrObj.ServerName)

    #$HotFixInfo = $HealthExSvrObj.OSVersion.HotFixes.Hotfixid
    $HotFixInfo = @() 
    foreach($Hotfix in $HealthExSvrObj.OSVersion.HotFixes)
    {
        $HotFixInfo += $Hotfix.HotfixId 
    }

    $serverOS = $HealthExSvrObj.OSVersion.OSVersion
    if($serverOS -eq ([HealthChecker.OSVersionName]::Windows2008))
    {
        Write-VerboseOutput("Windows 2008 detected")
        $KBHashTable = @{"KB4295656"="KB4345397"}
    }
    elseif($serverOS -eq ([HealthChecker.OSVersionName]::Windows2008R2))
    {
        Write-VerboseOutput("Windows 2008 R2 detected")
        $KBHashTable = @{"KB4338823"="KB4345459";"KB4338818"="KB4338821"}
    }
    elseif($serverOS -eq ([HealthChecker.OSVersionName]::Windows2012))
    {
        Write-VerboseOutput("Windows 2012 detected")
        $KBHashTable = @{"KB4338820"="KB4345425";"KB4338830"="KB4338816"}
    }
    elseif($serverOS -eq ([HealthChecker.OSVersionName]::Windows2012R2))
    {
        Write-VerboseOutput("Windows 2012 R2 detected")
        $KBHashTable = @{"KB4338824"="KB4345424";"KB4338815"="KB4338831"}
    }
    elseif($serverOS -eq ([HealthChecker.OSVersionName]::Windows2016))
    {
        Write-VerboseOutput("Windows 2016 detected")
        $KBHashTable = @{"KB4338814"="KB4345418"}
    }

    if($HotFixInfo -ne $null)
    {
        if($KBHashTable -ne $null)
        {
            foreach($key in $KBHashTable.Keys)
            {
                foreach($problemKB in $HotFixInfo)
                {
                    if($problemKB -eq $key)
                    {
                        Write-VerboseOutput("Found Impacted {0}" -f $key)
                        $foundFixKB = $false 
                        foreach($fixKB in $HotFixInfo)
                        {
                            if($fixKB -eq ($KBHashTable[$key]))
                            {
                                Write-VerboseOutput("Found {0} that fixes the issue" -f ($KBHashTable[$key]))
                                $foundFixKB = $true 
                            }

                        }
                        if(-not($foundFixKB))
                        {
                            Write-Break
                            Write-Break
                            Write-Red("July Update detected: Error --- Problem {0} detected without the fix {1}. This can cause odd issues to occur on the system. See https://blogs.technet.microsoft.com/exchange/2018/07/16/issue-with-july-updates-for-windows-on-an-exchange-server/" -f $key, ($KBHashTable[$key]))
                        }
                    }
                }
            }
        }
        else
        {
            Write-VerboseOutput("KBHashTable was null. July Update issue not checked.")
        }
    }
    else 
    {
        Write-VerboseOutput("No hotfixes were detected on the server")    
    }

    
}

Function Display-KBHotfixCheck {
param(
[Parameter(Mandatory=$true)][HealthChecker.HealthExchangeServerObject]$HealthExSvrObj
)
    Write-VerboseOutput("Calling: Display-KBHotfixCheck")
    Write-VerboseOutput("For Server: {0}" -f $HealthExSvrObj.ServerName)
    
    $HotFixInfo = $HealthExSvrObj.OSVersion.HotFixInfo
    $KBsToCheckAgainst = Get-HotFixListInfo -OS_Version $HealthExSvrObj.OSVersion.OSVersion 
    $FailSafe = $false
    if($KBsToCheckAgainst -ne $null)
    {
        foreach($KB in $HotFixInfo)
        {
            $KBName = $KB.KBName 
            foreach($KBInfo in $KB.KBInfo)
            {
                if(-not ($KBInfo.Error))
                {
                    #First need to find the correct KB to compare against 
                    $i = 0 
                    $iMax = $KBsToCheckAgainst.Count 
                    while($i -lt $iMax)
                    {
                        if($KBsToCheckAgainst[$i].KBName -eq $KBName)
                        {
                            break; 
                        }
                        else 
                        {
                            $i++ 
                        }
                    }
                    $allPass = $true 
                    foreach($CheckFile in $KBInfo)
                    {
                        $ii = 0 
                        $iMax = $KBsToCheckAgainst[$i].FileInformation.Count 
                        while($ii -lt $iMax)
                        {
                            if($KBsToCheckAgainst[$i].FileInformation[$ii].FriendlyFileName -eq $CheckFile.FriendlyName)
                            {
                                break; 
                            }
                            else 
                            {
                                $ii++    
                            }
                        }
                        
                        $ServerBuild = Get-BuildVersionObjectFromString -BuildString $CheckFile.BuildVersion 
                        $CheckVersion = Get-BuildVersionObjectFromString -BuildString $KBsToCheckAgainst[$i].FileInformation[$ii].BuildVersion
                        if(-not (Get-BuildLevelVersionCheck -ActualVersionObject $ServerBuild -CheckVersionObject $CheckVersion -DebugFunction $false))
                        {
                            $allPass = $false
                        }

                    }
                    
                    $KBInfo | Add-Member -MemberType NoteProperty -Name Passed -Value $allPass   
                }
                else 
                {
                    #If an error has occurred, that means we failed to find the files 
                    $FailSafe = $true 
                    break;    
                }
            }
        }
        if($FailSafe)
        {
            Display-KBHotfixCheckFailSafe -HealthExSvrObj $HealthExSvrObj
        }
        else 
        {
            Write-Grey("`r`nHotfix Check:")
            foreach($KBInfo in $HotFixInfo)
            {
                
                $allPass = $true 
                foreach($KBs in $KBInfo.KBInfo)
                {
                    if(-not ($KBs.Passed))
                    {
                        $allPass = $false
                    }
                }
                $dString = if($allPass){"is Installed"}else{"is recommended for this OS and was not detected.  Please consider installing it to prevent performance issues."}
                if($allPass)
                {
                    Write-Grey("{0} {1}" -f $KBInfo.KBName, ($dString))
                }
                else 
                {
                    Write-Yellow("{0} {1}" -f $KBInfo.KBName, ($dString))    
                }
                
            }
        }
    }

}

Function Display-ResultsToScreen {
param(
[Parameter(Mandatory=$true)][HealthChecker.HealthExchangeServerObject]$HealthExSvrObj
)
    Write-VerboseOutput("Calling: Display-ResultsToScreen")
    Write-VerboseOutput("For Server: " + $HealthExSvrObj.ServerName)

    ####################
    #Header information#
    ####################

    Write-Green("Exchange Health Checker version " + $healthCheckerVersion)
    Write-Green("System Information Report for " + $HealthExSvrObj.ServerName + " on " + $date) 
    Write-Break
    Write-Break
    ###############################
    #OS, System, and Exchange Info#
    ###############################

    if($HealthExSvrObj.HardwareInfo.ServerType -eq [HealthChecker.ServerType]::VMWare -or $HealthExSvrObj.HardwareInfo.ServerType -eq [HealthChecker.ServerType]::HyperV)
    {
        Write-Yellow($VirtualizationWarning) 
        Write-Break
        Write-Break
    }
    Write-Grey("Hardware/OS/Exchange Information:");
    Write-Grey("`tHardware Type: " + $HealthExSvrObj.HardwareInfo.ServerType.ToString())
    if($HealthExSvrObj.HardwareInfo.ServerType -eq [HealthChecker.ServerType]::Physical)
    {
        Write-Grey("`tManufacturer: " + $HealthExSvrObj.HardwareInfo.Manufacturer)
        Write-Grey("`tModel: " + $HealthExSvrObj.HardwareInfo.Model) 
    }

    Write-Grey("`tOperating System: " + $HealthExSvrObj.OSVersion.OperatingSystemName)
    Write-Grey("`tExchange: " + $HealthExSvrObj.ExchangeInformation.ExchangeFriendlyName)
    Write-Grey("`tBuild Number: " + $HealthExSvrObj.ExchangeInformation.ExchangeBuildNumber)
    #If IU or Security Hotfix detected
    if($HealthExSvrObj.ExchangeInformation.KBsInstalled -ne $null)
    {
        Write-Grey("`tExchange IU or Security Hotfix Detected")
        foreach($kb in $HealthExSvrObj.ExchangeInformation.KBsInstalled)
        {
            Write-Yellow("`t`t{0}" -f $kb)
        }
    }

    if($HealthExSvrObj.ExchangeInformation.SupportedExchangeBuild -eq $false -and $HealthExSvrObj.ExchangeInformation.ExchangeVersion -ge [HealthChecker.ExchangeVersion]::Exchange2013)
    {
        $Dif_Days = ($date - ([System.Convert]::ToDateTime([DateTime]$HealthExSvrObj.ExchangeInformation.BuildReleaseDate))).Days
        Write-Red("`tError: Out of date Cumulative Update.  Please upgrade to one of the two most recently released Cumulative Updates. Currently running on a build that is " + $Dif_Days + " Days old")
    }
    if($HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2013 -and ($HealthExSvrObj.ExchangeInformation.ExServerRole -ne [HealthChecker.ServerRole]::Edge -and $HealthExSvrObj.ExchangeInformation.ExServerRole -ne [HealthChecker.ServerRole]::MultiRole))
    {
        Write-Yellow("`tServer Role: " + $HealthExSvrObj.ExchangeInformation.ExServerRole.ToString() + " --- Warning: Multi-Role servers are recommended") 
    }
    else
    {
        Write-Grey("`tServer Role: " + $HealthExSvrObj.ExchangeInformation.ExServerRole.ToString())
    }

    #MAPI/HTTP
    if($HealthExSvrObj.ExchangeInformation.ExchangeVersion -ge [HealthChecker.ExchangeVersion]::Exchange2013)
    {
        Write-Grey("`tMAPI/HTTP Enabled: {0}" -f $HealthExSvrObj.ExchangeInformation.MapiHttpEnabled)
        if($HealthExSvrObj.ExchangeInformation.MapiHttpEnabled -eq $true -and $HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2013)
        {
            if($HealthExSvrObj.ExchangeInformation.MapiFEAppGCEnabled -eq "false" -and 
            $HealthExSvrObj.HardwareInfo.TotalMemory -ge 21474836480)
            {
                Write-Red("`t`tMAPI Front End App Pool GC Mode: Workstation --- Error")
                Write-Yellow("`t`tTo Fix this issue go into the file MSExchangeMapiFrontEndAppPool_CLRConfig.config in the Exchange Bin direcotry and change the GCServer to true and recycle the MAPI Front End App Pool")
            }
            elseif($HealthExSvrObj.ExchangeInformation.MapiFEAppGCEnabled -eq "false")
            {
                Write-Yellow("`t`tMapi Front End App Pool GC Mode: Workstation --- Warning")
                Write-Yellow("`t`tYou could be seeing some GC issues within the Mapi Front End App Pool. However, you don't have enough memory installed on the system to recommend switching the GC mode by default without consulting a support professional.")
            }
            elseif($HealthExSvrObj.ExchangeInformation.MapiFEAppGCEnabled -eq "true")
            {
                Write-Green("`t`tMapi Front End App Pool GC Mode: Server")
            }
            else 
            {
                Write-Yellow("Mapi Front End App Pool GC Mode: Unknown --- Warning")    
            }
        }
    }

    ###########
    #Page File#
    ###########

    Write-Grey("Pagefile Settings:")
    if($HealthExSvrObj.HardwareInfo.AutoPageFile) 
    {
        Write-Red("`tError: System is set to automatically manage the pagefile size. This is not recommended.") 
    }
    elseif($HealthExSvrObj.OSVersion.PageFile.PageFile.Count -gt 1)
    {
        Write-Red("`tError: Multiple page files detected. This has been known to cause performance issues please address this.")
    }
    elseif($HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2010) 
    {
        #Exchange 2010, we still recommend that we have the page file set to RAM + 10 MB 
        #Page File Size is Less than Physical Memory Size Plus 10 MB 
        #https://technet.microsoft.com/en-us/library/cc431357(v=exchg.80).aspx
        $sDisplay = Verify-PagefileEqualMemoryPlus10 -page_obj $HealthExSvrObj.OSVersion.PageFile -hardware_obj $HealthExSvrObj.HardwareInfo
        if($sDisplay -eq "Good")
        {
            Write-Grey("`tPagefile Size: " + $HealthExSvrObj.OSVersion.PageFile.MaxPageSize)
        }
        else
        {
            Write-Yellow("`tPagefile Size: {0} --- Warning: Article: https://technet.microsoft.com/en-us/library/cc431357(v=exchg.80).aspx" -f $sDisplay)
            Write-Yellow("`tNote: Please double check page file setting, as WMI Object Win32_ComputerSystem doesn't report the best value for total memory available") 
        }
    }
    #Exchange 2013+ with memory greater than 32 GB. Should be set to 32 + 10 MB for a value 
    #32GB = 1024 * 1024 * 1024 * 32 = 34,359,738,368 
    elseif($HealthExSvrObj.HardwareInfo.TotalMemory -ge 34359738368)
    {
        if($HealthExSvrObj.OSVersion.PageFile.MaxPageSize -eq 32778)
        {
            Write-Grey("`tPagefile Size: " + $HealthExSvrObj.OSVersion.PageFile.MaxPageSize)
        }
        else
        {
            Write-Yellow("`tPagefile Size: " + $HealthExSvrObj.OSVersion.PageFile.MaxPageSize + " --- Warning: Pagefile should be capped at 32778 MB for 32 GB Plus 10 MB - Article: https://technet.microsoft.com/en-us/library/dn879075(v=exchg.150).aspx")
        }
    }
    #Exchange 2013 with page file size that should match total memory plus 10 MB 
    else
    {
        $sDisplay = Verify-PagefileEqualMemoryPlus10 -page_obj $HealthExSvrObj.OSVersion.PageFile -hardware_obj $HealthExSvrObj.HardwareInfo
        if($sDisplay -eq "Good")
        {
            Write-Grey("`tPagefile Size: " + $HealthExSvrObj.OSVersion.PageFile.MaxPageSize)
        }
        else
        {
            Write-Yellow("`tPagefile Size: {0} --- Warning: Article: https://technet.microsoft.com/en-us/library/dn879075(v=exchg.150).aspx" -f $sDisplay)
            Write-Yellow("`tNote: Please double check page file setting, as WMI Object Win32_ComputerSystem doesn't report the best value for total memory available") 
        }
    }

    ################
    #.NET FrameWork#
    ################

    if($HealthExSvrObj.ExchangeInformation.ExchangeVersion -gt [HealthChecker.ExchangeVersion]::Exchange2010)
    {
        Write-Grey(".NET Framework:")
        
        if($HealthExSvrObj.NetVersionInfo.SupportedVersion)
        {
            if($HealthExSvrObj.ExchangeInformation.RecommendedNetVersion)
            {
                Write-Green("`tVersion: " + $HealthExSvrObj.NetVersionInfo.FriendlyName)
            }
            else
            {
                Write-Yellow("`tDetected Version: " + $HealthExSvrObj.NetVersionInfo.FriendlyName + " --- Warning: " + $HealthExSvrObj.NetVersionInfo.DisplayWording)
            }
        }
        else
        {
                Write-Red("`tDetected Version: " + $HealthExSvrObj.NetVersionInfo.FriendlyName + " --- Error: " + $HealthExSvrObj.NetVersionInfo.DisplayWording)
        }

    }

    ################
    #Power Settings#
    ################
    Write-Grey("Power Settings:")
    if($HealthExSvrObj.OSVersion.HighPerformanceSet)
    {
        Write-Green("`tPower Plan: " + $HealthExSvrObj.OSVersion.PowerPlanSetting)
    }
    elseif($HealthExSvrObj.OSVersion.PowerPlan -eq $null) 
    {
        Write-Red("`tPower Plan: Not Accessible --- Error")
    }
    else
    {
        Write-Red("`tPower Plan: " + $HealthExSvrObj.OSVersion.PowerPlanSetting + " --- Error: High Performance Power Plan is recommended")
    }

    ################
    #Pending Reboot#
    ################

    Write-Grey("Server Pending Reboot:")

    if($HealthExSvrObj.OSVersion.ServerPendingReboot)
    {
        Write-Red("`tTrue --- Error: This can cause issues if files haven't been properly updated.")
    }
    else 
    {
        Write-Green("`tFalse")    
    }

	#####################
	#Http Proxy Settings#
	#####################

	Write-Grey("Http Proxy Setting:")
	if($HealthExSvrObj.OSVersion.HttpProxy -eq "<None>")
	{
		Write-Green("`tSetting: {0}" -f $HealthExSvrObj.OSVersion.HttpProxy)
	}
	else
	{
		Write-Yellow("`tSetting: {0} --- Warning: This could cause connectivity issues." -f $HealthExSvrObj.OSVersion.HttpProxy)
	}

    ##################
    #Network Settings#
    ##################

    Write-Grey("NIC settings per active adapter:")
    if($HealthExSvrObj.OSVersion.OSVersion -ge [HealthChecker.OSVersionName]::Windows2012R2)
    {
        foreach($adapter in $HealthExSvrObj.OSVersion.NetworkAdapters)
        {
            Write-Grey(("`tInterface Description: {0} [{1}] " -f $adapter.Description, $adapter.Name))
            if($HealthExSvrObj.HardwareInfo.ServerType -eq [HealthChecker.ServerType]::Physical)
            {
                if((New-TimeSpan -Start $date -End $adapter.DriverDate).Days -lt [int]-365)
                {
                    Write-Yellow("`t`tWarning: NIC driver is over 1 year old. Verify you are at the latest version.")
                }
                Write-Grey("`t`tDriver Date: " + $adapter.DriverDate)
                Write-Grey("`t`tDriver Version: " + $adapter.DriverVersion)
                Write-Grey("`t`tLink Speed: " + $adapter.LinkSpeed)
            }
            else
            {
                Write-Yellow("`t`tLink Speed: Cannot be accurately determined due to virtualized hardware")
            }
            if($adapter.RSSEnabled -eq "NoRSS")
            {
                Write-Yellow("`t`tRSS: No RSS Feature Detected.")
            }
            elseif($adapter.RSSEnabled -eq "True")
            {
                Write-Green("`t`tRSS: Enabled")
            }
            else
            {
                Write-Yellow("`t`tRSS: Disabled --- Warning: Enabling RSS is recommended.")
            }
            
        }

    }
    else
    {
        Write-Grey("NIC settings per active adapter:")
        Write-Yellow("`tMore detailed NIC settings can be detected if both the local and target server are running on Windows 2012 R2 or later.")
        
        foreach($adapter in $HealthExSvrObj.OSVersion.NetworkAdapters)
        {
            Write-Grey("`tInterface Description: " + $adapter.Description)
            if($HealthExSvrObj.HardwareInfo.ServerType -eq [HealthChecker.ServerType]::Physical)
            {
                Write-Grey("`tLink Speed: " + $adapter.LinkSpeed)
            }
            else 
            {
                Write-Yellow("`tLink Speed: Cannot be accurately determined due to virtualization hardware")    
            }
        }
        
    }
    if($HealthExSvrObj.ExchangeInformation.ExchangeVersion -ne [HealthChecker.ExchangeVersion]::Exchange2010)
    {
        if($HealthExSvrObj.OSVersion.NetworkAdapters.Count -gt 1 -and ($HealthExSvrObj.ExchangeInformation.ExServerRole -eq [HealthChecker.ServerRole]::Mailbox -or $HealthExSvrObj.ExchangeInformation.ExServerRole -eq [HealthChecker.ServerRole]::MultiRole))
        {
            Write-Yellow("`t`tMultiple active network adapters detected. Exchange 2013 or greater may not need separate adapters for MAPI and replication traffic.  For details please refer to https://technet.microsoft.com/en-us/library/29bb0358-fc8e-4437-8feb-d2959ed0f102(v=exchg.150)#NR")
        }
    }

    #######################
    #Processor Information#
    #######################
    Write-Grey("Processor/Memory Information")
    Write-Grey("`tProcessor Type: " + $HealthExSvrObj.HardwareInfo.Processor.ProcessorName)
    #Hyperthreading check
    <#if($HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors -gt 24 -and $HealthExSvrObj.ExchangeInformation.ExchangeVersion -ne [HealthChecker.ExchangeVersion]::Exchange2010)
    {
        if($HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors -gt $HealthExSvrObj.HardwareInfo.Processor.NumberOfPhysicalCores)
        {
            Write-Red("`tHyper-Threading Enabled: Yes --- Error")
            Write-Red("`tError: More than 24 logical cores detected.  Please disable Hyper-Threading.  For details see`r`n`thttp://blogs.technet.com/b/exchange/archive/2015/06/19/ask-the-perf-guy-how-big-is-too-big.aspx")
        }
        else
        {
            Write-Green("`tHyper-Threading Enabled: No")
            Write-Red("`tError: More than 24 physical cores detected.  This is not recommended.  For details see`r`n`thttp://blogs.technet.com/b/exchange/archive/2015/06/19/ask-the-perf-guy-how-big-is-too-big.aspx")
        }
    }
    elseif($HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors -gt $HealthExSvrObj.HardwareInfo.Processor.NumberOfPhysicalCores)
    {
        if($HealthExSvrObj.HardwareInfo.Processor.ProcessorName.StartsWith("AMD"))
        {
            Write-Yellow("`tHyper-Threading Enabled: Yes --- Warning: Enabling Hyper-Threading is not recommended")
            Write-Yellow("`tThis script may incorrectly report that Hyper-Threading is enabled on certain AMD processors.  Check with the manufacturer to see if your model supports SMT.")
        }
        else
        {
            Write-Yellow("`tHyper-Threading Enabled: Yes --- Warning: Enabling Hyper-Threading is not recommended")
        }
    }
    #>

    Function Check-MaxCoresCount {
    param(
    [Parameter(Mandatory=$true)][HealthChecker.HealthExchangeServerObject]$HealthExSvrObj
    )
        if($HealthExSvrObj.ExchangeInformation.ExchangeVersion -ge [HealthChecker.ExchangeVersion]::Exchange2019 -and 
        $HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors -gt 48)
        {
            Write-Red("`tError: More than 48 cores detected, this goes against best practices. For details see `r`n`thttps://blogs.technet.microsoft.com/exchange/2018/07/24/exchange-server-2019-public-preview/")
        }
        elseif(($HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2013 -or 
        $HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2016) -and 
        $HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors -gt 24)
        {
            Write-Red("`tError: More than 24 cores detected, this goes against best practices. For details see `r`n`thttps://blogs.technet.microsoft.com/exchange/2015/06/19/ask-the-perf-guy-how-big-is-too-big/")
        }
    }

    #First, see if we are hyperthreading
    if($HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors -gt $HealthExSvrObj.HardwareInfo.Processor.NumberOfPhysicalCores)
    {
        #Hyperthreading enabled 
        Write-Red("`tHyper-Threading Enabled: Yes --- Error: Having Hyper-Threading enabled goes against best practices. Please disable as soon as possible.")
        #AMD might not have the correct logic here. Throwing warning about this. 
        if($HealthExSvrObj.HardwareInfo.Processor.ProcessorName.StartsWith("AMD"))
        {
            Write-Yellow("`tThis script may incorrectly report that Hyper-Threading is enabled on certain AMD processors.  Check with the manufacturer to see if your model supports SMT.")
        }
        Check-MaxCoresCount -HealthExSvrObj $HealthExSvrObj
    }
    else
    {
        Write-Green("`tHyper-Threading Enabled: No")
        Check-MaxCoresCount -HealthExSvrObj $HealthExSvrObj
    }
    #Number of Processors - Number of Processor Sockets. 
    if($HealthExSvrObj.HardwareInfo.Processor.NumberOfProcessors -gt 2)
    {
        Write-Red("`tNumber of Processors: {0} - Error: We recommend only having 2 Processor Sockets." -f $HealthExSvrObj.HardwareInfo.Processor.NumberOfProcessors)
    }
    else 
    {
        Write-Green("`tNumber of Processors: {0}" -f $HealthExSvrObj.HardwareInfo.Processor.NumberOfProcessors)
    }

    #Core count
    if(($HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors -gt 24 -and 
    $HealthExSvrObj.ExchangeInformation.ExchangeVersion -lt [HealthChecker.ExchangeVersion]::Exchange2019) -or 
    ($HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors -gt 48))
    {
        Write-Yellow("`tNumber of Physical Cores: " + $HealthExSvrObj.HardwareInfo.Processor.NumberOfPhysicalCores)
        Write-Yellow("`tNumber of Logical Cores: " + $HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors)
    }
    else
    {
        Write-Green("`tNumber of Physical Cores: " + $HealthExSvrObj.HardwareInfo.Processor.NumberOfPhysicalCores)
        Write-Green("`tNumber of Logical Cores: " + $HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors)
    }

    #NUMA BIOS CHECK - AKA check to see if we can properly see all of our cores on the box. 
	if($HealthExSvrObj.HardwareInfo.Model -like "*ProLiant*")
	{
		if($HealthExSvrObj.HardwareInfo.Processor.EnvProcessorCount -eq -1)
		{
			Write-Yellow("`tNUMA Group Size Optimization: Unable to determine --- Warning: If this is set to Clustered, this can cause multiple types of issues on the server")
		}
		elseif($HealthExSvrObj.HardwareInfo.Processor.EnvProcessorCount -ne $HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors)
		{
			Write-Red("`tNUMA Group Size Optimization: BIOS Set to Clustered --- Error: This setting should be set to Flat. By having this set to Clustered, we will see multiple different types of issues.")
		}
		else
		{
			Write-Green("`tNUMA Group Size Optimization: BIOS Set to Flat")
		}
	}
	else
	{
		if($HealthExSvrObj.HardwareInfo.Processor.EnvProcessorCount -eq -1)
		{
			Write-Yellow("`tAll Processor Cores Visible: Unable to determine --- Warning: If we aren't able to see all processor cores from Exchange, we could see performance related issues.")
		}
		elseif($HealthExSvrObj.HardwareInfo.Processor.EnvProcessorCount -ne $HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors)
		{
			Write-Red("`tAll Processor Cores Visible: Not all Processor Cores are visable to Exchange and this will cause a performance impact --- Error")
		}
		else
		{
			Write-Green("`tAll Processor Cores Visible: Passed")
		}
	}
    if($HealthExSvrObj.HardwareInfo.Processor.ProcessorIsThrottled)
    {
        #We are set correctly at the OS layer
        if($HealthExSvrObj.OSVersion.HighPerformanceSet)
        {
            Write-Red("`tError: Processor speed is being throttled. Power plan is set to `"High performance`", so it is likely that we are throttling in the BIOS of the computer settings")
        }
        else
        {
            Write-Red("`tError: Processor speed is being throttled. Power plan isn't set to `"High performance`". Change this ASAP because you are throttling your CPU and is likely causing issues.")
            Write-Yellow("`tNote: This change doesn't require a reboot and takes affect right away. Re-run the script after doing so")
        }
        Write-Red("`tCurrent Processor Speed: " + $HealthExSvrObj.HardwareInfo.Processor.CurrentMegacyclesPerCore + " --- Error: Processor appears to be throttled. This will cause performance issues. See Max Processor Speed to see what this should be at.")
        Write-Red("`tMax Processor Speed: " + $HealthExSvrObj.HardwareInfo.Processor.MaxMegacyclesPerCore )
    }
    else
    {
        Write-Grey("`tMegacycles Per Core: " + $HealthExSvrObj.HardwareInfo.Processor.MaxMegacyclesPerCore)
    }
    
    #Memory Going to check for greater than 96GB of memory for Exchange 2013
    #The value that we shouldn't be greater than is 103,079,215,104 (96 * 1024 * 1024 * 1024) 
    #Exchange 2016 we are going to check to see if there is over 192 GB https://blogs.technet.microsoft.com/exchange/2017/09/26/ask-the-perf-guy-update-to-scalability-guidance-for-exchange-2016/
    #For Exchange 2016 the value that we shouldn't be greater than is 206,158,430,208 (192 * 1024 * 1024 * 1024)
    #For Exchange 2019 the value that we shouldn't be greater than is 274,877,906,944 (256 * 1024 * 1024 * 1024) - https://blogs.technet.microsoft.com/exchange/2018/07/24/exchange-server-2019-public-preview/
    $totalPhysicalMemory = [System.Math]::Round($HealthExSvrObj.HardwareInfo.TotalMemory / 1024 /1024 /1024) 
    if($HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2019 -and
        $HealthExSvrObj.HardwareInfo.TotalMemory -gt 274877906944)
    {
        Write-Yellow("`tPhysical Memory: {0} GB --- We recommend for the best performance to be scaled at or below 256 GB of Memory." -f $totalPhysicalMemory)
    }
    elseif($HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2016 -and
        $HealthExSvrObj.HardwareInfo.TotalMemory -gt 206158430208)
    {
        Write-Yellow("`tPhysical Memory: {0} GB --- We recommend for the best performance to be scaled at or below 192 GB of Memory." -f $totalPhysicalMemory)
    }
    elseif($HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2013 -and
     $HealthExSvrObj.HardwareInfo.TotalMemory -gt 103079215104)
    {
        Write-Yellow ("`tPhysical Memory: " + $totalPhysicalMemory + " GB --- Warning: We recommend for the best performance to be scaled at or below 96GB of Memory. However, having higher memory than this has yet to be linked directly to a MAJOR performance issue of a server.")
    }
    else
    {
        Write-Grey("`tPhysical Memory: " + $totalPhysicalMemory + " GB") 
    }

    ################
	#Service Health#
	################
    #We don't want to run if the server is 2013 CAS role or if the Role = None
    if(-not(($HealthExSvrObj.ExchangeInformation.ExServerRole -eq [HealthChecker.ServerRole]::None) -or 
        (($HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2013) -and 
        ($HealthExSvrObj.ExchangeInformation.ExServerRole -eq [HealthChecker.ServerRole]::ClientAccess))))
    {
		if($HealthExSvrObj.ExchangeInformation.ExchangeServicesNotRunning)
	    {
		    Write-Yellow("`r`nWarning: The following services are not running:")
        $HealthExSvrObj.ExchangeInformation.ExchangeServicesNotRunning | %{Write-Grey($_)}
	    }

    }

    #################
	#TCP/IP Settings#
	#################
    Write-Grey("`r`nTCP/IP Settings:")
    if($HealthExSvrObj.OSVersion.TCPKeepAlive -eq 0)
    {
        Write-Red("Error: The TCP KeepAliveTime value is not specified in the registry.  Without this value the KeepAliveTime defaults to two hours, which can cause connectivity and performance issues between network devices such as firewalls and load balancers depending on their configuration.  To avoid issues, add the KeepAliveTime REG_DWORD entry under HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\Tcpip\Parameters and set it to a value between 900000 and 1800000 decimal.  You want to ensure that the TCP idle timeout value gets higher as you go out from Exchange, not lower.  For example if the Exchange server has a value of 30 minutes, the Load Balancer could have an idle timeout of 35 minutes, and the firewall could have an idle timeout of 40 minutes.  Please note that this change will require a restart of the system.  Refer to the sections `"CAS Configuration`" and `"Load Balancer Configuration`" in this blog post for more details:  https://blogs.technet.microsoft.com/exchange/2016/05/31/checklist-for-troubleshooting-outlook-connectivity-in-exchange-2013-and-2016-on-premises/")
    }
    elseif($HealthExSvrObj.OSVersion.TCPKeepAlive -lt 900000 -or $HealthExSvrObj.OSVersion.TCPKeepAlive -gt 1800000)
    {
        Write-Yellow("Warning: The TCP KeepAliveTime value is not configured optimally.  It is currently set to " + $HealthExSvrObj.OSVersion.TCPKeepAlive + ". This can cause connectivity and performance issues between network devices such as firewalls and load balancers depending on their configuration.  To avoid issues, set the HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\Tcpip\Parameters\KeepAliveTime registry entry to a value between 15 and 30 minutes (900000 and 1800000 decimal).  You want to ensure that the TCP idle timeout gets higher as you go out from Exchange, not lower.  For example if the Exchange server has a value of 30 minutes, the Load Balancer could have an idle timeout of 35 minutes, and the firewall could have an idle timeout of 40 minutes.  Please note that this change will require a restart of the system.  Refer to the sections `"CAS Configuration`" and `"Load Balancer Configuration`" in this blog post for more details:  https://blogs.technet.microsoft.com/exchange/2016/05/31/checklist-for-troubleshooting-outlook-connectivity-in-exchange-2013-and-2016-on-premises/")
    }
    else
    {
        Write-Green("The TCP KeepAliveTime value is configured optimally (" + $HealthExSvrObj.OSVersion.TCPKeepAlive + ")")
    }

    ###############################
	#LmCompatibilityLevel Settings#
	###############################
    Write-Grey("`r`nLmCompatibilityLevel Settings:")
    Write-Grey("`tLmCompatibilityLevel is set to: " + $HealthExSvrObj.OSVersion.LmCompat.LmCompatibilityLevel)
    Write-Grey("`tLmCompatibilityLevel Description: " + $HealthExSvrObj.OSVersion.LmCompat.LmCompatibilityLevelDescription)
    Write-Grey("`tLmCompatibilityLevel Ref: " + $HealthExSvrObj.OSVersion.LmCompat.LmCompatibilityLevelRef)

	##############
	#Hotfix Check#
	##############
    
    if($HealthExSvrObj.ExchangeInformation.ExchangeVersion -ne [HealthChecker.ExchangeVersion]::Exchange2010)
    {
        Display-KBHotfixCheck -HealthExSvrObj $HealthExSvrObj
    }
    Display-KBHotFixCompareIssues -HealthExSvrObj $HealthExSvrObj


    Write-Grey("`r`n`r`n")

}

Function Build-ServerObject
{
    param(
    [Parameter(Mandatory=$true)][HealthChecker.HealthExchangeServerObject]$HealthExSvrObj
    )

    $ServerObject = New-Object –TypeName PSObject

    $ServerObject | Add-Member –MemberType NoteProperty –Name ServerName –Value $HealthExSvrObj.ServerName

    if($HealthExSvrObj.HardwareInfo.ServerType -eq [HealthChecker.ServerType]::VMWare -or $HealthExSvrObj.HardwareInfo.ServerType -eq [HealthChecker.ServerType]::HyperV)
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name VirtualServer –Value "Yes"
    }
    else
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name VirtualServer –Value "No"
    }

    $ServerObject | Add-Member –MemberType NoteProperty –Name HardwareType –Value $HealthExSvrObj.HardwareInfo.ServerType.ToString()

    if($HealthExSvrObj.HardwareInfo.ServerType -eq [HealthChecker.ServerType]::Physical)
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name Manufacturer –Value $HealthExSvrObj.HardwareInfo.Manufacturer
        $ServerObject | Add-Member –MemberType NoteProperty –Name Model –Value $HealthExSvrObj.HardwareInfo.Model
    }

    $ServerObject | Add-Member –MemberType NoteProperty –Name OperatingSystem –Value $HealthExSvrObj.OSVersion.OperatingSystemName
    $ServerObject | Add-Member –MemberType NoteProperty –Name Exchange –Value $HealthExSvrObj.ExchangeInformation.ExchangeFriendlyName
    $ServerObject | Add-Member –MemberType NoteProperty –Name BuildNumber –Value $HealthExSvrObj.ExchangeInformation.ExchangeBuildNumber

    #If IU or Security Hotfix detected
    if($HealthExSvrObj.ExchangeInformation.KBsInstalled -ne $null)
    {
        $KBArray = @()
        foreach($kb in $HealthExSvrObj.ExchangeInformation.KBsInstalled)
        {
            $KBArray += $kb
        }

        $ServerObject | Add-Member –MemberType NoteProperty –Name InterimUpdatesInstalled -Value $KBArray
    }

    if($HealthExSvrObj.ExchangeInformation.SupportedExchangeBuild -eq $false -and $HealthExSvrObj.ExchangeInformation.ExchangeVersion -ge [HealthChecker.ExchangeVersion]::Exchange2013)
    {
        $Dif_Days = ((Get-Date) - ([System.Convert]::ToDateTime([DateTime]$HealthExSvrObj.ExchangeInformation.BuildReleaseDate))).Days
        $ServerObject | Add-Member –MemberType NoteProperty –Name BuildDaysOld –Value $Dif_Days
		$ServerObject | Add-Member –MemberType NoteProperty –Name SupportedExchangeBuild -Value $HealthExSvrObj.ExchangeInformation.SupportedExchangeBuild
    }
	else
	{
		$ServerObject | Add-Member –MemberType NoteProperty –Name SupportedExchangeBuild -Value $True
	}

    if($HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2013 -and ($HealthExSvrObj.ExchangeInformation.ExServerRole -ne [HealthChecker.ServerRole]::Edge -and $HealthExSvrObj.ExchangeInformation.ExServerRole -ne [HealthChecker.ServerRole]::MultiRole))
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name ServerRole -Value "Not Multirole"
    }
    else
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name ServerRole -Value $HealthExSvrObj.ExchangeInformation.ExServerRole.ToString()
    }


    ###########
    #Page File#
    ###########

    if($HealthExSvrObj.HardwareInfo.AutoPageFile) 
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name AutoPageFile -Value "Yes"
    }
    else
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name AutoPageFile -Value "No"
    }
    
    if($HealthExSvrObj.OSVersion.PageFile.PageFile.Count -gt 1)
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name MultiplePageFiles -Value "Yes"
    }
    else
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name MultiplePageFiles -Value "No"
    }
    
    
    if($HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2010) 
    {
        #Exchange 2010, we still recommend that we have the page file set to RAM + 10 MB 
        #Page File Size is Less than Physical Memory Size Plus 10 MB 
        #https://technet.microsoft.com/en-us/library/cc431357(v=exchg.80).aspx
        $sDisplay = Verify-PagefileEqualMemoryPlus10 -page_obj $HealthExSvrObj.OSVersion.PageFile -hardware_obj $HealthExSvrObj.HardwareInfo
        if($sDisplay -eq "Good")
        {

            $ServerObject | Add-Member –MemberType NoteProperty –Name PagefileSize -Value "$($HealthExSvrObj.OSVersion.PageFile.MaxPageSize)"
      			$ServerObject | Add-Member –MemberType NoteProperty –Name PagefileSizeSetRight -Value "Yes"

        }
        else
        {
            $ServerObject | Add-Member –MemberType NoteProperty –Name PagefileSize -Value "$($HealthExSvrObj.OSVersion.PageFile.MaxPageSize)"
			$ServerObject | Add-Member –MemberType NoteProperty –Name PagefileSizeSetRight -Value "No"
        }
    }

    #Exchange 2013+ with memory greater than 32 GB. Should be set to 32 + 10 MB for a value 
    #32GB = 1024 * 1024 * 1024 * 32 = 34,359,738,368 
    elseif($HealthExSvrObj.HardwareInfo.TotalMemory -ge 34359738368)

    {
        if($HealthExSvrObj.OSVersion.PageFile.MaxPageSize -eq 32778)
        {
            $ServerObject | Add-Member –MemberType NoteProperty –Name PagefileSize -Value "$($HealthExSvrObj.OSVersion.PageFile.MaxPageSize)"
			$ServerObject | Add-Member –MemberType NoteProperty –Name PagefileSizeSetRight -Value "Yes"
        }
        else
        {
            $ServerObject | Add-Member –MemberType NoteProperty –Name PagefileSize -Value "$($HealthExSvrObj.OSVersion.PageFile.MaxPageSize)"
			$ServerObject | Add-Member –MemberType NoteProperty –Name PagefileSizeSetRight -Value "No"
        }
    }
    #Exchange 2013 with page file size that should match total memory plus 10 MB 
    else
    {
        $sDisplay = Verify-PagefileEqualMemoryPlus10 -page_obj $HealthExSvrObj.OSVersion.PageFile -hardware_obj $HealthExSvrObj.HardwareInfo
        if($sDisplay -eq "Good")
        {
            $ServerObject | Add-Member –MemberType NoteProperty –Name PagefileSize -Value "$($HealthExSvrObj.OSVersion.PageFile.MaxPageSize)"
			$ServerObject | Add-Member –MemberType NoteProperty –Name PagefileSizeSetRight -Value "Yes"
        }
        else
        {
            $ServerObject | Add-Member –MemberType NoteProperty –Name PagefileSize -Value "$($HealthExSvrObj.OSVersion.PageFile.MaxPageSize)"
			$ServerObject | Add-Member –MemberType NoteProperty –Name PagefileSizeSetRight -Value "No"
        }
    }

    ################
    #.NET FrameWork#
    ################
    
    if($HealthExSvrObj.ExchangeInformation.ExchangeVersion -gt [HealthChecker.ExchangeVersion]::Exchange2010)
    {

        if($HealthExSvrObj.NetVersionInfo.SupportedVersion)
        {
            if($HealthExSvrObj.ExchangeInformation.RecommendedNetVersion)
            {
                $ServerObject | Add-Member –MemberType NoteProperty –Name DotNetVersion -Value $HealthExSvrObj.NetVersionInfo.FriendlyName
            }
            else
            {
                $ServerObject | Add-Member –MemberType NoteProperty –Name DotNetVersion -Value "$($HealthExSvrObj.NetVersionInfo.FriendlyName) $($HealthExSvrObj.NetVersionInfo.DisplayWording)"
            }
        }
        else
        {
                $ServerObject | Add-Member –MemberType NoteProperty –Name DotNetVersion -Value "$($HealthExSvrObj.NetVersionInfo.FriendlyName) $($HealthExSvrObj.NetVersionInfo.DisplayWording)"
        }

    }


    ################
    #Power Settings#
    ################

    if($HealthExSvrObj.OSVersion.HighPerformanceSet)
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name PowerPlan -Value $HealthExSvrObj.OSVersion.PowerPlanSetting
		$ServerObject | Add-Member –MemberType NoteProperty –Name PowerPlanSetRight -Value $True
    }
    elseif($HealthExSvrObj.OSVersion.PowerPlan -eq $null) 
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name PowerPlan -Value "Not Accessible"
		$ServerObject | Add-Member –MemberType NoteProperty –Name PowerPlanSetRight -Value $False
    }
    else
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name PowerPlan -Value "$($HealthExSvrObj.OSVersion.PowerPlanSetting)"
		$ServerObject | Add-Member –MemberType NoteProperty –Name PowerPlanSetRight -Value $False
    }



    #####################
	#Http Proxy Settings#
	#####################

    $ServerObject | Add-Member –MemberType NoteProperty –Name HTTPProxy -Value $HealthExSvrObj.OSVersion.HttpProxy


    ##################
    #Network Settings#
    ##################

    if($HealthExSvrObj.OSVersion.OSVersion -ge [HealthChecker.OSVersionName]::Windows2012R2)
    {
        if((($HealthExSvrObj.OSVersion.NetworkAdapters).count) -gt 1)
        {
			$i = 1
			
			$ServerObject | Add-Member –MemberType NoteProperty –Name NumberNICs ($HealthExSvrObj.OSVersion.NetworkAdapters).count

            foreach($adapter in $HealthExSvrObj.OSVersion.NetworkAdapters)
            {
                $ServerObject | Add-Member –MemberType NoteProperty –Name NIC_Name_$($i) -Value $adapter.Name
                $ServerObject | Add-Member –MemberType NoteProperty –Name NIC_Description_$($i) -Value $adapter.Description

                if($HealthExSvrObj.HardwareInfo.ServerType -eq [HealthChecker.ServerType]::Physical)
                {
                    if((New-TimeSpan -Start (Get-Date) -End $adapter.DriverDate).Days -lt [int]-365)
                    {
                        $ServerObject | Add-Member –MemberType NoteProperty –Name NIC_Driver_$($i) -Value "Outdated (>1 Year Old)"
                    }
                    $ServerObject | Add-Member –MemberType NoteProperty –Name NIC_DriverDate_$($i) -Value $adapter.DriverDate
                    $ServerObject | Add-Member –MemberType NoteProperty –Name NIC_DriverVersion_$($i) -Value $adapter.DriverVersion
                    $ServerObject | Add-Member –MemberType NoteProperty –Name NIC_LinkSpeed_$($i) -Value $adapter.LinkSpeed
                }
                else
                {
                    $ServerObject | Add-Member –MemberType NoteProperty –Name NIC_LinkSpeed_$($i) -Value "VM - Not Applicable"
                }
                if($adapter.RSSEnabled -eq "NoRSS")
                {
                    $ServerObject | Add-Member –MemberType NoteProperty –Name NIC_RSS_$($i) -Value "NoRSS"
                }
                elseif($adapter.RSSEnabled -eq "True")
                {
                    $ServerObject | Add-Member –MemberType NoteProperty –Name NIC_RSS_$($i) -Value  "Enabled"
                }
                else
                {
                    $ServerObject | Add-Member –MemberType NoteProperty –Name NIC_RSS_$($i) -Value "Disabled"
                }
				
				$i++
            }

               

             
        }
    }
    else
    {
        
        foreach($adapter in $HealthExSvrObj.OSVersion.NetworkAdapters)
        {
			$ServerObject | Add-Member –MemberType NoteProperty –Name NIC_Name_1 -Value $adapter.Name
            $ServerObject | Add-Member –MemberType NoteProperty –Name NIC_Description_1 -Value $adapter.Description
            if($HealthExSvrObj.HardwareInfo.ServerType -eq [HealthChecker.ServerType]::Physical)
            {
                $ServerObject | Add-Member –MemberType NoteProperty –Name NIC_LinkSpeed_1 -Value $adapter.LinkSpeed
            }
            else 
            {
                $ServerObject | Add-Member –MemberType NoteProperty –Name NIC_LinkSpeed_1 -Value "VM - Not Applicable"  
            }
        }
        
    }
    if($HealthExSvrObj.ExchangeInformation.ExchangeVersion -ne [HealthChecker.ExchangeVersion]::Exchange2010)
    {
        if($HealthExSvrObj.OSVersion.NetworkAdapters.Count -gt 1 -and ($HealthExSvrObj.ExchangeInformation.ExServerRole -eq [HealthChecker.ServerRole]::Mailbox -or $HealthExSvrObj.ExchangeInformation.ExServerRole -eq [HealthChecker.ServerRole]::MultiRole))
        {
            $ServerObject | Add-Member –MemberType NoteProperty –Name E2013MultipleNICs -Value "Yes"
        }
    }

    #######################
    #Processor Information#
    #######################

    $ServerObject | Add-Member –MemberType NoteProperty –Name ProcessorName -Value $HealthExSvrObj.HardwareInfo.Processor.ProcessorName

    #Recommendation by PG is no more than 24 cores (this should include logical with Hyper Threading
    if($HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors -gt 24 -and $HealthExSvrObj.ExchangeInformation.ExchangeVersion -ne [HealthChecker.ExchangeVersion]::Exchange2010)
    {
        if($HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors -gt $HealthExSvrObj.HardwareInfo.Processor.NumberOfPhysicalCores)
        {
            $ServerObject | Add-Member –MemberType NoteProperty –Name HyperThreading -Value "Enabled"
        }
        else
        {
            $ServerObject | Add-Member –MemberType NoteProperty –Name HyperThreading -Value "Disabled"
        }
    }
    elseif($HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors -gt $HealthExSvrObj.HardwareInfo.Processor.NumberOfPhysicalCores)
    {
        if($HealthExSvrObj.HardwareInfo.Processor.ProcessorName.StartsWith("AMD"))
        {
            $ServerObject | Add-Member –MemberType NoteProperty –Name AMD_HyperThreading -Value "Enabled"
        }
        else
        {
            $ServerObject | Add-Member –MemberType NoteProperty –Name AMD_HyperThreading -Value "Disabled"
        }
    }
    else
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name HyperThreading -Value "Disabled"
    }

    $ServerObject | Add-Member –MemberType NoteProperty –Name NumberOfProcessors -Value $HealthExSvrObj.HardwareInfo.Processor.NumberOfProcessors

    if($HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors -gt 24)
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name NumberOfPhysicalCores -Value $HealthExSvrObj.HardwareInfo.Processor.NumberOfPhysicalCores
        $ServerObject | Add-Member –MemberType NoteProperty –Name NumberOfLogicalProcessors -Value $HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors
    }
    else
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name NumberOfPhysicalCores -Value $HealthExSvrObj.HardwareInfo.Processor.NumberOfPhysicalCores
        $ServerObject | Add-Member –MemberType NoteProperty –Name NumberOfLogicalProcessors -Value $HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors
    }
	if($HealthExSvrObj.HardwareInfo.Model -like "*ProLiant*")
	{
		if($HealthExSvrObj.HardwareInfo.Processor.EnvProcessorCount -eq -1)
		{
			$ServerObject | Add-Member –MemberType NoteProperty –Name NUMAGroupSize -Value "Undetermined"
		}
		elseif($HealthExSvrObj.HardwareInfo.Processor.EnvProcessorCount -ne $HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors)
		{
			$ServerObject | Add-Member –MemberType NoteProperty –Name NUMAGroupSize -Value "Clustered"
		}
		else
		{
			$ServerObject | Add-Member –MemberType NoteProperty –Name NUMAGroupSize -Value "Flat"
		}
	}
	else
	{
		if($HealthExSvrObj.HardwareInfo.Processor.EnvProcessorCount -eq -1)
		{
			$ServerObject | Add-Member –MemberType NoteProperty –Name AllProcCoresVisible -Value "Undetermined"
		}
		elseif($HealthExSvrObj.HardwareInfo.Processor.EnvProcessorCount -ne $HealthExSvrObj.HardwareInfo.Processor.NumberOfLogicalProcessors)
		{
			$ServerObject | Add-Member –MemberType NoteProperty –Name AllProcCoresVisible -Value "No"
		}
		else
		{
			$ServerObject | Add-Member –MemberType NoteProperty –Name AllProcCoresVisible -Value "Yes"
		}
	}
    if($HealthExSvrObj.HardwareInfo.Processor.ProcessorIsThrottled)
    {
        #We are set correctly at the OS layer
        if($HealthExSvrObj.OSVersion.HighPerformanceSet)
        {
            $ServerObject | Add-Member –MemberType NoteProperty –Name ProcessorSpeed -Value "Throttled, Not Power Plan"
        }
        else
        {
            $ServerObject | Add-Member –MemberType NoteProperty –Name ProcessorSpeed -Value "Throttled, Power Plan"
        }
        $ServerObject | Add-Member –MemberType NoteProperty –Name CurrentProcessorSpeed -Value $HealthExSvrObj.HardwareInfo.Processor.CurrentMegacyclesPerCore
        $ServerObject | Add-Member –MemberType NoteProperty –Name MaxProcessorSpeed -Value $HealthExSvrObj.HardwareInfo.Processor.MaxMegacyclesPerCore
    }
    else
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name MaxMegacyclesPerCore -Value $HealthExSvrObj.HardwareInfo.Processor.MaxMegacyclesPerCore
    }


    #Memory Going to check for greater than 96GB of memory for Exchange 2013
    #The value that we shouldn't be greater than is 103,079,215,104 (96 * 1024 * 1024 * 1024) 
    #Exchange 2016 we are going to check to see if there is over 192 GB https://blogs.technet.microsoft.com/exchange/2017/09/26/ask-the-perf-guy-update-to-scalability-guidance-for-exchange-2016/
    #For Exchange 2016 the value that we shouldn't be greater than is 206,158,430,208 (192 * 1024 * 1024 * 1024)
    $totalPhysicalMemory = [System.Math]::Round($HealthExSvrObj.HardwareInfo.TotalMemory / 1024 /1024 /1024) 

    $ServerObject | Add-Member –MemberType NoteProperty –Name TotalPhysicalMemory -Value "$totalPhysicalMemory GB"
	
	if($HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2016 -and
        $HealthExSvrObj.HardwareInfo.TotalMemory -gt 206158430208)
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name E2016MemoryRight -Value $False
    }
    elseif($HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2013 -and
     $HealthExSvrObj.HardwareInfo.TotalMemory -gt 103079215104)
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name E2013MemoryRight -Value $False
    }
	else
	{
		$ServerObject | Add-Member –MemberType NoteProperty –Name E2016MemoryRight -Value $True
		$ServerObject | Add-Member –MemberType NoteProperty –Name E2013MemoryRight -Value $True
	}

    ################
	#Service Health#
	################
    #We don't want to run if the server is 2013 CAS role or if the Role = None
    if(-not(($HealthExSvrObj.ExchangeInformation.ExServerRole -eq [HealthChecker.ServerRole]::None) -or 
        (($HealthExSvrObj.ExchangeInformation.ExchangeVersion -eq [HealthChecker.ExchangeVersion]::Exchange2013) -and 
        ($HealthExSvrObj.ExchangeInformation.ExServerRole -eq [HealthChecker.ServerRole]::ClientAccess))))
    {
	    
	    if($HealthExSvrObj.ExchangeInformation.ExchangeServicesNotRunning)
	    {
		    $ServerObject | Add-Member –MemberType NoteProperty –Name ServiceHealth -Value "Impacted"
			$ServerObject | Add-Member –MemberType NoteProperty –Name ServicesImpacted -Value $HealthExSvrObj.ExchangeInformation.ExchangeServicesNotRunning
	    }
        else
        {
            $ServerObject | Add-Member –MemberType NoteProperty –Name ServiceHealth -Value "Healthy"
        }
    }

    #################
	#TCP/IP Settings#
	#################
    if($HealthExSvrObj.OSVersion.TCPKeepAlive -eq 0)
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name TCPKeepAlive -Value "Not Set" 
    }
    elseif($HealthExSvrObj.OSVersion.TCPKeepAlive -lt 900000 -or $HealthExSvrObj.OSVersion.TCPKeepAlive -gt 1800000)
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name TCPKeepAlive -Value "Not Optimal"
    }
    else
    {
        $ServerObject | Add-Member –MemberType NoteProperty –Name TCPKeepAlive -Value "Optimal"
    }

    ###############################
	#LmCompatibilityLevel Settings#
	###############################
    $ServerObject | Add-Member –MemberType NoteProperty –Name LmCompatibilityLevel -Value $HealthExSvrObj.OSVersion.LmCompat.LmCompatibilityLevel


	##############
	#Hotfix Check#
	##############
    
    #Issue: throws errors 
    <#
    Add-Member : Cannot add a member with the name "Passed" because a member with that name already exists. To overwrite
    the member anyway, add the Force parameter to your command.
    #>
    #if($HealthExSvrObj.ExchangeInformation.ExchangeVersion -ne [HealthChecker.ExchangeVersion]::Exchange2010)
    #{
        #If((Display-KBHotfixCheck -HealthExSvrObj $HealthExSvrObj) -like "*Installed*")
        #{
       #     $ServerObject | Add-Member –MemberType NoteProperty –Name KB3041832 -Value "Installed"
        #}
    #}


    Write-debug "Building ServersObject " 
	$ServerObject
    

}


Function Get-HealthCheckFilesItemsFromLocation{
    $items = Get-ChildItem $XMLDirectoryPath | Where-Object{$_.Name -like "HealthCheck-*-*.xml"}
    if($items -eq $null)
    {
        Write-Host("Doesn't appear to be any Health Check XML files here....stopping the script")
        exit 
    }
    return $items
}

Function Get-OnlyRecentUniqueServersXMLs {
param(
[Parameter(Mandatory=$true)][array]$FileItems 
)   

    $aObject = @() 
    foreach($item in $FileItems)
    {
        $obj = New-Object PSCustomobject 
        [string]$itemName = $item.Name
        $ServerName = $itemName.Substring(($itemName.IndexOf("-") + 1), ($itemName.LastIndexOf("-") - $itemName.IndexOf("-") - 1))
        $obj | Add-Member -MemberType NoteProperty -Name ServerName -Value $ServerName
        $obj | Add-Member -MemberType NoteProperty -Name FileName -Value $itemName
        $obj | Add-Member -MemberType NoteProperty -Name FileObject -Value $item 
        $aObject += $obj
    }

    $grouped = $aObject | Group-Object ServerName 

    $FilePathList = @()
    foreach($gServer in $grouped)
    {
        
        if($gServer.Count -gt 1)
        {
            #going to only use the most current file for this server providing that they are using the newest updated version of Health Check we only need to sort by name
            $groupData = $gServer.Group #because of win2008
            $FilePathList += ($groupData | Sort-Object FileName -Descending | Select-Object -First 1).FileObject.VersionInfo.FileName

        }
        else {
            $FilePathList += ($gServer.Group).FileObject.VersionInfo.FileName
        }
        
    }

    return $FilePathList
}

Function Import-MyData {
param(
[Parameter(Mandatory=$true)][array]$FilePaths
)
    [System.Collections.Generic.List[System.Object]]$myData = New-Object -TypeName System.Collections.Generic.List[System.Object]
    foreach($filePath in $FilePaths)
    {
        $importData = Import-Clixml -Path $filePath
        $myData.Add($importData)
    }
    return $myData
}

Function Build-HtmlServerReport {

    $Files = Get-HealthCheckFilesItemsFromLocation
    $FullPaths = Get-OnlyRecentUniqueServersXMLs $Files
    $ImportData = Import-MyData -FilePaths $FullPaths

    $AllServersOutputObject = @()
    foreach($data in $ImportData)
    {
        $AllServersOutputObject += Build-ServerObject $data
    }
    
    Write-Debug "Building HTML report from AllServersOutputObject" 
	#Write-Debug $AllServersOutputObject 
    
	
	
    $htmlhead="<html>
            <style>
            BODY{font-family: Arial; font-size: 8pt;}
            H1{font-size: 16px;}
            H2{font-size: 14px;}
            H3{font-size: 12px;}
            TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
            TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
            TD{border: 1px solid black; padding: 5px; }
            td.pass{background: #7FFF00;}
            td.warn{background: #FFE600;}
            td.fail{background: #FF0000; color: #ffffff;}
            td.info{background: #85D4FF;}
            </style>
            <body>
            <h1 align=""center"">Exchange Health Checker v$($Script:healthCheckerVersion)</h1>
            <p>This shows a breif overview of known areas of concern. Details about each server are below.</p>
            <p align='center'>Note: KBs that could be missing on the server are not included in this version of the script. Please check this in the .log file of the Health Checker script results</p>"
    

    $HtmlTableHeader = "<p>
                        <table>
                        <tr>
                        <th>Server Name</th>
                        <th>Virtual Server</th>
                        <th>Hardware Type</th>
                        <th>OS</th>
                        <th>Exchange Version</th>
                        <th>Build Number</th>
                        <th>Build Days Old</th>
                        <th>Server Role</th>
                        <th>Auto Page File</th>
						<th>System Memory</th>
                        <th>Multiple Page Files</th>
                        <th>Page File Size</th>
                        <th>.Net Version</th>
                        <th>Power Plan</th>
                        <th>Hyper-Threading</th>
                        <th>Processor Speed</th>
                        <th>Service Health</th>
                        <th>TCP Keep Alive</th>
                        <th>LmCompatibilityLevel</th>
                        </tr>"
                        
    $ServersHealthHtmlTable = $ServersHealthHtmlTable + $htmltableheader 
    
    $ServersHealthHtmlTable += "<H2>Servers Overview</H2>"
                        
    foreach($ServerArrayItem in $AllServersOutputObject)
    {
        Write-Debug $ServerArrayItem
        $HtmlTableRow = "<tr>"
        $HtmlTableRow += "<td>$($ServerArrayItem.ServerName)</td>"	
        $HtmlTableRow += "<td>$($ServerArrayItem.VirtualServer)</td>"	
        $HtmlTableRow += "<td>$($ServerArrayItem.HardwareType)</td>"	
        $HtmlTableRow += "<td>$($ServerArrayItem.OperatingSystem)</td>"	
        $HtmlTableRow += "<td>$($ServerArrayItem.Exchange)</td>"			
        $HtmlTableRow += "<td>$($ServerArrayItem.BuildNumber)</td>"	
        
        If(!$ServerArrayItem.SupportedExchangeBuild) 
        {
            $HtmlTableRow += "<td class=""fail"">$($ServerArrayItem.BuildDaysOld)</td>"	
        }
        Else
        {
            $HtmlTableRow += "<td>$($ServerArrayItem.BuildDaysOld)</td>"
        }

        
        $HtmlTableRow += "<td>$($ServerArrayItem.ServerRole)</td>"	
        
        If($ServerArrayItem.AutoPageFile -eq "Yes")
        {
            $HtmlTableRow += "<td class=""fail"">$($ServerArrayItem.AutoPageFile)</td>"	
        }
        Else
        {
            $HtmlTableRow += "<td>$($ServerArrayItem.AutoPageFile)</td>"	
        }
		
		
		If(!$ServerArrayItem.E2013MemoryRight)
        {
            $HtmlTableRow += "<td class=""warn"">$($ServerArrayItem.TotalPhysicalMemory)</td>"	
        }
        ElseIf (!$ServerArrayItem.E2016MemoryRight)
        {
            $HtmlTableRow += "<td class=""warn"">$($ServerArrayItem.TotalPhysicalMemory)</td>"	
        }
        Else
        {
            $HtmlTableRow += "<td>$($ServerArrayItem.TotalPhysicalMemory)</td>"	
        }
		
		
                    
        If($ServerArrayItem.MultiplePageFiles -eq "Yes")
        {
            $HtmlTableRow += "<td class=""fail"">$($ServerArrayItem.MultiplePageFiles)</td>"	
        }
        Else
        {
            $HtmlTableRow += "<td>$($ServerArrayItem.MultiplePageFiles)</td>"	
        }
        
        If($ServerArrayItem.PagefileSizeSetRight -eq "No")
        {
            $HtmlTableRow += "<td class=""fail"">$($ServerArrayItem.PageFileSize)</td>"	
        }
        ElseIf ($ServerArrayItem.PagefileSizeSetRight -eq "Yes")
        {
            $HtmlTableRow += "<td>$($ServerArrayItem.PageFileSize)</td>"	
        }
        ElseIf (!$ServerArrayItem.PagefileSizeSetRight)
        {
            $HtmlTableRow += "<td class=""warn"">Undetermined</td>"	
        }
        
        $HtmlTableRow += "<td>$($ServerArrayItem.DotNetVersion)</td>"			
        
        If($ServerArrayItem.PowerPlan -ne "High performance")
        {
            $HtmlTableRow += "<td class=""fail"">$($ServerArrayItem.PowerPlan)</td>"	
        }
        Else
        {
            $HtmlTableRow += "<td>$($ServerArrayItem.PowerPlan)</td>"	
        }
        
        If($ServerArrayItem.HyperThreading -eq "Yes" -or $ServerArrayItem.AMD_HyperThreading -eq "Yes")
        {
            $HtmlTableRow += "<td class=""fail"">$($ServerArrayItem.HyperThreading)$($ServerArrayItem.AMD_HyperThreading)</td>"	
        }
        Else
        {
            $HtmlTableRow += "<td>$($ServerArrayItem.HyperThreading)$($ServerArrayItem.AMD_HyperThreading)</td>"	
        }
        
        If($ServerArrayItem.ProcessorSpeed -like "Throttled*")
        {
            $HtmlTableRow += "<td class=""fail"">$($ServerArrayItem.MaxProcessorSpeed)/$($ServerArrayItem.CurrentProcessorSpeed)</td>"	
        }
        Else
        {
            $HtmlTableRow += "<td>$($ServerArrayItem.MaxMegacyclesPerCore)</td>"	
        }
        
        If($ServerArrayItem.ServiceHealth -like "Impacted*")
        {
            $HtmlTableRow += "<td class=""fail"">Impacted</td>"	
        }
        Else
        {
            $HtmlTableRow += "<td>Healthy</td>"	
        }
        
        If($ServerArrayItem.TCPKeepAlive -eq "Not Optimal")
        {
            $HtmlTableRow += "<td class=""warn"">$($ServerArrayItem.TCPKeepAlive)</td>"	
        }
		ElseIf($ServerArrayItem.TCPKeepAlive -eq "Not Set")
        {
            $HtmlTableRow += "<td class=""fail"">$($ServerArrayItem.TCPKeepAlive)</td>"	
        }
        Else
        {
            $HtmlTableRow += "<td>$($ServerArrayItem.TCPKeepAlive)</td>"	
        }
        
        $HtmlTableRow += "<td>$($ServerArrayItem.LmCompatibilityLevel)</td>"	

        $HtmlTableRow += "</tr>"
                    
                    
        $ServersHealthHtmlTable = $ServersHealthHtmlTable + $htmltablerow
        
    }
    
    $ServersHealthHtmlTable += "</table></p>"
    
    $WarningsErrorsHtmlTable += "<H2>Warnings/Errors in your environment.</H2><table>"
    
    If($AllServersOutputObject.PowerPlanSetRight -contains $False)
	{
		$WarningsErrorsHtmlTable += "<tr><td class=""fail"">Power Plan</td><td>Error: High Performance Power Plan is recommended</td></tr>"
	}	
	If($AllServersOutputObject.SupportedExchangeBuild -contains $False)
	{
		$WarningsErrorsHtmlTable += "<tr><td class=""fail"">Old Build</td><td>Error: Out of date Cumulative Update detected. Please upgrade to one of the two most recently released Cumulative Updates.</td></tr>"
	}
	If($AllServersOutputObject.TCPKeepAlive -contains "Not Set")
	{
		$WarningsErrorsHtmlTable += "<tr><td class=""fail"">TCP Keep Alive</td><td>Error: The TCP KeepAliveTime value is not specified in the registry.  Without this value the KeepAliveTime defaults to two hours, which can cause connectivity and performance issues between network devices such as firewalls and load balancers depending on their configuration.  To avoid issues, add the KeepAliveTime REG_DWORD entry under HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\Tcpip\Parameters and set it to a value between 900000 and 1800000 decimal.  You want to ensure that the TCP idle timeout value gets higher as you go out from Exchange, not lower.  For example if the Exchange server has a value of 30 minutes, the Load Balancer could have an idle timeout of 35 minutes, and the firewall could have an idle timeout of 40 minutes.  Please note that this change will require a restart of the system.  Refer to the sections `"CAS Configuration`" and `"Load Balancer Configuration`" in this blog post for more details:  https://blogs.technet.microsoft.com/exchange/2016/05/31/checklist-for-troubleshooting-outlook-connectivity-in-exchange-2013-and-2016-on-premises/</td></tr>"	
	}
	
	If($AllServersOutputObject.TCPKeepAlive -contains "Not Optimal")
	{
		$WarningsErrorsHtmlTable += "<tr><td class=""warn"">TCP Keep Alive</td><td>Warning: The TCP KeepAliveTime value is not configured optimally. This can cause connectivity and performance issues between network devices such as firewalls and load balancers depending on their configuration. To avoid issues, set the HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\Tcpip\Parameters\KeepAliveTime registry entry to a value between 15 and 30 minutes (900000 and 1800000 decimal).  You want to ensure that the TCP idle timeout gets higher as you go out from Exchange, not lower.  For example if the Exchange server has a value of 30 minutes, the Load Balancer could have an idle timeout of 35 minutes, and the firewall could have an idle timeout of 40 minutes.  Please note that this change will require a restart of the system.  Refer to the sections `"CAS Configuration`" and `"Load Balancer Configuration`" in this blog post for more details:  https://blogs.technet.microsoft.com/exchange/2016/05/31/checklist-for-troubleshooting-outlook-connectivity-in-exchange-2013-and-2016-on-premises/</td></tr>"	
	}
	
	If($AllServersOutputObject.PagefileSizeSetRight -contains "No")
	{
		$WarningsErrorsHtmlTable += "<tr><td class=""fail"">Pagefile Size</td><td>Page set incorrectly detected. See https://technet.microsoft.com/en-us/library/cc431357(v=exchg.80).aspx - Please double check page file setting, as WMI Object Win32_ComputerSystem doesn't report the best value for total memory available.</td></tr>"
	}

    If($AllServersOutputObject.VirtualServer -contains "Yes")
    {
        $WarningsErrorsHtmlTable += "<tr><td class=""warn"">Virtual Servers</td><td>$($VirtualizationWarning)</td></tr>" 
    }

	If($AllServersOutputObject.E2013MultipleNICs -contains "Yes")
	{
		$WarningsErrorsHtmlTable += "<tr><td class=""Warn"">Multiple NICs</td><td>Multiple active network adapters detected. Exchange 2013 or greater may not need separate adapters for MAPI and replication traffic.  For details please refer to https://technet.microsoft.com/en-us/library/29bb0358-fc8e-4437-8feb-d2959ed0f102(v=exchg.150)#NR</td></tr>"
	}
	
	$a = ($ServerArrayItem.NumberNICs)
	 while($a -ge 1)
	 {
		$rss = "NIC_RSS_{0}" -f $a 
		
		If($AllServersOutputObject.$rss -contains "Disabled")
		{
			$WarningsErrorsHtmlTable += "<tr><td class=""Warn"">RSS</td><td>Enabling RSS is recommended.</td></tr>"
			break;
		}	
		ElseIf($AllServersOutputObject.$rss -contains "NoRSS")
		{
			$WarningsErrorsHtmlTable += "<tr><td class=""Warn"">RSS</td><td>Enabling RSS is recommended.</td></tr>"
			break;
		}	
		
		$a--
	 }
	 
	If($AllServersOutputObject.NUMAGroupSize -contains "Undetermined")
	{
		$WarningsErrorsHtmlTable += "<tr><td class=""Warn"">NUMA Group Size Optimization</td><td>Unable to determine --- Warning: If this is set to Clustered, this can cause multiple types of issues on the server</td></tr>"
	}
	ElseIf($AllServersOutputObject.NUMAGroupSize -contains "Clustered")
	{
		$WarningsErrorsHtmlTable += "<tr><td class=""fail"">NUMA Group Size Optimization</td><td>BIOS Set to Clustered --- Error: This setting should be set to Flat. By having this set to Clustered, we will see multiple different types of issues.</td></tr>"
	}
	
	If($AllServersOutputObject.AllProcCoresVisible -contains "Undetermined")
	{
		$WarningsErrorsHtmlTable += "<tr><td class=""Warn"">All Processor Cores Visible</td><td>Unable to determine --- Warning: If we aren't able to see all processor cores from Exchange, we could see performance related issues.</td></tr>"
	}
	ElseIf($AllServersOutputObject.AllProcCoresVisible -contains "No")
	{
		$WarningsErrorsHtmlTable += "<tr><td class=""fail"">All Processor Cores Visible</td><td>Not all Processor Cores are visable to Exchange and this will cause a performance impact</td></tr>"
	}
	
	If($AllServersOutputObject.E2016MemoryRight -contains $False)
	{
		$WarningsErrorsHtmlTable += "<tr><td class=""Warn"">Exchange 2016 Memory</td><td>Memory greater than 192GB. We recommend for the best performance to be scaled at or below 192 GB of Memory.</td></tr>"
	}
	
	If($AllServersOutputObject.E2013MemoryRight -contains $False)
	{
		$WarningsErrorsHtmlTable += "<tr><td class=""Warn"">Exchange 2013 Memory</td><td>Memory greater than 96GB. We recommend for the best performance to be scaled at or below 96GB of Memory. However, having higher memory than this has yet to be linked directly to a MAJOR performance issue of a server.</td></tr>"
	}	
	
    $WarningsErrorsHtmlTable += "</table>"

	
    $ServerDetailsHtmlTable += "<p><H2>Server Details</H2><table>"
    
    Foreach($ServerArrayItem in $AllServersOutputObject)
    {

        $ServerDetailsHtmlTable += "<tr><th>Server Name</th><th>$($ServerArrayItem.ServerName)</th></tr>"
        $ServerDetailsHtmlTable += "<tr><td>Manufacturer</td><td>$($ServerArrayItem.Manufacturer)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>Model</td><td>$($ServerArrayItem.Model)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>Hardware Type</td><td>$($ServerArrayItem.HardwareType)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>Operating System</td><td>$($ServerArrayItem.OperatingSystem)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>Exchange</td><td>$($ServerArrayItem.Exchange)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>Build Number</td><td>$($ServerArrayItem.BuildNumber)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>Server Role</td><td>$($ServerArrayItem.ServerRole)</td></tr>"
		$ServerDetailsHtmlTable += "<tr><td>System Memory</td><td>$($ServerArrayItem.TotalPhysicalMemory)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>Page File Size</td><td>$($ServerArrayItem.PagefileSize)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>.Net Version Installed</td><td>$($ServerArrayItem.DotNetVersion)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>HTTP Proxy</td><td>$($ServerArrayItem.HTTPProxy)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>Processor</td><td>$($ServerArrayItem.ProcessorName)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>Number of Processors</td><td>$($ServerArrayItem.NumberOfProcessors)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>Logical/Physical Cores</td><td>$($ServerArrayItem.NumberOfLogicalProcessors)/$($ServerArrayItem.NumberOfPhysicalCores)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>Max Speed Per Core</td><td>$($ServerArrayItem.MaxMegacyclesPerCore)</td></tr>"
		$ServerDetailsHtmlTable += "<tr><td>NUMA Group Size</td><td>$($ServerArrayItem.NUMAGroupSize)</td></tr>"
		$ServerDetailsHtmlTable += "<tr><td>All Procs Visible</td><td>$($ServerArrayItem.AllProcCoresVisible)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>System Memory</td><td>$($ServerArrayItem.TotalPhysicalMemory)</td></tr>"
		$ServerDetailsHtmlTable += "<tr><td>Multiple NICs</td><td>$($ServerArrayItem.E2013MultipleNICs)</td></tr>"
        $ServerDetailsHtmlTable += "<tr><td>Services Down</td><td>$($ServerArrayItem.ServicesImpacted)</td></tr>"
		
		#NIC 
		$a = ($ServerArrayItem.NumberNICs)
		 while($a -ge 1)
		 {
            $name = "NIC_Name_{0}" -f $a 
		    $ServerDetailsHtmlTable += "<tr><td>NIC Name</td><td>$($ServerArrayItem.$name)</td></tr>"
			$description = "NIC_Description_{0}" -f $a 
		    $ServerDetailsHtmlTable += "<tr><td>NIC Description</td><td>$($ServerArrayItem.$description)</td></tr>"
			$driver = "NIC_Driver_{0}" -f $a 
		    $ServerDetailsHtmlTable += "<tr><td>NIC Driver</td><td>$($ServerArrayItem.$driver)</td></tr>"
			$linkspeed = "NIC_LinkSpeed_{0}" -f $a 
		    $ServerDetailsHtmlTable += "<tr><td>NIC LinkSpeed</td><td>$($ServerArrayItem.$linkspeed)</td></tr>"
			$rss = "NIC_RSS_{0}" -f $a 
		    $ServerDetailsHtmlTable += "<tr><td>RSS</td><td>$($ServerArrayItem.$rss)</td></tr>"
			$a--
		 }
		 
    }
    
    $ServerDetailsHtmlTable += "</table></p>"
    
    $htmltail = "</body>
    </html>"

    $htmlreport = $htmlhead  + $ServersHealthHtmlTable + $WarningsErrorsHtmlTable + $ServerDetailsHtmlTable  + $htmltail
    
    $htmlreport | Out-File $HtmlReportFile -Encoding UTF8
}


##############################################################
#
#           DC to Exchange cores Report Functions 
#
##############################################################

Function Get-ComputerCoresObject {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name
)
    Write-VerboseOutput("Calling: Get-ComputerCoresObject")
    Write-VerboseOutput("Passed: {0}" -f $Machine_Name)

    $returnObj = New-Object pscustomobject 
    $returnObj | Add-Member -MemberType NoteProperty -Name Error -Value $false
    $returnObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Machine_Name
    $returnObj | Add-Member -MemberType NoteProperty -Name NumberOfCores -Value ([int]::empty)
    $returnObj | Add-Member -MemberType NoteProperty -Name Exception -Value ([string]::empty)
    $returnObj | Add-Member -MemberType NoteProperty -Name ExceptionType -Value ([string]::empty)
    try {
        $wmi_obj_processor = Get-WmiObject -Class Win32_Processor -ComputerName $Machine_Name

        foreach($processor in $wmi_obj_processor)
        {
            $returnObj.NumberOfCores +=$processor.NumberOfCores
        }
        
        Write-Grey("Server {0} Cores: {1}" -f $Machine_Name, $returnObj.NumberOfCores)
    }
    catch {
        $thisError = $Error[0]
        if($thisError.Exception.Gettype().FullName -eq "System.UnauthorizedAccessException")
        {
            Write-Yellow("Unable to get processor information from server {0}. You do not have the correct permissions to get this data from that server. Exception: {1}" -f $Machine_Name, $thisError.ToString())
        }
        else 
        {
            Write-Yellow("Unable to get processor infomration from server {0}. Reason: {1}" -f $Machine_Name, $thisError.ToString())
        }
        $returnObj.Exception = $thisError.ToString() 
        $returnObj.ExceptionType = $thisError.Exception.Gettype().FullName
        $returnObj.Error = $true
    }
    
    return $returnObj
}

Function Get-ExchnageDCCoreRatio {

    $OutputFullPath = "{0}\HealthCheck-ExchangeDCCoreRatio-{1}.log" -f $OutputFilePath, $dateTimeStringFormat
    Write-VerboseOutput("Calling: Get-ExchnageDCCoreRatio")
    Write-Grey("Exchange Server Health Checker Report - AD GC Core to Exchange Server Core Ratio - v{0}" -f $healthCheckerVersion)
    $coreRatioObj = New-Object pscustomobject 
    try 
    {
        Write-VerboseOutput("Attempting to load Active Directory Module")
        Import-Module ActiveDirectory 
        Write-VerboseOutput("Successfully loaded")
    }
    catch 
    {
        Write-Red("Failed to load Active Directory Module. Stopping the script")
        exit 
    }

    $ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name
    [array]$DomainControllers = (Get-ADForest).Domains | %{ Get-ADDomainController -Filter {isGlobalCatalog -eq $true -and Site -eq $ADSite} -Server $_ }

    [System.Collections.Generic.List[System.Object]]$DCList = New-Object System.Collections.Generic.List[System.Object]
    $DCCoresTotal = 0
    Write-Break
    Write-Grey("Collecting data for the Active Directory Environment in Site: {0}" -f $ADSite)
    $iFailedDCs = 0 
    foreach($DC in $DomainControllers)
    {
        $DCCoreObj = Get-ComputerCoresObject -Machine_Name $DC.Name 
        $DCList.Add($DCCoreObj)
        if(-not ($DCCoreObj.Error))
        {
            $DCCoresTotal += $DCCoreObj.NumberOfCores
        }
        else 
        {
            $iFailedDCs++     
        } 
    }

    $coreRatioObj | Add-Member -MemberType NoteProperty -Name DCList -Value $DCList
    if($iFailedDCs -eq $DomainControllers.count)
    {
        #Core count is going to be 0, no point to continue the script
        Write-Red("Failed to collect data from your DC servers in site {0}." -f $ADSite)
        Write-Yellow("Because we can't determine the ratio, we are going to stop the script. Verify with the above errors as to why we failed to collect the data and address the issue, then run the script again.")
        exit 
    }

    [array]$ExchangeServers = Get-ExchangeServer | Where-Object {$_.Site -match $ADSite}
    $EXCoresTotal = 0
    [System.Collections.Generic.List[System.Object]]$EXList = New-Object System.Collections.Generic.List[System.Object]
    Write-Break
    Write-Grey("Collecting data for the Exchange Environment in Site: {0}" -f $ADSite)
    foreach($svr in $ExchangeServers)
    {
        $EXCoreObj = Get-ComputerCoresObject -Machine_Name $svr.Name 
        $EXList.Add($EXCoreObj)
        if(-not ($EXCoreObj.Error))
        {
            $EXCoresTotal += $EXCoreObj.NumberOfCores
        }
    }
    $coreRatioObj | Add-Member -MemberType NoteProperty -Name ExList -Value $EXList

    Write-Break
    $CoreRatio = $EXCoresTotal / $DCCoresTotal
    Write-Grey("Total DC/GC Cores: {0}" -f $DCCoresTotal)
    Write-Grey("Total Exchange Cores: {0}" -f $EXCoresTotal)
    Write-Grey("You have {0} Exchange Cores for every Domain Controller Global Catalog Server Core" -f $CoreRatio)
    if($CoreRatio -gt 8)
    {
        Write-Break
        Write-Red("Your Exchange to Active Directory Global Catalog server's core ratio does not meet the recommended guidelines of 8:1")
        Write-Red("Recommended guidelines for Exchange 2013/2016 for every 8 Exchange cores you want at least 1 Active Directory Global Catalog Core.")
        Write-Yellow("Documentation:")
        Write-Yellow("`thttps://blogs.technet.microsoft.com/exchange/2013/05/06/ask-the-perf-guy-sizing-exchange-2013-deployments/")
        Write-Yellow("`thttps://technet.microsoft.com/en-us/library/dn879075(v=exchg.150).aspx")

    }
    else 
    {
        Write-Break
        Write-Green("Your Exchange Environment meets the recommended core ratio of 8:1 guidelines.")    
    }
    
    $XMLDirectoryPath = $OutputFullPath.Replace(".log",".xml")
    $coreRatioObj | Export-Clixml $XMLDirectoryPath 
    Write-Grey("Output file written to {0}" -f $OutputFullPath)
    Write-Grey("Output XML Object file written to {0}" -f $XMLDirectoryPath)

}

Function Main {
    
    if(-not (Is-Admin))
	{
		Write-Warning "The script needs to be executed in elevated mode. Start the Exchange Mangement Shell as an Administrator." 
		sleep 2;
		exit
    }
    
    if($BuildHtmlServersReport)
    {
        Build-HtmlServerReport
        sleep 2;
        exit
    }

	Load-ExShell
    if((Test-Path $OutputFilePath) -eq $false)
    {
        Write-Host "Invalid value specified for -OutputFilePath." -ForegroundColor Red
        exit 
    }
    $iErrorStartCount = $Error.Count #useful for debugging 
    $Script:iErrorExcluded = 0 #this is a way to determine if the only errors occurred were in try catch blocks. If there is a combination of errors in and out, then i will just dump it all out to avoid complex issues. 
    $Script:date = (Get-Date)
    $Script:dateTimeStringFormat = $date.ToString("yyyyMMddHHmmss")
    $OutputFileName = "HealthCheck" + "-" + $Server + "-" + $dateTimeStringFormat + ".log"
    $OutputFullPath = $OutputFilePath + "\" + $OutputFileName
    Write-VerboseOutput("Calling: main Script Execution")

    if($LoadBalancingReport)
    {
        [int]$iMajor = (Get-ExchangeServer $Server).AdminDisplayVersion.Major
        if($iMajor -gt 14)
        {
            $OutputFileName = "LoadBalancingReport" + "-" + $dateTimeStringFormat + ".log"
            $OutputFullPath = $OutputFilePath + "\" + $OutputFileName
            Write-Green("Exchange Health Checker Script version: " + $healthCheckerVersion)
            Write-Green("Client Access Load Balancing Report on " + $date)
            Get-CASLoadBalancingReport
            Write-Grey("Output file written to " + $OutputFullPath)
            Write-Break
            Write-Break
        }
        else
        {
            Write-Yellow("-LoadBalancingReport is only supported for Exchange 2013 and greater")
        }
        #Load balancing report only needs to be the thing that runs
        exit
    }

    if($DCCoreRatio)
    {
        $oldErrorAction = $ErrorActionPreference
        $ErrorActionPreference = "Stop"
        try 
        {
            Get-ExchnageDCCoreRatio
        }
        finally
        {
            $ErrorActionPreference = $oldErrorAction
            exit 
        }
    }

   
    $OutputFileName = "HealthCheck" + "-" + $Server + "-" + $dateTimeStringFormat + ".log"
	$OutputFullPath = $OutputFilePath + "\" + $OutputFileName
	$OutXmlFullPath = $OutputFilePath + "\" + ($OutputFileName.Replace(".log",".xml"))
	$HealthObject = Build-HealthExchangeServerObject $Server
	Display-ResultsToScreen $healthObject 
	if($MailboxReport)
	{
	    Get-MailboxDatabaseAndMailboxStatistics -Machine_Name $Server
	}
	Write-Grey("Output file written to " + $OutputFullPath)
	if($Error.Count -gt $iErrorStartCount)
	{
	    Write-Grey(" ");Write-Grey(" ")
	    Function Write-Errors {
	        $index = 0; 
	        "Errors that occurred" | Out-File ($OutputFullPath) -Append
	        while($index -lt ($Error.Count - $iErrorStartCount))
	        {
	            $Error[$index++] | Out-File ($OutputFullPath) -Append
	        }
	    }
	    #Now to determine if the errors are expected or not 
	    if(($Error.Count - $iErrorStartCount) -ne $Script:iErrorExcluded)
	    {
	        Write-Red("There appears to have been some errors in the script. To assist with debugging of the script, please RE-RUN the script with -Verbose send the .log and .xml file to dpaul@microsoft.com.")
	        Write-Errors
	    }
	    elseif($Script:VerboseEnabled)
	    {
	        Write-Grey("All errors that occurred were in try catch blocks and was handled correctly.")
	        Write-Errors
	    }
        
	}
	Write-Grey("Exported Data Object written to " + $OutXmlFullPath)
	$HealthObject | Export-Clixml -Path $OutXmlFullPath -Encoding UTF8 -Depth 5
	
}

Main 