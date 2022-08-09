$user = read-host "Please input user iniital that need enable remote mailbox"

Write-Host "Connecting to Exchange Online... " -ForegroundColor Yellow -NoNewline
$credential = Get-Credential 
$ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $ExchangeOnlineSession -DisableNameChecking -AllowClobber | Out-Null
Write-Host "Done" -ForegroundColor Green
$host.ui.RawUI.WindowTitle = "Customer" + " :: Office 365"
$Office365Mailbox = Get-Mailbox $user

Remove-PSSession $ExchangeOnlineSession

Write-Host "Connecting to Exchange OnPremiese... " -ForegroundColor Yellow -NoNewline
$ComputerFQDN = "ExchangeServer.domain.com"
$OnPremSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ComputerFQDN/PowerShell/ -Authentication Kerberos -Credential $credential
Import-PSSession $OnPremSession -Allowclobber | Out-Null
$Host.UI.RawUI.WindowTitle = "Machine:$ComputerFQDN"
Write-Host "Done" -ForegroundColor Green

$OnPremMailbox = Get-Mailbox $User
$OnpremLegacyDN = $OnPremMailbox.LegacyExchangeDN
$SipAddress = "Sip:" + $User + "@domain.com"
$OnPremExchangeGuid = $Office365Mailbox.ExchangeGuid.Guid
Write-Host "
LegacyDN will add to remote mailbox
$OnpremLegacyDN
"
[Array]$OnpremAddresses = $OnPremMailbox.emailaddresses | select ProxyAddressString -ExpandProperty ProxyAddressString
Write-Host "
Following Address will be added into RemoteMailbox
$OnpremAddresses
"
$LegacyDNX500 = "X500:" + [String]$OnpremLegacyDN
$PrimaryAddress = $User + "@domain.com"
$RemoteRoutingAddress = $user + "@domain.mail.onmicrosoft.com"
Disable-Mailbox $User -Confirm:$false -Verbose
Enable-RemoteMailbox $user -PrimarySmtpAddress $PrimaryAddress -RemoteRoutingAddress $RemoteRoutingAddress
Set-RemoteMailbox $user -Emailaddresses @{Add="$LegacyDNX500"} 
Set-RemoteMailbox $User -Emailaddresses @{Add="$SipAddress"}
Set-RemoteMailbox $User -Emailaddresses @{Add="$RemoteRoutingAddress"}
Write-Host "$OnPremExchangeGuid will be unified with ExchangeOnline Guid"
Set-RemoteMailbox $User -ExchangeGuid $OnPremExchangeGuid
Foreach ($OnPrAddress in $OnpremAddresses){Set-RemoteMailbox $user -Emailaddresses @{Add="$OnPrAddress"}}

Get-RemoteMailbox $User | select *address*,ExchangeGuid

Remove-PSSession $OnPremSession