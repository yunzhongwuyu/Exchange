Function Get-O365Credential()
{
          #Write-Output "Enter Connect-O365"
          #Read-Host -prompt "Enter Password" -assecurestring | convertfrom-securestring |out-file D:\O365PASS.txt -Force
          $pass = cat "D:\O365PASS.txt" | Convertto-securestring
          return new-object -typename System.Management.Automation.PSCredential -argumentlist "adminuser@domain.com", $pass
}


$pass = cat "D:\O365PASS.txt" | ConvertTo-SecureString
$O365CREDS  = Get-O365Credential
$OnPremiseCreds = new-object -typename System.Management.Automation.PSCredential -argumentlist "domain\adminAccount", $pass
# $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $O365CREDS  -Authentication Basic -AllowRedirection
#Import-PSSession $Session |Out-Null

New-MoveRequest -Identity mailbox@domain.com -Remote -RemoteHostName migrationEndPoint.domain.com -TargetDeliveryDomain domain.mail.onmicrosoft.com -BatchName onPrem_mailbox -RemoteCredential $OnPremiseCreds 
