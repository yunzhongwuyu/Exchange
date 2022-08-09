Write-Host ""
$Sender = Read-Host "Enter the e-mail address to remove from whitelist"
$Servers = @("ExchangeSrv001","ExchangeSrv002","ExchangeSrv003","ExchangeSrv004","ExchangeSrv005","ExchangeSrv006","ExchangeSrv007","ExchangeSrv008","ExchangeSrv009","ExchangeSrv010")

Foreach ($Server in $Servers) {
    $Session = New-PSSession -ComputerName $Server
    Invoke-Command -Session $Session -ScriptBlock {
		Add-PSSnapin FSSPSSnapin -ErrorAction SilentlyContinue
        	$fse = Get-FseSpamContentFilter
		$fse.AllowedSender.Remove($Args[0])
		Set-FseSpamContentFilter -AllowedSender $fse.AllowedSender		
    } -ArgumentList $Sender
	
	Write-Host ""
	Write-Host "$Sender is now included in antispam scanning on $Server" -ForegroundColor Green
	Remove-PSSession $Session
}
