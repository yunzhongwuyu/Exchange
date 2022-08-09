Write-Host ""
$Domain = Read-Host "Enter the domain to remove from whitelist"
$Servers = @("ExchangeSrv001","ExchangeSrv002","ExchangeSrv003","ExchangeSrv004","ExchangeSrv005","ExchangeSrv006","ExchangeSrv007","ExchangeSrv008","ExchangeSrv009","ExchangeSrv010")

Foreach ($Server in $Servers) {
    $Session = New-PSSession -ComputerName $Server
    Invoke-Command -Session $Session -ScriptBlock {
		Add-PSSnapin FSSPSSnapin -ErrorAction SilentlyContinue
		$fse = Get-FseSpamContentFilter
		$fse.AllowedSenderDomain.Remove($args[0])
		Set-FseSpamContentFilter -AllowedSenderDomain $fse.AllowedSenderDomain
    } -ArgumentList $Domain

	Write-Host ""
	Write-Host "$Domain is now included in antispam scanning on $Server" -ForegroundColor Red
	Remove-PSSession $Session
}
