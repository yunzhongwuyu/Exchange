Write-host ""
$Service = Read-host "Enter Service"
$Servers = @("ExchangeSrv001","ExchangeSrv002","ExchangeSrv003","ExchangeSrv004","ExchangeSrv005","ExchangeSrv006","ExchangeSrv007","ExchangeSrv008","ExchangeSrv009","ExchangeSrv010")
If ($service -ne "MSExchangeADTopology" -or "MSExchangeIS") {
Foreach ($server in $Servers){
 Write-host "$service is restarted on $server" -f green
 Invoke-Command -ComputerName $server -ScriptBlock {Restart-Service $Args[0]} -ArgumentList $Service}
}
else {Write-host "Microsoft Exchange Active Directory Topology and Microsoft Exchange Information Store must be restarted in Maitenance Mode" -ForegroundColor red}