$Configuration = Get-SenderFilterConfig
$senders = Get-Content .\blocksenders.txt
Foreach ($sender in $senders) {$Configuration.BlockedSenders += $sender}
connect-exchangeserver ExchangeSRv01
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv02
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv03
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv04
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv05
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv06
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv07
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv08
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv09
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv10
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders