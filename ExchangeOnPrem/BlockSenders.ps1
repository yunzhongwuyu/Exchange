connect-exchangeserver ExchangeSRv01
$Configuration = Get-SenderFilterConfig
$senders = Get-Content .\blocksenders.txt
Foreach ($sender in $senders) {$Configuration.BlockedSenders += $sender}
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv02
$Configuration = Get-SenderFilterConfig
$senders = Get-Content .\blocksenders.txt
Foreach ($sender in $senders) {$Configuration.BlockedSenders += $sender}
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv03
$Configuration = Get-SenderFilterConfig
$senders = Get-Content .\blocksenders.txt
Foreach ($sender in $senders) {$Configuration.BlockedSenders += $sender}
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv04
$Configuration = Get-SenderFilterConfig
$senders = Get-Content .\blocksenders.txt
Foreach ($sender in $senders) {$Configuration.BlockedSenders += $sender}
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv05
$Configuration = Get-SenderFilterConfig
$senders = Get-Content .\blocksenders.txt
Foreach ($sender in $senders) {$Configuration.BlockedSenders += $sender}
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv06
$Configuration = Get-SenderFilterConfig
$senders = Get-Content .\blocksenders.txt
Foreach ($sender in $senders) {$Configuration.BlockedSenders += $sender}
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv07
$Configuration = Get-SenderFilterConfig
$senders = Get-Content .\blocksenders.txt
Foreach ($sender in $senders) {$Configuration.BlockedSenders += $sender}
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv08
$Configuration = Get-SenderFilterConfig
$senders = Get-Content .\blocksenders.txt
Foreach ($sender in $senders) {$Configuration.BlockedSenders += $sender}
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv09
$Configuration = Get-SenderFilterConfig
$senders = Get-Content .\blocksenders.txt
Foreach ($sender in $senders) {$Configuration.BlockedSenders += $sender}
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
connect-exchangeserver ExchangeSRv10
$Configuration = Get-SenderFilterConfig
$senders = Get-Content .\blocksenders.txt
Foreach ($sender in $senders) {$Configuration.BlockedSenders += $sender}
Set-SenderFilterConfig -BlockedSenders $Configuration.BlockedSenders
