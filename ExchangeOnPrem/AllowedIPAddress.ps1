get-content D:\scripts\allowedIP.txt | foreach { Add-IPAllowListEntry -IPaddress $_ -server ExchangeSRv01 }
get-content D:\scripts\allowedIP.txt | foreach { Add-IPAllowListEntry -IPaddress $_ -server ExchangeSRv02 }
get-content D:\scripts\allowedIP.txt | foreach { Add-IPAllowListEntry -IPaddress $_ -server ExchangeSRv03 }
get-content D:\scripts\allowedIP.txt | foreach { Add-IPAllowListEntry -IPaddress $_ -server ExchangeSRv04 }
get-content D:\scripts\allowedIP.txt | foreach { Add-IPAllowListEntry -IPaddress $_ -server ExchangeSRv05 }
get-content D:\scripts\allowedIP.txt | foreach { Add-IPAllowListEntry -IPaddress $_ -server ExchangeSRv06 }
get-content D:\scripts\allowedIP.txt | foreach { Add-IPAllowListEntry -IPaddress $_ -server ExchangeSRv07 }
get-content D:\scripts\allowedIP.txt | foreach { Add-IPAllowListEntry -IPaddress $_ -server ExchangeSRv08 }
get-content D:\scripts\allowedIP.txt | foreach { Add-IPAllowListEntry -IPaddress $_ -server ExchangeSRv09 }
get-content D:\scripts\allowedIP.txt | foreach { Add-IPAllowListEntry -IPaddress $_ -server ExchangeSRv10 }