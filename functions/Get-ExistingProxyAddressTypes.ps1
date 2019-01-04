    Function Get-ExistingProxyAddressTypes {
        
        param(
        [object[]]$proxyAddresses
        )
        $ProxyAddresses | ForEach-Object -Process {$_.split(':')[0]} | Sort-Object | Select-Object -Unique
    
    }
