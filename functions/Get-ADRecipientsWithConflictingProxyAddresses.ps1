    Function Get-ADRecipientsWithConflictingProxyAddresses {
        
        [cmdletbinding()]
        param
        (
            $SourceRecipients
            ,
            $TargetExchangeOrganization
        )
        foreach ($sr in $SourceRecipients)
        {
            $ProxyAddressesToCheck = $sr.proxyaddresses | Where-Object -FilterScript {$_ -ilike 'x500:*' -or $_ -ilike 'smtp:*'}
            foreach ($pa2c in $ProxyAddressesToCheck)
            {
            $type = $pa2c.split(':')[0]
            if (Test-ExchangeProxyAddress -ProxyAddress $pa2c -ProxyAddressType $type -ExchangeOrganization $TargetExchangeOrganization)
            {
                    Write-OneShellLog -Message "No Conflict for $pa2c" -EntryType Notification -Verbose
            }
            else
            {
                    $conflicts = @(Test-ExchangeProxyAddress -ProxyAddress $pa2c -ProxyAddressType $type -ExchangeOrganization $TargetExchangeOrganization -ReturnConflicts)
                    [pscustomobject]@{
                        SourceObjectGUID = $sr.ObjectGUID
                        ConflictingTargetObjectGUIDs = $conflicts
                        ProxyAddress = $pa2c
                    }
            }
            }
        }
    
    }
