    Function New-SourceTargetRecipientMap {
        
        [cmdletbinding()]
        param
        (
            $SourceRecipients
            ,
            $ExchangeSystem
            ,
            [hashtable]$DomainReplacement = @{}
        )
        $SourceTargetRecipientMap = @{}
        $TargetSourceRecipientMap = @{}
        foreach ($SR in $SourceRecipients)
        {
            Connect-OneShellSystem -Identity $ExchangeSystem
            $ExchangeSession = Get-OneShellSystemPSSession -id $ExchangeSystem
            $ProxyAddressesToCheck = $sr.proxyaddresses | Where-Object -FilterScript {$_ -ilike 'smtp:*'}
            $rawrecipientmatches =
            @(
                foreach ($pa2c in $ProxyAddressesToCheck)
                {
                    $domain = $pa2c.split('@')[1] 
                    if ($domain -in $DomainReplacement.Keys)
                    {
                        $pa2c = $pa2c.replace($domain,$($DomainReplacement.$domain))
                    }
                    if (Test-ExchangeProxyAddress -ProxyAddress $pa2c -ProxyAddressType SMTP -ExchangeSession $ExchangeSession)
                    {$null}
                    else
                    {
                        Test-ExchangeProxyAddress -ProxyAddress $pa2c -ProxyAddressType SMTP -ExchangeSession $ExchangeSession -ReturnConflicts
                    }
                }
            )
            $recipientmatches = @($rawrecipientmatches | Select-Object -Unique | Where-Object -FilterScript {$_ -ne $null})
            if ($recipientmatches.Count -eq 1)
            {
                $SourceTargetRecipientMap.$($SR.ObjectGUID.guid)=$recipientmatches
                $TargetSourceRecipientMap.$($recipientmatches[0])=$($SR.ObjectGUID.guid)
            }
            elseif ($recipientmatches.Count -eq 0) {
                $SourceTargetRecipientMap.$($SR.ObjectGUID.guid)=$null
            }
            else
            {
                $SourceTargetRecipientMap.$($SR.ObjectGUID.guid)=$recipientmatches
            }
        }#foreach
        $RecipientMap = @{
            SourceTargetRecipientMap = $SourceTargetRecipientMap
            TargetSourceRecipientMap = $TargetSourceRecipientMap
        }
        $RecipientMap
    
    }
