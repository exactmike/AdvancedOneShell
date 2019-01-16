    Function Get-DesiredProxyAddresses {
        
        [cmdletbinding()]
        param
        (
            [parameter()]
            [string[]]$CurrentProxyAddresses #Current proxy addresses to preserve or evaluate for preservation
            ,
            [string]$DesiredPrimaryAddress #replace existing primary smtp address with this value
            ,
            [string]$DesiredOrCurrentAlias #used for calculation of a TargetAddress if required.
            ,
            [string[]]$LegacyExchangeDNs #legacyexchangedn to convert to additional x500 address
            ,
            [psobject[]]$Recipients #Recipient objects to consume for their proxy addresses and legacyexchangedn
            ,
            [parameter()]
            [switch]$VerifyAddTargetAddress #have the function ensure inclusion of a targetdeliverydomain proxy address.  Requires the TargetDeliveryDomain and DesiredOrCurrentAlias parameters.
            ,
            [string]$TargetDeliveryDomain #specify the external delivery domain - usually for cross forest or cloud like contoso.mail.onmicrosoft.com
            ,
            [switch]$VerifySMTPAddressValidity #verifies that the SMTP address complies with basic format requirements to be valid. See documentation for Test-EmailAddress for more information.
            ,
            [string[]]$DomainsToRemove #specify the domains for which to remove the associated proxy addresses. Include only the domain name, like 'contoso.com'
            ,
            [string[]]$AddressesToRemove #specify the complete address including the type: prefix, like smtp: or x500:
            ,
            [string[]]$AddressesToAdd #specifcy the complete address including the type: prefix, like smtp: or x500:
            ,
            [switch]$TestAddressAvailability
            ,
            $TestAddressExchangeOrganizationSession
        )
        if ($PSBoundParameters.ContainsKey('CurrentProxyAddresses'))
        {
            $DesiredProxyAddresses = $CurrentProxyAddresses.Clone() | Select-Object -Unique
        }
        else
        {
            $DesiredProxyAddresses = @()
        }
        if ($LegacyExchangeDNs.Count -ge 1)
        {
            foreach ($LED in $LegacyExchangeDNs)
            {
                $existingProxyAddressTypes = Get-ExistingProxyAddressTypes -proxyAddresses $DesiredProxyAddresses
                $type = 'X500'
                if ($existingProxyAddressTypes -ccontains $type)
                {
                    $type = $type.ToLower()
                }
                $newX500 = "$type`:$LED"
                if ($newX500 -in $DesiredProxyAddresses) {}
                else
                {
                    $DesiredProxyAddresses += $newX500
                }
            }
        }
        if ($VerifyAddTargetAddress -eq $true)
        {
            if ($DesiredOrCurrentAlias -and $TargetDeliveryDomain)
            {
                $DesiredTargetAddress = "smtp:$DesiredOrCurrentAlias@$TargetDeliveryDomain"
                if (($DesiredProxyAddresses | Where-Object {$_ -eq $DesiredTargetAddress}).count -lt 1)
                {
                    $DesiredProxyAddresses += $DesiredTargetAddress
                }#end if
            }#if
            else
            {
                throw('ERROR: VerifyAddTargetAddress was specified but DesiredOrCurrentAlias and/or TargetDeliveryDomain were not specified.')
            }#else
        }#if
        if ($Recipients.Count -ge 1)
        {
            $RecipientProxyAddresses = @()
            foreach ($recipient in $Recipients)
            {
                $paProperty = if (Test-Member -InputObject $recipient -Name emailaddresses) {'EmailAddresses'} elseif (Test-Member -InputObject $recipient -Name proxyaddresses ) {'proxyAddresses'} else {$null}
                if ($paProperty)
                {
                $existingProxyAddressTypes = Get-ExistingProxyAddressTypes -proxyAddresses $DesiredProxyAddresses
                    $rpa = @($recipient.$paProperty)
                    foreach ($a in $rpa)
                    {
                        $type = $a.split(':')[0]
                        $address = $a.split(':')[1]
                        if ($existingProxyAddressTypes -ccontains $type)
                        {
                            $la = $type.tolower() + ':' +  $address
                        } #end if
                        else
                        {
                            $la = $a
                        } #end else
                        $RecipientProxyAddresses += $la
                    }#end foreach
                }#end if
            }#foreach
            if ($RecipientProxyAddresses.count -ge 1)
            {
                $add = @($RecipientProxyAddresses | Where-Object {$DesiredProxyAddresses -inotcontains $_})
                if ($add.Count -ge 1)
                {
                    $DesiredProxyAddresses += @($add)
                }
            }#if
        }#if
        if ($AddressesToAdd.Count -ge 1)
        {
            $add = @($AddressesToAdd | Where-Object {$DesiredProxyAddresses -inotcontains $_})
            if ($add.Count -ge 1)
            {
                $DesiredProxyAddresses += @($add)
            }
        }
        if($PSBoundParameters.ContainsKey('DesiredPrimaryAddress'))
        {
            $currentPrimary = @($DesiredProxyAddresses | Where-Object {$_ -clike 'SMTP:*'} | ForEach-Object {$_.split(':')[1]})
            switch ($currentPrimary.count)
            {
                1
                {
                    if (-not $currentPrimary[0] -ceq $DesiredPrimaryAddress)
                    {
                        $DesiredProxyAddresses = @($DesiredProxyAddresses | where-object {$_ -notlike "smtp:$DesiredPrimaryAddress"})
                        $DesiredProxyAddresses = @($DesiredProxyAddresses | where-object {$_ -notlike "SMTP:$($currentPrimary[0])"})
                        $DesiredProxyAddresses += $("smtp:$($currentPrimary[0])")
                        $DesiredProxyAddresses += $("SMTP:$DesiredPrimaryAddress")
                    }
                }#end 1
                0
                {
                    $DesiredProxyAddresses += $("SMTP:$DesiredPrimaryAddress")
                }
                {$_ -ge 2}
                {
                    throw('Multiple Primary SMTP addresses detected: Invalid Configuration')
                }
            }#end switch

        }#end if
        if ($VerifySMTPAddressValidity -eq $true)
        {
            $SMTPProxyAddresses = @($DesiredProxyAddresses | Where-Object {$_ -ilike 'smtp:*'})
            foreach ($spa in $SMTPProxyAddresses)
            {
                if (Test-EmailAddress -EmailAddress $spa.split(':')[1])
                {}
                else
                {
                    Write-OneShellLog -Message "SMTP Proxy Address $spa appears to be invalid." -ErrorLog -EntryType Failed
                    $DesiredProxyAddresses = $DesiredProxyAddresses | Where-Object {$_ -ne $spa}
                }
            }
        }
        switch ($DesiredProxyAddresses)
        {
            {$PSBoundParameters.ContainsKey('DomainsToRemove')}
            {
                $DesiredProxyAddresses = $DesiredProxyAddresses | Where-Object {$_.split('@')[1] -notin $DomainsToRemove}
            }
            {$PSBoundParameters.ContainsKey('AddressesToRemove')}
            {
                $DesiredProxyAddresses = $DesiredProxyAddresses | Where-Object {$_ -notin $AddressesToRemove}
            }
        }
        if ($true -eq $TestAddressAvailability)
        {
            $Passed = @()
            $Failed = @()
            foreach ($dpa in $DesiredProxyAddresses)
            {
                switch (Test-ExchangeProxyAddress -ProxyAddress $dpa -ProxyAddressType SMTP -ExchangeSession $TestAddressExchangeOrganizationSession)
                {
                    $true
                    {
                        $Passed += $dpa
                    }
                    $false
                    {
                        $Failed += $dpa
                    }
                }
            }
            $DesiredProxyAddresses = $Passed
        }
        $DesiredProxyAddresses = @($DesiredProxyAddresses | Select-Object -Unique)
        #test for one Primary
        $PrimaryCount = @($DesiredProxyAddresses | Where-object {$_ -clike 'SMTP:*'}).Count
        if ($PrimaryCount -ne 1)
        {throw ("$PrimaryCount Primary Addresses Generated")}
        else
        {
            $DesiredProxyAddresses
        }
    }
