Function Get-AltDesiredProxyAddresses
{
    [cmdletbinding()]
    param
    (
        [parameter()]
        [string[]]$CurrentProxyAddresses #Current proxy addresses to preserve or evaluate for preservation
        ,
        [parameter()]
        [ValidateScript({$_ -clike 'SMTP:*'})]
        [string]$DesiredPrimarySMTPAddress #replace existing primary smtp address with this value
        ,
        [string]$DesiredOrCurrentAlias #used for calculation of a TargetAddress if required.
        ,
        [parameter()]
        [ValidateScript({$_ -notlike 'x500:*'})]
        [string[]]$LegacyExchangeDNs #legacyexchangedn to convert to additional x500 address
        ,
        [psobject[]]$Recipients #Recipient objects to consume for their proxy addresses and legacyexchangedn
        ,
        [parameter()]
        [switch]$AddPrimarySMTPAddressForAlias
        ,
        [parameter()]
        [switch]$AddTargetSMTPAddress #have the function ensure inclusion of a targetdeliverydomain proxy address.  Requires the TargetDeliveryDomain and DesiredOrCurrentAlias parameters.
        ,
        [parameter()]
        [switch]$VerifyTargetSMTPAddress
        ,
        [string]$TargetDeliverySMTPDomain #specify the external/remote delivery domain - usually for cross forest or cloud like contoso.mail.onmicrosoft.com
        ,
        [string]$PrimarySMTPDomain #specify the Primary delivery domain - usually the main public name like 'contoso.com'
        ,
        [string[]]$DomainsToRemove #specify the domains for which to remove the associated proxy addresses. Include only the domain name, like 'contoso.com'
        ,
        [string[]]$DomainsToAdd #specify the domains for which to remove the associated proxy addresses. Include only the domain name, like 'contoso.com'
        ,
        [parameter()]
        [ValidateScript({@($_ | % {$_ -like '*:*'}) -notcontains $false}]
        [string[]]$AddressesToRemove #specify the complete address including the type: prefix, like smtp: or x500:
        ,
        [parameter()]
        [ValidateScript({@($_ | % {$_ -like '*:*'}) -notcontains $false}]
        [string[]]$AddressesToAdd #specifcy the complete address including the type: prefix, like smtp: or x500:
        ,
        [switch]$VerifySMTPAddressValidity #verifies that the SMTP address complies with basic format requirements to be valid. See documentation for Test-EmailAddress for more information.
        ,
        [System.Management.Automation.Runspaces.PSSession]$ExchangeSession
        ,
        [switch]$TestAddressAvailabilityInExchangeSession
    )
    #parameter validation(s)
    if (($PSBoundParameters.ContainsKey('AddTargetSMTPAddress') -or $PSBoundParameters.ContainsKey('VerifyTargetSMTPAddress') -or $PSBoundParameters.ContainsKey('AddPrimarySMTPAddressForAlias')) -and -not $PSBoundParameters.ContainsKey('DesiredOrCurrentAlias'))
    {
        throw('Parameters AddTargetSMTPAddressForAlias, VerifyTargetSMTPAddress, and AddPrimarySMTPAddressForAlias require a value for Parameter DesiredOrCurrentAlias. Please provide a value for parameter DesiredOrCurrentAlias and try again.')
        return $null
    }
    if ($PSBoundParameters.ContainsKey('AddPrimarySMTPAddressForAlias') -and -not $PSBoundParameters.ContainsKey('PrimarySMTPDomain'))
    {
        throw('Parameter AddPrimarySMTPAddressForAlias required a value for Parameter PrimarySMTPDomain. Please provide a value for parameter PrimarySMTPDomain and try again.')
        return $null
    }
    if (($PSBoundParameters.ContainsKey('AddTargetSMTPAddressForAlias') -or $PSBoundParameters.ContainsKey('VerifyTargetSMTPAddress')) -and -not $PSBoundParameters.ContainsKey('TargetDeliverySMTPDomain'))
    {
        throw('Parameters AddTargetSMTPAddressForAlias or VerifyTargetSMTPAddress require a value for Parameter TargetDeliverySMTPDomain. Please provide a value for parameter TargetDeliverySMTPDomain and try again.')
        return $null
    }
    # Gather All Addresses From Input
    $AllIncomingProxyAddresses = New-Object System.Collections.ArrayList
    if ($PSBoundParameters.ContainsKey('CurrentProxyAddresses'))
    {
        foreach ($cpa in $CurrentProxyAddresses)
        {
            if ($null -eq $cpa.split(':')[1] -and $null -ne $cpa.split('@')[1])
            {
                $cpa = 'smtp:' + $cpa
            }
            $null = $AllIncomingProxyAddresses.Add($cpa)
        }
    }
    if ($PSBoundParameters.ContainsKey('DesiredPrimarySMTPAddress'))
    {
        foreach ($cpa in $CurrentProxyAddresses)
        {
            $null = $AllIncomingProxyAddresses.Add($cpa)
        }
    }
    if ($LegacyExchangeDNs.Count -ge 1)
    {
        foreach ($l in $LegacyExchangeDNs)
        {
            $existingProxyAddressTypes = Get-ExistingProxyAddressTypes -proxyAddresses $AllIncomingProxyAddresses
            $type = 'X500'
            if ($existingProxyAddressTypes -ccontains $type)
            {
                $type = $type.ToLower()
            }
            $newX500 = "$type`:$l"
            if ($newX500 -notin $AllIncomingProxyAddresses)
            {
                $null = $AllIncomingProxyAddresses.Add($newX500)
            }
        }
    }
    if ($Recipients.Count -ge 1)
    {
        $RecipientProxyAddresses = @()
        foreach ($r in $Recipients)
        {
            $paProperty = if (Test-Member -InputObject $r -Name emailaddresses) {'EmailAddresses'} elseif (Test-Member -InputObject $r -Name proxyaddresses ) {'proxyAddresses'} else {$null}
            if ($paProperty)
            {
                $existingProxyAddressTypes = Get-ExistingProxyAddressTypes -proxyAddresses $AllIncomingProxyAddresses
                $rpa = @($r.$paProperty)
                foreach ($a in $rpa)
                {
                    $type = $a.split(':')[0]
                    $address = $a.split(':')[1]
                    if ($existingProxyAddressTypes -ccontains $type)
                    {
                        $la = $type.tolower() + ':' + $address
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
            foreach ($npa in $RecipientProxyAddresses)
            {
                if ($npa -cnotin $AllIncomingProxyAddresses)
                {
                    $null = $AllIncomingProxyAddresses.Add($npa)
                }
            }
        }#if
    }#if
    if ($AddressesToAdd.Count -ge 1)
    {
        foreach ($a in $AddressesToAdd)
        {
            if ($a -inotin $AllIncomingProxyAddresses)
            {
                $AllIncomingProxyAddresses.Add($a)
            }
        }
    }
    if ($VerifyTargetSMTPAddress -eq $true)
    {
        $existingdomains = @($AllIncomingProxyAddresses | ForEach-Object {$_.split('@')[1]} | Select-Object -Unique)
        if ($TargetDeliverySMTPDomain -notin $existingdomains)
        {
            $NewTargetDeliverySMTPAddress = 'smtp:' + $DesiredOrCurrentAlias + '@' + $TargetDeliverySMTPDomain
            $AllIncomingProxyAddresses.Add($NewTargetDeliverySMTPAddress)
        }
    }#if
    if ($AddTargetSMTPAddress -eq $true)
    {
        $NewTargetDeliverySMTPAddress = 'smtp:' + $DesiredOrCurrentAlias + '@' + $TargetDeliverySMTPDomain
        $AllIncomingProxyAddresses.Add($NewTargetDeliverySMTPAddress)
    }#if

    if ($PSBoundParameters.ContainsKey('DesiredPrimaryAddress'))
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
