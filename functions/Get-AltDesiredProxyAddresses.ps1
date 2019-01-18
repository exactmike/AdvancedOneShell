Set-StrictMode -Version Latest
Function Get-AltDesiredProxyAddresses
{
    [cmdletbinding()]
    param
    (
        [parameter()]
        [ValidateScript({@($_ | % {$_ -like '*:*'}) -notcontains $false})]
        [string[]]$CurrentProxyAddresses #Current proxy addresses to preserve or evaluate for preservation
        ,
        [parameter()]
        [ValidateScript({$_ -clike 'SMTP:*'})]
        [string]$DesiredPrimarySMTPAddress #replace existing primary smtp address with this value
        ,
        [parameter()]
        [string]$DesiredOrCurrentAlias #used for calculation of a TargetAddress if required.
        ,
        [parameter()]
        [ValidateScript({$_ -notlike 'x500:*'})]
        [string[]]$LegacyExchangeDNs #legacyexchangedn to convert to additional x500 address
        ,
        [parameter()]
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
        [parameter()]
        [string]$TargetDeliverySMTPDomain #specify the external/remote delivery domain - usually for cross forest or cloud like contoso.mail.onmicrosoft.com
        ,
        [parameter()]
        [ValidateScript({$_ -like '*.*'})]
        [string]$PrimarySMTPDomain #specify the Primary delivery domain - usually the main public name like 'contoso.com'
        ,
        [parameter()]
        [ValidateScript({@($_ | % {$_ -like '*:*'}) -notcontains $false})]
        [string[]]$AddressesToRemove #specify the complete address including the type: prefix, like smtp: or x500:
        ,
        [parameter()]
        [ValidateScript({@($_ | % {$_ -like '*:*'}) -notcontains $false})]
        [string[]]$AddressesToAdd #specifcy the complete address including the type: prefix, like smtp: or x500:
        ,
        [parameter()]
        [ValidateScript({@($_ | % {$_ -like '*.*'}) -notcontains $false})]
        [string[]]$DomainsToAdd #specify the domains for which to remove the associated proxy addresses. Include only the domain name, like 'contoso.com'
        ,
        [parameter()]
        [switch]$ForceDomainsToAdd #adds a new address for any domain in $DomainsToAdd using the DesiredOrCurrentAlias, even if another proxy address already exists in the domain
        ,
        [parameter()]
        [ValidateScript({@($_ | % {$_ -like '*.*'}) -notcontains $false})]
        [string[]]$DomainsToRemove #specify the domains for which to remove the associated proxy addresses. Include only the domain name, like 'contoso.com'
        ,
        [parameter()]
        [switch]$VerifySMTPAddressValidity #verifies that the SMTP address complies with basic format requirements to be valid. See documentation for Test-EmailAddress for more information.
        ,
        [parameter()]
        [System.Management.Automation.Runspaces.PSSession]$ExchangeSession
        ,
        [parameter()]
        [switch]$TestAddressAvailabilityInExchangeSession
        ,
        [parameter()]
        [string[]]$TestAddressAvailabilityExemptGUID
    )
    #parameter validation(s)
    if (($true -eq $AddTargetSMTPAddress -or $true -eq $VerifyTargetSMTPAddress -or $true -eq $AddPrimarySMTPAddressForAlias) -and -not $PSBoundParameters.ContainsKey('DesiredOrCurrentAlias'))
    {
        throw('Parameters AddTargetSMTPAddressForAlias, VerifyTargetSMTPAddress, and AddPrimarySMTPAddressForAlias require a value for Parameter DesiredOrCurrentAlias. Please provide a value for parameter DesiredOrCurrentAlias and try again.')
        return $null
    }
    if ($true -eq $AddPrimarySMTPAddressForAlias -and -not $PSBoundParameters.ContainsKey('PrimarySMTPDomain'))
    {
        throw('Parameter AddPrimarySMTPAddressForAlias required a value for Parameter PrimarySMTPDomain. Please provide a value for parameter PrimarySMTPDomain and try again.')
        return $null
    }
    if (($true -eq $AddTargetSMTPAddress -or $true -eq $VerifyTargetSMTPAddress) -and -not $PSBoundParameters.ContainsKey('TargetDeliverySMTPDomain'))
    {
        throw('Parameters AddTargetSMTPAddressForAlias or VerifyTargetSMTPAddress require a value for Parameter TargetDeliverySMTPDomain. Please provide a value for parameter TargetDeliverySMTPDomain and try again.')
        return $null
    }
    if ($PSBoundParameters.ContainsKey('DomainsToAdd') -and -not $PSBoundParameters.ContainsKey('DesiredorCurrentAlias'))
    {
        throw('Parameter DomainsToAdd requires a value for Parameter DesiredOrCurrentAlias. Please provide a value for parameter DesiredOrCurrentAlias and try again.')
        return $null
    }
    # First Add all specified/requested addresses.  We won't validate single primary of each type until the end
    $newPrimarySMTPAddressForAlias = $null
    $PromotedPrimarySMTPAddress = $null
    $AllIncomingProxyAddresses = New-Object System.Collections.ArrayList
    if ($PSBoundParameters.ContainsKey('CurrentProxyAddresses'))
    {
        foreach ($cpa in $CurrentProxyAddresses)
        {
            $null = $AllIncomingProxyAddresses.Add($cpa)
        }
    }
    if ($PSBoundParameters.ContainsKey('LegacyExchangeDNs'))
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
    if ($PSBoundParameters.ContainsKey('Recipients'))
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
    if ($PSBoundParameters.ContainsKey('AddressesToAdd'))
    {
        foreach ($a in $AddressesToAdd)
        {
            if ($a -inotin $AllIncomingProxyAddresses)
            {
                $null = $AllIncomingProxyAddresses.Add($a)
            }
        }
    }
    if ($true -eq $VerifyTargetSMTPAddress)
    {
        $existingdomains = @($AllIncomingProxyAddresses | ForEach-Object {$_.split('@')[1]} | Select-Object -Unique)
        if ($TargetDeliverySMTPDomain -notin $existingdomains)
        {
            [string]$NewTargetDeliverySMTPAddress = 'smtp:' + $DesiredOrCurrentAlias + '@' + $TargetDeliverySMTPDomain
            $null = $AllIncomingProxyAddresses.Add($NewTargetDeliverySMTPAddress)
        }
    }#if
    if ($true -eq $AddTargetSMTPAddress)
    {
        $NewTargetDeliverySMTPAddress = 'smtp:' + $DesiredOrCurrentAlias + '@' + $TargetDeliverySMTPDomain
        $null = $AllIncomingProxyAddresses.Add($NewTargetDeliverySMTPAddress)
    }#if
    if ($PSBoundParameters.ContainsKey('DomainsToAdd'))
    {
        foreach ($d in $DomainsToAdd)
        {
            if (($AllIncomingProxyAddresses | Where-Object {$_ -like "*@$d"}).count -eq 0 -or $true -eq $ForceDomainsToAdd)
            {
                [string]$newDomainSMTPAddress = 'smtp:' + $DesiredOrCurrentAlias + '@' + $d
                if ($AllIncomingProxyAddresses -notcontains $newDomainSMTPAddress)
                {
                    $null = $AllIncomingProxyAddresses.Add($newDomainSMTPAddress)
                }
            }
        }
    }
    if ($PSBoundParameters.ContainsKey('DesiredPrimarySMTPAddress') -or $PSBoundParameters.ContainsKey('PrimarySMTPDomain') -or $true -eq $AddPrimarySMTPAddressForAlias)
    {
        if ($PSBoundParameters.ContainsKey('DesiredPrimarySMTPAddress'))
        {
            if ($AllIncomingProxyAddresses -inotcontains $DesiredPrimarySMTPAddress)
            {
                $null = $AllIncomingProxyAddresses.Add($DesiredPrimarySMTPAddress)
            }
            elseif (@($AllIncomingProxyAddresses | ForEach-Object {$_ -ilike $DesiredPrimarySMTPAddress}) -contains $true)
            {
                #do until it's gone then put back in
                do
                {
                    $removethis = $AllIncomingProxyAddresses | Where-Object {$_ -ilike $DesiredPrimarySMTPAddress} | Select-Object -First 1
                    $AllIncomingProxyAddresses.Remove($removethis)
                }
                until (@($AllIncomingProxyAddresses | ForEach-Object {$_ -ilike $DesiredPrimarySMTPAddress}) -notcontains $true)
                $null = $AllIncomingProxyAddresses.Add($DesiredPrimarySMTPAddress)
            }
        }
        if ($PSBoundParameters.ContainsKey('PrimarySMTPDomain'))
        {
            if (($AllIncomingProxyAddresses | ForEach-Object {$_ -clike "SMTP:*@*" -and $_ -ilike "*@$PrimarySMTPDomain"}) -notcontains $true)
            {
                if (($AllIncomingProxyAddresses | ForEach-Object {$_ -clike "smtp:*@*" -and $_ -ilike "*@$PrimarySMTPDomain"}) -notcontains $true)
                {
                    [string]$newPrimarySMTPAddressForAlias = 'SMTP:' + $DesiredOrCurrentAlias + '@' + $PrimarySMTPDomain
                    $null = $AllIncomingProxyAddresses.Add($newPrimarySMTPAddressForAlias)
                }
                else
                {
                    $potentialPrimarySMTPAddresses = @($AllIncomingProxyAddresses | Where-Object {$_ -like "*@$PrimarySMTPDomain"})
                    switch ($potentialPrimarySMTPAddresses.count)
                    {
                        1
                        {
                            do
                            {
                                $AllIncomingProxyAddresses.Remove($potentialPrimarySMTPAddresses[0])
                            }
                            until ($AllIncomingProxyAddresses -notcontains $potentialPrimarySMTPAddresses[0])
                            [string]$PromotedPrimarySMTPAddress = 'SMTP:' + $potentialPrimarySMTPAddresses[0].split(':')[1]
                            $null = $AllIncomingProxyAddresses.Add($PromotedPrimarySMTPAddress)
                        }
                        {$_ -gt 1}
                        {
                            if ($PSBoundParameters.ContainsKey('DesiredOrCurrentAlias'))
                            {
                                [string]$AliasAndPrimarySMTPDomainAddress = $('SMTP:' + $DesiredOrCurrentAlias + '@' + $PrimarySMTPDomain)
                                if ($potentialPrimarySMTPAddresses -contains $AliasAndPrimarySMTPDomainAddress)
                                {
                                    do
                                    {
                                        $removethis = $potentialPrimarySMTPAddresses | Where-Object {$_ -ilike $AliasAndPrimarySMTPDomainAddress} | Select-Object -First 1
                                        $AllIncomingProxyAddresses.Remove($removethis)
                                    }
                                    until ($AllIncomingProxyAddresses -notcontains $AliasAndPrimarySMTPDomainAddress)
                                    [string]$PromotedPrimarySMTPAddress = $AliasAndPrimarySMTPDomainAddress
                                    $null = $AllIncomingProxyAddresses.Add($PromotedPrimarySMTPAddress)
                                }
                            }
                            else
                            {
                                $PromotedPrimarySMTPAddress = 'SMTP:' + $potentialPrimarySMTPAddresses[0].split(':')[1]
                                do
                                {
                                    $AllIncomingProxyAddresses.Remove($potentialPrimarySMTPAddresses[0])
                                }
                                until ($AllIncomingProxyAddresses -notcontains $PromotedPrimarySMTPAddress)
                                $null = $AllIncomingProxyAddresses.Add($PromotedPrimarySMTPAddress)
                            }
                        }
                        0
                        {
                            if ($PSBoundParameters.ContainsKey('DesiredOrCurrentAlias'))
                            {
                                [string]$newPrimarySMTPAddressForAlias = $('SMTP:' + $DesiredOrCurrentAlias + '@' + $PrimarySMTPDomain)
                                $null = $AllIncomingProxyAddresses.Add($newPrimarySMTPAddressForAlias)
                            }
                            else
                            {
                                throw ('PrimarySMTPDomain was specified but no proxy addresses are in that namespace and DesiredOrCurrentAlias parameter was not specified.')
                            }
                        }
                    }
                }
            }
        }
        if ($PSBoundParameters.ContainsKey('AddPrimarySMTPAddressForAlias'))
        {
            [string]$AliasAndPrimarySMTPDomainAddress = $('SMTP:' + $DesiredOrCurrentAlias + '@' + $PrimarySMTPDomain)
            if ($AllIncomingProxyAddresses -contains $AliasAndPrimarySMTPDomainAddress)
            {
                do
                {
                    $removethis = $AllIncomingProxyAddresses | Where-Object {$_ -ilike $AliasAndPrimarySMTPDomainAddress} | Select-Object -First 1
                    $AllIncomingProxyAddresses.Remove($removethis)
                }
                until ($AllIncomingProxyAddresses -notcontains $AliasAndPrimarySMTPDomainAddress)
                [string]$PromotedPrimarySMTPAddress = $AliasAndPrimarySMTPDomainAddress
                $null = $AllIncomingProxyAddresses.Add($PromotedPrimarySMTPAddress)
            }
            else
            {
                [string]$newPrimarySMTPAddressForAlias = $AliasAndPrimarySMTPDomainAddress
                $null = $AllIncomingProxyAddresses.Add($newPrimarySMTPAddressForAlias)
            }
        }
    }
    # Now Remove all specified addresses and domains
    if ($PSBoundParameters.ContainsKey('AddressesToRemove'))
    {
        foreach ($a in $AddressesToRemove)
        {
            if ($AllIncomingProxyAddresses -icontains $a)
            {
                do
                {
                    $removeThis = $AllIncomingProxyAddresses | Where-Object {$_ -ilike $a} | Select-Object -First 1
                    $AllIncomingProxyAddresses.Remove($removethis)
                }
                until ($AllIncomingProxyAddresses -inotcontains $a)
            }
        }
    }
    if ($PSBoundParameters.ContainsKey('DomainsToRemove'))
    {
        foreach ($d in $DomainsToRemove)
        {
            $dAddresses = @($AllIncomingProxyAddresses | Where-Object -FilterScript {$_ -ilike "*@$d"})
            foreach ($da in $dAddresses)
            {
                do
                {
                    $AllIncomingProxyAddresses.Remove($da)
                }
                until ($AllIncomingProxyAddresses -notcontains $da)
            }
        }
    }
    #Various validations
    #Basic smtp email address validitity check
    if ($VerifySMTPAddressValidity -eq $true)
    {
        $SMTPProxyAddresses = @($AllIncomingProxyAddresses | Where-Object {$_ -ilike 'smtp:*'})
        foreach ($spa in $SMTPProxyAddresses)
        {
            if (-not (Test-EmailAddress -EmailAddress $spa.split(':')[1]))
            {
                $AllIncomingProxyAddresses.Remove($spa)
            }
        }
    }
    #Availability of smtp email address in a target Exchange Org
    if ($true -eq $TestAddressAvailabilityInExchangeSession)
    {
        foreach ($i in $($AllIncomingProxyAddresses | Where-Object -FilterScript {$_ -ilike 'smtp:*'}))
        {
            $TestParams = @{
                ProxyAddress = $i
                ProxyAddressType = 'SMTP'
                ExchangeSession = $ExchangeSession
            }
            if ($PSBoundParameters.ContainsKey('TestAddressAvailabilityExemptGUID'))
            {
                $TestParams.ExemptObjectGUIDs = $TestAddressAvailabilityExemptGUID
            }
            if (-not (Test-ExchangeProxyAddress @TestParams))
            {
                $AllIncomingProxyAddresses.Remove($i)
            }
        }
    }
    #Duplicate check
    $Duplicates = @($AllIncomingProxyAddresses | Group-Object | Where-Object -FilterScript {$_.count -gt 1})
    if ($Duplicates.Count -ne 0)
    {
        foreach ($d in $Duplicates)
        {
            $da = $d.Group[0]
            do
            {
                $RemoveThis = $AllIncomingProxyAddresses | Where-Object {$_ -ilike $da} | Select-Object -First 1
                $AllIncomingProxyAddresses.Remove($RemoveThis)
            }
            until ($AllIncomingProxyAddresses -inotcontains $da)
            switch ($da)
            {
                {$DesiredPrimarySMTPAddress -eq $_}
                {
                    $null = $AllIncomingProxyAddresses.Add($DesiredPrimarySMTPAddress)
                    break
                }
                {$newPrimarySMTPAddressForAlias -eq $_}
                {
                    $null = $AllIncomingProxyAddresses.Add($newPrimarySMTPAddressForAlias)
                    break
                }
                {$PromotedPrimarySMTPAddress -eq $_}
                {
                    $null = $AllIncomingProxyAddresses.Add($PromotedPrimarySMTPAddress)
                    break
                }
                {$AddressesToAdd -contains $_}
                {
                    $add = $AddressesToAdd | Where-Object {$_ -eq $_} | Select-Object -First 1
                    $null = $AllIncomingProxyAddresses.Add($add)
                    break
                }
                {$CurrentProxyAddresses -contains $_}
                {
                    $add = $CurrentProxyAddresses | Where-Object {$_ -eq $_} | Select-Object -First 1
                    $null = $AllIncomingProxyAddresses.Add($add)
                    break
                }
                default
                {
                    $add = $da.split(':')[0].tolower() + ':' + $da.split(':')[1]
                    $null = $AllIncomingProxyAddresses.Add($add)
                }
            }
        }
    }
    #Multiple Primary Check for SMTP (x500 doesn't matter - maybe we should add a SIP check too?)
    $MultiplePrimarySMTP = @($AllIncomingProxyAddresses | ForEach-Object {$_.split(':')[0]} | Where-Object {$_ -eq 'smtp' -and $_ -ceq $_.toupper()} | Group-Object | Where-Object {$_.count -gt 1})
    if ($MultiplePrimarySMTP.count -ne 0)
    {
        foreach ($a in @($AllIncomingProxyAddresses | Where-Object {$_ -clike 'SMTP:*'}))
        {
            [string]$demoteda = 'smtp:' + $a.split(':')[1]
            $AllIncomingProxyAddresses.Remove($a)
            $null = $AllIncomingProxyAddresses.Add($demoteda)
        }
        switch ($true)
        {
            {$PSBoundParameters.ContainsKey('DesiredPrimarySMTPAddress')}
            {
                do
                {
                    $RemoveThis = $AllIncomingProxyAddresses | Where-Object {$_ -ilike $DesiredPrimarySMTPAddress} | Select-Object -First 1
                    $AllIncomingProxyAddresses.Remove($RemoveThis)
                }
                until ($AllIncomingProxyAddresses -inotcontains $DesiredPrimarySMTPAddress)
                $null = $AllIncomingProxyAddresses.Add($DesiredPrimarySMTPAddress)
                break
            }
            {$null -ne $newPrimarySMTPAddressForAlias}
            {
                do
                {
                    $RemoveThis = $AllIncomingProxyAddresses | Where-Object {$_ -ilike $newPrimarySMTPAddressForAlias} | Select-Object -First 1
                    $AllIncomingProxyAddresses.Remove($RemoveThis)
                }
                until ($AllIncomingProxyAddresses -inotcontains $newPrimarySMTPAddressForAlias)
                $null = $AllIncomingProxyAddresses.Add($newPrimarySMTPAddressForAlias)
                break
            }
            {$null -ne $PromotedPrimarySMTPAddress}
            {
                do
                {
                    $RemoveThis = $AllIncomingProxyAddresses | Where-Object {$_ -ilike $PromotedPrimarySMTPAddress} | Select-Object -First 1
                    $AllIncomingProxyAddresses.Remove($RemoveThis)
                }
                until ($AllIncomingProxyAddresses -inotcontains $PromotedPrimarySMTPAddress)
                $null = $AllIncomingProxyAddresses.Add($PromotedPrimarySMTPAddress)
                break
            }
            #Order is important here - give preference to the CurrentProxyAddresses for determining Primary SMTP Address if there was no explicit request to change it via params
            {$PSBoundParameters.ContainsKey('CurrentProxyAddresses')}
            {
                $Add = $CurrentProxyAddresses | Where-Object {$_ -clike 'SMTP:*'} | Select-Object -First 1
                if ($null -ne $Add)
                {
                    do
                    {
                        $RemoveThis = $CurrentProxyAddresses | Where-Object {$_ -ilike $Add} | Select-Object -First 1 | ForEach-Object {'smtp:' + $_.split(':')[1]}
                        $AllIncomingProxyAddresses.Remove($RemoveThis)
                    }
                    until ($AllIncomingProxyAddresses -inotcontains $Add)
                    $null = $AllIncomingProxyAddresses.Add($Add)
                    break
                }
            }
            {$PSBoundParameters.ContainsKey('AddressesToAdd')}
            {
                $Add = $AddressesToAdd | Where-Object {$_ -clike 'SMTP:*'} | Select-Object -First 1
                if ($null -ne $Add)
                {
                    do
                    {
                        $RemoveThis = $AddressesToAdd | Where-Object {$_ -ilike $Add} | Select-Object -First 1 | ForEach-Object {'smtp:' + $_.split(':')[1]}
                        $AllIncomingProxyAddresses.Remove($RemoveThis)
                    }
                    until ($AllIncomingProxyAddresses -inotcontains $Add)
                    $null = $AllIncomingProxyAddresses.Add($Add)
                    break
                }
            }
        }
    }
    $AllIncomingProxyAddresses
}


#Get-AltDesiredProxyAddresses -CurrentProxyAddresses 'X500:/o=NotReallyNeeded','x500:/o=webmd/ou=External (FYDIBOHF25SPDLT)/cn=Recipients/cn=846992148f6a4b0da7d6d0092dbdf5db','smtp:messagelabsdataprotection@emdeon.com','smtp:AMAIN@emdeon.mail.onmicrosoft.com','x500:/o=webmd/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=Main, Andrew31e','X500:/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=553e77e80d224ef1a945ed561ac0226c-Main, Andre','smtp:andrewmain@emdeon.com','smtp:andrew.main@changehealthcare.com','smtp:anmain@emdeon.com','SMTP:anmain@changehealthcare.com','smtp:amain@changehealthcare.com','smtp:amain@patientbusinessoffice.com','smtp:amain@patientaccountsoffice.org','smtp:amain@chapinrcm.com','smtp:amain@chamberlinedmonds.com','smtp:amain@ce-a.com','smtp:amain@erxnetwork.com','smtp:maina@tc3health.com','smtp:amaintest@changehealthcare.com','smtp:amain@htms.com','smtp:amain1@htms.com','smtp:AMAIN317@emdeon.mail.onmicrosoft.com','smtp:anmain@equiclaim.com','smtp:anmain@ghsinc.com','smtp:amain@capario.com','smtp:anmain@myadvocatehelps.com','smtp:anmain@altegrahealth.com','smtp:amain@fvtech.com','smtp:amain@iadur.com','smtp:amain@iowamedicaidpdl.com','smtp:amain@iowamedicaidpos.com','smtp:amain@msmedicaidrxportal.com','smtp:amain@rxssdc.com','smtp:amain@sentinelgrouponline.com','smtp:relayer_testing@ce-a.com' -PrimarySMTPDomain 'emdeon.com' | Group-Object
#Get-AltDesiredProxyAddresses -CurrentProxyAddresses 'smtp:user@contoso.com','SMTP:user.name@contoso.com','smtp:user@fabrikam.com' -PrimarySMTPDomain 'fabrikam'
#Get-AltDesiredProxyAddresses -CurrentProxyAddresses 'smtp:user@contoso.com','SMTP:user.name@contoso.com','smtp:user@fabrikam.com' -PrimarySMTPDomain 'fabrikam.com' -AddressesToAdd 'SMTP:user@northwindtraders.com'
Get-AltDesiredProxyAddresses -CurrentProxyAddresses 'smtp:user@contoso.com','SMTP:user.name@contoso.com','smtp:user@fabrikam.com' -AddressesToAdd 'SMTP:user@northwindtraders.com'