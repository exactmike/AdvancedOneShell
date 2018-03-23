###############################################################################################
#Module Variables and Variable Functions
###############################################################################################
function Get-AOSVariable
    {
        param
        (
        [string]$Name
        )
        Get-Variable -Scope Script -Name $name
    }
#end function Get-AOSVariable
function Get-AOSVariableValue
    {
        param
        (
        [string]$Name
        )
        Get-Variable -Scope Script -Name $name -ValueOnly
    }
#end function Get-AOSVariableValue
function Set-AOSVariable
    {
        param
        (
        [string]$Name
        ,
        $Value
        )
        Set-Variable -Scope Script -Name $Name -Value $value
    }
#end function Set-AOSVariable
function New-AOSVariable
    {
        param
        (
        [string]$Name
        ,
        $Value
        )
        New-Variable -Scope Script -Name $name -Value $Value
    }
#end function New-AOSVariable
function Remove-AOSVariable
    {
        param
        (
        [string]$Name
        )
        Remove-Variable -Scope Script -Name $name
    }
#end function Remove-AOSVariable
###############################################################################################
#Core Advanced OneShell Functions
###############################################################################################
function Get-ExistingProxyAddressTypes
    {
        param(
        [object[]]$proxyAddresses
        )
        $ProxyAddresses | ForEach-Object -Process {$_.split(':')[0]} | Sort-Object | Select-Object -Unique
    }
#end function Get-ExistingProxyAddressTypes
function Get-DesiredProxyAddresses
    {
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
        )
        if ($PSBoundParameters.ContainsKey('CurrentProxyAddresses'))
        {
            $DesiredProxyAddresses = $CurrentProxyAddresses.Clone()
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
                    Write-Log -Message "SMTP Proxy Address $spa appears to be invalid." -ErrorLog -EntryType Failed
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
        $DesiredProxyAddresses
    }
#end function get-desiredproxyaddresses
Function Export-FailureRecord
    {
        [cmdletbinding()]
        param(
        [string]$Identity
        ,
        [string]$ExceptionCode
        ,
        [string]$FailureGroup
        ,
        [string]$ExceptionDetails
        ,
        [string]$RelatedObjectIdentifier
        ,
        [string]$RelatedObjectIdentifierType
        )#Param
        $Exception=[ordered]@{
            Identity = $Identity
            ExceptionCode = $ExceptionCode
            ExceptionDetails = $ExceptionDetails
            FailureGroup = $FailureGroup
            RelatedObjectIdentifier = $RelatedObjectIdentifier
            RelatedObjectIdentifierType = $RelatedObjectIdentifierType
            TimeStamp = Get-TimeStamp
        }
        try {
        $ExceptionObject = $Exception | Convert-HashTableToObject
        Export-Data -DataToExportTitle $FailureGroup -DataToExport $ExceptionObject -Append -DataType csv -ErrorAction Stop
        $Global:SEATO_Exceptions += $ExceptionObject
        }
        catch {
        Write-Log -Message "FAILED: to write Exception Record for $identity with Exception Code $ExceptionCode and Failure Group $FailureGroup" -ErrorLog
        }
    }
#end Function Export-FailureRecord
function Move-StagedADObjectToOperationalOU
    {
        param(
        [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [string[]]$Identity
        ,
        [string]$DestinationOU
        )
        begin {}
        process {
            foreach ($I in $Identity) {
                try {
                    $message = "Find AD Object: $I"
                    Write-Log -Message $message -EntryType Attempting
                    $aduser = Get-ADObject -Identity $I -ErrorAction Stop
                    Write-Log -Message $message -EntryType Succeeded
                }#try
                catch {
                    Write-Log -Message $message -Verbose -EntryType Failed -ErrorLog
                    Write-Log -Message $_.tostring() -ErrorLog
                }#catch
                try {
                    $message = "Move-ADObject -Identity $I -TargetPath $DestinationOU"
                    Write-Log -Message $message -EntryType Attempting
                    $aduser | Move-ADObject -TargetPath $DestinationOU -ErrorAction Stop
                    Write-Log -Message $message -EntryType Succeeded
                }#try
                catch {
                    Write-Log -Message $message -Verbose -ErrorLog -EntryType Failed
                    Write-Log -Message $_.tostring() -ErrorLog
                }#catch
            }#foreach
        }
        end{}
    }
#end function Move-StagedADObjectToOperationalOU
function Update-PostMigrationMailboxUser
    {
        [cmdletbinding()]
        param(
        [parameter(
            Mandatory=$true,
            ParameterSetName = 'Individual'
            )]
        [string[]]$Identity
        ,
        [parameter(
            Mandatory=$true,
            ParameterSetName = 'InputList'
            )]
        [array]$InputList
        ,
        [parameter(Mandatory=$true)]
        [string[]]$SourceAD
        ,
        [parameter(Mandatory=$true)]
        [string]$TargetAD
        ,
        [parameter(Mandatory=$true)]
        [string]$TargetExchangeOrg
        ,
        [parameter(Mandatory = $true)]
        [validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid')]
        [string]$TargetLookupAttribute
        ,
        [parameter(Mandatory = $true)]
        [validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid')]
        [string]$SourceLookupAttribute
        ,
        [parameter(Mandatory = $true)]
        [validateset('UserMailbox','LinkedMailbox','LegacyMailbox','MailUser','RemoteUserMailbox','RemoteSharedMailbox','RemoteRoomMailbox','RemoteEquipmentMailbox')]
        [string]$TargetRecipientTypeDetails
        ,
        [parameter(Mandatory = $true)]
        [string]$TargetDeliveryDomain
        )
        begin {
            switch ($PSCmdlet.ParameterSetName) {
                'Individual' {
                    $recordcount = $Identity.Count
                }#Individual
                'InputList' {
                    if ($TargetLookupAttribute -in ($InputList | get-member -MemberType Properties | select-object -ExpandProperty Name)) {
                        $Identity = @($InputList | Select-Object -ExpandProperty $TargetLookupAttribute)
                        $RecordCount = $Identity.count
                    }#if
                    else {
                        Write-Log -Message "FAILED: InputList does not contain the Target Lookup Attribute $TargetLookupAttribute." -Verbose -ErrorLog
                        throw("FAILED: InputList does not contain the Target Lookup Attribute $TargetLookupAttribute.")
                    }#else
                }#InputList
            }#switch
            $Global:Exceptions = @()
            $Global:ProcessedUsers = @()
            $msExchRecipientTypeDetails = switch ($TargetRecipientTypeDetails) {'UserMailbox' {1} 'LinkedMailbox' {2} 'LegacyMailbox'{8} 'MailUser' {128} 'RemoteUserMailbox' {2147483648} 'RemoteEquipmentMailbox' {17179869184} 'RemoteSharedMailbox' {34359738368} 'RemoteRoomMailbox' {8589934592}}
        }#begin
        process{
            $cr = 0
            foreach ($ID in $Identity) {
                try {
                    $cr++
                    $writeProgressParams = @{
                    Activity = "Update-PostMailboxMigrationUser: TargetForest $TargetAD"
                    CurrentOperation = "Processing Record $cr of $recordcount : $ID"
                    Status = "Lookup User with $ID by $TargetLookupAttribute in Target AD"
                    PercentComplete = $cr/$RecordCount*100
                    }
                    Write-Progress @writeProgressParams
                ################################################################################################################################################################
                #lookup users in source Active Directory Environments
                ################################################################################################################################################################
                #Lookup Target AD User
                ################################################################################################################################################################
                    try {
                        Write-Log -Message "Attempting: Find AD User $ID in Target AD Forest $TargetAD" -Verbose
                        $TADU = @(Find-Aduser -Identity $ID -IdentityType $TargetLookupAttribute -ADInstance $TargetAD -ErrorAction Stop)
                        Write-Log -Message "Succeeded: Find AD User $ID in Target AD Forest $TargetAD" -Verbose
                    }#try
                    catch {
                        Write-Log -Message "FAILED: Find AD User $ID in Target AD Forest $TargetAD" -Verbose -ErrorLog
                        Write-Log -Message $_.tostring() -ErrorLog
                        $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'TargetADUserNotFound'}}
                        Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                        throw("User Object for value $ID in Attribute $TargetLookupAttribute in Target AD $TargetAD not found.")
                    }#catch
                    if ($TADU.count -gt 1) {#check for ambiguous results
                        Write-Log -Message "FAILED: Find AD User $ID in Target AD Forest $TargetAD returned multiple objects/ambiguous results." -Verbose -ErrorLog
                        $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'TargetADUserAmbiguous'}}
                        Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                        throw("User Object for value $ID in Attribute $TargetLookupAttribute in Target AD $TargetAD was ambiguous.")
                    }#if
                    else {
                        $TADU = $TADU[0]
                        $TADUGUID = $TADU.objectguid
                        Write-Log -Message "NOTE: Target AD User in $TargetAD Identified with ObjectGUID: $TADUGUID" -Verbose
                    }
                    ################################################################################################################################################################
                    #Lookup Matching Source AD User
                    ################################################################################################################################################################
                    $writeProgressParams.status = "Lookup User with $($TADU.$SourceLookupAttribute) by $SourceLookupAttribute in Source AD"
                    Write-Progress @writeProgressParams
                    $SADU = @()
                    foreach ($ad in $SourceAD) {
                        try {
                            Write-Log -message "Attempting: Find Matching User for $ID in Source AD $ad by Lookup Attribute $SourceLookupAttribute" -Verbose
                            $SADU += Find-Aduser -Identity $($TADU.$SourceLookupAttribute) -IdentityType $SourceLookupAttribute -ADInstance $ad -ErrorAction Stop
                            Write-Log -message "Succeeded: Find Matching User for $ID in Source AD $ad by Lookup Attribute $SourceLookupAttribute" -Verbose
                        }#try
                        catch {
                            Write-Log -message "FAILED: Find Matching User for $ID in Source AD $ad by Lookup Attribute $SourceLookupAttribute" -Verbose -ErrorLog
                            Write-Log -Message $_.tostring() -ErrorLog
                        }
                    }#foreach
                    #check for no results or ambiguous results
                    switch ($SADU.count) {
                        1 {
                            Write-Log -message "Succeeded: Found exactly 1 Matching User for $ID in $($SourceAD -join ' & ') by Lookup Attribute $SourceLookupAttribute" -Verbose
                            $SADU = $SADU[0]
                            $SADUGUID = $SADU.objectguid
                            Write-Log -Message "NOTE: Source AD User Identified in with ObjectGUID: $SADUGUID" -Verbose
                        }#1
                        0 {
                            Write-Log -message "FAILED: Found 0 Matching User for $ID in Source AD $($SourceAD -join ' & ') by Lookup Attribute $SourceLookupAttribute" -Verbose
                            $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'SourceADUserNotFound'}}
                            Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                            throw("User Object for value $ID in Attribute $SourceLookupAttribute in Source AD $($SourceAD -join ' & ') not found.")
                        }#0
                        Default {
                            Write-Log -message "FAILED: Found multiple ambiguous matching User for $ID in Source AD $($SourceAD -join ' & ') by Lookup Attribute $SourceLookupAttribute" -Verbose
                            $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'SourceADUserAmbiguous'}}
                            Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                            throw("User Object for value $ID in Attribute $SourceLookupAttribute in Source AD $($SourceAD -join ' & ') was ambiguous.")
                        }#Default
                    }#switch $SADU.count
                    ################################################################################################################################################################
                    #Calculate Address Changes
                    ################################################################################################################################################################
                    $writeProgressParams.status = "Calculate Proxy Address and Target Address Changes"
                    Write-Progress @writeProgressParams
                    try {
                        Write-Log -Message "Attempting: Find Current proxy $TargetDeliveryDomain SMTP Address for Target AD User $TADUGUID" -Verbose
                        $TargetDeliveryDomainAddress = ($TADU.proxyaddresses | Where-Object {$_ -like "smtp:*@$TargetDeliveryDomain"} | Select-Object -First 1).split(':')[1]
                        Write-Log -Message "Succeeded: Find Current proxy $TargetDeliveryDomain SMTP Address for Target AD User $TADUGUID : $TargetDeliveryDomainAddress" -Verbose
                    }#try
                    catch {
                        Write-Log -Message "FAILED: Find Current proxy $TargetDeliveryDomain SMTP Address for Target AD User $TADUGUID" -Verbose -ErrorLog
                        Write-Log -Message $_.tostring() -ErrorLog
                        Write-Log -Message "NOTE: $TargetDeliveryDomain SMTP Proxy Address for Target AD User $TADUGUID will be added." -Verbose -ErrorLog
                        $AddTargetDeliveryProxyAddress = $true
                    }#catch
                    #setup for get-desiredproxyaddresses function to calculate updated addresses
                    $GetDesiredProxyAddressesParams = @{
                        CurrentProxyAddresses=$TADU.proxyAddresses
                        LegacyExchangeDNs=@($SADU.legacyExchangeDN)
                        Recipients = $SADU
                        DesiredOrCurrentAlias = $TADU.mailNickName
                    }
                    if ($AddTargetDeliveryProxyAddress) {$GetDesiredProxyAddressesParams.VerifyAddTargetAddress = $true}
                    $DesiredProxyAddresses = Get-DesiredProxyAddresses @GetDesiredProxyAddressesParams
                    if ($AddTargetDeliveryProxyAddress) {$TargetDeliveryDomainAddress = ($DesiredProxyAddresses | Where-Object {$_ -like "smtp:*@$TargetDeliveryDomain"} | Select-Object -First 1).split(':')[1]}
                    #preparation activities complete, time to write changes to Target AD
                    Write-Log -message "Using AD Cmdlets to set attributes for $TADUGUID in $TargetAD" -Verbose
                    $writeProgressParams.status = "Updating Attributes for $TADUGUID in $TargetAD using AD Cmdlets"
                    Write-Progress @writeProgressParams
                    #ClearTargetAttributes
                    $setaduserparams1 = @{
                        Identity=$TADUGUID
                        clear='proxyaddresses','targetaddress','msExchRecipientDisplayType','msExchRecipientTypeDetails','msExchUsageLocation'
                        Server=$TADU.CanonicalName.split('/')[0] #get's the Target AD Users Domain to use as the server value with set-aduser
                        ErrorAction = 'Stop'
                    }#setaduserparams1
                    try {
                        Write-Log -message "Attempting: Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD" -Verbose
                        set-aduser @setaduserparams1
                        Write-Log -message "Succeeded: Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD" -Verbose
                    }#try
                    catch {
                        Write-Log -message "FAILED: Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD" -Verbose -ErrorLog
                        Write-Log -Message $_.tostring() -ErrorLog
                        $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'FailedToClearTargetAttributes'}}
                        Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                        throw("Failed to clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD")
                    }#catch
                    #SetNewValuesOnTargetAttributes
                    $setaduserparams2 = @{
                        identity=$TADUGUID
                        add=@{
                            targetaddress = "SMTP:$TargetDeliveryDomainAddress"
                            proxyaddresses = [string[]]$DesiredProxyAddresses
                            msExchRecipientDisplayType = -2147483642
                            msExchRecipientTypeDetails = 2147483648
                            msExchRemoteRecipientType = 4
                        }
                        Server=$TADU.CanonicalName.split('/')[0]
                        ErrorAction = 'Stop'
                    }#setaduserparams2
                    if ($TADU.c) {$setaduserparams1.msExchangeUsageLocation = $TADU.c}
                    try {
                        Write-Log -message "Attempting: SET target attributes $($setaduserparams2.'Add'.keys -join ';') for $TADUGUID in $TargetAD" -Verbose
                        set-aduser @setaduserparams2
                        Write-Log -message "Succeeded: SET target attributes $($setaduserparams2.'Add'.keys -join ';') for $TADUGUID in $TargetAD" -Verbose
                    }#try
                    catch {
                        Write-Log -message "FAILED: SET target attributes $($setaduserparams2.'Add'.keys -join ';')  for $ID in $TargetAD" -Verbose -ErrorLog
                        Write-Log -Message $_.tostring() -ErrorLog
                        $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'FailedToSetTargetAttributes'}}
                        Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                        throw("Failed to set target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD")
                    }#catch
                    #have exchange clean up and set version/legacyexchangedn
                    #wait until Exchange "sees" the new attributes in the Global Catalog
                    $writeProgressParams.status = "Enabling ADUser $TADUGUID in $TargetAD"
                    Write-Progress @writeProgressParams
                    $EnableADAccountParams = @{
                        identity=$TADUGUID
                        Server=$TADU.CanonicalName.split('/')[0]
                        ErrorAction = 'Stop'
                    }#EnableADAccountParams
                    try {
                        Write-Log -message "Attempting: Enable-ADAccount $TADUGUID in $TargetAD" -Verbose
                        Enable-ADAccount @EnableADAccountParams
                        Write-Log -message "Succeeded: Enable-ADAccount $TADUGUID in $TargetAD" -Verbose
                    }#try
                    catch {
                        Write-Log -message "FAILED: Enable-ADAccount $TADUGUID in $TargetAD" -Verbose -ErrorLog
                        Write-Log -Message $_.tostring() -ErrorLog
                        $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'FailedToEnableAccount'}}
                        Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                        throw("Failed to set target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD")
                    }#catch

                    $writeProgressParams.status = "Updating Recipient $TADUGUID in $TargetExchangeOrg"
                    Write-Progress @writeProgressParams
                    do {
                        $count++
                        start-sleep -Seconds 3
                        Connect-Exchange -ExchangeOrganization $TargetExchangeOrg
                        $Recipient = @(Invoke-ExchangeCommand -cmdlet get-recipient -ExchangeOrganization $TargetExchangeOrg -string "-Identity $TADUGUID -ErrorAction SilentlyContinue")
                    }
                    until ($Recipient.count -ge 1 -or $count -ge 15)
                    #now that we found the object as a recipient, go ahead and run Update-Recipient against the object
                    try {
                        Write-Log -message "Attempting: Update Recipient $DesiredUPNAndPrimarySMTPAddress in $TargetExchangeOrg" -Verbose
                        $Global:ErrorActionPreference = 'Stop'
                        Connect-Exchange -ExchangeOrganization $TargetExchangeOrg
                        Invoke-ExchangeCommand -cmdlet Update-Recipient -ExchangeOrganization $TargetExchangeOrg -string "-Identity $TADUGUID -ErrorAction Stop"
                        $Global:ErrorActionPreference = 'Continue'
                        Write-Log -message "Succeeded: Update Recipient $DesiredUPNAndPrimarySMTPAddress in $TargetExchangeOrg" -Verbose
                    }
                    catch {
                        $Global:ErrorActionPreference = 'Continue'
                        Write-Log -message "FAILED: Update Recipient $DesiredUPNAndPrimarySMTPAddress in $TargetExchangeOrg" -Verbose -ErrorLog
                        Write-Log -message $_.tostring() -ErrorLog
                        $Global:Exceptions += $DesiredUPNAndPrimarySMTPAddress | Select-Object *,@{n='Exception';e={'FailedToUpdateRecipient'}}
                        Export-Data -DataToExportTitle TargetForestExceptionsUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                        throw("Failed to Update Recipient for $TADUGUID in $TargetExchangeOrg")
                    }
                    $ProcessedUser = $TADU | Select-Object -Property SAMAccountName,DistinguishedName,@{n='UserPrincipalname';e={$DesiredUPNAndPrimarySMTPAddress}},@{n='ObjectGUID';e={$TADUGUID}}
                    $Global:ProcessedUsers += $ProcessedUser
                    Write-Log -Message "NOTE: Processing for $DesiredUPNAndPrimarySMTPAddress with GUID $TADUGUID in $TargetAD and $TargetExchangeOrg has completed successfully." -Verbose
                Export-Data -DataToExportTitle PostMailboxMigrationProcessedUsers -DataToExport $ProcessedUser -DataType csv -Append
            }#try
            catch {
                $_
            }
            }#foreach
        }#process
        end{
        if ($Global:ProcessedUsers.count -ge 1) {
            Write-Log -Message "Successfully Processed $($Global:ProcessedUsers.count) Users." -Verbose
        }
        if ($Global:Exceptions.count -ge 1) {
            Write-Log -Message "Processed $($Global:Exceptions.count) Users with Exceptions." -Verbose
        }
        }#end
    }
#end function Update-PostMigraitonMailboxUser
function Add-MSOLLicenseToUser
    {
        [cmdletbinding()]
        param(
            [parameter(Mandatory=$true,ParameterSetName = "Migration Wave Source Data")]
            [string]$Wave
            ,
            [parameter(ParameterSetName = "Migration Wave Source Data")]
            [ValidateSet('Full','Sub','All')]
            [string]$WaveType='Sub'
            ,
            [parameter(Mandatory=$true,ParameterSetName = "Custom Source Data")]
            $CustomSource
            ,
            [parameter(Mandatory=$true,ParameterSetName = "Single User",ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
            [string]$UserPrincipalName
            ,
            [switch]$AssignUsageLocation
            ,
            [string]$StaticUsagelocation
            ,
            [switch]$SelectServices
            ,
            [validateset('Exchange','Lync','SharePoint','OfficeWebApps','OfficeProPlus','LyncEV','AzureRMS')]
            [string[]]$DisabledServices
            ,
            [parameter(ParameterSetName = "Migration Wave Source Data")]
            [parameter(ParameterSetName = "Custom Source Data")]
            [switch]$DisabledServicesFromSourceColumn
            ,
            [parameter(Mandatory=$true,ParameterSetName = "Single User")]
            [validateset('E4','E3','E2','E1','K1')]
            [string]$LicenseTypeDesired
        )
        #initial setup
        #Build License Strings for Available License Plans
            [string]$E4AddLicenses = ($Global:TenantSubDomain + ':' + $Global:E4_SkuPartNumber)
            [string]$E3AddLicenses = ($Global:TenantSubDomain + ':' + $Global:E3_SkuPartNumber)
            [string]$E2AddLicenses = ($Global:TenantSubDomain + ':' + $Global:E2_SkuPartNumber)
            [string]$E1AddLicenses = ($Global:TenantSubDomain + ':' + $Global:E1_SkuPartNumber)
            [string]$K1AddLicenses = ($Global:TenantSubDomain + ':' + $Global:K1_SkuPartNumber)

        $InterchangeableLicenses = @($E4AddLicenses,$E3AddLicenses,$E2AddLicenses,$E1AddLicenses,$K1AddLicenses)

        #input and output files
        $stamp = Get-TimeStamp
        [string]$LogPath = $trackingfolder + $stamp + '-Apply-MSOLLicense-' + "$wave$BaseLogFilePath.log"
        [string]$ErrorLogPath = $trackingfolder + $stamp + '-ERRORS_Apply-MSOLLicense-' + "$wave$BaseLogFilePath.log"

        switch ($PSCmdlet.ParameterSetName) {
            'Migration Wave Source Data' {
                switch ($wavetype) {
                    'Full' {$WaveSourceData = @($SourceData | Where-Object {$_.wave -like "$wave*"})}
                    'Sub' {$WaveSourceData = @($SourceData | Where-Object {$_.wave -eq $wave})}
                    'All' {$WaveSourceData = $sourceData}
                }
            }
            'Custom Source Data' {
                #Validate Custom Source
                $CustomSourceColumns = $CustomSource | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
                Write-Log -verbose -message "Custom Source Columns are $CustomSourceColumns." -logpath $logpath
                $RequiredColumns = @('UserPrincipalName','LicenseTypeDesired')
                Write-Log -verbose -message "Required Columns are $RequiredColumns." -logpath $logpath
                $proceed = $true
                foreach ($reqcol in $RequiredColumns) {
                    if ($reqcol -notin $CustomSourceColumns) {
                        $Proceed = $false
                        Write-Log -errorlog -verbose -message "Required Column $reqcol is not found in the Custom Source data provided.  Processing cannot proceed." -logpath $logpath -errorlogpath $errorlogpath
                    }
                }
                if ($Proceed) {
                    $UsersToLicense = $CustomSource
                    Write-Log -verbose -message "Custom Source Data Columns Validated." -logpath $logpath
                }
                else {
                    $UsersToLicense = $null
                    Write-Log -errorlog -verbose -message "ERROR: Custom Source Data Colums failed validation.  Processing cannot proceed." -logpath $logpath -errorlogpath $errorlogpath
                }
            }
            'Single User' {
                    $UsersToLicense = @('' | Select-Object -Property @{n='UserPrincipalName';e={$UserPrincipalName}},@{n='DisableServices';e={$DisabledServices}},@{n='LicenseTypeDesired';e={[string]$LicenseTypeDesired}})
            }

        }
        #record record count to be processed
        $RecordCount = $UsersToLicense.count
        $b=0
        $WriteProgressParams = @{}
        $writeProgressParams.Activity = 'Processing User License and License Option Assignments'

        #main processing
        foreach ($Record in $UsersToLicense) {
            $b++
            #Modify the following lines if the organization has discrepancies between Azure UPN and AD UPN
            $CurrentADUPN = $Record.UserPrincipalName
            $CurrentAzureUPN = $Record.UserPrincipalName
            $WriteProgressParams.PercentComplete = ($b/$RecordCount*100)
            $writeProgressParams.Status = "Processing Record $b of $($RecordCount): $CurrentADUPN."
            $WriteProgressParams.CurrentOperation = "Reading Desired Licensing and Licensing Options."
            Write-Progress @WriteProgressParams
            #calculate any services to disable
            if ($SelectServices) {
                switch ($DisabledServicesFromSourceColumn) {
                    $true {
                        $DisableServices = if ([string]::IsNullOrWhiteSpace($record.DisableServices)) {$null} else {@($Record.DisableServices.split(';'))}
                    }
                    $false {
                        $DisableServices = $disabledServices
                    }
                }
            }
            #calculate license and license options to apply
            $msollicenseoptionsparams = @{}
            $msollicenseoptionsparams.ErrorAction = 'Stop'
            switch ($Record.LicensetypeDesired) {
                'E4' {
                    $DesiredLicense = $E4AddLicenses
                    $msollicenseoptionsparams.AccountSkuID = $DesiredLicense
                    if ($SelectServices) {
                        $DisabledPlans = @()
                        foreach ($service in $DisableServices) {
                            switch ($service) {
                                'Exchange' {$DisabledPlans += 'EXCHANGE_S_ENTERPRISE'}
                                'Lync' {$DisabledPlans += 'MCOSTANDARD'}
                                'SharePoint' {$DisabledPlans += 'SHAREPOINTENTERPRISE'}
                                'OfficeWebApps' {$DisabledPlans += 'SHAREPOINTWAC'}
                                'OfficeProPlus' {$DisabledPlans += 'OFFICESUBSCRIPTION'}
                                'LyncEV' {$DisabledPlans += 'MCOVOICECONF'}
                                'AzureRMS' {$DisabledPlans += 'RMS_S_ENTERPRISE'}
                            }
                        }
                        if ($DisabledPlans.Count -gt 0) {
                            Write-Log -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                            $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                        }
                    }
                    else {$msollicenseoptionsparams.DisabledPlans = $Null}
                    Write-Log -Message "Desired E4 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
                    #Create License Options Object
                    $LicenseOptions = New-MsolLicenseOptions @msollicenseoptionsparams
                    $Proceed = $true
                }
                'E3' {
                    $DesiredLicense = $E3AddLicenses
                    $msollicenseoptionsparams.AccountSkuID = $DesiredLicense
                    if ($SelectServices) {
                        $DisabledPlans = @()
                        foreach ($service in $DisabledServices) {
                            switch ($service) {
                                'Exchange' {$DisabledPlans += 'EXCHANGE_S_ENTERPRISE'}
                                'Lync' {$DisabledPlans += 'MCOSTANDARD'}
                                'SharePoint' {$DisabledPlans += 'SHAREPOINTENTERPRISE'}
                                'OfficeWebApps' {$DisabledPlans += 'SHAREPOINTWAC'}
                                'OfficeProPlus' {$DisabledPlans += 'OFFICESUBSCRIPTION'}
                                'AzureRMS' {$DisabledPlans += 'RMS_S_ENTERPRISE'}
                            }
                        }
                        if ($DisabledPlans.Count -gt 0) {
                            Write-Log -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                            $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                        }
                    }
                    else {$msollicenseoptionsparams.DisabledPlans = $Null}
                    Write-Log -Message "Desired E3 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
                    #Create License Options Object
                    $LicenseOptions = New-MsolLicenseOptions @msollicenseoptionsparams
                    $Proceed = $true
                }
                'E2' {
                    $DesiredLicense = $E2AddLicenses
                    $msollicenseoptionsparams.AccountSkuID = $DesiredLicense
                    if ($SelectServices) {
                        $DisabledPlans = @()
                        foreach ($service in $DisabledServices) {
                            switch ($service) {
                                'Exchange' {$DisabledPlans += 'EXCHANGE_S_STANDARD'}
                                'Lync' {$DisabledPlans += 'MCOSTANDARD'}
                                'SharePoint' {$DisabledPlans += 'SHAREPOINTSTANDARD'}
                                'OfficeWebApps' {$DisabledPlans += 'SHAREPOINTWAC'}
                            }
                        }
                        if ($DisabledPlans.Count -gt 0) {
                            Write-Log -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                            $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                        }
                    }
                    else {$msollicenseoptionsparams.DisabledPlans = $Null}
                    Write-Log -Message "Desired E2 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
                    #Create License Options Object
                    $LicenseOptions = New-MsolLicenseOptions @msollicenseoptionsparams

                    $Proceed = $true
                }
                'E1' {
                    $DesiredLicense = $E1AddLicenses
                    $msollicenseoptionsparams.AccountSkuID = $DesiredLicense
                    if ($SelectServices) {
                        $DisabledPlans = @()
                        foreach ($service in $DisabledServices) {
                            switch ($service) {
                            'Exchange' {$DisabledPlans += 'EXCHANGE_S_STANDARD'}
                            'Lync' {$DisabledPlans += 'MCOSTANDARD'}
                            'SharePoint' {$DisabledPlans += 'SHAREPOINTSTANDARD'}
                            }
                        }
                        if ($DisabledPlans.Count -gt 0) {
                            Write-Log -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                            $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                        }
                    }
                    else {$msollicenseoptionsparams.DisabledPlans = $Null}
                    Write-Log -Message "Desired E1 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
                    #Create License Options Object
                    $LicenseOptions = New-MsolLicenseOptions @msollicenseoptionsparams

                    $Proceed = $true
                }
                'K1' {
                    $DesiredLicense = $K1AddLicenses
                    $msollicenseoptionsparams.AccountSkuID = $DesiredLicense
                    if ($SelectServices) {
                        $DisabledPlans = @()
                        foreach ($service in $DisabledServices) {
                            switch ($service) {
                            'Exchange' {$DisabledPlans += 'EXCHANGE_S_DESKLESS'}
                            'SharePoint' {$DisabledPlans += 'SHAREPOINTDESKLESS'}
                            }
                        }
                        if ($DisabledPlans.Count -gt 0) {
                            Write-Log -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                            $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                        }
                    }
                    else {$msollicenseoptionsparams.DisabledPlans = $Null}
                    Write-Log -Message "Desired K1 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
                    #Create License Options Object
                    $LicenseOptions = New-MsolLicenseOptions @msollicenseoptionsparams
                    $Proceed = $true
                }
                Default {
                    $Proceed = $false
                    Write-Log -Message "No License Desired (non E4,E3,E2,E1,K1) Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
                }
            }
            #Lookup MSOL User Object
            if ($proceed) {
                $WriteProgressParams.CurrentOperation = "Looking up MSOL User Object."
                Write-Progress @WriteProgressParams
                Try {
                    Write-Log -Message "Looking up MSOL User Object $CurrentAzureUPN for AD User Object $CurrentADUPN" -Verbose -LogPath $LogPath
                    $CurrentMSOLUser = Get-MsolUser -UserPrincipalName $CurrentAzureUPN -ErrorAction Stop
                    Write-Log -Message "Found MSOL User for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                    $Proceed = $true
                }
                Catch {
                    $Proceed = $false
                    Write-Log -Message "ERROR: MSOL User for $CurrentAzureUPN not found." -Verbose -LogPath $LogPath
                    Write-Log -Message "ERROR: MSOL User for $CurrentAzureUPN not found." -LogPath $ErrorLogPath
                    Write-Log -Message $_.tostring() -LogPath $ErrorLogPath
                }
            }

            #Check Usage Location Assignment of User Object
            if ($Proceed) {

                switch -Wildcard ($CurrentMSOLUser.usagelocation) {
                    $null {
                        Try {
                            if ($AssignUsageLocation -and ($StaticUsagelocation -ne $null)) {Set-MsolUser -UserPrincipalName $CurrentAzureUPN -UsageLocation $StaticUsagelocation -ErrorAction Stop}
                            elseif ($AssignUsageLocation -and ($StaticUsagelocation -eq $null)) {Set-MsolUserUsageLocation -UserPrincipalName $CurrentAzureUPN -ErrorAction Stop}
                        }
                        Catch {
                            $Proceed = $false
                        }
                    }
                    '' {
                        Try {
                            if ($AssignUsageLocation -and ($StaticUsagelocation -ne $null)) {Set-MsolUser -UserPrincipalName $CurrentAzureUPN -UsageLocation $StaticUsagelocation -ErrorAction Stop}
                            elseif ($AssignUsageLocation -and ($StaticUsagelocation -eq $null)) {Set-MsolUserUsageLocation -UserPrincipalName $CurrentAzureUPN -ErrorAction Stop}
                        }
                        Catch {
                            $Proceed = $false
                        }
                    }
                    Default {
                        Write-Log -Message "Usage Location for MSOL User $CurrentAzureUPN is set to $($CurrentMSOLUser.UsageLocation)." -LogPath $LogPath -Verbose
                        $Proceed = $true
                    }

                }
            }
            #Determine License Operation Required
            if ($Proceed) {
            $WriteProgressParams.currentoperation = "Determining License Operation Required"
            Write-Progress @WriteProgressParams
            #Correct License Already Applied?
            if ($CurrentMSOLUser.IsLicensed) {$LicenseAssigned = $true} else {$LicenseAssigned = $false}
            Write-Log -Message "$CurrentADUPN license assignment status = $LicenseAssigned" -Verbose -LogPath $LogPath
            if ($CurrentMSOLUser.Licenses.AccountSkuId -contains $DesiredLicense) {$CorrectLicenseType = $True}
            else {
                $CorrectLicenseType = $false
                $LicenseToReplace = $CurrentMSOLUser.Licenses.AccountSkuID | where-object {$_ -in ($InterchangeableLicenses)}
            }
            Write-Log -Message "$CurrentADUPN correct license applied status = $CorrectLicenseType" -Verbose -LogPath $LogPath

            #Correct License Options Already Applied?
            if (-not $CorrectLicenseType) {$correctLicenseOptions = $false}
            else {
                #get current user's disabled plans
                $currentUserDisabledPlans = @($currentMSOLUser.Licenses.servicestatus | ? ProvisioningStatus -eq 'Disabled' | % {$_.servicePlan.ServiceName})
                #compare intended disabled plans to current user's disabled plans:
                $unintendedDisabledPlans = @($currentUserDisabledPlans | ?  {$_ -notin $DisabledPlans})
                $unintendedEnabledPlans = @($DisabledPlans | ? {$_ -notin $currentUserDisabledPlans})
                if ($unintendedDisabledPlans.count -gt 0 -or $unintendedEnabledPlans.Count -gt 0) {$correctLicenseOptions = $false}
                else {$correctLicenseOptions = $true}
            }
            Write-Log -Message "$CurrentADUPN correct license options applied status = $correctLicenseOptions" -Verbose -LogPath $LogPath
            #Set Operation To Process on User
            $MSOLUserLicenseParams = @{}
            $MSOLUserLicenseParams.ErrorAction = 'Stop'
            $MSOLUserLicenseParams.UserPrincipalName = $CurrentAzureUPN
            if (-not $LicenseAssigned) {$LicenseOperation = 'Assign'}
            if ($licenseAssigned -and $CorrectLicenseType -and $correctLicenseOptions) {$LicenseOperation = 'None'}
            if ($licenseAssigned -and $CorrectLicenseType -and -not $correctLicenseOptions) {$LicenseOperation = 'Options'}
            if ($LicenseAssigned -and -not $CorrectLicenseType) {$LicenseOperation = 'Replace'}


            Write-Log -Message "$CurrentADUPN license operation selected = $LicenseOperation" -Verbose -LogPath $LogPath

            #Process License Operation
            switch ($LicenseOperation) {
                'None' {Write-Log -Message "$CurrentAzureUPN is already correctly licensed." -Verbose -LogPath $LogPath
                }
                'Assign'{
                    Try {
                        $MSOLUserLicenseParams.AddLicenses = $DesiredLicense
                        $MSOLUserLicenseParams.LicenseOptions = $LicenseOptions
                        Write-Log -Message "Setting User License for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                        Set-MsolUserLicense @MSOLUserLicenseParams
                        Write-Log -Message "Success: Assigned User License for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                    }
                    Catch {
                        Write-Log -Message "ERROR: License could not be assigned for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                        Write-Log -Message "ERROR: License could not be assigned for $CurrentAzureUPN" -LogPath $ErrorLogPath
                        Write-Log -Message $_.tostring() -Verbose -errorlogpath $ErrorLogPath
                    }
                }
                'Replace' {
                    Try {
                        $MSOLUserLicenseParams.AddLicenses = $DesiredLicense
                        $MSOLUserLicenseParams.LicenseOptions = $LicenseOptions
                        $MSOLUserLicenseParams.RemoveLicenses = $LicenseToReplace
                        Write-Log -Message "Replacing User License for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                        Set-MsolUserLicense @MSOLUserLicenseParams
                        Write-Log -Message "Success: Replaced User License for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                    }
                    Catch {
                        Write-Log -Message "ERROR: License could not be replaced for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                        Write-Log -Message "ERROR: License could not be replaced for $CurrentAzureUPN" -LogPath $ErrorLogPath
                        Write-Log -Message $_.tostring() -Verbose -LogPath $ErrorLogPath
                    }
                }
                'Options' {
                    Try {
                        #$MSOLUserLicenseParams.AddLicenses = $DesiredLicense
                        $MSOLUserLicenseParams.LicenseOptions = $LicenseOptions
                        Write-Log -Message "Setting User License Options for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                        Set-MsolUserLicense @MSOLUserLicenseParams
                        Write-Log -Message "Success: Set User License Options for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                    }
                    Catch {
                        Write-Log -Message "ERROR: License options could not be set for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                        Write-Log -Message "ERROR: License options could not be set for $CurrentAzureUPN" -LogPath $ErrorLogPath
                        Write-Log -Message $_.tostring() -Verbose -LogPath $ErrorLogPath
                    }
                }
            }

        }
        }
    }
#end function Add-MSOLLicenseToUser
function Add-LicenseToMSOLUser
    {
        [cmdletbinding()]
        param(
        [parameter(Mandatory)]
        [ValidatePattern(".:.")]
        [string]$AccountSKUID
        ,
        [parameter()]
        [string[]]$DisabledPlans
        ,
        [Parameter(Mandatory,ParameterSetName='UserPrincipalName',ValueFromPipelineByPropertyName)]
        [string[]]$UserPrincipalName
        ,
        [Parameter(Mandatory,ParameterSetName='ObjectID',ValueFromPipeline)]
        [guid[]]$ObjectID
        ,
        [Parameter(Mandatory)]
        [bool]$CheckForExchangeOnlineRecipient = $true
        ,
        [parameter(Mandatory)]
        [string]$UsageLocation
        ,
        [parameter()]
        [string]$ExchangeOrganization
        )
        <#
        DynamicParam {
                $NewDynamicParameterParams=@{
                    Name = 'ExchangeOrganization'
                    ValidateSet = @(Get-OneShellVariableValue -name 'CurrentOrgAdminProfileSystems' | Where-Object SystemType -eq 'ExchangeOrganizations' | Select-Object -ExpandProperty Name)
                    Alias = @('Org','ExchangeOrg')
                    Position = 2
                    ParameterSetName = 'Organization'
                }
                New-DynamicParameter @NewDynamicParameterParams -Mandatory $false
            }#DynamicParam
        #>
        begin
        {
            if ($PSBoundParameters['CheckForExchangeOnlineRecipient'] -eq $true -and $PSBoundParameters.ContainsKey('ExchangeOrganization') -eq $false)
            {throw "ExchangeOrganization parameter required when -CheckForExchangeOnlineRecipient is True"}
            $newLicenseOptionsParams = @{
                AccountSkuID = $AccountSKUID
                ErrorAction = 'Stop'
            }
            if ($PSBoundParameters.ContainsKey('DisabledPlans'))
            {
                $newLicenseOptionsParams.DisabledPlans = $DisabledPlans
            }
            try
            {
                $message = "Build License Options Object"
                Write-Log -Message $message -EntryType Attempting
                $LicenseOptions = New-MsolLicenseOptions @newLicenseOptionsParams
                Write-Log -Message $message -EntryType Succeeded
            }
            catch
            {
                $myerror = $_
                Write-Log -Message $message -EntryType Failed -Verbose
                Write-Log -Message $_.tostring() -ErrorLog -Verbose
                throw("Failed:$message")
            }
            $IdentityParameter = $PSCmdlet.ParameterSetName
        }
        process
        {
            switch ($PSCmdlet.ParameterSetName)
            {
                'UserPrincipalName'
                {
                    $Identities = @($UserPrincipalName)
                    $GetMSOLUserParams = @{
                        UserPrincipalName = ''
                    }
                }

                'ObjectID'
                {
                    $Identities = @($ObjectID)
                    $GetMSOLUserParams = @{
                        ObjectID = ''
                    }
                }
            }
            $GetMSOLUserParams.ErrorAction = 'Stop'
            :nextID foreach ($ID in $Identities)
            {
                try
                {
                    $GetMSOLUserParams.$IdentityParameter = $ID.ToString()
                    $message = "Get MSOL User Object for $ID"
                    Write-Log -Message $message -EntryType Attempting
                    $MSOLUser = Get-MsolUser @GetMSOLUserParams
                    Write-Log -Message $message -EntryType Succeeded
                }
                catch
                {
                    $myerror = $_
                    Write-Log -Message $message -EntryType Failed -Verbose
                    Write-Log -Message $_.tostring() -ErrorLog -Verbose
                    continue nextID
                }
                if ($CheckForExchangeOnlineRecipient -eq $true)
                {
                    $message = "Lookup Exchange Online Recipient for Identity $ID"
                    $getRecipientParams = @{
                        Identity = $MSOLUser.objectID.guid
                        ErrorAction = 'Stop'
                    }
                    try
                    {
                        Write-Log -Message $message -EntryType Attempting
                        $EOLRecipient = Invoke-ExchangeCommand -cmdlet Get-Recipient -ExchangeOrganization $psboundparameters['exchangeOrganization'] -splat $getRecipientParams -ErrorAction Stop
                        Write-Log -Message $message -EntryType Succeeded
                    }
                    catch
                    {
                        $myerror = $_
                        Write-Log -Message $message -EntryType Failed -ErrorLog
                        Write-Log -Message $_.tostring() -ErrorLog -Verbose
                        continue nextID
                    }
                }
                $AssignedLicenseAccountSKUIDs = @($MSOLUser.licenses | Select-Object -ExpandProperty AccountSkuID)
                if ($AccountSKUID -notin $AssignedLicenseAccountSKUIDs)
                {
                    if ($MSOLUser.UsageLocation -eq $null)
                    {
                        $message = "UsageLocation for $ID is current NULL"
                        Write-Log -Message $message -EntryType Notification
                        $setMSOLUserParams = @{
                            ObjectID = $MSOLUser.ObjectID.guid
                            UsageLocation = $UsageLocation
                            ErrorAction = 'Stop'
                        }
                        $message = "Set UsageLocation for $ID to $UsageLocation"
                        try
                        {
                            Write-Log -Message $message -EntryType Attempting
                            Set-MsolUser @setMSOLUserParams
                            Write-Log -Message $message -EntryType Succeeded
                        }
                        catch
                        {
                            $myerror = $_
                            Write-Log -Message $message -EntryType Failed -ErrorLog
                            Write-Log -Message $_.tostring() -ErrorLog -Verbose
                            continue nextID
                        }
                    }#if usage location is null
                    $message = "Add $AccountSKUID license to MSOL User $ID"
                    $setMSOLUserLicenseParams = @{
                        ObjectID = $MSOLUser.ObjectID.guid
                        LicenseOptions = $LicenseOptions
                        AddLicenses = $AccountSKUID
                        ErrorAction = 'Stop'
                    }
                    try
                    {
                        Write-Log -Message $message -EntryType Attempting
                        Set-MsolUserLicense @setMSOLUserLicenseParams
                        Write-Log -Message $message -EntryType Succeeded
                    }
                    catch
                    {
                        $myerror = $_
                        Write-Log -Message $message -EntryType Failed -ErrorLog
                        Write-Log -Message $_.tostring() -ErrorLog -Verbose
                    }
                }
            }
        }
    }
#end function Add-LicenseToMSOLUser
function Set-UsageLocationForMSOLUser
    {
        [cmdletbinding()]
        param(
        [parameter(Mandatory)]
        [string]$UsageLocation
        ,
        [Parameter(Mandatory,ParameterSetName='UPN')]
        [string]$UserPrincipalName
        ,
        [Parameter(Mandatory,ParameterSetName='ObjectID')]
        [string]$ObjectID
        )

    }
#end function Set-UsageLocaitonForMSOLUser
function Set-ExchangeAttributesOnTargetObject
    {
        [cmdletbinding()]
        param
        (
            [Parameter(ParameterSetName='SourceLookup')]
            [switch]$SourceLookup
            ,
            [Parameter(ParameterSetName='SourceDataProvided')]
            [switch]$SourceDataProvided
            ,
            [parameter(Mandatory=$true,ParameterSetName='SourceDataProvided')]
            $SourceData
            ,
            [parameter(Mandatory=$true,ParameterSetName='SourceLookup')]
            [ValidateScript({$_ -in $(Get-PSDrive -PSProvider ActiveDirectory | Select-Object -ExpandProperty Name)})]
            [string]$SourceAD
            ,
            [parameter(Mandatory = $true,ParameterSetName='SourceLookup')]
            [validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','employeeNumber','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid','SID','msExchMasterAccountSID')]
            [string]$SourceLookupAttribute
            ,
            [parameter(Mandatory = $true,ParameterSetName='SourceLookup',ValueFromPipeline=$true)]
            [string[]]$SourceLookupValue
            ,
            <#
            [parameter(ParameterSetName='SourceLookup')]
            [validateset('UserMailbox','LinkedMailbox','LegacyMailbox','MailUser','RemoteUserMailbox')]
            [string]$SourceRecipientTypeDetails
            ,
            #>
            [parameter(Mandatory=$true)]
            [ValidateScript({$_ -in $(Get-PSDrive -PSProvider ActiveDirectory | Select-Object -ExpandProperty Name)})]
            [string]$TargetAD
            ,
            [parameter(Mandatory = $true)]
            [validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','employeeNumber','employeeID','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid','SID','msExchMasterAccountSID','GivenNameSurname')]
            [string]$TargetLookupPrimaryAttribute
            ,
            [parameter()]
            [validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','employeeNumber','employeeID','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid','SID','msExchMasterAccountSID','GivenNameSurname')]
            [string]$TargetLookupSecondaryAttribute
            ,
            [parameter(Mandatory = $true)]
            [validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','employeeNumber','employeeID','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid','SID','msExchMasterAccountSID','GivenNameSurname')]
            [string]$TargetLookupPrimaryValue
            ,
            [parameter(Mandatory = $true)]
            [validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','employeeNumber','employeeID','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid','SID','msExchMasterAccountSID','GivenNameSurname')]
            [string]$TargetLookupSecondaryValue
            ,
            [parameter(Mandatory = $true)]
            [string]$TargetDeliveryDomain = $CurrentOrgProfile.office365tenants[0].TargetDomain
            ,
            [parameter()]
            [string]$ForceTargetPrimarySMTPDomain
            ,
            [switch]$AddAdditionalSMTPProxyAddress
            ,
            [parameter()]
            [ValidateScript({$_.split('@')[0] -in @('Alias','PrimaryPrefix')})]
            [string[]]$AdditionalSMTPProxyAddressPattern
            ,
            [boolean]$DeleteContact = $false
            ,
            [boolean]$DeleteSourceObject = $false
            ,
            [boolean]$UpdateSourceObject = $false
            ,
            [boolean]$DisableEmailAddressPolicyInTarget = $true
            ,
            [boolean]$ReplaceUPN = $false
            ,
            [parameter(Mandatory=$false)]
            [string]$TargetExchangeOrganization
            ,
            [switch]$UpdateTargetRecipient
            ,
            [string[]]$TargetAttributestoClear =
                $(@(
                    'Mail'
                    'mailNickName'
                    'msExchArchiveGUID'
                    'msExchArchiveName'
                    'msExchMailboxGUID'
                    'msExchPoliciesExcluded'
                    'msExchRecipientDisplayType'
                    'msExchRecipientTypeDetails'
                    'msExchRemoteRecipientType'
                    'msExchUserCulture'
                    'msExchVersion'
                    'msExchUsageLocation'
                    'proxyaddresses'
                    'targetaddress'
                    'extensionattribute5'
                    'c'
                    'co'
                    #'countrycode'
                ))
            ,
            [int]$ADSyncDelayInSeconds = 90
            ,
            [switch]$ClearGlobalTrackingVariables
            ,
            [switch]$TestOnly
            ,
            $DestinationOU
            ,
            [string]$MoveRequestWaveBatchName
            ,
            [switch]$postCutover
            ,
            [switch]$PreserveSourceMailbox
            ,
            [switch]$AddTargetToSourceGroups
            ,
            [string[]]$DomainsToRemove
            ,
            [switch]$perUserDirSyncTest
        )#Param
        begin
        {
            #Set up the global tracking/reporting variables if needed and/or clear them if requested
            $GlobalTrackingVariables =
                @(
                    'SEATO_Exceptions'
                    ,'SEATO_ProcessedUsers'
                    ,'SEATO_FullProcessedUsers'
                    ,'SEATO_MailContactsFound'
                    ,'SEATO_MailContactDeletionFailures'
                    ,'SEATO_OLMailboxSummary'
                    ,'SEATO_OriginalTargetUsers'
                    ,'SEATO_OriginalSourceUsers'
                    ,'SEATO_MembershipAddFailures'
                )
            if ($ClearGlobalTrackingVariables)
            {
                foreach ($var in $GlobalTrackingVariables)
                {
                    Set-Variable -Scope Global -Name $var -Value @()
                }
            }
            else {
                foreach ($var in $GlobalTrackingVariables) {
                    if (-not (Test-Path -Path "variable:global:$var")) {
                        Set-Variable -Scope Global -Name $var -Value @()
                    }
                }
            }
        }#begin
        process
        {
        #Process processes all incoming Source Lookup Values or Source Data and outputs Source Object Data
        #setup for Source Operation specified
            switch ($PSCmdlet.ParameterSetName)
            {
                'SourceLookup'
                {
                    #Populate SourceData
                    #############################################################
                    #lookup user in the Source AD
                    #############################################################
                    $SourceData =
                    @(
                        $recordcount = $SourceLookupValue.Count
                        $cr = 0
                        foreach ($value in $SourceLookupValue)
                        {
                            $cr++
                            $writeProgressParams =
                                @{
                                    Activity = "Find Source Object"
                                    Status = "Processing Record $cr of $recordcount : $value"
                                    PercentComplete = $cr/$RecordCount*100
                                }#writeProgressParams
                            $writeProgressParams.currentoperation = "Find User with value $value in $SourceLookupAttribute in Source Object Forest $SourceAD"
                            Write-Progress @writeProgressParams
                            $TrialSADU = $null
                            $TrialSADU =
                            @(
                                try {
                                    Write-Log -message $writeProgressParams.currentoperation -EntryType Attempting
                                    Find-ADUser -Identity $value -IdentityType $sourceLookupAttribute -AmbiguousAllowed -ActiveDirectoryInstance $SourceAD -ErrorAction Stop
                                    Write-Log -message $writeProgressParams.currentoperation -EntryType Succeeded
                                }#try
                                catch {
                                    Write-Log -message $writeProgressParams.currentoperation -Verbose -EntryType Failed -ErrorLog
                                    Write-Log -Message $_.tostring() -ErrorLog
                                    Export-FailureRecord -Identity $ID -ExceptionCode 'SourceADUserNotFound' -FailureGroup NotProcessed -RelatedObjectIdentifier $value -RelatedObjectIdentifierType $SourceLookupAttribute
                                }
                            )#TrialSADU
                            #Determine action based on the results of the lookup attempt in the target AD
                            switch ($TrialSADU.count) {
                                1 {
                                    Write-Log -message "Succeeded: Found exactly 1 Matching User with value $value in $SourceLookupAttribute in Source Object Forest $SourceAD"
                                    #output the object into $SourceData
                                    $TrialSADU[0]
                                    Write-Log -Message "Source AD User Identified in with ObjectGUID: $($TrialSADU[0].objectguid)" -EntryType Notification
                                }#1
                                0 {
                                    Write-Log -message "FAILED: Found 0 Matching Users with value $value in $SourceLookupAttribute in Source Object Forest $SourceAD" -Verbose
                                    Export-FailureRecord -Identity $ID -ExceptionCode 'SourceADUserNotFound' -FailureGroup NotProcessed -RelatedObjectIdentifier $ID -RelatedObjectIdentifierType $SourceLookupAttribute
                                }#0
                                Default {
                                    Write-Log -message "FAILED: Found multiple ambiguous Matching Users with value $value in $SourceLookupAttribute in Source Object Forest $SourceAD" -Verbose
                                    Export-FailureRecord -Identity $ID -ExceptionCode 'SourceADUserAmbiguous' -FailureGroup NotProcessed -RelatedObjectIdentifier $ID -RelatedObjectIdentifierType $SourceLookupAttribute
                                }#Default
                            }#switch $SADU.count
                        }#foreach $value in $SourceLookupValue
                    )#SourceData Array
                }#SourceLookup
                'SourceDataProvided'
                {
                    #validate attributes in SourceData?
                }#SourceDataProvided
            }#switch $PSCmdlet.ParameterSetName
            Write-Log -Message "Completed Source Object Lookup/Validation Operations" -EntryType Notification -Verbose
        }#process
        end
        {
        #End processes the Source Object Data, finds target and related object (such as contacts) and makes determinations about processing in the Intermediate Object stage
        #End then processes the objects, adding/adjusting attribute values on the target objects and deleting source and related objects if specified
            $IntermediateObjects =
            @(
                $recordcount = $SourceData.Count
                $cr = 0
                :nextID foreach ($SADU in $SourceData)
                {
                    #region Preparing To Generate Intermediate Object
                    #region FindTADU
                    $cr++
                    $SADUGUID = $SADU.objectguid.guid
                    #value is the value to use in lookup.  attribute is the attribute to lookup against.
                    $ID = $SADU.$TargetLookupPrimaryValue
                    #check for secondary lookup attempt parameters
                    if ($PSBoundParameters.ContainsKey('TargetLookupSecondaryAttribute') -and $PSBoundParameters.ContainsKey('targetLookupSecondaryValue'))
                    {
                        $trySecondary = $true
                        $SecondaryID = $SADU.$targetLookupSecondaryValue
                    }#if
                    else
                    {
                        $trySecondary = $false
                    }#else
                    $writeProgressParams =
                    @{
                        Activity = "Preparing for User Updates"
                        CurrentOperation = "Find Target Object with $ID in $TargetLookupPrimaryAttribute in $targetAD"
                        Status = "Processing Record $cr of $recordcount : $ID"
                        PercentComplete = $cr/$RecordCount*100
                    }#writeProgressParams
                    Write-Progress @writeProgressParams
                    try
                    {
                        Write-Log -Message $writeProgressParams.CurrentOperation -EntryType Attempting
                        $TrialTADU =
                        @(
                            if (-not [string]::IsNullOrWhiteSpace($id))
                            {
                                Find-Aduser -Identity $ID -IdentityType $TargetLookupPrimaryAttribute -AD $TargetAD -ErrorAction Stop -AmbiguousAllowed
                            }
                        )
                        $TrialTADU = @($TrialTADU | Where-Object {$_.ObjectGUID -ne $SADUGUID})
                        if ($TrialTADU.Count -gt 0)
                        {
                            $TrialTADU | Add-Member -MemberType NoteProperty -Name MatchAttribute -Value $TargetLookupPrimaryAttribute
                        }
                        if ($TrialTADU.Count -lt 1 -and $trySecondary)
                        {
                            if ($TargetLookupSecondaryAttribute -eq 'GivenNameSurname')
                            {
                                $GivenName = $SADU.GivenName
                                $SurName = $SADU.Surname
                                $message = "Attempting Secondary Attribute Lookup using GivenName: $givenName Surname: $SurName"
                                $writeProgressParams.CurrentOperation = $message
                                Write-log -Message $message -EntryType Notification
                                $TrialTADU = @(Find-ADUser -GivenName $GivenName -SurName $SurName -IdentityType GivenNameSurname -AmbiguousAllowed -AD $TargetAD -ErrorAction Stop)
                                $TrialTADU = @($TrialTADU | Where-Object {$_.ObjectGUID -ne $SADUGUID})
                            }
                            else
                            {
                                $message = "Attempting Secondary Attribute Lookup using $secondaryID in $TargetLookupSecondaryAttribute"
                                $writeProgressParams.CurrentOperation = $message
                                Write-log -Message $message -EntryType Notification
                                $TrialTADU = @(Find-Aduser -Identity $SecondaryID -IdentityType $TargetLookupSecondaryAttribute -AD $TargetAD -ErrorAction Stop -AmbiguousAllowed)
                                $TrialTADU = @($TrialTADU | Where-Object {$_.ObjectGUID -ne $SADUGUID})
                            }
                            if ($TrialTADU.Count -ge 1)
                            {
                                $TrialTADU | Add-Member -MemberType NoteProperty -Name MatchAttribute -Value $TargetLookupSecondaryAttribute
                            }
                        }#if
                        Write-Log -Message $writeProgressParams.CurrentOperation -EntryType Succeeded
                    }#try
                    catch
                    {
                        Write-Log -Message $writeProgressParams.CurrentOperation -EntryType Failed -Verbose -ErrorLog
                        Write-Log -Message $_.tostring() -ErrorLog
                        Export-FailureRecord -Identity $ID -ExceptionCode 'TargetADUserNotFound' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType 'ObjectGUID'
                        continue nextID
                    }#catch
                    #Determine action based on the results of the lookup attempt in the target AD
                    #filter SADU out of TADU results
                    $TrialTADU = @($TrialTADU | Where-Object {$_.ObjectGUID -ne $SADUGUID})
                    switch ($TrialTADU.count)
                    {
                        1
                        {
                            Write-Log -message "Succeeded: Found exactly 1 Matching User" -Verbose
                            $TADU = $TrialTADU[0]
                            $TADUGUID = $TADU.objectguid.guid
                            Write-Log -Message "Target AD User Identified in $TargetAD with ObjectGUID: $TADUGUID" -Verbose -EntryType Notification
                        }#1
                        0
                        {
                            if ($SADU.enabled)
                            {
                                Write-Log -Message "Found 0 Matching Users for User $ID, but Source User Object is Enabled." -Verbose -EntryType Notification
                                $TADU = $SADU
                                $TADUGUID = $SADUGUID
                            }
                            else {
                                Write-Log -message "Found 0 Matching Users for User $ID" -Verbose -EntryType Failed
                                Export-FailureRecord -Identity $ID -ExceptionCode 'TargetADUserNotFound' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType 'ObjectGUID'
                                continue nextID
                            }
                        }#0
                        Default
                        {#check for ambiguous results
                            Write-Log -Message "Find AD User returned multiple objects/ambiguous results for User $ID." -Verbose -ErrorLog -EntryType Failed
                            Export-FailureRecord -Identity $ID -ExceptionCode 'TargetADUserAmbiguous' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType 'ObjectGUID'
                            continue nextID
                        }
                    }#switch
                    #endregion FindTADU
                    #region FindSADUExchangeDetails
                    $writeProgressParams.CurrentOperation = "Get Source Object $ID Exchange Details"
                    Write-Progress @writeProgressParams
                    #Determine Source Object Exchange Recipient Status
                    $SourceUserObjectIsExchangeRecipient = $(
                        if ($SADU.msExchRecipientTypeDetails -ne $null -or $SADU.msExchRecipientDisplayType -ne $null)
                        {
                            $true
                        }
                        else
                        {
                            $false
                        }
                    )#SourceUserObjectIsExchangeRecipient
                    if ($SourceUserObjectIsExchangeRecipient)
                    {
                        try
                        {
                            $SADUCurrentPrimarySmtpAddress = Find-PrimarySMTPAddress -ProxyAddresses $SADU.proxyaddresses -Identity $ID -ErrorAction Stop
                        }#try
                        catch
                        {
                            Write-Log -Message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity $ID -ExceptionCode 'SourceADUserPrimarySMTPNotFound' -FailureGroup NotProcessed
                            continue nextID
                        }#catch
                        $SADUUserObjectExchangeRecipientType = Get-RecipientType -msExchRecipientTypeDetails $SADU.msExchRecipientTypeDetails | Select-Object -ExpandProperty Name
                    }
                    else
                    {
                        $SADUCurrentPrimarySmtpAddress = $null
                        $SADUUserObjectExchangeRecipientType = $null
                    }
                    #endregion FindSADUExchangeDetails
                    #region FindTADUExchangeDetails
                    #Determine Target Object Exchange Recipient Status
                    $TargetUserObjectIsExchangeRecipient = $(
                        if ($TADUGUID -eq $SADUGUID)
                        {
                            $false
                        }
                        else
                        {
                            if ($TADU.msExchRecipientTypeDetails -ne $null -or $TADU.msExchRecipientDisplayType -ne $null)
                            {
                                $true
                            }
                            else
                            {
                                $false
                            }
                        }
                    )
                    if ($TargetUserObjectIsExchangeRecipient) {
                        try
                        {
                            $TADUCurrentPrimarySmtpAddress = Find-PrimarySMTPAddress -ProxyAddresses $TADU.proxyaddresses -Identity $ID -ErrorAction Stop
                        }#try
                        catch
                        {
                            Write-Log -Message $_.tostring() -ErrorLog
                            #Export-FailureRecord -Identity $ID -ExceptionCode 'TargetADUserPrimarySMTPNotFound' -FailureGroup NotProcessed
                            #continue nextID
                        }#catch
                        $TargetUserObjectExchangeRecipientType = Get-RecipientType -msExchRecipientTypeDetails $TADU.msExchRecipientTypeDetails | Select-Object -ExpandProperty Name
                    }
                    else {
                        $TADUCurrentPrimarySmtpAddress = $null
                        $TargetUserObjectExchangeRecipientType = $null
                    }
                    #endregion FindTADUExchangeDetails
                    #region FindContacts
                    #lookup mail contacts in the Target AD (using Source AD Proxy addresses, target address, and altRecipient)
                    $writeProgressParams.currentOperation = "Get any mail contacts for $ID in target AD $TargetAD"
                    Write-Progress @writeProgressParams
                    $MailContacts = @()
                    $addr = $null
                    $MailContact = $null
                    #look for contacts via proxy addresses
                    :nextAddr foreach ($addr in $SADU.proxyaddresses)
                    {
                        try
                        {
                            #Write-Log -message "Find Mail Contact for $addr in $TargetAD" -EntryType Attempting
                            $MailContact = @(Find-ADContact -Identity $addr -IdentityType ProxyAddress -AmbiguousAllowed -ActiveDirectoryInstance $TargetAD -ErrorAction Stop)
                            #Write-Log -message "No Errors: Find Mail Contact for $addr in $TargetAD" -EntryType Succeeded
                        }#try
                        catch
                        {
                            Write-Log -message "Unexpected Error: Find Mail Contact for $addr in $TargetAD" -EntryType Failed -Verbose -ErrorLog
                            Write-Log -message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity "$ID`:$addr" -ExceptionCode 'UnexpectedFailureDuringMailContactLookup' -FailureGroup ContactLookupFailure -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                            continue nextAddr
                        }#catch
                        If ($MailContact.count -ge 1)
                        {
                            Write-Log -Message "NOTE: A mail contact was found for $addr in $TargetAD." -Verbose
                            if ($MailContacts.distinguishedname -notcontains $MailContact.Distinguishedname)
                            {
                                $MailContacts += $MailContact
                                $Global:SEATO_MailContactsFound += $MailContact
                            }
                        }#if
                    }#foreach
                    #look for contacts via target address if target was not in proxy addresses
                    if ($SADU.TargetAddress -ne $null -and $SADU.proxyaddresses -notcontains $SADU.TargetAddress)
                    {
                        $addr = $SADU.targetaddress
                        try
                        {
                            #Write-Log -message "Find Mail Contact for $addr in $TargetAD" -EntryType Attempting
                            $MailContact = @(Find-ADContact -Identity $addr -IdentityType ProxyAddress -AmbiguousAllowed -ActiveDirectoryInstance $TargetAD -ErrorAction Stop)
                            #Write-Log -message "No Errors: Find Mail Contact for $addr in $TargetAD" -EntryType Succeeded
                        }#try
                        catch
                        {
                            Write-Log -message "Unexpected Error: Find Mail Contact for $addr in $TargetAD" -EntryType Failed -Verbose -ErrorLog
                            Write-Log -message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity "$ID`:$addr" -ExceptionCode 'UnexpectedFailureDuringMailContactLookup' -FailureGroup ContactLookupFailure -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                        }#catch
                        If ($MailContact.count -ge 1)
                        {
                            Write-Log -Message "NOTE: A mail contact was found for $addr in $TargetAD." -Verbose
                            if ($MailContacts.distinguishedname -notcontains $MailContact.Distinguishedname) {
                                $MailContacts += $MailContact
                                $Global:SEATO_MailContactsFound += $MailContact
                            }
                        }#if
                    }
                    #look for contacts via altRecipient if altRecipient has a value and it is not equal to the TADU DN
                    if (-not [string]::IsNullOrWhiteSpace($SADU.altRecipient) -and $SADU.altRecipent -ne $TADU.distinguishedName)
                    {
                        $addr = $SADU.altRecipient
                        try
                        {
                            #Write-Log -message "Find Mail Contact for $addr in $TargetAD" -EntryType Attempting
                            $MailContact = @(Find-ADContact -Identity $addr -IdentityType DistinguishedName -ActiveDirectoryInstance $TargetAD -ErrorAction Stop)
                            Write-Log -message "No Errors: Find Mail Contact for $addr in $TargetAD" -EntryType Succeeded
                        }#try
                        catch
                        {
                            Write-Log -message "Unexpected Error: Find Mail Contact for $addr in $TargetAD" -EntryType Failed -Verbose -ErrorLog
                            Write-Log -message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity "$ID`:$addr" -ExceptionCode 'UnexpectedFailureDuringMailContactLookup' -FailureGroup ContactLookupFailure -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                        }#catch
                        If ($MailContact.count -ge 1)
                        {
                            Write-Log -Message "NOTE: A mail contact was found for $addr in $TargetAD." -Verbose
                            if ($MailContacts.distinguishedname -notcontains $MailContact.Distinguishedname)
                            {
                                $MailContacts += $MailContact
                                $Global:SEATO_MailContactsFound += $MailContact
                            }
                        }#if
                    }
                    Write-Log -Message "A total of $($MailContacts.count) mail contacts were found for $ID in $TargetAD" -Verbose -EntryType Notification
                    #endregion FindContacts
                    #region BuildDesiredProxyAddresses
                    #First, check desired Alias and desired PrimarySMTPAddress for conflicts
                    $AliasAndPrimarySMTPAttemptCount = 0
                    $ExemptObjectGUIDs = @($SADUGUID,$TADUGUID)
                    If ($MailContacts.Count -ge 1)
                    {
                        $ExemptObjectGUIDs += @($MailContacts | foreach {$_.ObjectGUID.guid})
                    }
                    Do {
                        $AliasPass = $false
                        $PrimarySMTPPass = $false
                        $AliasAndPrimarySMTPAttemptCount++
                        switch ($AliasAndPrimarySMTPAttemptCount) {
                            1
                            {
                                $DesiredAlias = $SADU.givenname.substring(0,1) + $SADU.surname
                                #remove spaces and other special characters
                                $DesiredAlias = $DesiredAlias -replace '\s|[^1-9a-zA-Z_-]',''
                            }
                            2
                            {
                                $DesiredAlias = $SADU.givenname.substring(0,2) + $SADU.surname
                                #remove spaces and other special characters
                                $DesiredAlias = $DesiredAlias -replace '\s|[^1-9a-zA-Z_-]',''
                            }
                            3
                            {
                                $DesiredAlias = $SADU.givenname + '.' + $SADU.surname
                                #remove spaces and other special characters
                                $DesiredAlias = $DesiredAlias -replace '\s|[^1-9a-zA-Z_-]',''
                            }
                        }
                        [string]$DesiredUPNAndPrimarySMTPAddress = ''
                        #using switch here to accomodate other scenarios that are likely to arise
                        switch ($DesiredUPNAndPrimarySMTPAddress)
                        {
                            #if ForceTargetPrimarySMTPDomain is used
                            {$PSBoundParameters.ContainsKey('ForceTargetPrimarySMTPDomain')}
                            {
                                $DesiredUPNAndPrimarySMTPAddress = $DesiredAlias + '@' + $ForceTargetPrimarySMTPDomain
                                Break
                            }
                            #otherwise use existing PrimarySMTPAddress
                            {-not [string]::IsNullOrWhiteSpace($SADUCurrentPrimarySmtpAddress)}
                            {
                                $DesiredUPNAndPrimarySMTPAddress = $SADUCurrentPrimarySmtpAddress
                            }
                        }
                        if (Test-ExchangeAlias -Alias $DesiredAlias -ExemptObjectGUIDs $ExemptObjectGUIDs -ExchangeOrganization $TargetExchangeOrganization)
                        {
                            $AliasPass = $true
                        }
                        if (Test-ExchangeProxyAddress -ProxyAddress $DesiredUPNAndPrimarySMTPAddress -ProxyAddressType SMTP -ExemptObjectGUIDs $ExemptObjectGUIDs -ExchangeOrganization $TargetExchangeOrganization)
                        {
                            $PrimarySMTPPass = $true
                        }
                    }
                    Until (($AliasPass -and $PrimarySMTPPass) -or $AliasAndPrimarySMTPAttemptCount -gt 3)
                    if ($AliasAndPrimarySMTPAttemptCount -gt 3) {
                            Write-Log -message "Was not able to find a valid alias and/or PrimarySMTPAddress to Assign to the target: $ID" -Verbose -EntryType Failed
                            Export-FailureRecord -Identity $ID -ExceptionCode 'InvalidAliasOrPrimarySMTPAddress' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                            continue nextID
                    }
                    else {
                        $null = Add-ExchangeAliasToTestExchangeAlias -Alias $DesiredAlias -ObjectGUID $TADUGUID
                        $null = Add-ExchangeProxyAddressToTestExchangeProxyAddress -ProxyAddress $DesiredUPNAndPrimarySMTPAddress -ObjectGUID $TADUGUID -ProxyAddressType SMTP
                    }
                    if ($AddAdditionalSMTPProxyAddress)
                    {
                        $AddressesToAdd = @(
                            foreach ($pattern in $AdditionalSMTPProxyAddressPattern)
                            {
                                $PrefixPattern,$SMTPDomain = $pattern.split('@')
                                switch ($PrefixPattern)
                                {
                                    'Alias'
                                    {
                                        $ProposedAddress = 'smtp:' + $DesiredAlias + '@' + $SMTPDomain
                                    }
                                    'PrimaryPrefix'
                                    {
                                        $Prefix = $DesiredUPNAndPrimarySMTPAddress.Split('@')[0]
                                        $ProposedAddress = 'smtp:' + $Prefix  + '@' + $SMTPDomain
                                    }
                                }#switch
                                if (Test-ExchangeProxyAddress -ProxyAddress $ProposedAddress -ProxyAddressType SMTP -ExemptObjectGUIDs $ExemptObjectGUIDs -ExchangeOrganization $TargetExchangeOrganization)
                                {
                                    $ProposedAddress
                                    $null = Add-ExchangeProxyAddressToTestExchangeProxyAddress -ProxyAddress $ProposedAddress -ObjectGUID $TADUGUID -ProxyAddressType SMTP
                                }
                                else
                                {
                                    Export-FailureRecord -Identity $ID -ExceptionCode 'InvalidAdditionalProxyAddress' -FailureGroup FailedToAddAdditionalProxyAddress -RelatedObjectIdentifier $TADUGUID -RelatedObjectIdentifierType ObjectGUID -ExceptionDetails $ProposedAddress
                                }
                            }
                        )
                    }
                    #Build Proxy Address Array to use in Target AD
                    $writeProgressParams.currentOperation = "Build Proxy Addresses Array for $ID"
                    Write-Progress @writeProgressParams
                    #setup for get-desiredproxyaddresses function to calculate updated addresses
                    $GetDesiredProxyAddressesParams = @{
                        CurrentProxyAddresses=$SADU.proxyAddresses
                        LegacyExchangeDNs=@($SADU.legacyExchangeDN)
                        Recipients = @($TADU)
                        VerifyAddTargetAddress = $true
                        DesiredOrCurrentAlias = $DesiredAlias
                        TargetDeliveryDomain = $TargetDeliveryDomain
                    }
                    if (-not [string]::IsNullOrWhiteSpace($ForceTargetPrimarySMTPDomain)) {
                        $GetDesiredProxyAddressesParams.DesiredPrimaryAddress=$DesiredUPNAndPrimarySMTPAddress
                    }
                    if ($AddressesToAdd.Count -ge 1) {$GetDesiredProxyAddressesParams.AddressesToAdd = $AddressesToAdd}
                    #include contacts legacyexchangedn as x500 and proxy addresses if contacts were found
                    if ($MailContacts.Count -ge 1) {$GetDesiredProxyAddressesParams.legacyExchangeDNs += $MailContacts.LegacyExchangeDN; $GetDesiredProxyAddressesParams.Recipients += $MailContacts}
                    if ($PSBoundParameters.ContainsKey('DomainsToRemove')){$GetDesiredProxyAddressesParams.DomainsToRemove = $DomainsToRemove}
                    $DesiredProxyAddresses = Get-DesiredProxyAddresses @GetDesiredProxyAddressesParams
                    #endregion BuildDesiredProxyAddresses
                    #endregion Preparing To Generate Intermediate Object
                    #region IntermediateObjectGeneration
                    $IntermediateObject = [pscustomobject]@{
                        SourceUserObjectCN = $SADU.CanonicalName
                        SourceUserObjectGUID = $SADUGUID
                        SourceUserObjectEnabled = $SADU.Enabled
                        SourceUserGivenName = $SADU.GivenName
                        SourceUserSurName = $SADU.Surname
                        SourceUserMail = $SADU.Mail
                        SourceUserAlias = $SADU.mailNickname
                        SourceUserObjectIsExchangeRecipient = $SourceUserObjectIsExchangeRecipient
                        SourceUserObjectExchangeRecipientType = $SADUUserObjectExchangeRecipientType
                        SourceUserPrimarySMTPAddress = $SADUCurrentPrimarySmtpAddress
                        SourceUserObject = $SADU
                        MatchAttribute = $TADU.MatchAttribute
                        TargetUserObjectIsSourceUserObject = if ($TADUGUID -eq $SADUGUID) {$true} else {$false}
                        TargetUserObjectCN = $TADU.CanonicalName
                        TargetUserObjectGUID = $TADUGUID
                        TargetUserObjectEnabled = $TADU.Enabled
                        TargetUserGivenName = $TADU.GivenName
                        TargetUserSurName = $TADU.Surname
                        TargetUserMail = $TADU.Mail
                        TargetUserAlias = $TADU.mailNickname
                        TargetUserObjectIsExchangeRecipient = $TargetUserObjectIsExchangeRecipient
                        TargetUserObjectExchangeRecipientType = $TargetUserObjectExchangeRecipientType
                        TargetUserObjectPrimarySMTPAddress = $TADUCurrentPrimarySmtpAddress
                        TargetUserObject = $TADU
                        MatchingContactObject = $MailContacts
                        MatchedContactCount = $MailContacts.Count
                        DesiredProxyAddresses = $DesiredProxyAddresses
                        DesiredUPNAndPrimarySMTPAddress = $DesiredUPNAndPrimarySMTPAddress
                        DesiredAlias = $DesiredAlias
                        DesiredCoexistenceRoutingAddress = $forwardingAddressHash.$SADUCurrentPrimarySmtpAddress
                        #"O365_$($DesiredAlias)@$($SADUCurrentPrimarySmtpAddress.split('@')[1])"
                        DesiredTargetAddress = "SMTP:$($DesiredAlias)@$($TargetDeliveryDomain)"
                        ShouldUpdateUPN = if ($DesiredUPNAndPrimarySMTPAddress -ne $TADU.UserPrincipalName) {$true} else {$false}
                    }
                    $IntermediateObject
                    #endregion IntermediateObjectGeneration
                }#foreach $SADU in $SourceData
                $writeProgressParams.currentOperation = "Completing Intermediate Objects Operations"
                Write-Progress @writeProgressParams -Completed
            )#IntermediateObjects
            Write-Log -Message "$($IntermediateObjects.count) Object(s) Processed (Lookup of Source and Target Objects and Attribute Calculations)." -EntryType Notification
            #region CYABackup
            #depth must be 2 or greater to capture and restore MV attributes like proxy addresses correctly
            Export-Data -DataToExport $IntermediateObjects -DataToExportTitle IntermediateObjects -Depth 3 -DataType json
            #endregion CYABackup
            #############################################################
            #preparation activities complete, time to write changes to Target AD Objects
            #############################################################
            #region write
            $recordcount = $IntermediateObjects.Count
            $cr = 0
            #Write to Target Objects, Delete Source Objects, Update Exchange Recipient for Target Objects
            $ProcessedObjects =
            @(
                :nextIntObj
                foreach ($IntObj in $IntermediateObjects) {
                    #region PrepareForTargetOperation
                    $cr++
                    $TADUGUID = $IntObj.TargetUserObjectGUID
                    $SADUGUID = $IntObj.SourceUserObjectGUID
                    $TADU = $IntObj.TargetUserObject
                    $SADU = $IntObj.SourceUserObject
                    $writeProgressParams = @{
                            Activity = "Update Target Object"
                            Status = "Processing Record $cr of $recordcount : $TADUGUID"
                            PercentComplete = $cr/$RecordCount*100
                    }
                    $TargetDomain = Get-ADObjectDomain -adobject $TADU
                    $writeProgressParams.currentoperation = "Updating Attributes for $TADUGUID in $TargetAD using AD Cmdlets"
                    Write-Progress @writeProgressParams
                    #endregion PrepareForTargetOperation
                    #region DetermineTargetOperation
                    $TargetOperation = 'None'
                    if ($intobj.TargetUserObjectIsExchangeRecipient)
                    {
                        switch -Wildcard ($IntObj.TargetUserObjectExchangeRecipientType)
                        {
                            UserMailbox
                            {
                                $TargetOperation = 'UpdateAndMigrateOnPremisesMailbox'
                            }
                            Default
                            {
                                $TargetOperation = 'None'
                            }
                        }
                    }
                    elseif ($intobj.TargetUserObjectIsSourceUserObject -and $intobj.SourceUserObjectIsExchangeRecipient)
                    {
                        if ($intobj.SourceUserObjectExchangeRecipientType -like '*Mailbox*')
                        {
                            $TargetOperation = 'SourceIsTarget:UpdateAndMigrateOnPremisesMailbox'
                        }
                        else {$TargetOperation = 'None'}
                    }
                    elseif ($PreserveSourceMailbox)
                    {
                        if (($intobj.SourceUserObjectIsExchangeRecipient) -and ($intobj.SourceUserObjectExchangeRecipientType -like '*Mailbox*') -and (-not $intobj.TargetUserObjectIsExchangeRecipient))
                        {
                            $TargetOperation = 'ConnectSourceMailboxToTarget:UpdateAndMigrateOnPremisesMailbox'
                        }
                        else {$TargetOperation = 'None'}
                    }
                    elseif (($intobj.SourceUserObjectIsExchangeRecipient) -and ($intobj.SourceUserObjectExchangeRecipientType -like '*Mailbox*') -and (-not $intobj.TargetUserObjectIsExchangeRecipient))
                    {
                        $TargetOperation = 'EnableMailUserWithMailboxGUID'
                    }
                    else
                    {
                        if ([string]::IsNullOrWhiteSpace($intobj.DesiredCoexistenceRoutingAddress))
                        {
                            $TargetOperation = 'None'
                        }
                        else
                        {
                            $TargetOperation = 'EnableRemoteMailbox'
                        }
                    }
                    $intobj | Add-Member -MemberType NoteProperty -Name TargetOperation -Value $TargetOperation
                    #endregion DetermineTargetOperation
                    #region PerformTargetAttributeUpdate
                    if ($TestOnly)
                    {
                        Write-Output -InputObject $IntObj
                    }
                    else
                    {
                        Push-Location
                        Set-Location -Path $($TargetAD + ':\')
                    switch ($TargetOperation)
                    {
                        'None'
                        {
                            Write-Output -InputObject $IntObj
                            Write-Log -Message "Target Operation Could Not Be Determined for $SADUGUID" -Verbose -ErrorLog -EntryType Failed
                            Export-FailureRecord -Identity $ID -ExceptionCode 'TargetOperationNotDetermined' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType 'ObjectGUID'
                            continue nextIntObj
                        }
                        'EnableRemoteMailbox'
                        {
                            #############################################################
                            #clear the target attributes in target AD
                            #############################################################
                            #ClearTargetAttributes
                            $setaduserparams1 = @{
                                Identity=$TADUGUID
                                clear=$TargetAttributestoClear
                                Server=$TargetDomain
                                ErrorAction = 'Stop'
                            }#setaduserparams1
                            #add UPN to clear list if UPN needs to be replaced
                            if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                                $setaduserparams1.'Clear' += 'UserPrincipalName'
                            }
                            $message = "Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD"
                            try {
                                Write-Log -message $message -EntryType Attempting
                                set-aduser @setaduserparams1
                                Write-Log -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-Log -message $message -EntryType Failed -Verbose -ErrorLog
                                Write-Log -Message $_.tostring() -ErrorLog
                                Export-FailureRecord -Identity $TADUGUID -ExceptionCode 'FailedToClearTargetAttributes' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                                Continue nextIntObj
                            }#catch
                            #############################################################
                            #set new values on the target attributes in target AD
                            #############################################################
                            $setaduserparams2 = @{
                                identity=$TADUGUID
                                add=@{
                                    #Adjust this section to use parameters depending on the recipient type which should be created.  Following is currently set for Remote Mailbox.
                                    msExchRecipientDisplayType = -2147483642 #RemoteUserMailbox
                                    msExchRecipientTypeDetails = 2147483648 #RemoteUserMailbox
                                    msExchRemoteRecipientType = 1 #ProvisionMailbox
                                    msExchVersion = 44220983382016
                                    targetaddress = $intobj.DesiredTargetAddress
                                    #hard coding these for a customer for now
                                    msExchUsageLocation = 'US'
                                    c = 'US'
                                    co = 'United States'
                                    #countrycode = 840
                                    extensionattribute5 = $($SADUGUID.tostring())
                                }
                                Server=$TargetDomain
                                ErrorAction = 'Stop'
                            }#setaduserparams2
                            #Disable Email Address Policy if admin user specified the parameter, otherwise leave status quo of source object
                            if ($DisableEmailAddressPolicyInTarget) {
                                $setaduserparams2.'add'.msExchPoliciesExcluded = '{26491cfc-9e50-4857-861b-0cb8df22b5d7}'
                            }
                            elseif (-not [string]::IsNullOrWhiteSpace($SADU.msExchPoliciesExcluded)) {
                                $setaduserparams2.'add'.msExchPoliciesExcluded = $SADU.msExchPoliciesExcluded
                            }
                            <#region NotUsing
                                    if(-not [string]::IsNullOrWhiteSpace($sadu.displayname)) {
                                    $setaduserparams2.'add'.DisplayName = [string]$($SADU.displayName)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($sadu.department)) {
                                    $setaduserparams2.'add'.department = [string]$($SADU.department)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($SADU.msExchMailboxGUID)) {
                                    $setaduserparams2.'add'.msExchMailboxGUID = [byte[]]$($SADU.msExchMailboxGUID)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveGUID)) {
                                    $setaduserparams2.'add'.msExchArchiveGUID = [byte[]]$($SADU.msExchArchiveGUID)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveName)) {
                                    $setaduserparams2.'add'.msExchArchiveName = [byte[]]$($SADU.msExchArchiveName)
                                    }
                                    if (-not [string]::IsNullOrWhiteSpace($TADU.c)) {
                                        $setaduserparams2.'add'.msExchangeUsageLocation = $TADU.c
                                    }
                            endRegion NotUsing#>
                            if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredUPNAndPrimarySMTPAddress)) {
                                $setaduserparams2.'add'.Mail = $intobj.DesiredUPNAndPrimarySMTPAddress
                            }
                            if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredAlias)) {
                                $setaduserparams2.'add'.mailNickName = [string]$intobj.DesiredAlias
                            }
                            if(-not [string]::IsNullOrWhiteSpace($SADU.msExchangeUserCulture)) {
                                $setaduserparams2.'add'.msExchangeUserCulture = [string]$SADU.msExchangeUserCulture
                            }
                            if(-not [string]::IsNullOrWhiteSpace($IntObj.DesiredProxyAddresses)) {
                                $setaduserparams2.'add'.proxyaddresses = [string[]]$($IntObj.DesiredProxyAddresses)
                            }
                            if(-not [string]::IsNullOrWhiteSpace($sadu.msExchMasterAccountSID)) {
                                $setaduserparams2.'add'.msExchMasterAccountSID = [string]$($SADU.msExchMasterAccountSID)
                            }
                            if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                                $setaduserparams2.'add'.UserPrincipalName = $intobj.DesiredUPNAndPrimarySMTPAddress
                            }
                            try {
                                $message = "SET target attributes $($setaduserparams2.'Add'.keys -join ';') for $TADUGUID in $TargetAD"
                                Write-Log -message $message -EntryType Attempting
                                set-aduser @setaduserparams2
                                Write-Log -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-Log -message "FAILED: SET target attributes $($setaduserparams2.'Add'.keys -join ';')  for $TADUGUID in $TargetAD" -Verbose -ErrorLog
                                Write-Log -Message $_.tostring() -ErrorLog
                                Export-FailureRecord -Identity $id -ExceptionCode 'FailedToSetTargetAttributes' -FailureGroup PartiallyProcessed -RelatedObjectIdentifier $TADUGUID -RelatedObjectIdentifierType ObjectGUID
                                continue nextIntObj
                            }#catch
                        }#EnableRemoteMailbox
                        'UpdateAndMigrateOnPremisesMailbox'
                        {
                            #############################################################
                            #clear the target attributes in target AD
                            #############################################################
                            #ClearTargetAttributes
                            #Since this is an existing mailbox, we won't clear all the mail attributes
                            $ExemptTargetAttributesForExistingMailbox =
                            @(
                                'msExchArchiveGUID'
                                'msExchArchiveName'
                                'msExchMailboxGUID'
                                'msExchRecipientDisplayType'
                                'msExchRecipientTypeDetails'
                                'msExchRemoteRecipientType'
                                'msExchUserCulture'
                                'msExchVersion'
                                'targetAddress'
                            )
                            $TargetAttributestoClearForExistingMailbox = $TargetAttributestoClear | Where-Object {$_ -notin $ExemptTargetAttributesForExistingMailbox}
                            $setaduserparams1 = @{
                                Identity=$TADUGUID
                                clear=$TargetAttributestoClearForExistingMailbox
                                Server=$TargetDomain
                                ErrorAction = 'Stop'
                            }#setaduserparams1
                            #add UPN to clear list if UPN needs to be replaced
                            if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                                $setaduserparams1.'Clear' += 'UserPrincipalName'
                            }
                            $message = "Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD"
                            try {
                                Write-Log -message $message -EntryType Attempting
                                set-aduser @setaduserparams1
                                Write-Log -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-Log -message $message -EntryType Failed -Verbose -ErrorLog
                                Write-Log -Message $_.tostring() -ErrorLog
                                Export-FailureRecord -Identity $TADUGUID -ExceptionCode 'FailedToClearTargetAttributes' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                                Continue nextIntObj
                            }#catch
                            #############################################################
                            #set new values on the target attributes in target AD
                            #############################################################
                            $setaduserparams2 = @{
                                identity=$TADUGUID
                                add=@{
                                    #Adjust this section to use parameters depending on the recipient type which should be created.  Following is currently set for Existing Mailbox.
                                    #msExchRecipientDisplayType = -2147483642 #RemoteUserMailbox
                                    #msExchRecipientTypeDetails = 2147483648 #RemoteUserMailbox
                                    #msExchRemoteRecipientType = 1 #ProvisionMailbox
                                    #msExchVersion = 44220983382016
                                    #targetaddress = $intobj.DesiredTargetAddress #can't do this here - must be post mailbox move
                                    #hard coding these for a customer for now
                                    msExchUsageLocation = 'US'
                                    c = 'US'
                                    co = 'United States'
                                    #countrycode = 840
                                    extensionattribute5 = $($SADUGUID.tostring())
                                }
                                Server=$TargetDomain
                                ErrorAction = 'Stop'
                            }#setaduserparams2
                            #Disable Email Address Policy if admin user specified the parameter, otherwise leave status quo of source object
                            if ($DisableEmailAddressPolicyInTarget) {
                                $setaduserparams2.'add'.msExchPoliciesExcluded = '{26491cfc-9e50-4857-861b-0cb8df22b5d7}'
                            }
                            elseif (-not [string]::IsNullOrWhiteSpace($SADU.msExchPoliciesExcluded)) {
                                $setaduserparams2.'add'.msExchPoliciesExcluded = $SADU.msExchPoliciesExcluded
                            }
                            <#region NotUsing
                                    if(-not [string]::IsNullOrWhiteSpace($sadu.displayname)) {
                                    $setaduserparams2.'add'.DisplayName = [string]$($SADU.displayName)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($sadu.department)) {
                                    $setaduserparams2.'add'.department = [string]$($SADU.department)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($SADU.msExchMailboxGUID)) {
                                    $setaduserparams2.'add'.msExchMailboxGUID = [byte[]]$($SADU.msExchMailboxGUID)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveGUID)) {
                                    $setaduserparams2.'add'.msExchArchiveGUID = [byte[]]$($SADU.msExchArchiveGUID)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveName)) {
                                    $setaduserparams2.'add'.msExchArchiveName = [byte[]]$($SADU.msExchArchiveName)
                                    }
                                    if (-not [string]::IsNullOrWhiteSpace($TADU.c)) {
                                        $setaduserparams2.'add'.msExchangeUsageLocation = $TADU.c
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($SADU.msExchangeUserCulture)) {
                                        $setaduserparams2.'add'.msExchangeUserCulture = [string]$SADU.msExchangeUserCulture
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($sadu.msExchMasterAccountSID)) {
                                        $setaduserparams2.'add'.msExchMasterAccountSID = [string]$($SADU.msExchMasterAccountSID)
                                    }
                            endRegion NotUsing#>
                            if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredUPNAndPrimarySMTPAddress)) {
                                $setaduserparams2.'add'.Mail = $intobj.DesiredUPNAndPrimarySMTPAddress
                            }
                            if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredAlias)) {
                                $setaduserparams2.'add'.mailNickName = [string]$intobj.DesiredAlias
                            }
                            if(-not [string]::IsNullOrWhiteSpace($IntObj.DesiredProxyAddresses)) {
                                $setaduserparams2.'add'.proxyaddresses = [string[]]$($IntObj.DesiredProxyAddresses)
                            }
                            if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                                $setaduserparams2.'add'.UserPrincipalName = $intobj.DesiredUPNAndPrimarySMTPAddress
                            }
                            try {
                                $message = "SET target attributes $($setaduserparams2.'Add'.keys -join ';') for $TADUGUID in $TargetAD"
                                Write-Log -message $message -EntryType Attempting
                                set-aduser @setaduserparams2
                                Write-Log -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                                Write-Log -Message $_.tostring() -ErrorLog
                                Export-FailureRecord -Identity $id -ExceptionCode 'FailedToSetTargetAttributes' -FailureGroup PartiallyProcessed -RelatedObjectIdentifier $TADUGUID -RelatedObjectIdentifierType ObjectGUID
                                continue nextIntObj
                            }#catch
                        }#UpdateAndMigrateOnPremisesMailbox
                        'SourceIsTarget:UpdateAndMigrateOnPremisesMailbox'
                        {
                            #############################################################
                            #clear the target attributes in target AD
                            #############################################################
                            #ClearTargetAttributes
                            #Since this is an existing mailbox, we won't clear all the mail attributes
                            $ExemptTargetAttributesForExistingMailbox =
                            @(
                                'msExchArchiveGUID'
                                'msExchArchiveName'
                                'msExchMailboxGUID'
                                'msExchRecipientDisplayType'
                                'msExchRecipientTypeDetails'
                                'msExchRemoteRecipientType'
                                'msExchUserCulture'
                                'msExchVersion'
                                'targetAddress'
                            )
                            $TargetAttributestoClearForExistingMailbox = $TargetAttributestoClear | Where-Object {$_ -notin $ExemptTargetAttributesForExistingMailbox}
                            $setaduserparams1 = @{
                                Identity=$TADUGUID
                                clear=$TargetAttributestoClearForExistingMailbox
                                Server=$TargetDomain
                                ErrorAction = 'Stop'
                            }#setaduserparams1
                            #add UPN to clear list if UPN needs to be replaced
                            if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                                $setaduserparams1.'Clear' += 'UserPrincipalName'
                            }
                            $message = "Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD"
                            try {
                                Write-Log -message $message -EntryType Attempting
                                set-aduser @setaduserparams1
                                Write-Log -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-Log -message $message -EntryType Failed -Verbose -ErrorLog
                                Write-Log -Message $_.tostring() -ErrorLog
                                Export-FailureRecord -Identity $TADUGUID -ExceptionCode 'FailedToClearTargetAttributes' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                                Continue nextIntObj
                            }#catch
                            #############################################################
                            #set new values on the target attributes in target AD
                            #############################################################
                            $setaduserparams2 = @{
                                identity=$TADUGUID
                                add=@{
                                    #Adjust this section to use parameters depending on the recipient type which should be created.  Following is currently set for Existing Mailbox.
                                    #msExchRecipientDisplayType = -2147483642 #RemoteUserMailbox
                                    #msExchRecipientTypeDetails = 2147483648 #RemoteUserMailbox
                                    #msExchRemoteRecipientType = 1 #ProvisionMailbox
                                    #msExchVersion = 44220983382016
                                    #targetaddress = $intobj.DesiredTargetAddress #can't do this here - must be post mailbox move
                                    #hard coding these for a customer for now
                                    msExchUsageLocation = 'US'
                                    c = 'US'
                                    co = 'United States'
                                    #countrycode = 840
                                    extensionattribute5 = $($SADUGUID.tostring())
                                }
                                Server=$TargetDomain
                                ErrorAction = 'Stop'
                            }#setaduserparams2
                            #Disable Email Address Policy if admin user specified the parameter, otherwise leave status quo of source object
                            if ($DisableEmailAddressPolicyInTarget) {
                                $setaduserparams2.'add'.msExchPoliciesExcluded = '{26491cfc-9e50-4857-861b-0cb8df22b5d7}'
                            }
                            elseif (-not [string]::IsNullOrWhiteSpace($SADU.msExchPoliciesExcluded)) {
                                $setaduserparams2.'add'.msExchPoliciesExcluded = $SADU.msExchPoliciesExcluded
                            }
                            <#region NotUsing
                                    if(-not [string]::IsNullOrWhiteSpace($sadu.displayname)) {
                                    $setaduserparams2.'add'.DisplayName = [string]$($SADU.displayName)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($sadu.department)) {
                                    $setaduserparams2.'add'.department = [string]$($SADU.department)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($SADU.msExchMailboxGUID)) {
                                    $setaduserparams2.'add'.msExchMailboxGUID = [byte[]]$($SADU.msExchMailboxGUID)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveGUID)) {
                                    $setaduserparams2.'add'.msExchArchiveGUID = [byte[]]$($SADU.msExchArchiveGUID)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveName)) {
                                    $setaduserparams2.'add'.msExchArchiveName = [byte[]]$($SADU.msExchArchiveName)
                                    }
                                    if (-not [string]::IsNullOrWhiteSpace($TADU.c)) {
                                        $setaduserparams2.'add'.msExchangeUsageLocation = $TADU.c
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($SADU.msExchangeUserCulture)) {
                                        $setaduserparams2.'add'.msExchangeUserCulture = [string]$SADU.msExchangeUserCulture
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($sadu.msExchMasterAccountSID)) {
                                        $setaduserparams2.'add'.msExchMasterAccountSID = [string]$($SADU.msExchMasterAccountSID)
                                    }
                            endRegion NotUsing#>
                            if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredUPNAndPrimarySMTPAddress)) {
                                $setaduserparams2.'add'.Mail = $intobj.DesiredUPNAndPrimarySMTPAddress
                            }
                            if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredAlias)) {
                                $setaduserparams2.'add'.mailNickName = [string]$intobj.DesiredAlias
                            }
                            if(-not [string]::IsNullOrWhiteSpace($IntObj.DesiredProxyAddresses)) {
                                $setaduserparams2.'add'.proxyaddresses = [string[]]$($IntObj.DesiredProxyAddresses)
                            }
                            if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                                $setaduserparams2.'add'.UserPrincipalName = $intobj.DesiredUPNAndPrimarySMTPAddress
                            }
                            try {
                                $message = "SET target attributes $($setaduserparams2.'Add'.keys -join ';') for $TADUGUID in $TargetAD"
                                Write-Log -message $message -EntryType Attempting
                                set-aduser @setaduserparams2
                                Write-Log -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                                Write-Log -Message $_.tostring() -ErrorLog
                                Export-FailureRecord -Identity $id -ExceptionCode 'FailedToSetTargetAttributes' -FailureGroup PartiallyProcessed -RelatedObjectIdentifier $TADUGUID -RelatedObjectIdentifierType ObjectGUID
                                continue nextIntObj
                            }#catch
                        }#SourceIsTarget:UpdateAndMigrateOnPremisesMailbox
                        'ConnectSourceMailboxToTarget:UpdateAndMigrateOnPremisesMailbox'
                        {
                            #############################################################
                            #clear the target attributes in target AD
                            #############################################################
                            #ClearTargetAttributes
                            #Since this is an existing mailbox, we won't clear all the mail attributes
                            $ExemptTargetAttributesForExistingMailbox =
                            @(
                                'msExchArchiveGUID'
                                'msExchArchiveName'
                                'msExchMailboxGUID'
                                'msExchRecipientDisplayType'
                                'msExchRecipientTypeDetails'
                                'msExchRemoteRecipientType'
                                'msExchUserCulture'
                                'msExchVersion'
                                'targetAddress'
                            )
                            $TargetAttributestoClearForExistingMailbox = $TargetAttributestoClear | Where-Object {$_ -notin $ExemptTargetAttributesForExistingMailbox}
                            $setaduserparams1 = @{
                                Identity=$TADUGUID
                                clear=$TargetAttributestoClearForExistingMailbox
                                Server=$TargetDomain
                                ErrorAction = 'Stop'
                            }#setaduserparams1
                            #add UPN to clear list if UPN needs to be replaced
                            if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                                $setaduserparams1.'Clear' += 'UserPrincipalName'
                            }
                            $message = "Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD"
                            try {
                                Write-Log -message $message -EntryType Attempting
                                set-aduser @setaduserparams1
                                Write-Log -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-Log -message $message -EntryType Failed -Verbose -ErrorLog
                                Write-Log -Message $_.tostring() -ErrorLog
                                Export-FailureRecord -Identity $TADUGUID -ExceptionCode 'FailedToClearTargetAttributes' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                                Continue nextIntObj
                            }#catch
                            #############################################################
                            #set new values on the target attributes in target AD
                            #############################################################
                            $setaduserparams2 = @{
                                identity=$TADUGUID
                                add=@{
                                    #Adjust this section to use parameters depending on the recipient type which should be created.  Following is currently set for Existing Mailbox.
                                    msExchRecipientDisplayType = 1073741824 #UserMailbox
                                    msExchRecipientTypeDetails = 1 #UserMailbox
                                    #msExchRemoteRecipientType = 1 #ProvisionMailbox
                                    msExchVersion = 44220983382016
                                    msExchHomeServerName = $($SADU.msExchHomeServerName)
                                    homeMDB = $($SADU.homeMDB)
                                    homeMTA = $($SADU.homeMTA)
                                    #targetaddress = $intobj.DesiredTargetAddress #can't do this here - must be post mailbox move
                                    #hard coding these for a customer for now
                                    msExchUsageLocation = 'US'
                                    c = 'US'
                                    co = 'United States'
                                    #countrycode = 840
                                    extensionattribute5 = $($SADUGUID.tostring())
                                }
                                Server=$TargetDomain
                                ErrorAction = 'Stop'
                            }#setaduserparams2
                            #Disable Email Address Policy if admin user specified the parameter, otherwise leave status quo of source object
                            if ($DisableEmailAddressPolicyInTarget) {
                                $setaduserparams2.'add'.msExchPoliciesExcluded = '{26491cfc-9e50-4857-861b-0cb8df22b5d7}'
                            }
                            elseif (-not [string]::IsNullOrWhiteSpace($SADU.msExchPoliciesExcluded)) {
                                $setaduserparams2.'add'.msExchPoliciesExcluded = $SADU.msExchPoliciesExcluded
                            }
                            #Move Source Mailbox Attributes to Target
                            if(-not [string]::IsNullOrWhiteSpace($SADU.msExchMailboxGUID)) {
                            $setaduserparams2.'add'.msExchMailboxGUID = [byte[]]$($SADU.msExchMailboxGUID)
                            }
                            if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveGUID)) {
                            $setaduserparams2.'add'.msExchArchiveGUID = [byte[]]$($SADU.msExchArchiveGUID)
                            }
                            if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveName)) {
                            $setaduserparams2.'add'.msExchArchiveName = [byte[]]$($SADU.msExchArchiveName)
                            }
                            if (-not [string]::IsNullOrWhiteSpace($TADU.c)) {
                                $setaduserparams2.'add'.msExchangeUsageLocation = $TADU.c
                            }
                            if(-not [string]::IsNullOrWhiteSpace($SADU.msExchangeUserCulture)) {
                                $setaduserparams2.'add'.msExchangeUserCulture = [string]$SADU.msExchangeUserCulture
                            }
                            if(-not [string]::IsNullOrWhiteSpace($sadu.msExchMasterAccountSID)) {
                                $setaduserparams2.'add'.msExchMasterAccountSID = [string]$($SADU.msExchMasterAccountSID)
                            }
                            <#region NotUsing
                                    if(-not [string]::IsNullOrWhiteSpace($sadu.displayname)) {
                                    $setaduserparams2.'add'.DisplayName = [string]$($SADU.displayName)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($sadu.department)) {
                                    $setaduserparams2.'add'.department = [string]$($SADU.department)
                                    }
                            endRegion NotUsing#>
                            if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredUPNAndPrimarySMTPAddress)) {
                                $setaduserparams2.'add'.Mail = $intobj.DesiredUPNAndPrimarySMTPAddress
                            }
                            if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredAlias)) {
                                $setaduserparams2.'add'.mailNickName = [string]$intobj.DesiredAlias
                            }
                            if(-not [string]::IsNullOrWhiteSpace($IntObj.DesiredProxyAddresses)) {
                                $setaduserparams2.'add'.proxyaddresses = [string[]]$($IntObj.DesiredProxyAddresses)
                            }
                            if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                                $setaduserparams2.'add'.UserPrincipalName = $intobj.DesiredUPNAndPrimarySMTPAddress
                            }
                            try {
                                $message = "SET target attributes $($setaduserparams2.'Add'.keys -join ';') for $TADUGUID in $TargetAD"
                                Write-Log -message $message -EntryType Attempting
                                set-aduser @setaduserparams2
                                Write-Log -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                                Write-Log -Message $_.tostring() -ErrorLog
                                Export-FailureRecord -Identity $id -ExceptionCode 'FailedToSetTargetAttributes' -FailureGroup PartiallyProcessed -RelatedObjectIdentifier $TADUGUID -RelatedObjectIdentifierType ObjectGUID
                                continue nextIntObj
                            }#catch
                        }#ConnectSourceMailboxToTarget:UpdateAndMigrateOnPremisesMailbox
                        'EnableMailUserWithMailboxGUID'
                        {
                            #############################################################
                            #clear the target attributes in target AD
                            #############################################################
                            #ClearTargetAttributes
                            #Since this is an existing mailbox, we won't clear all the mail attributes
                            $ExemptTargetAttributesForExistingMailbox =
                            @(
                            )
                            $TargetAttributestoClearForExistingMailbox = $TargetAttributestoClear | Where-Object {$_ -notin $ExemptTargetAttributesForExistingMailbox}
                            $setaduserparams1 = @{
                                Identity=$TADUGUID
                                clear=$TargetAttributestoClearForExistingMailbox
                                Server=$TargetDomain
                                ErrorAction = 'Stop'
                            }#setaduserparams1
                            #add UPN to clear list if UPN needs to be replaced
                            if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                                $setaduserparams1.'Clear' += 'UserPrincipalName'
                            }
                            $message = "Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD"
                            try {
                                Write-Log -message $message -EntryType Attempting
                                set-aduser @setaduserparams1
                                Write-Log -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-Log -message $message -EntryType Failed -Verbose -ErrorLog
                                Write-Log -Message $_.tostring() -ErrorLog
                                Export-FailureRecord -Identity $TADUGUID -ExceptionCode 'FailedToClearTargetAttributes' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                                Continue nextIntObj
                            }#catch
                            #############################################################
                            #set new values on the target attributes in target AD
                            #############################################################
                            $setaduserparams2 = @{
                                identity=$TADUGUID
                                add=@{
                                    #Adjust this section to use parameters depending on the recipient type which should be created.  Following is currently set for Existing Mailbox.
                                    msExchRecipientDisplayType = -2147483642 #MailUser or MailContact
                                    msExchRecipientTypeDetails = 128 #MailUser
                                    #msExchRemoteRecipientType = 2 #On Premises Mailbox #this breaks it don't set to this value
                                    msExchVersion = 44220983382016
                                    #targetaddress = $intobj.DesiredTargetAddress #can't do this here - must be post mailbox move
                                    #hard coding these for a customer for now
                                    msExchUsageLocation = 'US'
                                    c = 'US'
                                    co = 'United States'
                                    #countrycode = 840
                                    extensionattribute5 = $($SADUGUID.tostring())
                                }
                                Server=$TargetDomain
                                ErrorAction = 'Stop'
                            }#setaduserparams2
                            #Disable Email Address Policy if admin user specified the parameter, otherwise leave status quo of source object
                            if ($DisableEmailAddressPolicyInTarget) {
                                $setaduserparams2.'add'.msExchPoliciesExcluded = '{26491cfc-9e50-4857-861b-0cb8df22b5d7}'
                            }
                            elseif (-not [string]::IsNullOrWhiteSpace($SADU.msExchPoliciesExcluded)) {
                                $setaduserparams2.'add'.msExchPoliciesExcluded = $SADU.msExchPoliciesExcluded
                            }
                            #Move Source Mailbox Attributes to Target
                            if(-not [string]::IsNullOrWhiteSpace($SADU.msExchMailboxGUID)) {
                            $setaduserparams2.'add'.msExchMailboxGUID = [byte[]]$($SADU.msExchMailboxGUID)
                            }
                            if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveGUID)) {
                            $setaduserparams2.'add'.msExchArchiveGUID = [byte[]]$($SADU.msExchArchiveGUID)
                            }
                            if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveName)) {
                            $setaduserparams2.'add'.msExchArchiveName = [byte[]]$($SADU.msExchArchiveName)
                            }
                            if (-not [string]::IsNullOrWhiteSpace($TADU.c)) {
                                $setaduserparams2.'add'.msExchangeUsageLocation = $TADU.c
                            }
                            if(-not [string]::IsNullOrWhiteSpace($SADU.msExchangeUserCulture)) {
                                $setaduserparams2.'add'.msExchangeUserCulture = [string]$SADU.msExchangeUserCulture
                            }
                            <#region NotUsing
                                    if(-not [string]::IsNullOrWhiteSpace($sadu.displayname)) {
                                    $setaduserparams2.'add'.DisplayName = [string]$($SADU.displayName)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($sadu.department)) {
                                    $setaduserparams2.'add'.department = [string]$($SADU.department)
                                    }
                                    if(-not [string]::IsNullOrWhiteSpace($sadu.msExchMasterAccountSID)) {
                                    $setaduserparams2.'add'.msExchMasterAccountSID = [string]$($SADU.msExchMasterAccountSID)
                            }
                            endRegion NotUsing#>
                            if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredUPNAndPrimarySMTPAddress)) {
                                $setaduserparams2.'add'.Mail = $intobj.DesiredUPNAndPrimarySMTPAddress
                                $setaduserparams2.'add'.TargetAddress = 'SMTP:' + $intobj.DesiredUPNAndPrimarySMTPAddress
                            }
                            if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredAlias)) {
                                $setaduserparams2.'add'.mailNickName = [string]$intobj.DesiredAlias
                            }
                            if(-not [string]::IsNullOrWhiteSpace($IntObj.DesiredProxyAddresses)) {
                                $setaduserparams2.'add'.proxyaddresses = [string[]]$($IntObj.DesiredProxyAddresses)
                            }
                            if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                                $setaduserparams2.'add'.UserPrincipalName = $intobj.DesiredUPNAndPrimarySMTPAddress
                            }
                            try {
                                $message = "SET target attributes $($setaduserparams2.'Add'.keys -join ';') for $TADUGUID in $TargetAD"
                                Write-Log -message $message -EntryType Attempting
                                set-aduser @setaduserparams2
                                Write-Log -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                                Write-Log -Message $_.tostring() -ErrorLog
                                Export-FailureRecord -Identity $id -ExceptionCode 'FailedToSetTargetAttributes' -FailureGroup PartiallyProcessed -RelatedObjectIdentifier $TADUGUID -RelatedObjectIdentifierType ObjectGUID
                                continue nextIntObj
                            }#catch
                        }#ConnectSourceMailboxToTarget:UpdateAndMigrateOnPremisesMailbox
                    }#Switch $TargetOperation
                        Pop-Location
                    #endregion PerformTargetAttributeUpdate
                    #region GroupMemberships
                    #############################################################
                    #add TADU to memberof groups from SADU
                    #############################################################
                    if ($AddTargetToSourceGroups)
                    {
                        $writeProgressParams = @{
                            Activity = "Update Target Object Group Memberships"
                            Status = "Processing Record $cr of $recordcount : $TADUGUID"
                            PercentComplete = $cr/$RecordCount*100
                        }
                        $writeProgressParams.currentoperation = "Adding $TADUGUID to Groups using AD Cmdlets"
                        Write-Progress @writeProgressParams
                        if ($SADU.memberof.count -ge 1 -and ($intobj.TargetUserObjectIsSourceUserObject -eq $false)) {
                            foreach ($groupDN in $SADU.memberof) {
                                try {
                                    $message = "Add $TADUGUID to group $groupDN"
                                    $GroupObject = Get-ADGroup -Identity $groupDN -Properties CanonicalName
                                    $Domain = Get-ADObjectDomain -adobject $GroupObject
                                    Write-Log -message $message -EntryType Attempting
                                    Add-ADGroupMember -Identity $groupDN -Members $TADUGUID -ErrorAction Stop -Confirm:$false -Server $Domain
                                    Write-Log -message $message -EntryType Succeeded
                                }#try
                                catch {
                                    Write-Log -message $message -EntryType Failed -Verbose -ErrorLog
                                    Write-Log -Message $_.tostring() -ErrorLog
                                    Export-FailureRecord -Identity $TADUGUID -ExceptionCode "GroupMembershipFailure:$groupDN" -FailureGroup GroupMembership -RelatedObjectIdentifier $GroupDN -RelatedObjectIdentifierType DistinguishedName
                                }#catch
                            }#Foreach
                        }#If
                    }#if $AddTargetToSourceGroups
                    #endregion GroupMemberships
                    #region ContactDeletionAndAttributeCopy
                    #############################################################
                    #delete contact objects found in the Target AD
                    #############################################################
                    $writeProgressParams = @{
                        Activity = "Process Target Object Related Contacts"
                        Status = "Processing Record $cr of $recordcount : $TADUGUID"
                        PercentComplete = $cr/$RecordCount*100
                    }
                    $writeProgressParams.currentoperation = "Processing Contacts for $TADUGUID"
                    Write-Progress @writeProgressParams
                    if ($deletecontact -and $intobj.MatchingContactObject.count -ge 1) {
                        Write-Log -message "Attempting: Delete $($intobj.MatchingContactObject.count) Mail Contact(s) from $TargetAD" -Verbose
                        foreach ($c in $intobj.MatchingContactObject) {
                            try {
                                Write-Log -message "Attempting: Delete $($c.distinguishedname) Mail Contact from $TargetAD" -Verbose
                                Push-Location
                                Set-Location $($TargetAD + ':\')
                                $Domain = Get-AdObjectDomain -adobject $c -ErrorAction Stop
                                $splat = @{Identity = $c.distinguishedname;Confirm=$false;ErrorAction='Stop';Server=$Domain}
                                Remove-ADObject @splat
                                Pop-Location
                                Write-Log -message "Succeeded: Delete $($c.distinguishedname) Mail Contact from $TargetAD" -Verbose
                            }#try
                            catch {
                                Pop-Location
                                #$Global:ErrorActionPreference = 'Continue'
                                Write-Log -message "FAILED: Delete $($c.distinguishedname) Mail Contact from $TargetAD" -Verbose -ErrorLog
                                Write-Log -Message $_.tostring() -ErrorLog
                                $Global:SEATO_MailContactDeletionFailures+=$c
                            }#catch
                        }#foreach
                    }#if
                    #############################################################
                    #copy contact object memberships to Target AD User
                    #############################################################
                    if ($deletecontact -and $intobj.MatchingContactObject.count -ge 1) {
                        Write-Log -message "Attempting: Add $TADUGUID to Contacts' Distribution Groups in $TargetAD" -Verbose
                        $ContactGroupMemberships = @($intobj.MatchingContactObject | Select-Object -ExpandProperty MemberOf)
                        foreach ($group in $ContactGroupMemberships) {
                            try {
                                $message = "Add-ADGroupMember -Members $TADUGUID -Identity $group"
                                Write-Log -message $message -EntryType Attempting
                                Push-Location
                                Set-Location $($TargetAD + ':\')
                                $ADGroup = Get-ADGroup -Identity $group -ErrorAction Stop -Properties CanonicalName
                                $Domain = Get-AdObjectDomain -adobject $ADGroup -ErrorAction Stop
                                $splat = @{Identity = $group;Confirm=$false;ErrorAction='Stop';Members=$TADUGUID;Server=$Domain}
                                Add-ADGroupMember @splat
                                Pop-Location
                                Write-Log -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Pop-Location
                                Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                                Write-Log -Message $_.tostring() -ErrorLog
                                Export-FailureRecord -Identity $TADUGUID -ExceptionCode "GroupMembershipFailure:$group" -FailureGroup GroupMembership -RelatedObjectIdentifier $Group -RelatedObjectIdentifierType DistinguishedName
                            }#catch
                        }#foreach
                    }#if
                    #endregion ContactDeletionAndAttributeCopy
                    #region DeleteSourceObject
                    #############################################################
                    #Delete SADU from AD
                    #############################################################
                    if ($DeleteSourceObject -and ($intobj.TargetUserObjectIsSourceUserObject -eq $false)) {
                        try {
                            $message = "Remove Object $SADUGUID from AD $SourceAD"
                            Write-Log -message $message -EntryType Attempting
                            $Splat = @{
                                Identity = $SADUGUID
                                ErrorAction = 'Stop'
                                Confirm = $false
                                Server = Get-ADObjectDomain -adobject $SADU
                            }
                            Remove-ADObject @splat
                            Write-Log -message $message -EntryType Succeeded
                        }
                        catch {
                            Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                            Write-Log -Message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity $SADUGUID -ExceptionCode "SourceObjectRemovalFailure:$SADUGUID" -FailureGroup SourceObjectRemoval -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                        }
                    }
                    #endregion DeleteSourceObject
                    #region UpdateSourceObject
                    #############################################################
                    #Update SADU in Source AD
                    #############################################################
                    if ($UpdateSourceObject -and ($intobj.TargetUserObjectIsSourceUserObject -eq $false)) {
                        try {
                            $message = "Update Object $SADUGUID in AD $SourceAD"
                            Write-Log -message $message -EntryType Attempting
                            $Splat = @{
                                Identity = $SADUGUID
                                ErrorAction = 'Stop'
                                Confirm = $false
                                Server = Get-ADObjectDomain -adobject $SADU
                                Replace = @{extensionAttribute5 = $TADUGUID}
                                Add = @{proxyaddresses = $($IntObj.DesiredTargetAddress.tolower())}
                            }
                            Push-Location
                            Set-Location -Path $($SourceAD + ':\')
                            Set-ADObject @splat
                            Write-Log -message $message -EntryType Succeeded
                            Pop-Location
                        }
                        catch {
                            Pop-Location
                            Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                            Write-Log -Message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity $SADUGUID -ExceptionCode "SourceObjectRemovalFailure:$SADUGUID" -FailureGroup SourceObjectRemoval -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                        }
                    }
                    #endregion UpdateSourceObject
                    #region MoveTargetObject
                    #############################################################
                    #Move Target Object if Target Object was Source Object
                    #############################################################
                    if ($TargetOperation -eq 'SourceIsTarget:UpdateAndMigrateOnPremisesMailbox' -and $PSBoundParameters.ContainsKey('DestinationOU'))
                    {
                        $message = "Target User Object is the Source User Object.  Move to Destination OU: $DestinationOU"
                        try
                        {
                            Write-Log -Message $message -EntryType Attempting -Verbose
                            $domain = Get-AdObjectDomain -adobject $TADU -ErrorAction Stop
                            Move-ADObject -Server $domain -Identity $TADUGUID -TargetPath $DestinationOU -ErrorAction Stop
                            Write-Log -Message $message -EntryType Succeeded -Verbose
                        }
                        catch
                        {
                            Write-Log -Message $message -EntryType Failed -Verbose
                            Write-Log -Message $_.tostring() -ErrorLog
                        }
                    }
                    #endregion MoveTargetObject
                    #region RefreshTargetObjectRecipient
                    #############################################################
                    #Refresh Exchange recipient object for TADU
                    #############################################################
                    if ($UpdateTargetRecipient) {
                        $RecipientFound = $false
                        do {
                            Connect-Exchange -ExchangeOrganization $TargetExchangeOrganization > $null
                            $RecipientFound = Invoke-ExchangeCommand -cmdlet 'Get-Recipient' -ExchangeOrganization $TargetExchangeOrganization -string "-Identity $TADUGUID -ErrorAction SilentlyContinue" -ErrorAction SilentlyContinue
                            Start-Sleep -Seconds 1
                        }
                        until ($RecipientFound)
                        $UpdateRecipientFailedCount = 0
                        $RecipientUpdated = $false
                        do {
                        try {
                            $message = "Update-Recipient $TADUGUID in Exchange Organization $TargetExchangeOrganization"
                            Write-Log -message $message -EntryType Attempting
                            $Splat = @{
                                Identity = $TADUGUID
                                ErrorAction = 'Stop'
                            }
                            $ErrorActionPreference = 'Stop'
                            Connect-Exchange -ExchangeOrganization $TargetExchangeOrganization > $null
                            Invoke-ExchangeCommand -cmdlet 'Update-Recipient' -splat $Splat -ExchangeOrganization $TargetExchangeOrganization -ErrorAction Stop
                            Write-Log -message $message -EntryType Succeeded
                            $RecipientUpdated = $true
                            $ErrorActionPreference = 'Continue'
                        }
                        catch {
                            $UpdateRecipientFailedCount++
                            $ErrorActionPreference = 'Continue'
                            Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                            Write-Log -Message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity $TADUGUID -ExceptionCode "UpdateRecipientFailure:$TADUGUID" -FailureGroup UpdateRecipient -RelatedObjectIdentifier $TADUGUID -RelatedObjectIdentifierType ObjectGUID
                            Start-Sleep -Seconds 5
                        }
                        }
                        until ($RecipientUpdated -or $UpdateRecipientFailedCount -ge 3)
                    }
                    #endregion RefreshTargetObjectRecipient
                    Write-Output -InputObject $IntObj
                    }#else (when -not $TestOnly)
                }#foreach
            )#ProcessedObjects
        #endregion write
            if ($testOnly)
            {
                $RecordCount = $ProcessedObjects.Count
                Write-Log -Message "$recordcount Objects Processed for Test Only" -EntryType Notification -Verbose
                foreach ($intObj in $ProcessedObjects)
                {
                    $SADUGUID = $IntObj.SourceUserObjectGUID
                    $TADUGUID = $IntObj.TargetUserObjectGUID
                    Write-Log -Message "Processed Object SADU $SADUGUID and TADU $TADUGUID" -EntryType Notification -Verbose
                }
                Write-Output -InputObject $ProcessedObjects
            }
            else
            {
                $RecordCount = $ProcessedObjects.Count
                $cr = 0
                Write-Log -Message "$recordcount Objects Processed Locally" -EntryType Notification -Verbose
                if ($ProcessedObjects.Count -ge 1) {
                    #Start a Directory Synchronization to Azure AD Tenant
                    #Wait first for AD replication
                    Write-Log -Message "Waiting for $ADSyncDelayInSeconds seconds for AD Synchronization before starting an Azure AD Directory Synchronization." -Verbose -EntryType Notification
                    New-Timer -units Seconds -length $ADSyncDelayInSeconds -showprogress -Frequency 5 -voice
                    #Write-Log -Message "Starting an Azure AD Directory Synchronization." -Verbose -EntryType Notification
                    #Start-DirectorySynchronization
                    #Build Properties for CSV Output
                }
                foreach ($IntObj in $ProcessedObjects) {
                    $cr++
                    $writeProgressParams =
                    @{
                        Activity = "Performing Post-Attribute/Object Update Operations"
                        CurrentOperation = "Processing Object $($IntObj.DesiredUPNAndPrimarySMTPAddress)"
                        Status = "Processing Record $cr of $recordcount"
                        PercentComplete = $cr/$RecordCount*100
                    }#writeProgressParams
                    Write-Progress @writeProgressParams
                    $SADUGUID = $IntObj.SourceUserObjectGUID
                    $TADUGUID = $IntObj.TargetUserObjectGUID
                    $TADU = Find-ADUser -Identity $TADUGUID -IdentityType ObjectGUID -ActiveDirectoryInstance $TargetAD
                    $PropertySet = Get-CSVExportPropertySet -Delimiter '|' -MultiValuedAttributes $MultiValuedADAttributesToRetrieve -ScalarAttributes $ScalarADAttributesToRetrieve -SuppressCommonADProperties
                    $Global:SEATO_FullProcessedUsers += $TADU | Select-Object -Property $PropertySet -ExcludeProperty msExchPoliciesExcluded
                    #region WaitforDirectorySynchronization
                    #############################################################
                    #Request Directory Synchronization and Wait for Completion to Set Forwarding
                    #############################################################
                    if ($perUserDirSyncTest)
                    {
                        $TestDirectorySynchronizationParams = @{
                            Identity = $IntObj.DesiredUPNAndPrimarySMTPAddress
                            MaxSyncWaitMinutes = 5
                            DeltaSyncExpectedMinutes = 2
                            SyncCheckInterval = 15
                            ExchangeOrganization = 'OL'
                            RecipientAttributeToCheck = 'CustomAttribute5'
                            RecipientAttributeValue = $SADUGUID
                            InitiateSynchronization = $true
                        }
                        $DirSyncTest = Test-DirectorySynchronization @TestDirectorySynchronizationParams
                        #endregion WaitforDirectorySynchronization
                        #region SetMailboxForwarding
                        if ($DirSyncTest) {
                            switch -Wildcard ($IntObj.TargetOperation)
                            {
                                'EnableRemoteMailbox'
                                {
                                    if ($postCutover)
                                    {
                                        #don't forward, cutover already happened
                                    }
                                    else
                                    {
                                        try {
                                            $message = "Set Exchange Online Mailbox $($IntObj.DesiredUPNAndPrimarySMTPAddress) for forwarding to $($IntObj.DesiredCoexistenceRoutingAddress)."
                                            Connect-Exchange -ExchangeOrganization OL
                                            $ErrorActionPreference = 'Stop'
                                            Write-Log -message $message -EntryType Attempting
                                            Invoke-ExchangeCommand -cmdlet 'Set-Mailbox' -ExchangeOrganization OL -string "-Identity $($IntObj.DesiredUPNAndPrimarySMTPAddress) -ForwardingSmtpAddress $($IntObj.DesiredCoexistenceRoutingAddress)" -ErrorAction Stop
                                            Write-Log -message $message -EntryType Succeeded
                                            $ErrorActionPreference = 'Continue'
                                            $SetMailboxForwardingStatus = $true
                                        }
                                        catch {
                                            Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                                            Write-Log -Message $_.tostring() -ErrorLog
                                            Export-FailureRecord -Identity $($IntObj.DesiredUPNAndPrimarySMTPAddress) -ExceptionCode "SetCoexistenceForwardingFailure:$($IntObj.DesiredUPNAndPrimarySMTPAddress)" -FailureGroup SetCoexistenceForwarding
                                            $SetMailboxForwardingStatus = $false
                                            $ErrorActionPreference = 'Continue'
                                        }
                                    }#else
                                }#'EnableRemoteMailbox'
                                '*UpdateAndMigrateOnPremisesMailbox'
                                {
                                    $SourceDataProperties = @(
                                        @{
                                            name='SourceSystem'
                                            expression={$TargetExchangeOrganization}
                                        }
                                        @{
                                            name='Alias'
                                            expression={$_.DesiredAlias}
                                        }
                                        @{
                                            name='Wave'
                                            expression = {$MoveRequestWaveBatchName}
                                        }
                                        @{
                                            name='UserPrincipalName'
                                            expression = {$_.DesiredUPNAndPrimarySMTPAddress}
                                        }
                                    )
                                    try
                                    {
                                        $message = "Create Move Request for $TADUGUID"
                                        Write-Log -Message $message -EntryType Attempting
                                        $MRSourceData = @($IntObj | Select-Object $SourceDataProperties)
                                        $MR = @(New-MRMMoveRequest -SourceData $MRSourceData -wave $MoveRequestWaveBatchName -wavetype Sub -SuspendWhenReadyToComplete $true -ExchangeOrganization OL -LargeItemLimit 50 -BadItemLimit 50 -ErrorAction Stop)
                                        if ($MR.Count -eq 1)
                                        {
                                            Write-Log -Message $message -EntryType Succeeded
                                        } else {
                                            Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                                            #Write-Log -Message $_.tostring() -ErrorLog
                                            Export-FailureRecord -Identity $($IntObj.DesiredUPNAndPrimarySMTPAddress) -ExceptionCode "CreateMoveRequestFailure" -FailureGroup MailboxMove -ExceptionDetails $_.tostring()
                                        }

                                    }
                                    catch
                                    {
                                        Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                                        Write-Log -Message $_.tostring() -ErrorLog
                                        Export-FailureRecord -Identity $($IntObj.DesiredUPNAndPrimarySMTPAddress) -ExceptionCode "CreateMoveRequestFailure" -FailureGroup MailboxMove -ExceptionDetails $_.tostring()
                                    }
                                }#'UpdateAndMigrateOnPremisesMailbox'
                            }#switch
                        }
                        else {
                            $message = "Sync Related Failure for $($IntObj.DesiredUPNAndPrimarySMTPAddress)."
                            Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                            Export-FailureRecord -Identity $($IntObj.DesiredUPNAndPrimarySMTPAddress) -ExceptionCode "Synchronization:$($IntObj.DesiredUPNAndPrimarySMTPAddress)" -FailureGroup Synchronization
                        }
                    }
                    if ($SetMailboxForwardingStatus -and $IntObj.TargetOperation -eq 'EnableRemoteMailbox') {
                        $OLMailbox = Invoke-ExchangeCommand -cmdlet 'Get-Mailbox' -ExchangeOrganization OL -string "-Identity $($IntObj.DesiredUPNAndPrimarySMTPAddress)"
                        $propertyset = Get-CSVExportPropertySet -Delimiter '|' -MultiValuedAttributes EmailAddresses -ScalarAttributes PrimarySMTPAddress,ForwardingSmtpAddress
                        $OLMailboxSummary = $OLMailbox | Select-Object -Property $PropertySet
                        $Global:SEATO_OLMailboxSummary += $OLMailboxSummary
                    }
                    #endregion SetMailboxForwarding
                    #############################################################
                    #Processing Complete: Report Results
                    #############################################################
                    $ProcessedUserSummary = $TADU | Select-Object -Property SAMAccountName,DistinguishedName,MailNickName,UserPrincipalname,@{n='OriginalPrimarySMTPAddress';e={$IntObj.SourceUserMail}},@{n='CoexistenceForwardingAddress';e={$IntObj.DesiredCoexistenceRoutingAddress}},@{n='ObjectGUID';e={$_.ObjectGUID.GUID}},@{n='TargetOperation';e={$intobj.TargetOperation}},@{n='TimeStamp';e={Get-TimeStamp}}
                    $Global:SEATO_ProcessedUsers += $ProcessedUserSummary
                    Write-Log -Message "NOTE: Processing for $($TADU.UserPrincipalName) with GUID $TADUGUID in $TargetAD has completed successfully." -Verbose
                }#foreach IntObj in ProcessedObjects
                $writeProgressParams.currentOperation = "Completed Post Attribute/Object Update Operations"
                Write-Progress @writeProgressParams -Completed
                #region ReportAllResults
                if ($Global:SEATO_ProcessedUsers.count -ge 1) {
                    Write-Log -Message "Successfully Processed $($Global:SEATO_ProcessedUsers.count) Users."
                    Export-Data -DataToExportTitle TargetForestProcessedUsers -DataToExport $Global:SEATO_ProcessedUsers -DataType csv #-Append
                    Export-Data -DataToExportTitle TargetForestFullProcessedUsers -DataToExport $Global:SEATO_FullProcessedUsers -DataType csv
                }
                if ($Global:SEATO_Exceptions.count -ge 1) {
                    Write-Log -Message "Processed $($Global:SEATO_Exceptions.count) Users with Exceptions."
                }
                if ($Global:SEATO_MailContactsFound.count -ge 1) {
                    Write-Log -Message "$($Global:SEATO_MailContactsFound.count) Contacts were found and are being exported."
                    Export-Data -DataToExportTitle FoundMailContacts -DataToExport $Global:SEATO_MailContactsFound -Depth 2 -DataType xml
                }
                if ($Global:SEATO_OriginalTargetUsers.count -ge 1) {
                    Write-Log -Message "$($Global:SEATO_OriginalTargetUsers.count) Original Target Users were attempted for processing and are being exported."
                    Export-Data -DataToExportTitle OriginalTargetUsers -DataToExport $Global:SEATO_OriginalTargetUsers -Depth 2 -DataType xml
                }
                if ($Global:SEATO_MailContactDeletionFailures.Count -ge 1) {
                    Write-Log -Message "$($Global:SEATO_MailContactDeletionFailures.Count) Mail Contact(s) NOT successfully deleted.  Exporting them for review."
                    Export-Data -DataToExportTitle MailContactsNOTDeleted -DataToExport $Global:SEATO_MailContactDeletionFailures -DataType csv
                }
                if ($Global:SEATO_OLMailboxSummary.count -ge 1) {
                    Write-Log -Message "$($Global:SEATO_OLMailboxSummary.Count) Online Mailboxes Configured for Forwarding.  Exporting summary details for review."
                    Export-Data -DataToExportTitle OnlineMailboxForwarding -DataToExport $Global:SEATO_OLMailboxSummary -DataType csv
                }
                #endregion ReportAllResults
            }# else when NOT -TestOnly
        }#end end
    }
#end function Set-ExchangeAttributesOnTargetObject
function Add-EmailAddress
    {
        [cmdletbinding()]
        param
        (
        [string]$Identity
        ,
        [string[]]$EmailAddresses
        ,
        [string]$ExchangeOrganization
        )
        #Get the Recipient Object for the specified Identity
        try
        {
            $message = "Get Recipient for Identity $Identity"
            #Write-Log -Message $message -EntryType Attempting -Verbose
            $Splat = @{
                Identity = $Identity
                ErrorAction = 'Stop'
            }
            $Recipient = Invoke-ExchangeCommand -cmdlet Get-Recipient -splat $Splat -ErrorAction Stop -ExchangeOrganization $ExchangeOrganization
            #Write-Log -Message $message -EntryType Succeeded -Verbose
        }
        catch
        {
            Write-Log -Message $message -EntryType Failed -Verbose -ErrorLog
            Write-Log -Message $_.tostring() -ErrorLog
            Return
        }
        #Determine the Set cmdlet to use based on the Recipient Object
        $cmdlet = Get-RecipientCmdlet -Recipient $Recipient -verb Set -ErrorAction Stop
        try
        {
            $message = "Add Email Address $($EmailAddresses -join ',') to recipient $Identity"
            Write-Log -Message $message -EntryType Attempting -Verbose
            $splat = @{
                Identity = $Identity
                EmailAddresses = @{Add = $EmailAddresses}
                ErrorAction = 'Stop'
            }
            Invoke-ExchangeCommand -cmdlet $cmdlet -splat $splat -ExchangeOrganization $ExchangeOrganization -ErrorAction Stop
            Write-Log -Message $message -EntryType Succeeded -Verbose
        }
        catch
        {
            Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
            Write-Log -Message $_.tostring() -ErrorLog
        }
    }
#end function Add-EmailAddress
function Remove-EmailAddress
    {
        [cmdletbinding()]
        param
        (
        [string]$Identity
        ,
        [string[]]$EmailAddresses
        ,
        [string]$ExchangeOrganization
        )
        #Get the Recipient Object for the specified Identity
        try
        {
            $message = "Get Recipient for Identity $Identity"
            #Write-Log -Message $message -EntryType Attempting -Verbose
            $Splat = @{
                Identity = $Identity
                ErrorAction = 'Stop'
            }
            $Recipient = Invoke-ExchangeCommand -cmdlet Get-Recipient -splat $Splat -ErrorAction Stop -ExchangeOrganization $ExchangeOrganization
            #Write-Log -Message $message -EntryType Succeeded -Verbose
        }
        catch
        {
            Write-Log -Message $message -EntryType Failed -Verbose -ErrorLog
            Write-Log -Message $_.tostring() -ErrorLog
            Return
        }
        #Determine the Set cmdlet to use based on the Recipient Object
        $cmdlet = Get-RecipientCmdlet -Recipient $Recipient -verb Set -ErrorAction Stop
        try
        {
            $message = "Remove Email Address $($EmailAddresses -join ',') from recipient $Identity"
            Write-Log -Message $message -EntryType Attempting -Verbose
            $splat = @{
                Identity = $Identity
                EmailAddresses = @{Remove = $EmailAddresses}
                ErrorAction = 'Stop'
            }
            Invoke-ExchangeCommand -cmdlet $cmdlet -splat $splat -ExchangeOrganization $ExchangeOrganization -ErrorAction Stop
            Write-Log -Message $message -EntryType Succeeded -Verbose
        }
        catch
        {
            Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
            Write-Log -Message $_.tostring() -ErrorLog
        }
    }
#end function Remove-EmailAddress
function Reset-AzureADUserPrincipalName
    {
        [cmdletbinding()]
        param
        (
            [Parameter(Mandatory,ParameterSetName='ExistingUPN')]
            [string]$UserPrincipalName
            ,
            [Parameter(Mandatory,ParameterSetName='ObjectID')]
            [string]$ObjectID
            ,
            [parameter(Mandatory)]
            [string]$TenantName
            ,
            [Parameter(Mandatory)]
            [string]$DesiredUserPrincipalName
            ,
            [switch]$Verify
        )
        if ((Connect-MSOnlineTenant -Tenant $TenantName) -ne $true)
        {throw {"Could not connect to MSOnline Tenant $($TenantName)"}}
        $message = "Get Azure AD User using $($PSCmdlet.ParameterSetName) "
        $splat = @{
            ErrorAction = 'Stop'
        }#splat
        switch ($PSCmdlet.ParameterSetName)
        {
            'ExistingUPN'
            {
                $message = $message + $UserPrincipalName
                $splat.UserPrincipalName = $UserPrincipalName
            }#ExistingUPN
            'ObjectID'
            {
                $message = $message + $ObjectID
                $splat.ObjectID = $ObjectID
            }#objectID
        }#switch
        try
        {
            Write-Log -Message $message -EntryType Attempting
            $OriginalAzureADUser = Get-MsolUser @splat
            Write-Log -Message $message -EntryType Succeeded
        }#try
        catch
        {
            $myerror = $_
            Write-Log -Message $message -EntryType Failed -ErrorLog
            Write-Log -Message $myerror.tostring() -ErrorLog
            throw {$myerror}
        }#catch
        $message = "Get Tenant domain to use for temporary UPN value"
        try
        {
            Write-Log -Message $message -EntryType Attempting
            $TenantDomain = Get-MsolDomain -ErrorAction Stop | Where-Object -FilterScript {$_.Name -like '*.onmicrosoft.com' -and $_.name -notlike '*.mail.onmicrosoft.com'} | Select-Object -ExpandProperty Name
            Write-Log -Message $message -EntryType Succeeded
        }#try
        catch
        {
            $myerror = $_
            Write-Log -Message $message -EntryType Failed -ErrorLog
            Write-Log -Message $myerror.tostring() -ErrorLog
            throw {$myerror}
        }#catch
        $temporaryUPN = $OriginalAzureADUser.ObjectID.guid + '@' + $TenantDomain
        $message = "Set Azure AD User $($OriginalAzureADUser.ObjectID.guid) UserPrincipalName to temporary value $temporaryUPN"
        $splat = @{
            ObjectID = $OriginalAzureADUser.objectID.guid
            NewUserPrincipalName = $temporaryUPN
            ErrorAction = 'Stop'
        }#splat
        try
        {
            Write-Log -Message $message -EntryType Attempting
            Set-MsolUserPrincipalName @splat | Out-Null #temporary password output thrown away
            Write-Log -Message $message -EntryType Succeeded
        }#try
        catch
        {
            $myerror = $_
            Write-Log -Message $message -EntryType Failed -ErrorLog
            Write-Log -Message $myerror.tostring() -ErrorLog
            throw {$myerror}
        }#catch
        $message = "Set Azure AD User $($OriginalAzureADUser.ObjectID.guid) UserPrincipalName to Desired value $DesiredUserPrincipalName"
        $splat = @{
            ObjectID = $OriginalAzureADUser.objectID.guid
            NewUserPrincipalName = $DesiredUserPrincipalName
            ErrorAction = 'Stop'
        }#splat
        try
        {
            Write-Log -Message $message -EntryType Attempting
            Set-MsolUserPrincipalName @splat
            Write-Log -Message $message -EntryType Succeeded
        }#try
        catch
        {
            $myerror = $_
            Write-Log -Message $message -EntryType Failed -ErrorLog
            Write-Log -Message $myerror.tostring() -ErrorLog
            throw {$myerror}
        }#catch
        if ($PSBoundParameters.ContainsKey('Verify'))
        {
            $splat = @{
                ObjectID = $OriginalAzureADUser.objectID.guid
                ErrorAction = 'Stop'
            }#splat
            Get-MsolUser @splat
        }
    }
#end function Reset-AzureADUserPrincipalName
function Get-AllADRecipientObjects
    {
        [cmdletbinding()]
        param
        (
            [Parameter()]
            [AllowNull()]
            [int]$ResultSetSize = $null
            ,
            [switch]$Passthrough
            ,
            [switch]$ExportData
        )
        $ADUserAttributes = Get-OneShellVariableValue -Name ADUserAttributes
        $ADGroupAttributesWMembership = Get-OneShellVariableValue -Name ADGroupAttributesWMembership
        $ADContactAttributes = Get-OneShellVariableValue -Name ADContactAttributes
        $ADPublicFolderAttributes = Get-OneShellVariableValue -Name ADPublicFolderAttributes
        $AllGroups = Get-ADGroup -ResultSetSize $ResultSetSize -Properties $ADGroupAttributesWMembership -Filter * | Select-Object -Property * -ExcludeProperty Property*,Item
        $AllMailEnabledGroups = $AllGroups | Where-Object -FilterScript {$_.legacyExchangeDN -ne $NULL -or $_.mailNickname -ne $NULL -or $_.proxyAddresses -ne $NULL}
        $AllContacts = Get-ADObject -Filter {objectclass -eq 'contact'} -Properties $ADContactAttributes -ResultSetSize $ResultSetSize | Select-Object -Property * -ExcludeProperty Property*,Item
        $AllMailEnabledContacts = $AllContacts | Where-Object -FilterScript {$_.legacyExchangeDN -ne $NULL -or $_.mailNickname -ne $NULL -or $_.proxyAddresses -ne $NULL}
        $AllUsers = Get-ADUser -ResultSetSize $ResultSetSize -Filter * -Properties $ADUserAttributes | Select-Object -Property * -ExcludeProperty Property*,Item
        $AllMailEnabledUsers = $AllUsers  | Where-Object -FilterScript {$_.legacyExchangeDN -ne $NULL -or $_.mailNickname -ne $NULL -or $_.proxyAddresses -ne $NULL}
        $AllPublicFolders = Get-ADObject -Filter {objectclass -eq 'publicFolder'} -ResultSetSize $ResultSetSize -Properties $ADPublicFolderAttributes | Select-Object -Property * -ExcludeProperty Property*,Item
        $AllMailEnabledPublicFolders = $AllPublicFolders  | Where-Object -FilterScript {$_.legacyExchangeDN -ne $NULL -or $_.mailNickname -ne $NULL -or $_.proxyAddresses -ne $NULL}
        $AllMailEnabledADObjects = $AllMailEnabledGroups + $AllMailEnabledContacts + $AllMailEnabledUsers + $AllMailEnabledPublicFolders
        if ($Passthrough) {$AllMailEnabledADObjects}
        if ($ExportData) {Export-Data -DataToExport $AllMailEnabledADObjects -DataToExportTitle 'AllADRecipientObjects' -Depth 3 -DataType xml}
    }
#end function Get-AllADRecipientObjects
function Get-ADRecipientsWithConflictingProxyAddresses
    {
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
                    Write-Log -Message "No Conflict for $pa2c" -EntryType Notification -Verbose
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
#end function Get-AdRecipientWithConflictingProxyAddress
function Get-ADRecipientsWithConflictingAlias
    {
        [cmdletbinding()]
        param
        (
        [parameter(Mandatory=$true)]
        $SourceRecipients
        ,
        [parameter(Mandatory=$true)]
        $TargetExchangeOrganization
        ,
        [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
        [string]$ReplacementPrefix
        ,
        [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
        [string]$SourcePrefix
        )
        foreach ($sr in $SourceRecipients)
        {
            $Alias = $sr.mailNickName
            $Alias = $Alias -replace '\s|[^1-9a-zA-Z_-]',''
            if ($PSCmdlet.ParameterSetName -eq 'ReplacePrefix')
            {
                $NewAlias = $Alias -replace "\b$($sourcePrefix)_",''
                $NewAlias = $NewAlias -replace "\b$($SourcePrefix)", ''
                $NewAlias = $NewAlias -replace "$($SourcePrefix)\b", ''
                $NewAlias = "$($ReplacementPrefix)_$($NewAlias)"
                $Alias = $NewAlias
            }
            if (Test-ExchangeAlias -Alias $Alias -ExchangeOrganization $TargetExchangeOrganization)
            {
                Write-Log -Message "No Conflict for $Alias" -EntryType Notification -Verbose
            }
            else
            {
                $conflicts = @(Test-ExchangeAlias -Alias $Alias -ExchangeOrganization $TargetExchangeOrganization -ReturnConflicts)
                [pscustomobject]@{
                    SourceObjectGUID = $sr.ObjectGUID
                    ConflictingTargetObjectGUIDs = $conflicts
                    OriginalAlias = $sr.mailNickName
                    TestedAlias = $Alias
                }
            }

        }
    }
#end function Get-ADRecipientsWithConflictingAlias
function New-SourceTargetRecipientMap
    {
        [cmdletbinding()]
        param
        (
        $SourceRecipients
        ,
        $TargetExchangeOrganization
        )
        $SourceTargetRecipientMap = @{}
        $TargetSourceRecipientMap = @{}
        foreach ($SR in $SourceRecipients)
        {
            $ProxyAddressesToCheck = $sr.proxyaddresses | Where-Object -FilterScript {$_ -ilike 'smtp:*'}
            $rawrecipientmatches =
            @(
                foreach ($pa2c in $ProxyAddressesToCheck)
                {
                    if (Test-ExchangeProxyAddress -ProxyAddress $pa2c -ProxyAddressType SMTP -ExchangeOrganization $TargetExchangeOrganization)
                    {$null}
                    else
                    {
                        Test-ExchangeProxyAddress -ProxyAddress $pa2c -ProxyAddressType SMTP -ExchangeOrganization $TargetExchangeOrganization -ReturnConflicts
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
#end function New-SourceTargetRecipientMap

function Get-TargetRecipientFromMap
    {
        [cmdletbinding()]
        param
        (
            $SourceObjectGUID
            ,
            $TargetExchangeOrganization
        )
        $TargetRecipientGUID = @($RecipientMaps.SourceTargetRecipientMap.$SourceObjectGUID)
        if ([string]::IsNullOrWhiteSpace($TargetRecipientGUID))
        {$null}
        else
        {
            $TargetRecipients =
            @(
                foreach ($id in $TargetRecipientGUID)
                {
                    $cmdlet = Get-RecipientCmdlet -Identity $id -verb Get -ExchangeOrganization $TargetExchangeOrganization
                    Invoke-ExchangeCommand -cmdlet $cmdlet -string "-Identity $id" -ExchangeOrganization $TargetExchangeOrganization -ErrorAction Stop
                }
            )
            $TargetRecipients
        }
    }
#end function Get-TargetRecipientFromMap
function New-SourceRecipientDNHash
    {
        [cmdletbinding()]
        param(
        [parameter(Mandatory=$true)]
        $SourceRecipients
        )
        $SourceRecipientsDNHash = @{}
        foreach ($recip in $SourceRecipients)
        {
            $SourceRecipientsDNHash.$($recip.DistinguishedName)=$recip
        }
        $SourceRecipientsDNHash
    }
#end function New-SourceRecipientDNHash
function New-NestingOrderedGroupArray
    {
        [cmdletbinding()]
        param(
        $Groups
        )
        $GroupsDNHash = @{}
        $groups | Select-Object -ExpandProperty DistinguishedName | ForEach-Object {$GroupsDNHash.$($_) = $true}
        $OutputGroups = @{}
        $NestingLevel = 0
        Do {
            foreach ($group in $Groups)
            {
                if ($NestingLevel -eq 0 -and $group.memberof.Count -eq 0)
                {
                    #these groups have no memberships in other groups and can only be containting groups so we create/populate them last
                    $Group | Add-Member -MemberType NoteProperty -Name NestingLevel -Value $NestingLevel
                    $OutputGroups.$($Group.DistinguishedName) = $Group
                    Write-Verbose -Message "added Group $($Group.DistinguishedName) to Output at Nesting Level $NestingLevel"
                }
                elseif ($NestingLevel -ge 1 -and $Group.memberof.Count -ge 1 -and (-not $OutputGroups.ContainsKey($($Group.DistinguishedName))))
                {
                    $testGroupMemberships = @{}
                    foreach ($membership in $group.memberof)
                    {
                        #if the member of is not in the Groups array then we ignore it
                        if ($GroupsDNHash.ContainsKey($membership))
                        {
                            #if the member of is in the groups array then we make sure that the group we would be a member of is created and populated after the member group
                            $testGroupMemberships.$($membership) = ($OutputGroups.ContainsKey($membership) -and $OutputGroups.$($membership).NestingLevel -lt $NestingLevel)
                        }
                    }
                    if ($testGroupMemberships.ContainsValue($false))
                    #do nothing yet - wait until no $false values appear
                    {} else
                    {
                        #add the group to the output at the current nesting level
                        $Group | Add-Member -MemberType NoteProperty -Name NestingLevel -Value $NestingLevel
                        $OutputGroups.$($Group.DistinguishedName) = $Group
                        Write-Verbose -Message "added Group $($Group.DistinguishedName) to Output at Nesting Level $NestingLevel"
                    }
                }
            }
            if ($OutputGroups.Keys.Count -eq $Groups.count)
            {
                Write-Verbose -Message "No More Nests Required"
                $NoMoreNests = $true
            }
            $NestingLevel++
        }
        Until
        ($NoMoreNests)
        $OrderedGroups = $OutputGroups.Values | Sort-Object -Property NestingLevel -Descending
        Write-Output $OrderedGroups
    }
#end function New-NestingOrderedGroupArray
function Publish-Groups
    {
        [cmdletbinding()]
        param
        (
        $TargetExchangeOrganization
        ,
        $TargetGroupOU
        ,
        $TargetContactOU
        ,
        $TargetSMTPDomain
        ,
        $SourcePrefix
        ,
        $ReplacementPrefix
        ,
        $SourceGroups
        ,
        $SourceRecipients
        ,
        [switch]$TestOnly
        ,
        [switch]$RefreshRecipientMaps
        ,
        [switch]$HideContacts
        )
        if (-not (Test-Path variable:\IntermediateGroupObjects)) {
            New-Variable -Name IntermediateGroupObjects -Value @() -Scope Global
        }
        $csgCount = 0
        $sgCount = $SourceGroups.Count
        $stopwatch = [system.diagnostics.stopwatch]::startNew()
        foreach ($sg in $SourceGroups)
        {
            $csgCount++
            Write-Log -Message "Processing Source Group $($sg.mailnickname)" -EntryType Notification
        #region Prepare
            $desiredAlias = Get-DesiredTargetAlias -SourceAlias $sg.mailNickName -TargetExchangeOrganization $TargetExchangeOrganization -ReplacementPrefix $ReplacementPrefix -SourcePrefix $SourcePrefix
            Write-Log -Message "Processing Source Group $($sg.mailnickname). Target Group alias will be $desiredAlias." -EntryType Notification
            $WriteProgressParams =
            @{
                Activity = "Provisioning $($SourceGroups.count) Groups into $TargetExchangeOrganization, $TargetGroupOU"
                Status = "Working $csgCount of $($SourceGroups.count)"
                CurrentOperation = $desiredAlias
                PercentComplete = $csgCount/$sgCount*100
            }
            if ($csgCount -gt 1){$WriteProgressParams.SecondsRemaining = ($($stopwatch.Elapsed.TotalSeconds.ToInt32($null))/($csgCount - 1)) * ($sgCount - ($csgCount - 1))}
            Write-Progress @WriteProgressParams
            $desiredPrimarySMTPAddress = Get-DesiredTargetPrimarySMTPAddress -DesiredAlias $desiredAlias -TargetExchangeOrganization $TargetExchangeOrganization -TargetSMTPDomain $TargetSMTPDomain
            $desiredName = Get-DesiredTargetName -SourceName $sg.DisplayName -TargetExchangeOrganization $TargetExchangeOrganization -ReplacementPrefix $ReplacementPrefix -SourcePrefix $SourcePrefix
            $targetRecipientGUIDs = @($RecipientMaps.SourceTargetRecipientMap.$($sg.ObjectGUID.Guid))
            $targetRecipients = Get-TargetRecipientFromMap -SourceObjectGUID $($sg.ObjectGUID.Guid) -TargetExchangeOrganization $TargetExchangeOrganization
            $GetDesiredProxyAddressesParams = @{
                CurrentProxyAddresses = $sg.proxyAddresses
                DesiredPrimaryAddress = $desiredPrimarySMTPAddress
                DesiredOrCurrentAlias = $desiredAlias
                Recipients = $targetRecipients
                LegacyExchangeDNs = $targetRecipients | Select-Object -ExpandProperty LegacyExchangeDN
            }
            $DesiredProxyAddresses = Get-DesiredProxyAddresses @GetDesiredProxyAddressesParams
        #endregion Prepare
        #region GetAndMapGroupMembers
            $AllSourceMembers =@($sg.Members | foreach {if ($SourceRecipientDNHash.ContainsKey($_)) {$SourceRecipientDNHash.$($_)}})
            $AllSourceUserMembers = @($AllSourceMembers | ? ObjectClass -eq 'User')
            $AllSourceGroupMembers =@($AllSourceMembers | ? ObjectClass -eq 'Group')
            $AllSourceContactMembers = @($AllSourceMembers | ? ObjectClass -eq 'Contact')
            $AllSourcePublicFolderMembers = @($AllSourceMembers | ? ObjectClass -eq 'publicFolder')
            $mappedTargetMemberUsers = @($AllSourceUserMembers | Select-Object @{n='GUIDString';e={$_.ObjectGUID.guid}} | Where-Object {$RecipientMaps.SourceTargetRecipientMap.ContainsKey($_.GUIDString)} | foreach {$RecipientMaps.SourceTargetRecipientMap.$($_.GUIDString) | Where-Object {$_ -ne $null}})
            $mappedTargetMemberContacts = @($AllSourceContactMembers | Select-Object @{n='GUIDString';e={$_.ObjectGUID.guid}} | Where-Object {$RecipientMaps.SourceTargetRecipientMap.ContainsKey($_.GUIDString)} | foreach {$RecipientMaps.SourceTargetRecipientMap.$($_.GUIDString) | Where-Object {$_ -ne $null}})
            $mappedTargetMemberGroups = @($AllSourceGroupMembers | Select-Object @{n='GUIDString';e={$_.ObjectGUID.guid}} | Where-Object {$RecipientMaps.SourceTargetRecipientMap.ContainsKey($_.GUIDString)} | foreach {$RecipientMaps.SourceTargetRecipientMap.$($_.GUIDString) | Where-Object {$_ -ne $null}})
            $AllMappedMembersToAddAtCreation = @($mappedTargetMemberUsers + $mappedTargetMemberContacts + $mappedTargetMemberGroups)
            $nonMappedTargetMemberGroups = @($AllSourceGroupMembers | Where-Object {$RecipientMaps.SourceTargetRecipientMap.$($_.ObjectGUID.guid) -eq $null})
            $nonMappedTargetMemberUsers = @($AllSourceUserMembers | Where-Object {$RecipientMaps.SourceTargetRecipientMap.$($_.ObjectGUID.guid) -eq $null})
            $nonMappedTargetMemberContacts = @($AllSourceContactMembers | Where-Object {$RecipientMaps.SourceTargetRecipientMap.$($_.ObjectGUID.guid) -eq $null})
        #endregion GetAndMapGroupMembers
        #region IntermediateGroupObject
            $intermediateGroupObject =
            [pscustomobject]@{
                DesiredAlias = $desiredAlias
                DesiredName = $desiredName
                DesiredPrimarySMTPAddress = $desiredPrimarySMTPAddress
                DesiredProxyAddresses = $DesiredProxyAddresses
                TargetRecipientGUIDs = @($targetRecipientGUIDs)
                TargetRecipients = @($targetRecipients)
                MappedTargetMemberUsers = @($mappedTargetMemberUsers)
                MappedTargetMemberContacts = @($mappedTargetMemberContacts)
                MappedTargetMemberGroups = @($mappedTargetMemberGroups)
                AllMappedMembersToAddAtCreation = @($AllMappedMembersToAddAtCreation)
                NonMappedMemberUsers = @($nonMappedTargetMemberUsers | Select-Object -ExpandProperty DistinguishedName)
                NonMappedMemberContacts = @($nonMappedTargetMemberContacts | Select-Object -ExpandProperty DistinguishedName)
                NonMappedMemberGroups = @($nonMappedTargetMemberGroups | Select-Object -ExpandProperty DistinguishedName)
                SourcePublicFolderMembers = @($AllSourcePublicFolderMembers | Select-Object -ExpandProperty Mail)
                SourceObject = $sg
            }
            $Global:intermediateGroupObjects += $intermediateGroupObject
            Export-Data -DataToExportTitle $("Group-" + $DesiredAlias) -DataToExport $intermediateGroupObject -Depth 3 -DataType json
        #endregion IntermediateGroupObject
            if ($TestOnly)
            {
                $intermediateGroupObject
            }
        #region RemoveTargetRecipients
            else {
                foreach ($tr in $targetRecipients)
                {
                    $message = "Remove target recipient $($tr.Alias) for Group $DesiredAlias"
                    Connect-Exchange -ExchangeOrganization $TargetExchangeOrganization
                    $cmdlet = Get-RecipientCmdlet -Recipient $tr -verb Remove
                    $rrParams =
                    @{
                        Identity = $($tr.Guid.guid)
                        Confirm = $false
                        ErrorAction = 'Stop'
                    }
                    try
                    {
                        Write-Log -Message $message -EntryType Attempting
                        Invoke-ExchangeCommand -cmdlet $cmdlet -splat $rrParams -ExchangeOrganization $TargetExchangeOrganization -ErrorAction Stop
                        Write-Log -Message $message -EntryType Succeeded
                    }
                    catch
                    {
                        Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                        Write-Log -Message $_.tostring() -ErrorLog -Verbose
                    }
                }
            }
        #endregion RemoveTargetRecipients
            #region CreateNeededContacts
            foreach ($nmc in $nonMappedTargetMemberContacts)
            {
                try {
                    $ContactDesiredName = Get-DesiredTargetName -SourceName $nmc.DisplayName -TargetExchangeOrganization $TargetExchangeOrganization -ReplacementPrefix $ReplacementPrefix -SourcePrefix $SourcePrefix
                    $ContactDesiredAlias = Get-DesiredTargetAlias -SourceAlias $nmc.MailNickName -TargetExchangeOrganization $TargetExchangeOrganization -ReplacementPrefix $ReplacementPrefix -SourcePrefix $SourcePrefix
                    $ContactDesiredProxyAddresses = Get-DesiredProxyAddresses -CurrentProxyAddresses $nmc.proxyAddresses -DesiredOrCurrentAlias $ContactDesiredAlias -LegacyExchangeDNs $nmc.legacyExchangeDN
                }
                catch {
                    Export-Data -DataToExport $nmc -DataToExportTitle "ContactCreationFailure-$($nmc.MailNickName)" -DataType json -Depth 3
                    Continue
                }
                $intermediateContactObject =
                [pscustomobject]@{
                    DesiredAlias = $ContactDesiredAlias
                    DesiredName = $ContactDesiredName
                    TargetAddress = $nmc.targetAddress
                    DesiredProxyAddresses = $ContactDesiredProxyAddresses
                    SourceObject = $nmc
                }
                Export-Data -DataToExportTitle $("Contact-" + $ContactDesiredAlias) -DataToExport $intermediateContactObject -Depth 3 -DataType json
                if ($TestOnly)
                {
                }#if
                else
                {
                    $newMailContactParams =
                    @{
                        Name = $ContactDesiredName
                        DisplayName = $ContactDesiredName
                        ExternalEmailAddress = $nmc.targetAddress
                        Alias = $ContactDesiredAlias
                        OrganizationalUnit = $TargetContactOU
                        ErrorAction = 'Stop'
                    }
                    $setMailContactParams =
                    @{
                        Identity = $ContactDesiredAlias
                        EmailAddressPolicyEnabled = $false
                        EmailAddresses = $ContactDesiredProxyAddresses
                        ErrorAction = 'Stop'
                    }
                    if ($HideContacts)
                    {
                        $setMailContactParams.HiddenFromAddressListsEnabled = $true
                    }
                    $message = "Create Contact $ContactDesiredAlias for group $desiredAlias."
                    try
                    {
                        Write-Log -Message $message -EntryType Attempting
                        Connect-Exchange -ExchangeOrganization $TargetExchangeOrganization
                        $newContact = Invoke-ExchangeCommand -cmdlet 'New-MailContact' -ExchangeOrganization $TargetExchangeOrganization -splat $newMailContactParams
                        $mappedTargetMemberContacts += $newContact.guid.guid
                        $AllMappedMembersToAddAtCreation += $newContact.guid.guid
                        Write-Log -Message $message -EntryType Failed
                        $message = "Find Newly Created Contact $ContactDesiredAlias."
                        $found = $false
                        do
                        {
                            #Write-Log -Message $message -EntryType Attempting
                            $Contact = @(Invoke-ExchangeCommand -cmdlet 'Get-MailContact' -string "-Identity $ContactDesiredAlias" -ExchangeOrganization $TargetExchangeOrganization)
                            if ($Contact.Count -eq 1)
                            {
                                Write-Log -Message $message -EntryType Succeeded
                                $found = $true
                            }
                            Start-Sleep -Seconds 10
                        }
                        until
                        (
                            $found -eq $true
                        )
                        $message = "Set Newly Created Contact $ContactDesiredName Attributes"
                        Write-Log -Message $message -EntryType Attempting
                        Invoke-ExchangeCommand -cmdlet 'Set-MailContact' -splat $setMailContactParams -exchangeOrganization $TargetExchangeOrganization -ErrorAction Stop
                        Write-Log -Message $message -EntryType Succeeded
                        foreach ($pa in $ContactDesiredProxyAddresses) {
                            $type = $pa.split(':')[0]
                            if ($type -in 'SMTP','x500')
                            {
                                    Add-ExchangeProxyAddressToTestExchangeProxyAddress -ProxyAddress $pa -ProxyAddressType $type -ObjectGUID $newContact.guid.guid
                            }

                                }
                                Add-ExchangeAliasToTestExchangeAlias -Alias $ContactDesiredAlias -ObjectGUID $newContact.guid.guid
                    }
                    catch
                    {
                        Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                        Write-Log -Message $_.tostring() -ErrorLog -Verbose
                    }
                }#else
            }#foreach $NMC
            #endregion CreateNeededContacts
            #region ProvisionDistributionGroup
            if ($TestOnly)
            {}
            else
            {
            $AliasLength = [math]::Min($desiredAlias.length,20)
            $newDistributionGroupParams =
            @{
                DisplayName = $desiredName
                Name = $desiredName
                IgnoreNamingPolicy = $true
                Members = @($AllMappedMembersToAddAtCreation | Where-Object {$_ -ne $null})
                Type = 'Distribution'
                Alias = $desiredAlias
                SAMAccountName = $desiredAlias.substring(0,$AliasLength)#"$($ReplacementPrefix)_" + $sg.ObjectGUID.guid.Substring(24,12)
                PrimarySmtpAddress = $desiredPrimarySMTPAddress
                OrganizationalUnit = $TargetGroupOU
                ErrorAction = 'Stop'
            }
            $setDistributionGroupParams =
            @{
                Identity = $desiredAlias
                EmailAddresses = $DesiredProxyAddresses
                EmailAddressPolicyEnabled = $false
                errorAction = 'Stop'
            }
            try
            {
                $message = "Create Group $desiredAlias"
                Write-Log -Message $message -EntryType Attempting
                $newgroup = Invoke-ExchangeCommand -cmdlet 'New-DistributionGroup' -splat $newDistributionGroupParams -exchangeOrganization $TargetExchangeOrganization -ErrorAction Stop
                Write-Log -Message $message -EntryType Succeeded
                Start-Sleep -Seconds 1
                $message = "Find Newly Created Group $desiredAlias"
                $found = $false
                Do
                {
                    #Write-Log -Message $message -EntryType Attempting
                    $group = @(Invoke-ExchangeCommand -cmdlet 'Get-DistributionGroup' -string "-Identity $desiredAlias -ErrorAction SilentlyContinue" -ExchangeOrganization $TargetExchangeOrganization -ErrorAction SilentlyContinue)
                    if ($group.Count -eq 1)
                    {
                        Write-Log -Message $message -EntryType Succeeded
                        $found = $true
                    }
                    Start-Sleep -Seconds 1
                }
                Until
                ($found -eq $true)
                $message = "Set Group $desiredAlias Attributes"
                Write-Log -Message $message -EntryType Attempting
                Invoke-ExchangeCommand -cmdlet 'Set-DistributionGroup' -splat $setDistributionGroupParams -ExchangeOrganization $TargetExchangeOrganization
                Write-Log -Message $message -EntryType Succeeded
                foreach ($pa in $DesiredProxyAddresses) {
                    $type = $pa.split(':')[0]
                    if ($type -in 'SMTP','x500')
                    {
                        Add-ExchangeProxyAddressToTestExchangeProxyAddress -ProxyAddress $pa -ProxyAddressType $type -ObjectGUID $newgroup.guid.guid
                    }
                }
                Add-ExchangeAliasToTestExchangeAlias -Alias $desiredAlias -ObjectGUID $newgroup.guid.guid
                Write-Log -Message "Provisioning Complete for Group $desiredAlias." -EntryType Notification -Verbose
            }
            catch
            {
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $_.tostring() -ErrorLog -Verbose
            }
            #endregion ProvisionDistributionGroup
            }#else
        }#foreach
        if ($TestOnly)
        {}
        else
        {
            if ($RefreshRecipientMaps)
            {
                New-SourceTargetRecipientMap -SourceRecipients $sourceRecipients -TargetExchangeOrganization $TargetExchangeOrganization -OutVariable RecipientMaps
            }
        }#else
        #$Global:TheLocalVariables = Get-Variable -Scope Local
    }
#end function Publish-Groups
function Get-GroupMemberMapping
    {
        [cmdletbinding()]
        param
        (
            $sourcemembers
        )
        $AllSourceMembers =@($sourcemembers | foreach {if ($SourceRecipientDNHash.ContainsKey($_)) {$SourceRecipientDNHash.$($_)}})
        $AllSourceUserMembers = @($AllSourceMembers | ? ObjectClass -eq 'User')
        $AllSourceGroupMembers =@($AllSourceMembers | ? ObjectClass -eq 'Group')
        $AllSourceContactMembers = @($AllSourceMembers | ? ObjectClass -eq 'Contact')
        $AllSourcePublicFolderMembers = @($AllSourceMembers | ? ObjectClass -eq 'publicFolder')
        $mappedTargetMemberUsers = @($AllSourceUserMembers | Select-Object @{n='GUIDString';e={$_.ObjectGUID.guid}} | Where-Object {$RecipientMaps.SourceTargetRecipientMap.ContainsKey($_.GUIDString)} | foreach {$RecipientMaps.SourceTargetRecipientMap.$($_.GUIDString) | Where-Object {$_ -ne $null}})
        $mappedTargetMemberContacts = @($AllSourceContactMembers | Select-Object @{n='GUIDString';e={$_.ObjectGUID.guid}} | Where-Object {$RecipientMaps.SourceTargetRecipientMap.ContainsKey($_.GUIDString)} | foreach {$RecipientMaps.SourceTargetRecipientMap.$($_.GUIDString) | Where-Object {$_ -ne $null}})
        $mappedTargetMemberGroups = @($AllSourceGroupMembers | Select-Object @{n='GUIDString';e={$_.ObjectGUID.guid}} | Where-Object {$RecipientMaps.SourceTargetRecipientMap.ContainsKey($_.GUIDString)} | foreach {$RecipientMaps.SourceTargetRecipientMap.$($_.GUIDString) | Where-Object {$_ -ne $null}})
        $AllMappedMembersToAddAtCreation = @($mappedTargetMemberUsers + $mappedTargetMemberContacts + $mappedTargetMemberGroups)
        $nonMappedTargetMemberGroups = @($AllSourceGroupMembers | Where-Object {$RecipientMaps.SourceTargetRecipientMap.$($_.ObjectGUID.guid) -eq $null})
        $nonMappedTargetMemberUsers = @($AllSourceUserMembers | Where-Object {$RecipientMaps.SourceTargetRecipientMap.$($_.ObjectGUID.guid) -eq $null})
        $nonMappedTargetMemberContacts = @($AllSourceContactMembers | Where-Object {$RecipientMaps.SourceTargetRecipientMap.$($_.ObjectGUID.guid) -eq $null})
        $membershipMap = @{
            MappedTargetMemberUsers = $mappedTargetMemberUsers
            MappedTargetMemberContacts = $mappedTargetMemberContacts
            MappedTargetMemberGroups = $mappedTargetMemberGroups
            AllMappedTargetMembers = $AllMappedMembersToAddAtCreation
            NonMappedTargetMemberUsers = $nonMappedTargetMemberUsers
            NonMappedTargetMemberContacts = $nonMappedTargetMemberContacts
            NonMappedTargetMemberGroups = $nonMappedTargetMemberGroups
        }
        $membershipMap
    }
#end function Get-GroupMembermapping

Function New-MailFlowContactFromMailbox
    {
        [cmdletbinding()]
        param(
        [parameter(Mandatory)]
        [string]$Identity
        ,
        [parameter(Mandatory)]
        [string]$SourceExchangeOrganization
        ,
        [parameter(Mandatory)]
        [string]$TargetExchangeOrganization
        ,
        [parameter(Mandatory)]
        [string]$OrganizationalUnit
        ,
        [parameter()]
        [string[]]$DomainsToRemove
        ,
        [parameter()]
        [string[]]$AddressesToRemove
        ,
        [parameter()]
        [string[]]$AddressesToAdd
        ,
        [parameter()]
        [string]$TargetAddress
        ,
        [parameter()]
        [string]$TargetDeliveryDomain
        ,
        [string]$AliasPrefix
        ,
        [bool]$testOnly = $true
        ,
        [switch]$AddExternalEmailAddressToSourceObject
        )
        #region GetSourceObject
        $GetRecipientParams = @{
            cmdlet = 'Get-Recipient'
            ExchangeOrganization = $SourceExchangeOrganization
            ErrorAction = 'Stop'
            splat = @{
                Identity = $Identity
                ErrorAction = 'Stop'
            }
        }
        $message = "Get Recipient Object for Identity ($Identity) from Source Exchange Organization ($SourceExchangeOrganization)"
        Try
        {
            Write-Log -Message $message -EntryType Attempting
            $SourceRecipientObject = @(Invoke-ExchangeCommand @GetRecipientParams)
            Write-Log -Message $message -EntryType Succeeded
        }
        Catch
        {
            $MyError = $_
            Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
            Write-Log -Message $myerror.tostring() -ErrorLog
            Throw $MyError
        }
        $GetFullRecipientCmdlet = Get-RecipientCmdlet -Recipient $SourceRecipientObject -verb Get -ErrorAction Stop
        $GetFullRecipientCmdletParams = @{
            cmdlet = $GetFullRecipientCmdlet
            ExchangeOrganization = $SourceExchangeOrganization
            ErrorAction = 'Stop'
            splat = @{
                Identity = $Identity
                ErrorAction = 'Stop'
            }
        }
        Try
        {
            Write-Log -Message $message -EntryType Attempting
            $SourceRecipientObject = @(Invoke-ExchangeCommand @GetFullRecipientCmdletParams)
            Write-Log -Message $message -EntryType Succeeded
        }
        Catch
        {
            $MyError = $_
            Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
            Write-Log -Message $myerror.tostring() -ErrorLog
            Throw $MyError
        }
        #endregion
        #region CreateIntermediateObject
        $GetDesiredTargetAliasParams = @{
            SourceAlias = $aliasPrefix + $SourceRecipientObject.Alias
            TargetExchangeOrganization = $TargetExchangeOrganization
            ErrorAction = 'Stop'
        }
        $message = "Get Desired Alias for Identity ($Identity) from Target Exchange Organization ($TargetExchangeOrganization)"
        Try
        {
            Write-Log -Message $message -EntryType Attempting
            $DesiredAlias = Get-DesiredTargetAlias @GetDesiredTargetAliasParams
            Write-Log -Message $message -EntryType Succeeded
        }
        Catch
        {
            $MyError = $_
            Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
            Write-Log -Message $myerror.tostring() -ErrorLog
            Throw $MyError
        }
        $GetDesiredProxyAddressesParams = @{
            CurrentProxyAddresses = $SourceRecipientObject.emailAddresses
            DesiredOrCurrentAlias = $DesiredAlias
            LegacyExchangeDNs = $SourceRecipientObject.LegacyExchangeDN
            VerifyAddTargetAddress = $true
            TargetDeliveryDomain = $TargetDeliveryDomain
        }
        if ($DomainsToRemove.Count -ge 1)
        {
            $GetDesiredProxyAddressesParams.DomainsToRemove = $DomainsToRemove
        }
        if ($AddressesToAdd.Count -ge 1)
        {
            $GetDesiredProxyAddressesParams.AddressesToAdd = $AddressesToAdd
        }
        if ($AddressesToRemove.Count -ge 1)
        {
            $GetDesiredProxyAddressesParams.AddressesToAdd = $AddressesToRemove
        }
        $message = "Get Desired ProxyAddresses for Identity ($Identity)"
        Try
        {
            Write-Log -Message $message -EntryType Attempting
            $DesiredProxyAddresses = Get-DesiredProxyAddresses @GetDesiredProxyAddressesParams
            Write-Log -Message $message -EntryType Succeeded
        }
        Catch
        {
            $MyError = $_
            Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
            Write-Log -Message $myerror.tostring() -ErrorLog
            Throw $MyError
        }
        $IntermediateObject = [pscustomobject]@{
            Name = $SourceRecipientObject.Name
            Alias = $DesiredAlias
            EmailAddresses = $DesiredProxyAddresses
            ExternalEmailAddress = ($DesiredProxyAddresses | ? {$_ -like $('*@' + $targetDeliveryDomain)} | Select-Object -First 1).split(':')[1]
            PrimarySMTPAddress = ($DesiredProxyAddresses | ? {$_ -clike 'SMTP:*'} | Select-Object -first 1).split(':')[1]
            DisplayName = $SourceRecipientObject.DisplayName
            OrganizationalUnit = $OrganizationalUnit
            CustomAttribute5 = $SourceRecipientObject.guid.guid
            CustomAttribute6 = 'Temporary Mail Routing Contact'
            EmailAddressPolicyEnabled = $false
            AddExternalEmailAddressToSourceObject = $AddExternalEmailAddressToSourceObject
        }
        #endregion
        #region CreateTargetObject
        if ($testOnly -eq $true)
        {
            $IntermediateObject
        }
        else
        {
            #region UpdateSourceObject
            if ($AddExternalEmailAddressToSourceObject)
            {
                Add-EmailAddress -Identity $Identity -EmailAddresses $IntermediateObject.ExternalEmailAddress -ExchangeOrganization $SourceExchangeOrganization -ErrorAction Stop
            }
            #endregion
            $newMailContactParams = @{
                Cmdlet = 'New-MailContact'
                ExchangeOrganization = $TargetExchangeOrganization
                ErrorAction = 'Stop'
                Splat = @{
                    ErrorAction = 'Stop'
                    Name = $IntermediateObject.Name
                    Alias = $IntermediateObject.Alias
                    ExternalEmailaddress = $IntermediateObject.ExternalEmailAddress
                    PrimarySMTPAddress = $IntermediateObject.PrimarySMTPAddress
                    DisplayName = $IntermediateObject.DisplayName
                    OrganizationalUnit = $IntermediateObject.OrganizationalUnit
                }
            }
            $setMailContactParams = @{
                Cmdlet = 'Set-MailContact'
                ExchangeOrganization = $TargetExchangeOrganization
                ErrorAction = 'Stop'
                Splat = @{
                    Identity = $IntermediateObject.Alias
                    ErrorAction = 'Stop'
                    EmailAddresses = $IntermediateObject.EmailAddresses
                    CustomAttribute5 = $IntermediateObject.CustomAttribute5
                    CustomAttribute6 = $IntermediateObject.CustomAttribute6
                    EmailAddressPolicyEnabled = $IntermediateObject.EmailAddressPolicyEnabled
                }
            }
            $message = "Create New Mail Contact for Identity ($identity) in Target Exchange Organization ($TargetExchangeOrganization)"
            Try
            {
                Write-Log -Message $message -EntryType Attempting
                $NewMailContactOutput = Invoke-ExchangeCommand @newMailContactParams
                Write-Log -Message $message -EntryType Succeeded
            }
            Catch
            {
                $MyError = $_
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $myerror.tostring() -ErrorLog
                Throw $MyError
            }
            $message = "Get New Mail Contact for Identity ($identity) in Target Exchange Organization ($TargetExchangeOrganization)"
            Write-Log -Message $message -EntryType Attempting
            $FindAttemptCount = 0
            Do
            {
                Start-Sleep -Seconds 5
                $GetMailContactParams = @{
                    Cmdlet = 'Get-MailContact'
                    ExchangeOrganization = $TargetExchangeOrganization
                    ErrorAction = 'SilentlyContinue'
                    Splat = @{
                        ErrorAction = 'SilentlyContinue'
                        Identity = $IntermediateObject.Alias
                    }
                }
                $NewMailContact = @(Invoke-ExchangeCommand @GetMailContactParams)
                $FindAttemptCount++
                if ($FindAttemptCount -ge 15)
                {
                    Throw "Failed: $message"
                }
            }
            Until ($NewMailContact.Count -eq 1)
            $message  = "Set additional attributes for Mail Contact for Identity ($Identity) in Target Exchange Organization ($TargetExchangeOrganization)"
            Try
            {
                Write-Log -Message $message -EntryType Attempting
                Invoke-ExchangeCommand @setMailContactParams
                Write-Log -Message $message -EntryType Succeeded
            }
            Catch
            {
                $MyError = $_
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $myerror.tostring() -ErrorLog
                Throw $MyError
            }
        }
        #endregion
    }
#end function New-MailFlowContactFromMailbox
###################################################################
. $(Join-Path $PSScriptRoot 'ProvisioningFunctions.ps1')
. $(Join-Path $PSScriptRoot 'ExchangeGetRecipientFunctions.ps1')