    Function Update-PostMigrationMailboxUser {
        
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
                        Write-OneShellLog -Message "FAILED: InputList does not contain the Target Lookup Attribute $TargetLookupAttribute." -Verbose -ErrorLog
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
                        Write-OneShellLog -Message "Attempting: Find AD User $ID in Target AD Forest $TargetAD" -Verbose
                        $TADU = @(Find-Aduser -Identity $ID -IdentityType $TargetLookupAttribute -ADInstance $TargetAD -ErrorAction Stop)
                        Write-OneShellLog -Message "Succeeded: Find AD User $ID in Target AD Forest $TargetAD" -Verbose
                    }#try
                    catch {
                        Write-OneShellLog -Message "FAILED: Find AD User $ID in Target AD Forest $TargetAD" -Verbose -ErrorLog
                        Write-OneShellLog -Message $_.tostring() -ErrorLog
                        $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'TargetADUserNotFound'}}
                        Export-OneShellData -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                        throw("User Object for value $ID in Attribute $TargetLookupAttribute in Target AD $TargetAD not found.")
                    }#catch
                    if ($TADU.count -gt 1) {#check for ambiguous results
                        Write-OneShellLog -Message "FAILED: Find AD User $ID in Target AD Forest $TargetAD returned multiple objects/ambiguous results." -Verbose -ErrorLog
                        $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'TargetADUserAmbiguous'}}
                        Export-OneShellData -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                        throw("User Object for value $ID in Attribute $TargetLookupAttribute in Target AD $TargetAD was ambiguous.")
                    }#if
                    else {
                        $TADU = $TADU[0]
                        $TADUGUID = $TADU.objectguid
                        Write-OneShellLog -Message "NOTE: Target AD User in $TargetAD Identified with ObjectGUID: $TADUGUID" -Verbose
                    }
                    ################################################################################################################################################################
                    #Lookup Matching Source AD User
                    ################################################################################################################################################################
                    $writeProgressParams.status = "Lookup User with $($TADU.$SourceLookupAttribute) by $SourceLookupAttribute in Source AD"
                    Write-Progress @writeProgressParams
                    $SADU = @()
                    foreach ($ad in $SourceAD) {
                        try {
                            Write-OneShellLog -message "Attempting: Find Matching User for $ID in Source AD $ad by Lookup Attribute $SourceLookupAttribute" -Verbose
                            $SADU += Find-Aduser -Identity $($TADU.$SourceLookupAttribute) -IdentityType $SourceLookupAttribute -ADInstance $ad -ErrorAction Stop
                            Write-OneShellLog -message "Succeeded: Find Matching User for $ID in Source AD $ad by Lookup Attribute $SourceLookupAttribute" -Verbose
                        }#try
                        catch {
                            Write-OneShellLog -message "FAILED: Find Matching User for $ID in Source AD $ad by Lookup Attribute $SourceLookupAttribute" -Verbose -ErrorLog
                            Write-OneShellLog -Message $_.tostring() -ErrorLog
                        }
                    }#foreach
                    #check for no results or ambiguous results
                    switch ($SADU.count) {
                        1 {
                            Write-OneShellLog -message "Succeeded: Found exactly 1 Matching User for $ID in $($SourceAD -join ' & ') by Lookup Attribute $SourceLookupAttribute" -Verbose
                            $SADU = $SADU[0]
                            $SADUGUID = $SADU.objectguid
                            Write-OneShellLog -Message "NOTE: Source AD User Identified in with ObjectGUID: $SADUGUID" -Verbose
                        }#1
                        0 {
                            Write-OneShellLog -message "FAILED: Found 0 Matching User for $ID in Source AD $($SourceAD -join ' & ') by Lookup Attribute $SourceLookupAttribute" -Verbose
                            $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'SourceADUserNotFound'}}
                            Export-OneShellData -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                            throw("User Object for value $ID in Attribute $SourceLookupAttribute in Source AD $($SourceAD -join ' & ') not found.")
                        }#0
                        Default {
                            Write-OneShellLog -message "FAILED: Found multiple ambiguous matching User for $ID in Source AD $($SourceAD -join ' & ') by Lookup Attribute $SourceLookupAttribute" -Verbose
                            $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'SourceADUserAmbiguous'}}
                            Export-OneShellData -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                            throw("User Object for value $ID in Attribute $SourceLookupAttribute in Source AD $($SourceAD -join ' & ') was ambiguous.")
                        }#Default
                    }#switch $SADU.count
                    ################################################################################################################################################################
                    #Calculate Address Changes
                    ################################################################################################################################################################
                    $writeProgressParams.status = "Calculate Proxy Address and Target Address Changes"
                    Write-Progress @writeProgressParams
                    try {
                        Write-OneShellLog -Message "Attempting: Find Current proxy $TargetDeliveryDomain SMTP Address for Target AD User $TADUGUID" -Verbose
                        $TargetDeliveryDomainAddress = ($TADU.proxyaddresses | Where-Object {$_ -like "smtp:*@$TargetDeliveryDomain"} | Select-Object -First 1).split(':')[1]
                        Write-OneShellLog -Message "Succeeded: Find Current proxy $TargetDeliveryDomain SMTP Address for Target AD User $TADUGUID : $TargetDeliveryDomainAddress" -Verbose
                    }#try
                    catch {
                        Write-OneShellLog -Message "FAILED: Find Current proxy $TargetDeliveryDomain SMTP Address for Target AD User $TADUGUID" -Verbose -ErrorLog
                        Write-OneShellLog -Message $_.tostring() -ErrorLog
                        Write-OneShellLog -Message "NOTE: $TargetDeliveryDomain SMTP Proxy Address for Target AD User $TADUGUID will be added." -Verbose -ErrorLog
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
                    Write-OneShellLog -message "Using AD Cmdlets to set attributes for $TADUGUID in $TargetAD" -Verbose
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
                        Write-OneShellLog -message "Attempting: Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD" -Verbose
                        set-aduser @setaduserparams1
                        Write-OneShellLog -message "Succeeded: Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD" -Verbose
                    }#try
                    catch {
                        Write-OneShellLog -message "FAILED: Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD" -Verbose -ErrorLog
                        Write-OneShellLog -Message $_.tostring() -ErrorLog
                        $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'FailedToClearTargetAttributes'}}
                        Export-OneShellData -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
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
                        Write-OneShellLog -message "Attempting: SET target attributes $($setaduserparams2.'Add'.keys -join ';') for $TADUGUID in $TargetAD" -Verbose
                        set-aduser @setaduserparams2
                        Write-OneShellLog -message "Succeeded: SET target attributes $($setaduserparams2.'Add'.keys -join ';') for $TADUGUID in $TargetAD" -Verbose
                    }#try
                    catch {
                        Write-OneShellLog -message "FAILED: SET target attributes $($setaduserparams2.'Add'.keys -join ';')  for $ID in $TargetAD" -Verbose -ErrorLog
                        Write-OneShellLog -Message $_.tostring() -ErrorLog
                        $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'FailedToSetTargetAttributes'}}
                        Export-OneShellData -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
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
                        Write-OneShellLog -message "Attempting: Enable-ADAccount $TADUGUID in $TargetAD" -Verbose
                        Enable-ADAccount @EnableADAccountParams
                        Write-OneShellLog -message "Succeeded: Enable-ADAccount $TADUGUID in $TargetAD" -Verbose
                    }#try
                    catch {
                        Write-OneShellLog -message "FAILED: Enable-ADAccount $TADUGUID in $TargetAD" -Verbose -ErrorLog
                        Write-OneShellLog -Message $_.tostring() -ErrorLog
                        $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'FailedToEnableAccount'}}
                        Export-OneShellData -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
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
                        Write-OneShellLog -message "Attempting: Update Recipient $DesiredUPNAndPrimarySMTPAddress in $TargetExchangeOrg" -Verbose
                        $Global:ErrorActionPreference = 'Stop'
                        Connect-Exchange -ExchangeOrganization $TargetExchangeOrg
                        Invoke-ExchangeCommand -cmdlet Update-Recipient -ExchangeOrganization $TargetExchangeOrg -string "-Identity $TADUGUID -ErrorAction Stop"
                        $Global:ErrorActionPreference = 'Continue'
                        Write-OneShellLog -message "Succeeded: Update Recipient $DesiredUPNAndPrimarySMTPAddress in $TargetExchangeOrg" -Verbose
                    }
                    catch {
                        $Global:ErrorActionPreference = 'Continue'
                        Write-OneShellLog -message "FAILED: Update Recipient $DesiredUPNAndPrimarySMTPAddress in $TargetExchangeOrg" -Verbose -ErrorLog
                        Write-OneShellLog -message $_.tostring() -ErrorLog
                        $Global:Exceptions += $DesiredUPNAndPrimarySMTPAddress | Select-Object *,@{n='Exception';e={'FailedToUpdateRecipient'}}
                        Export-OneShellData -DataToExportTitle TargetForestExceptionsUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                        throw("Failed to Update Recipient for $TADUGUID in $TargetExchangeOrg")
                    }
                    $ProcessedUser = $TADU | Select-Object -Property SAMAccountName,DistinguishedName,@{n='UserPrincipalname';e={$DesiredUPNAndPrimarySMTPAddress}},@{n='ObjectGUID';e={$TADUGUID}}
                    $Global:ProcessedUsers += $ProcessedUser
                    Write-OneShellLog -Message "NOTE: Processing for $DesiredUPNAndPrimarySMTPAddress with GUID $TADUGUID in $TargetAD and $TargetExchangeOrg has completed successfully." -Verbose
                Export-OneShellData -DataToExportTitle PostMailboxMigrationProcessedUsers -DataToExport $ProcessedUser -DataType csv -Append
            }#try
            catch {
                $_
            }
            }#foreach
        }#process
        end{
        if ($Global:ProcessedUsers.count -ge 1) {
            Write-OneShellLog -Message "Successfully Processed $($Global:ProcessedUsers.count) Users." -Verbose
        }
        if ($Global:Exceptions.count -ge 1) {
            Write-OneShellLog -Message "Processed $($Global:Exceptions.count) Users with Exceptions." -Verbose
        }
        }#end
    
    }
