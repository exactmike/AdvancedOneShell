    Function Set-ExchangeAttributesOnTargetObject {
        
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
                                    Write-OneShellLog -message $writeProgressParams.currentoperation -EntryType Attempting
                                    Find-ADUser -Identity $value -IdentityType $sourceLookupAttribute -AmbiguousAllowed -ActiveDirectoryInstance $SourceAD -ErrorAction Stop
                                    Write-OneShellLog -message $writeProgressParams.currentoperation -EntryType Succeeded
                                }#try
                                catch {
                                    Write-OneShellLog -message $writeProgressParams.currentoperation -Verbose -EntryType Failed -ErrorLog
                                    Write-OneShellLog -Message $_.tostring() -ErrorLog
                                    Export-FailureRecord -Identity $ID -ExceptionCode 'SourceADUserNotFound' -FailureGroup NotProcessed -RelatedObjectIdentifier $value -RelatedObjectIdentifierType $SourceLookupAttribute
                                }
                            )#TrialSADU
                            #Determine action based on the results of the lookup attempt in the target AD
                            switch ($TrialSADU.count) {
                                1 {
                                    Write-OneShellLog -message "Succeeded: Found exactly 1 Matching User with value $value in $SourceLookupAttribute in Source Object Forest $SourceAD"
                                    #output the object into $SourceData
                                    $TrialSADU[0]
                                    Write-OneShellLog -Message "Source AD User Identified in with ObjectGUID: $($TrialSADU[0].objectguid)" -EntryType Notification
                                }#1
                                0 {
                                    Write-OneShellLog -message "FAILED: Found 0 Matching Users with value $value in $SourceLookupAttribute in Source Object Forest $SourceAD" -Verbose
                                    Export-FailureRecord -Identity $ID -ExceptionCode 'SourceADUserNotFound' -FailureGroup NotProcessed -RelatedObjectIdentifier $ID -RelatedObjectIdentifierType $SourceLookupAttribute
                                }#0
                                Default {
                                    Write-OneShellLog -message "FAILED: Found multiple ambiguous Matching Users with value $value in $SourceLookupAttribute in Source Object Forest $SourceAD" -Verbose
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
            Write-OneShellLog -Message "Completed Source Object Lookup/Validation Operations" -EntryType Notification -Verbose
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
                        Write-OneShellLog -Message $writeProgressParams.CurrentOperation -EntryType Attempting
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
                                Write-OneShellLog -Message $message -EntryType Notification
                                $TrialTADU = @(Find-ADUser -GivenName $GivenName -SurName $SurName -IdentityType GivenNameSurname -AmbiguousAllowed -AD $TargetAD -ErrorAction Stop)
                                $TrialTADU = @($TrialTADU | Where-Object {$_.ObjectGUID -ne $SADUGUID})
                            }
                            else
                            {
                                $message = "Attempting Secondary Attribute Lookup using $secondaryID in $TargetLookupSecondaryAttribute"
                                $writeProgressParams.CurrentOperation = $message
                                Write-OneShellLog -Message $message -EntryType Notification
                                $TrialTADU = @(Find-Aduser -Identity $SecondaryID -IdentityType $TargetLookupSecondaryAttribute -AD $TargetAD -ErrorAction Stop -AmbiguousAllowed)
                                $TrialTADU = @($TrialTADU | Where-Object {$_.ObjectGUID -ne $SADUGUID})
                            }
                            if ($TrialTADU.Count -ge 1)
                            {
                                $TrialTADU | Add-Member -MemberType NoteProperty -Name MatchAttribute -Value $TargetLookupSecondaryAttribute
                            }
                        }#if
                        Write-OneShellLog -Message $writeProgressParams.CurrentOperation -EntryType Succeeded
                    }#try
                    catch
                    {
                        Write-OneShellLog -Message $writeProgressParams.CurrentOperation -EntryType Failed -Verbose -ErrorLog
                        Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                            Write-OneShellLog -message "Succeeded: Found exactly 1 Matching User" -Verbose
                            $TADU = $TrialTADU[0]
                            $TADUGUID = $TADU.objectguid.guid
                            Write-OneShellLog -Message "Target AD User Identified in $TargetAD with ObjectGUID: $TADUGUID" -Verbose -EntryType Notification
                        }#1
                        0
                        {
                            if ($SADU.enabled)
                            {
                                Write-OneShellLog -Message "Found 0 Matching Users for User $ID, but Source User Object is Enabled." -Verbose -EntryType Notification
                                $TADU = $SADU
                                $TADUGUID = $SADUGUID
                            }
                            else {
                                Write-OneShellLog -message "Found 0 Matching Users for User $ID" -Verbose -EntryType Failed
                                Export-FailureRecord -Identity $ID -ExceptionCode 'TargetADUserNotFound' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType 'ObjectGUID'
                                continue nextID
                            }
                        }#0
                        Default
                        {#check for ambiguous results
                            Write-OneShellLog -Message "Find AD User returned multiple objects/ambiguous results for User $ID." -Verbose -ErrorLog -EntryType Failed
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
                            Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                            Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                            #Write-OneShellLog -message "Find Mail Contact for $addr in $TargetAD" -EntryType Attempting
                            $MailContact = @(Find-ADContact -Identity $addr -IdentityType ProxyAddress -AmbiguousAllowed -ActiveDirectoryInstance $TargetAD -ErrorAction Stop)
                            #Write-OneShellLog -message "No Errors: Find Mail Contact for $addr in $TargetAD" -EntryType Succeeded
                        }#try
                        catch
                        {
                            Write-OneShellLog -message "Unexpected Error: Find Mail Contact for $addr in $TargetAD" -EntryType Failed -Verbose -ErrorLog
                            Write-OneShellLog -message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity "$ID`:$addr" -ExceptionCode 'UnexpectedFailureDuringMailContactLookup' -FailureGroup ContactLookupFailure -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                            continue nextAddr
                        }#catch
                        If ($MailContact.count -ge 1)
                        {
                            Write-OneShellLog -Message "NOTE: A mail contact was found for $addr in $TargetAD." -Verbose
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
                            #Write-OneShellLog -message "Find Mail Contact for $addr in $TargetAD" -EntryType Attempting
                            $MailContact = @(Find-ADContact -Identity $addr -IdentityType ProxyAddress -AmbiguousAllowed -ActiveDirectoryInstance $TargetAD -ErrorAction Stop)
                            #Write-OneShellLog -message "No Errors: Find Mail Contact for $addr in $TargetAD" -EntryType Succeeded
                        }#try
                        catch
                        {
                            Write-OneShellLog -message "Unexpected Error: Find Mail Contact for $addr in $TargetAD" -EntryType Failed -Verbose -ErrorLog
                            Write-OneShellLog -message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity "$ID`:$addr" -ExceptionCode 'UnexpectedFailureDuringMailContactLookup' -FailureGroup ContactLookupFailure -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                        }#catch
                        If ($MailContact.count -ge 1)
                        {
                            Write-OneShellLog -Message "NOTE: A mail contact was found for $addr in $TargetAD." -Verbose
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
                            #Write-OneShellLog -message "Find Mail Contact for $addr in $TargetAD" -EntryType Attempting
                            $MailContact = @(Find-ADContact -Identity $addr -IdentityType DistinguishedName -ActiveDirectoryInstance $TargetAD -ErrorAction Stop)
                            Write-OneShellLog -message "No Errors: Find Mail Contact for $addr in $TargetAD" -EntryType Succeeded
                        }#try
                        catch
                        {
                            Write-OneShellLog -message "Unexpected Error: Find Mail Contact for $addr in $TargetAD" -EntryType Failed -Verbose -ErrorLog
                            Write-OneShellLog -message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity "$ID`:$addr" -ExceptionCode 'UnexpectedFailureDuringMailContactLookup' -FailureGroup ContactLookupFailure -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                        }#catch
                        If ($MailContact.count -ge 1)
                        {
                            Write-OneShellLog -Message "NOTE: A mail contact was found for $addr in $TargetAD." -Verbose
                            if ($MailContacts.distinguishedname -notcontains $MailContact.Distinguishedname)
                            {
                                $MailContacts += $MailContact
                                $Global:SEATO_MailContactsFound += $MailContact
                            }
                        }#if
                    }
                    Write-OneShellLog -Message "A total of $($MailContacts.count) mail contacts were found for $ID in $TargetAD" -Verbose -EntryType Notification
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
                            Write-OneShellLog -message "Was not able to find a valid alias and/or PrimarySMTPAddress to Assign to the target: $ID" -Verbose -EntryType Failed
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
            Write-OneShellLog -Message "$($IntermediateObjects.count) Object(s) Processed (Lookup of Source and Target Objects and Attribute Calculations)." -EntryType Notification
            #region CYABackup
            #depth must be 2 or greater to capture and restore MV attributes like proxy addresses correctly
            Export-OneShellData -DataToExport $IntermediateObjects -DataToExportTitle IntermediateObjects -Depth 3 -DataType json
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
                            Write-OneShellLog -Message "Target Operation Could Not Be Determined for $SADUGUID" -Verbose -ErrorLog -EntryType Failed
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
                                Write-OneShellLog -message $message -EntryType Attempting
                                set-aduser @setaduserparams1
                                Write-OneShellLog -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-OneShellLog -message $message -EntryType Failed -Verbose -ErrorLog
                                Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                                Write-OneShellLog -message $message -EntryType Attempting
                                set-aduser @setaduserparams2
                                Write-OneShellLog -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-OneShellLog -message "FAILED: SET target attributes $($setaduserparams2.'Add'.keys -join ';')  for $TADUGUID in $TargetAD" -Verbose -ErrorLog
                                Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                                Write-OneShellLog -message $message -EntryType Attempting
                                set-aduser @setaduserparams1
                                Write-OneShellLog -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-OneShellLog -message $message -EntryType Failed -Verbose -ErrorLog
                                Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                                Write-OneShellLog -message $message -EntryType Attempting
                                set-aduser @setaduserparams2
                                Write-OneShellLog -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-OneShellLog -message $message -Verbose -ErrorLog -EntryType Failed
                                Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                                Write-OneShellLog -message $message -EntryType Attempting
                                set-aduser @setaduserparams1
                                Write-OneShellLog -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-OneShellLog -message $message -EntryType Failed -Verbose -ErrorLog
                                Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                                Write-OneShellLog -message $message -EntryType Attempting
                                set-aduser @setaduserparams2
                                Write-OneShellLog -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-OneShellLog -message $message -Verbose -ErrorLog -EntryType Failed
                                Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                                Write-OneShellLog -message $message -EntryType Attempting
                                set-aduser @setaduserparams1
                                Write-OneShellLog -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-OneShellLog -message $message -EntryType Failed -Verbose -ErrorLog
                                Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                                Write-OneShellLog -message $message -EntryType Attempting
                                set-aduser @setaduserparams2
                                Write-OneShellLog -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-OneShellLog -message $message -Verbose -ErrorLog -EntryType Failed
                                Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                                Write-OneShellLog -message $message -EntryType Attempting
                                set-aduser @setaduserparams1
                                Write-OneShellLog -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-OneShellLog -message $message -EntryType Failed -Verbose -ErrorLog
                                Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                                Write-OneShellLog -message $message -EntryType Attempting
                                set-aduser @setaduserparams2
                                Write-OneShellLog -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Write-OneShellLog -message $message -Verbose -ErrorLog -EntryType Failed
                                Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                                    Write-OneShellLog -message $message -EntryType Attempting
                                    Add-ADGroupMember -Identity $groupDN -Members $TADUGUID -ErrorAction Stop -Confirm:$false -Server $Domain
                                    Write-OneShellLog -message $message -EntryType Succeeded
                                }#try
                                catch {
                                    Write-OneShellLog -message $message -EntryType Failed -Verbose -ErrorLog
                                    Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                        Write-OneShellLog -message "Attempting: Delete $($intobj.MatchingContactObject.count) Mail Contact(s) from $TargetAD" -Verbose
                        foreach ($c in $intobj.MatchingContactObject) {
                            try {
                                Write-OneShellLog -message "Attempting: Delete $($c.distinguishedname) Mail Contact from $TargetAD" -Verbose
                                Push-Location
                                Set-Location $($TargetAD + ':\')
                                $Domain = Get-AdObjectDomain -adobject $c -ErrorAction Stop
                                $splat = @{Identity = $c.distinguishedname;Confirm=$false;ErrorAction='Stop';Server=$Domain}
                                Remove-ADObject @splat
                                Pop-Location
                                Write-OneShellLog -message "Succeeded: Delete $($c.distinguishedname) Mail Contact from $TargetAD" -Verbose
                            }#try
                            catch {
                                Pop-Location
                                #$Global:ErrorActionPreference = 'Continue'
                                Write-OneShellLog -message "FAILED: Delete $($c.distinguishedname) Mail Contact from $TargetAD" -Verbose -ErrorLog
                                Write-OneShellLog -Message $_.tostring() -ErrorLog
                                $Global:SEATO_MailContactDeletionFailures+=$c
                            }#catch
                        }#foreach
                    }#if
                    #############################################################
                    #copy contact object memberships to Target AD User
                    #############################################################
                    if ($deletecontact -and $intobj.MatchingContactObject.count -ge 1) {
                        Write-OneShellLog -message "Attempting: Add $TADUGUID to Contacts' Distribution Groups in $TargetAD" -Verbose
                        $ContactGroupMemberships = @($intobj.MatchingContactObject | Select-Object -ExpandProperty MemberOf)
                        foreach ($group in $ContactGroupMemberships) {
                            try {
                                $message = "Add-ADGroupMember -Members $TADUGUID -Identity $group"
                                Write-OneShellLog -message $message -EntryType Attempting
                                Push-Location
                                Set-Location $($TargetAD + ':\')
                                $ADGroup = Get-ADGroup -Identity $group -ErrorAction Stop -Properties CanonicalName
                                $Domain = Get-AdObjectDomain -adobject $ADGroup -ErrorAction Stop
                                $splat = @{Identity = $group;Confirm=$false;ErrorAction='Stop';Members=$TADUGUID;Server=$Domain}
                                Add-ADGroupMember @splat
                                Pop-Location
                                Write-OneShellLog -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Pop-Location
                                Write-OneShellLog -message $message -Verbose -ErrorLog -EntryType Failed
                                Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                            Write-OneShellLog -message $message -EntryType Attempting
                            $Splat = @{
                                Identity = $SADUGUID
                                ErrorAction = 'Stop'
                                Confirm = $false
                                Server = Get-ADObjectDomain -adobject $SADU
                            }
                            Remove-ADObject @splat
                            Write-OneShellLog -message $message -EntryType Succeeded
                        }
                        catch {
                            Write-OneShellLog -message $message -Verbose -ErrorLog -EntryType Failed
                            Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                            Write-OneShellLog -message $message -EntryType Attempting
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
                            Write-OneShellLog -message $message -EntryType Succeeded
                            Pop-Location
                        }
                        catch {
                            Pop-Location
                            Write-OneShellLog -message $message -Verbose -ErrorLog -EntryType Failed
                            Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                            Write-OneShellLog -Message $message -EntryType Attempting -Verbose
                            $domain = Get-AdObjectDomain -adobject $TADU -ErrorAction Stop
                            Move-ADObject -Server $domain -Identity $TADUGUID -TargetPath $DestinationOU -ErrorAction Stop
                            Write-OneShellLog -Message $message -EntryType Succeeded -Verbose
                        }
                        catch
                        {
                            Write-OneShellLog -Message $message -EntryType Failed -Verbose
                            Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                            Write-OneShellLog -message $message -EntryType Attempting
                            $Splat = @{
                                Identity = $TADUGUID
                                ErrorAction = 'Stop'
                            }
                            $ErrorActionPreference = 'Stop'
                            Connect-Exchange -ExchangeOrganization $TargetExchangeOrganization > $null
                            Invoke-ExchangeCommand -cmdlet 'Update-Recipient' -splat $Splat -ExchangeOrganization $TargetExchangeOrganization -ErrorAction Stop
                            Write-OneShellLog -message $message -EntryType Succeeded
                            $RecipientUpdated = $true
                            $ErrorActionPreference = 'Continue'
                        }
                        catch {
                            $UpdateRecipientFailedCount++
                            $ErrorActionPreference = 'Continue'
                            Write-OneShellLog -message $message -Verbose -ErrorLog -EntryType Failed
                            Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                Write-OneShellLog -Message "$recordcount Objects Processed for Test Only" -EntryType Notification -Verbose
                foreach ($intObj in $ProcessedObjects)
                {
                    $SADUGUID = $IntObj.SourceUserObjectGUID
                    $TADUGUID = $IntObj.TargetUserObjectGUID
                    Write-OneShellLog -Message "Processed Object SADU $SADUGUID and TADU $TADUGUID" -EntryType Notification -Verbose
                }
                Write-Output -InputObject $ProcessedObjects
            }
            else
            {
                $RecordCount = $ProcessedObjects.Count
                $cr = 0
                Write-OneShellLog -Message "$recordcount Objects Processed Locally" -EntryType Notification -Verbose
                if ($ProcessedObjects.Count -ge 1) {
                    #Start a Directory Synchronization to Azure AD Tenant
                    #Wait first for AD replication
                    Write-OneShellLog -Message "Waiting for $ADSyncDelayInSeconds seconds for AD Synchronization before starting an Azure AD Directory Synchronization." -Verbose -EntryType Notification
                    New-Timer -units Seconds -length $ADSyncDelayInSeconds -showprogress -Frequency 5 -voice
                    #Write-OneShellLog -Message "Starting an Azure AD Directory Synchronization." -Verbose -EntryType Notification
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
                                            Write-OneShellLog -message $message -EntryType Attempting
                                            Invoke-ExchangeCommand -cmdlet 'Set-Mailbox' -ExchangeOrganization OL -string "-Identity $($IntObj.DesiredUPNAndPrimarySMTPAddress) -ForwardingSmtpAddress $($IntObj.DesiredCoexistenceRoutingAddress)" -ErrorAction Stop
                                            Write-OneShellLog -message $message -EntryType Succeeded
                                            $ErrorActionPreference = 'Continue'
                                            $SetMailboxForwardingStatus = $true
                                        }
                                        catch {
                                            Write-OneShellLog -message $message -Verbose -ErrorLog -EntryType Failed
                                            Write-OneShellLog -Message $_.tostring() -ErrorLog
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
                                        Write-OneShellLog -Message $message -EntryType Attempting
                                        $MRSourceData = @($IntObj | Select-Object $SourceDataProperties)
                                        $MR = @(New-MRMMoveRequest -SourceData $MRSourceData -wave $MoveRequestWaveBatchName -wavetype Sub -SuspendWhenReadyToComplete $true -ExchangeOrganization OL -LargeItemLimit 50 -BadItemLimit 50 -ErrorAction Stop)
                                        if ($MR.Count -eq 1)
                                        {
                                            Write-OneShellLog -Message $message -EntryType Succeeded
                                        } else {
                                            Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                                            #Write-OneShellLog -Message $_.tostring() -ErrorLog
                                            Export-FailureRecord -Identity $($IntObj.DesiredUPNAndPrimarySMTPAddress) -ExceptionCode "CreateMoveRequestFailure" -FailureGroup MailboxMove -ExceptionDetails $_.tostring()
                                        }

                                    }
                                    catch
                                    {
                                        Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                                        Write-OneShellLog -Message $_.tostring() -ErrorLog
                                        Export-FailureRecord -Identity $($IntObj.DesiredUPNAndPrimarySMTPAddress) -ExceptionCode "CreateMoveRequestFailure" -FailureGroup MailboxMove -ExceptionDetails $_.tostring()
                                    }
                                }#'UpdateAndMigrateOnPremisesMailbox'
                            }#switch
                        }
                        else {
                            $message = "Sync Related Failure for $($IntObj.DesiredUPNAndPrimarySMTPAddress)."
                            Write-OneShellLog -message $message -Verbose -ErrorLog -EntryType Failed
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
                    Write-OneShellLog -Message "NOTE: Processing for $($TADU.UserPrincipalName) with GUID $TADUGUID in $TargetAD has completed successfully." -Verbose
                }#foreach IntObj in ProcessedObjects
                $writeProgressParams.currentOperation = "Completed Post Attribute/Object Update Operations"
                Write-Progress @writeProgressParams -Completed
                #region ReportAllResults
                if ($Global:SEATO_ProcessedUsers.count -ge 1) {
                    Write-OneShellLog -Message "Successfully Processed $($Global:SEATO_ProcessedUsers.count) Users."
                    Export-OneShellData -DataToExportTitle TargetForestProcessedUsers -DataToExport $Global:SEATO_ProcessedUsers -DataType csv #-Append
                    Export-OneShellData -DataToExportTitle TargetForestFullProcessedUsers -DataToExport $Global:SEATO_FullProcessedUsers -DataType csv
                }
                if ($Global:SEATO_Exceptions.count -ge 1) {
                    Write-OneShellLog -Message "Processed $($Global:SEATO_Exceptions.count) Users with Exceptions."
                }
                if ($Global:SEATO_MailContactsFound.count -ge 1) {
                    Write-OneShellLog -Message "$($Global:SEATO_MailContactsFound.count) Contacts were found and are being exported."
                    Export-OneShellData -DataToExportTitle FoundMailContacts -DataToExport $Global:SEATO_MailContactsFound -Depth 2 -DataType xml
                }
                if ($Global:SEATO_OriginalTargetUsers.count -ge 1) {
                    Write-OneShellLog -Message "$($Global:SEATO_OriginalTargetUsers.count) Original Target Users were attempted for processing and are being exported."
                    Export-OneShellData -DataToExportTitle OriginalTargetUsers -DataToExport $Global:SEATO_OriginalTargetUsers -Depth 2 -DataType xml
                }
                if ($Global:SEATO_MailContactDeletionFailures.Count -ge 1) {
                    Write-OneShellLog -Message "$($Global:SEATO_MailContactDeletionFailures.Count) Mail Contact(s) NOT successfully deleted.  Exporting them for review."
                    Export-OneShellData -DataToExportTitle MailContactsNOTDeleted -DataToExport $Global:SEATO_MailContactDeletionFailures -DataType csv
                }
                if ($Global:SEATO_OLMailboxSummary.count -ge 1) {
                    Write-OneShellLog -Message "$($Global:SEATO_OLMailboxSummary.Count) Online Mailboxes Configured for Forwarding.  Exporting summary details for review."
                    Export-OneShellData -DataToExportTitle OnlineMailboxForwarding -DataToExport $Global:SEATO_OLMailboxSummary -DataType csv
                }
                #endregion ReportAllResults
            }# else when NOT -TestOnly
        }#end end
    
    }
