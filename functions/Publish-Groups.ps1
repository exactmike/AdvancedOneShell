    Function Publish-Groups {
        
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
        [switch]$PrefixOnlyIfNecessary
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
        ,
        [switch]$CreateContacts
        ,
        $RequireSenderAuthenticationEnabledLookup
        )
        Connect-OneShellSystem -Identity $TargetExchangeOrganization -ErrorAction Stop
        $TargetExchangeOrganizationSession = Get-OneShellSystemPSSession -Identity $TargetExchangeOrganization
        if (-not (Test-Path variable:\IntermediateGroupObjects)) {
            New-Variable -Name IntermediateGroupObjects -Value @() -Scope Global
        }
        $csgCount = 0
        $sgCount = $SourceGroups.Count
        $stopwatch = [system.diagnostics.stopwatch]::startNew()
        foreach ($sg in $SourceGroups)
        {
            $csgCount++
            #Write-OneShellLog -Message "Processing Source Group $($sg.mailnickname)" -EntryType Notification
            #Write-Verbose -Message "Processing Source Group $($sg.mailnickname)" -Verbose
        #region Prepare
            $GetDesiredTargetAliasParams = @{
                sourceAlias = $sg.mailNickName
                TargetExchangeOrganizationSession = $TargetExchangeOrganizationSession
                ReplacementPrefix = $ReplacementPrefix
                SourcePrefix = $SourcePrefix
            }
            if ($true -eq $PrefixOnlyIfNecessary) {$GetDesiredTargetAliasParams.PrefixOnlyIfNecessary = $true}
            Connect-OneShellSystem -Identity $TargetExchangeOrganization
            $desiredAlias = Get-DesiredTargetAlias @GetDesiredTargetAliasParams
            #Write-OneShellLog -Message "Processing Source Group $($sg.mailnickname). Target Group alias will be $desiredAlias." -EntryType Notification
            $WriteProgressParams =
            @{
                Activity = "Provisioning $($SourceGroups.count) Groups into $TargetExchangeOrganization, $TargetGroupOU"
                Status = "Working $csgCount of $($SourceGroups.count)"
                CurrentOperation = $desiredAlias
                PercentComplete = $csgCount/$sgCount*100
            }
            if ($csgCount -gt 1 -and ($csgCount % 10) -eq 0)
            {
                $WriteProgressParams.SecondsRemaining = ($($stopwatch.Elapsed.TotalSeconds.ToInt32($null))/($csgCount - 1)) * ($sgCount - ($csgCount - 1))
                Write-Progress @WriteProgressParams
            }
            $GetDesiredPrimarySMTPAddressParams = @{
                DesiredAlias = $desiredAlias
                TargetExchangeOrganizationSession = $TargetExchangeOrganizationSession 
                TargetSMTPDomain = $TargetSMTPDomain
            }
            Connect-OneShellSystem -Identity $TargetExchangeOrganization
            $desiredPrimarySMTPAddress = 'SMTP:' + $(Get-DesiredTargetPrimarySMTPAddress @GetDesiredPrimarySMTPAddressParams)
            $desiredName = Get-DesiredTargetName -SourceName $sg.DisplayName  -SourcePrefix $SourcePrefix #-ReplacementPrefix $ReplacementPrefix
            $targetRecipientGUIDs = @($RecipientMaps.SourceTargetRecipientMap.$($sg.ObjectGUID.Guid))
            $targetRecipients = Get-TargetRecipientFromMap -SourceObjectGUID $($sg.ObjectGUID.Guid) -ExchangeSession $TargetExchangeOrganizationSession
            $GetDesiredProxyAddressesParams = @{
                CurrentProxyAddresses = $sg.proxyAddresses | Where-Object {$_ -like 'smtp:*'} | ForEach-Object {$($_.split('@')[0]) + '@' + $TargetSMTPDomain}
                DesiredPrimarySMTPAddress = $desiredPrimarySMTPAddress
                TestAddressAvailabilityInExchangeSession = $true
                ExchangeSession = $TargetExchangeOrganizationSession
            }
            Connect-OneShellSystem -Identity $TargetExchangeOrganization
            $DesiredProxyAddresses = Get-AltDesiredProxyAddresses @GetDesiredProxyAddressesParams
            $OriginalPrimarySMTPAddress = $sg.proxyAddresses | Where-Object {$_ -clike 'SMTP:*'} | Select-Object -First 1 | ForEach-Object {$_.split(':')[1]}
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
            Connect-OneShellSystem -Identity $TargetExchangeOrganization
            $ManagedBy = Get-TargetManagedBy -SourceGroup $sg -MappedTargetMemberUsers $mappedTargetMemberUsers -SourceRecipientDNHash $SourceRecipientDNHash -SourceTargetRecipientMap $SourceTargetRecipientMap -TargetExchangeOrganizationSession $TargetExchangeOrganizationSession
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
                ManagedBy = $ManagedBy.ManagedBy
                ManagedBySource = $ManagedBy.ManagedBySource
                SourcePrimarySMTPAddress = $OriginalPrimarySMTPAddress
                SourceRequireSenderAuthenticationEnabled = $RequireSenderAuthenticationEnabledLookup.$OriginalPrimarySMTPAddress
            }
            $Global:intermediateGroupObjects += $intermediateGroupObject
            Export-OneShellData -DataToExportTitle $("Group-" + $DesiredAlias) -DataToExport $intermediateGroupObject -Depth 3 -DataType json
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
                    Connect-OneShellSystem -Identity $TargetExchangeOrganization
                    $cmdlet = Get-RecipientCmdlet -Recipient $tr -verb Remove
                    $rrParams =
                    @{
                        Identity = $($tr.Guid.guid)
                        Confirm = $false
                        ErrorAction = 'Stop'
                    }
                    try
                    {
                        Write-OneShellLog -Message $message -EntryType Attempting
                        Invoke-ExchangeCommand -cmdlet $cmdlet -splat $rrParams -ExchangeOrganization $TargetExchangeOrganization -ErrorAction Stop
                        Write-OneShellLog -Message $message -EntryType Succeeded
                    }
                    catch
                    {
                        Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                        Write-OneShellLog -Message $_.tostring() -ErrorLog -Verbose
                    }
                }
            }
        #endregion RemoveTargetRecipients
            #region CreateNeededContacts
            foreach ($nmc in $nonMappedTargetMemberContacts)
            {
                try {
                    $ContactDesiredName = Get-DesiredTargetName -SourceName $nmc.DisplayName -SourcePrefix $SourcePrefix #-ReplacementPrefix $ReplacementPrefix
                    $ContactDesiredAlias = Get-DesiredTargetAlias -SourceAlias $nmc.MailNickName -TargetExchangeOrganizationSession $TargetExchangeOrganizationSession -ReplacementPrefix $ReplacementPrefix -SourcePrefix $SourcePrefix -PrefixOnlyIfNecessary
                    #$ContactDesiredProxyAddresses = Get-DesiredProxyAddresses -CurrentProxyAddresses $nmc.proxyAddresses -DesiredOrCurrentAlias $ContactDesiredAlias -LegacyExchangeDNs $nmc.legacyExchangeDN
                    $ContactDesiredProxyAddresses = @($nmc.TargetAddress)
                }
                catch {
                    Export-OneShellData -DataToExport $nmc -DataToExportTitle "ContactCreationFailure-$($nmc.MailNickName)" -DataType json -Depth 3
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
                Export-OneShellData -DataToExportTitle $("Contact-" + $ContactDesiredAlias) -DataToExport $intermediateContactObject -Depth 3 -DataType json
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
                        Write-OneShellLog -Message $message -EntryType Attempting
                        Connect-OneShellSystem -Identity $TargetExchangeOrganization -ErrorAction Stop
                        $TargetExchangeOrganizationSession = Get-OneShellSystemPSSession -identity $TargetExchangeOrganization -ErrorAction Stop
                        $newContact = Invoke-Command -ScriptBlock {New-MailContact @Using:newMailContactParams}  -Session $TargetExchangeOrganizationSession
                        $mappedTargetMemberContacts += $newContact.guid.guid
                        $AllMappedMembersToAddAtCreation += $newContact.guid.guid
                        $message = "Find Newly Created Contact $ContactDesiredAlias."
                        $found = $false
                        do
                        {
                            #Write-OneShellLog -Message $message -EntryType Attempting
                            $Contact = @(Invoke-Command -scriptblock {Get-MailContact -Identity $($using:NewContact).guid.guid} -Session $TargetExchangeOrganizationSession)
                            if ($Contact.Count -eq 1)
                            {
                                Write-OneShellLog -Message $message -EntryType Succeeded
                                $found = $true
                            }
                            Start-Sleep -Seconds 10
                        }
                        until
                        (
                            $found -eq $true
                        )
                        $message = "Set Newly Created Contact $ContactDesiredName Attributes"
                        Write-OneShellLog -Message $message -EntryType Attempting
                        Invoke-Command -ScriptBlock {Set-MailContact $using:setMailContactParams} -Session $TargetExchangeOrganizationSession -ErrorAction Stop
                        Write-OneShellLog -Message $message -EntryType Succeeded
                    }
                    catch
                    {
                        Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                        Write-OneShellLog -Message $_.tostring() -ErrorLog -Verbose
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
                try
                {
                    Connect-OneShellSystem -Identity $TargetExchangeOrganization
                    $message = "Create Group $desiredAlias"
                    Write-OneShellLog -Message $message -EntryType Attempting
                    $newgroup = Invoke-Command -ScriptBlock {New-DistributionGroup @using:newDistributionGroupParams} -Session $TargetExchangeOrganizationSession -ErrorAction Stop
                    Write-OneShellLog -Message $message -EntryType Succeeded
                    Start-Sleep -Seconds 1
                    $message = "Find Newly Created Group $desiredAlias"
                    $found = $false
                    Do
                    {
                        #Write-OneShellLog -Message $message -EntryType Attempting
                        $group = @(Invoke-Command -ScriptBlock {Get-DistributionGroup -Identity $($using:NewGroup).guid.guid -ErrorAction SilentlyContinue} -Session $TargetExchangeOrganizationSession -ErrorAction SilentlyContinue)
                        if ($group.Count -eq 1)
                        {
                            Write-OneShellLog -Message $message -EntryType Succeeded
                            $found = $true
                        }
                        Start-Sleep -Seconds 1
                    }
                    Until
                    ($found -eq $true)
                    $message = "Set Group $desiredAlias Attributes"
                    $setDistributionGroupParams =
                    @{
                        Identity = $newgroup.guid.guid
                        #EmailAddresses = $DesiredProxyAddresses
                        EmailAddressPolicyEnabled = $false
                        errorAction = 'Stop'
                    }
                    Write-OneShellLog -Message $message -EntryType Attempting
                    Invoke-Command -ScriptBlock {Set-DistributionGroup @using:setDistributionGroupParams} -Session $TargetExchangeOrganizationSession
                    Write-OneShellLog -Message $message -EntryType Succeeded
                    Write-OneShellLog -Message "Provisioning Complete for Group $desiredAlias." -EntryType Notification -Verbose
                }
                catch
                {
                    Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                    Write-OneShellLog -Message $_.tostring() -ErrorLog -Verbose
                }
                #endregion ProvisionDistributionGroup
            }#else
        }#foreach
    
    }
