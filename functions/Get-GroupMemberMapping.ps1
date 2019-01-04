    Function Get-GroupMemberMapping {
        
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
