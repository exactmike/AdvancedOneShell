    Function Get-msExchRemoteRecipientTypeValue {
        
        [cmdletbinding()]
        param
        (
            [parameter(Mandatory)]
            [string]$RecipientTypeDetails
        )
        switch ($RecipientTypeDetails)
        {
            'LinkedMailbox' {$Value = $null}
            'RemoteRoomMailbox'{$value = 36}
            'RemoteSharedMailbox' {$value = 100}
            'RemoteUserMailbox' {$value = 4}
            'RemoteEquipmentMailbox' {$value = 68}
            'RoomMailbox' {$value = 32}
            'SharedMailbox' {$value = 96}
            'DiscoveryMailbox' {$value = $null}
            'ArbitrationMailbox' {$value = $null}
            'UserMailbox' {$value = $null}
            'LegacyMailbox' {$value = $null}
            'EquipmentMailbox' {$value = 64}
            'MailContact' {$value = $null}
            'MailForestContact' {$value = $null}
            'MailUser' {$value = $null}
            'MailUniversalDistributionGroup' {$value = $null}
            'MailUniversalSecurityGroup' {$value = $null}
            'DynamicDistributionGroup' {$value = $null}
            'PublicFolder' {$value = $null}
        }
        Write-Output -InputObject $Value
    
    }
