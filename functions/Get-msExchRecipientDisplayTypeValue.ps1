    Function Get-msExchRecipientDisplayTypeValue {
        
        [cmdletbinding()]
        param
        (
            [parameter(Mandatory)]
            [string]$RecipientTypeDetails
        )
        switch ($RecipientTypeDetails)
        {
            'LinkedMailbox' {$Value = 1073741824}
            'RemoteRoomMailbox'{$value = -2147481850}
            'RemoteSharedMailbox' {$value = -2147483642}
            'RemoteUserMailbox' {$value = -2147483642}
            'RemoteEquipmentMailbox' {$value = -2147481594}
            'RoomMailbox' {$value = 7}
            'SharedMailbox' {$value = 1073741824}
            'DiscoveryMailbox' {$value = $null}
            'ArbitrationMailbox' {$value = $null}
            'UserMailbox' {$value = 1073741824}
            'LegacyMailbox' {$value = $null}
            'EquipmentMailbox' {$value = 8}
            'MailContact' {$value = 6}
            'MailForestContact' {$value = $null}
            'MailUser' {$value = 6}
            'MailUniversalDistributionGroup' {$value = 1}
            'MailUniversalSecurityGroup' {$value = 1073741833}
            'DynamicDistributionGroup' {$value = 3}
            'PublicFolder' {$value = 2}
        }
        $Value
    
    }
