    Function Get-msExchRecipientTypeDetailsValue {
        
        [cmdletbinding()]
        param(
            [parameter(Mandatory)]
            [string]$RecipientTypeDetails
        )
        switch ($RecipientTypeDetails)
        {
            'LinkedMailbox' {$Value = 2}
            'RemoteRoomMailbox'{$value = 8589934592}
            'RemoteSharedMailbox' {$value = 34359738368}
            'RemoteUserMailbox' {$value = 2147483648}
            'RemoteEquipmentMailbox' {$value = 17173869184}
            'RoomMailbox' {$value = 16}
            'SharedMailbox' {$value = 4}
            'DiscoveryMailbox' {$value = 536870912}
            'ArbitrationMailbox' {$value = 536870912}
            'UserMailbox' {$value = 1}
            'LegacyMailbox' {$value = 8}
            'EquipmentMailbox' {$value = 32}
            'MailContact' {$value = 64}
            'MailForestContact' {$value = 32768}
            'MailUser' {$value = 128}
            'MailUniversalDistributionGroup' {$value = 256}
            'MailUniversalSecurityGroup' {$value = 1024}
            'DynamicDistributionGroup' {$value = 2048}
            'PublicFolder' {$value = 4096}
        }
        Write-Output -InputObject $Value
    
    }
