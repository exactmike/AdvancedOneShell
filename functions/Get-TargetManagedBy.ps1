function Get-TargetManagedBy
{
    [cmdletbinding()]
    param
    (
        $SourceGroup
        ,
        $MappedTargetMemberUsers
        ,
        $SourceRecipientDNHash
        ,
        $SourceTargetRecipientMap
        ,
        $TargetExchangeOrganizationSession
    )
    $TargetManagedBy = $(
        #Test if ManagedBy Exists in the Exported Recipients
        if  ($SourceRecipientDNHash.ContainsKey($sg.ManagedBy)) 
        {
            #Test if ManagedBy Exists in the Mapped Target Recipients
            $SourceManagedByGUID = $($SourceRecipientDNHash.$($sg.ManagedBy).ObjectGUID.guid) 
            if ($SourceTargetRecipientMap.ContainsKey($SourceManagedByGUID))
            {
                $ObjectGUIDString = @($SourceTargetRecipientMap.$SourceManagedByGUID)[0]
                try
                {
                    Invoke-Command -Session $TargetExchangeOrganizationSession -ScriptBlock {Get-Recipient -identity $using:ObjectGUIDString -erroraction Stop | Select-Object -ExpandProperty DistinguishedName} -ErrorAction Stop
                }
                catch
                {
                    $false
                }
            }
            else
            {
                $false
            }
        }
        else
        {
            $false
        }
    )
    if ($false -ne $TargetManagedBy)
    {
        $Result = @{
            ManagedBy = $TargetManagedBy
            ManagedBySource = 'Source'
        }
        Return $Result
    }
    else
    {
        if ($MappedTargetMemberUsers.count -ge 1)
        {
            $TargetManagedBy = Invoke-Command -Session $TargetExchangeOrganizationSession -ScriptBlock {$Using:MappedTargetMemberUsers | Get-Recipient -errorAction Stop | Sort-Object -Property PrimarySMTPAddress | Select-Object -First 1 -ExpandProperty DistinguishedName} 
            $Result = @{
                ManagedBy = $TargetManagedBy
                ManagedBySource = 'Membership'
            }
        }
        else {
            $Result = @{
                ManagedBy = $null
                ManagedBySource = $null
            }
        }
        Return $Result
    }
}
