    Function New-SourceRecipientDNHash {
        
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
