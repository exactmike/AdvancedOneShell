    Function Get-TargetRecipientFromMap {
        
        [cmdletbinding()]
        param
        (
            $SourceObjectGUID
            ,
            $ExchangeSession
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
                    $cmdlet = Get-RecipientCmdlet -Identity $id -verb Get -ExchangeSession $ExchangeSession
                    $scriptblock = [scriptblock]::create("$cmdlet -Identity $id -ErrorAction Stop")
                    Invoke-Command -session $ExchangeSession -scriptblock $scriptblock -ErrorAction Stop
                }
            )
            $TargetRecipients
        }
    
    }
