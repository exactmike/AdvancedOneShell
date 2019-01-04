    Function Get-ADRecipientsWithConflictingAlias {
        
        [cmdletbinding()]
        param
        (
        [parameter(Mandatory=$true)]
        $SourceRecipients
        ,
        [parameter(Mandatory=$true)]
        $TargetExchangeOrganization
        ,
        [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
        [string]$ReplacementPrefix
        ,
        [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
        [string]$SourcePrefix
        )
        foreach ($sr in $SourceRecipients)
        {
            $Alias = $sr.mailNickName
            $Alias = $Alias -replace '\s|[^1-9a-zA-Z_-]',''
            if ($PSCmdlet.ParameterSetName -eq 'ReplacePrefix')
            {
                $NewAlias = $Alias -replace "\b$($sourcePrefix)_",''
                $NewAlias = $NewAlias -replace "\b$($SourcePrefix)", ''
                $NewAlias = $NewAlias -replace "$($SourcePrefix)\b", ''
                $NewAlias = "$($ReplacementPrefix)_$($NewAlias)"
                $Alias = $NewAlias
            }
            if (Test-ExchangeAlias -Alias $Alias -ExchangeOrganization $TargetExchangeOrganization)
            {
                Write-OneShellLog -Message "No Conflict for $Alias" -EntryType Notification -Verbose
            }
            else
            {
                $conflicts = @(Test-ExchangeAlias -Alias $Alias -ExchangeOrganization $TargetExchangeOrganization -ReturnConflicts)
                [pscustomobject]@{
                    SourceObjectGUID = $sr.ObjectGUID
                    ConflictingTargetObjectGUIDs = $conflicts
                    OriginalAlias = $sr.mailNickName
                    TestedAlias = $Alias
                }
            }

        }
    
    }
