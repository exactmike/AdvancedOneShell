    Function Get-DesiredTargetAlias {
        
        [cmdletbinding()]
        param
        (
            [parameter(ParameterSetName = 'NewPrefix',Mandatory=$true)]
            [parameter(ParameterSetName = 'Standard',Mandatory=$true)]
            [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
            $SourceAlias
            ,
            [parameter(ParameterSetName = 'NewPrefix')]
            [parameter(ParameterSetName = 'Standard')]
            [parameter(ParameterSetName = 'ReplacePrefix')]
            [System.Management.Automation.Runspaces.PSSession]$TargetExchangeOrganizationSession
            ,
            [parameter(ParameterSetName = 'NewPrefix')]
            [parameter(ParameterSetName = 'Standard')]
            [parameter(ParameterSetName = 'ReplacePrefix')]
            [hashtable]$AliasHashtable
            ,
            [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
            [string]$ReplacementPrefix
            ,
            [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
            [string]$SourcePrefix
            ,
            [parameter(ParameterSetName = 'NewPrefix',Mandatory=$true)]
            [string]$NewPrefix
            ,
            [parameter(ParameterSetName = 'NewPrefix',Mandatory=$true)]
            [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
            [switch]$PrefixOnlyIfNecessary

        )
        $Alias = $SourceAlias
        $Alias = $Alias -replace '\s|[^.0-9a-zA-Z_-]|\.$',''
        $Alias = $Alias -replace '\*',''
        switch ($PSCmdlet.ParameterSetName)
        {
            'ReplacePrefix'
            {
                $NewAlias = $Alias -replace "\b$($sourcePrefix)_",''
                $NewAlias = $NewAlias -replace "\b$($SourcePrefix)", ''
                $NewAlias = $NewAlias -replace "$($SourcePrefix)\b", ''
                if ($false -eq $PrefixOnlyIfNecessary) {
                    $NewAlias = "$($ReplacementPrefix)_$($NewAlias)"
                }
                $Alias = $NewAlias
            }
            'NewPrefix'
            {
                if ($PrefixOnlyIfNecessary -eq $true)
                {
                    $TestExchangeAliasParams = @{
                        Alias = $Alias
                    }
                    if ($PSBoundParameters.ContainsKey('TargetExchangeOrganizationSession'))
                    {
                        $TestExchangeAliasParams.ExchangeSession = $TargetExchangeOrganizationSession
                    }
                    elseif ($PSBoundParameters.ContainsKey('AliasHashtable'))
                    {
                        $TestExchangeAliasParams.AliasHashtable = $AliasHashtable
                    }
                    if (-not (Test-ExchangeAlias @TestExchangeAliasParams))
                    {
                        $Alias = $NewPrefix + '_' + $Alias
                    }
                }
                else
                {
                    $Alias = $NewPrefix + '_' + $Alias
                }
            }
            'Standard'
            {
                $Alias = $SourceAlias
            }
        }
        Try
        {
            $TestExchangeAliasParams = @{
                Alias = $Alias
            }
            if ($PSBoundParameters.ContainsKey('TargetExchangeOrganizationSession'))
            {
                $TestExchangeAliasParams.ExchangeSession = $TargetExchangeOrganizationSession
            }
            elseif ($PSBoundParameters.ContainsKey('AliasHashtable'))
            {
                $TestExchangeAliasParams.AliasHashtable = $AliasHashtable
            }
            if (Test-ExchangeAlias @TestExchangeAliasParams)
            {
                $Alias
            }
            else
            {
                $(new-guid).guid
                #throw "Desired Alias $Alias, derived from Source Alias $SourceAlias is not available."
            }
        }
        Catch
        {
            $(new-guid).guid            
        }    
    }
