    Function Get-DesiredTargetName {
        
        [cmdletbinding()]
        param
        (
        [parameter(ParameterSetName = 'NewPrefix',Mandatory=$true)]
        [parameter(ParameterSetName = 'Standard',Mandatory=$true)]
        [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
        $SourceName
        ,
        [parameter(ParameterSetName = 'ReplacePrefix')]
        [string]$ReplacementPrefix
        ,
        [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
        [string]$SourcePrefix
        ,
        [parameter(ParameterSetName = 'NewPrefix',Mandatory=$true)]
        [string]$NewPrefix
        )
        $Name = $SourceName
        $Name = ($Name -replace '|[^1-9a-zA-Z_-]','') -replace '\*',''
        switch ($PSCmdlet.ParameterSetName)
        {
            'ReplacePrefix'
            {
                $NewName = $Name -replace "\b$($sourcePrefix)_",''
                $NewName = $NewName -replace "\b$($SourcePrefix)", ''
                $NewName = $NewName -replace "$($SourcePrefix)\b", ''
                if ($null -ne $ReplacementPrefix -and -not [string]::IsNullOrEmpty($ReplacementPrefix))
                {$NewName = "$($ReplacementPrefix)_$($NewName)"}
                $Name = $NewName.Trim()
            }
            'NewPrefix'
            {
                $Name = $NewPrefix + '_' + $Name
            }
            'Standard'
            {
                #nothing needed here
            }
        }
        $Name
    
    }
