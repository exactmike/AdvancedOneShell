Function Test-ExchangeAlias 
{
    [cmdletbinding(DefaultParameterSetName = 'ExchangeSession')]
    param(
        [string]$Alias
        ,
        [string[]]$ExemptObjectGUIDs
        ,
        [switch]$ReturnConflicts
        ,
        [parameter(Mandatory = $true, ParameterSetName = 'ExchangeSession')]
        [System.Management.Automation.Runspaces.PSSession]$ExchangeSession
        ,
        [parameter(Mandatory = $true, ParameterSetName = 'Hashtable')]
        [hashtable]$AliasHashtable
    )
    #Test the Alias
    $ReturnedObjects = @(
        switch ($PSCmdlet.ParameterSetName)
        {
            'ExchangeSession'
            {
                try
                {
                    invoke-command -Session $ExchangeSession -ScriptBlock {Get-Recipient -identity $using:Alias -ErrorAction Stop} -ErrorAction Stop
                    Write-Verbose -Message "Existing object(s) Found for Alias $Alias"
                }
                catch {
                    if ($_.categoryinfo -like '*ManagementObjectNotFoundException*') {
                        Write-Verbose -Message "No existing object(s) Found for Alias $Alias"
                    }
                    else {
                        throw($_)
                    }
                }    
            }
            'Hashtable'
            {
                if ($AliasHashtable.ContainsKey($Alias))
                {
                    $AliasHashtable.$Alias
                    Write-Verbose -Message "Existing object(s) Found for Alias $Alias"
                }
                else
                {
                    Write-Verbose -Message "No existing object(s) Found for Alias $Alias"
                }
            }
        }
    )
    if ($ReturnedObjects.Count -ge 1) {
        $ConflictingGUIDs = @($ReturnedObjects | ForEach-Object {$_.guid.guid} | Where-Object {$_ -notin $ExemptObjectGUIDs})
        if ($ConflictingGUIDs.count -gt 0) {
            if ($ReturnConflicts) {
                Return $ConflictingGUIDs
            }
            else {
                $false
            }
        }
        else {
            $true
        }
    }
    else {
        $true
    }

}
