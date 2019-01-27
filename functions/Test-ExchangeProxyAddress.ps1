Function Test-ExchangeProxyAddress
{        
    [cmdletbinding(DefaultParameterSetName = 'ExchangeSession')]
    param
    (
        [string]$ProxyAddress
        ,
        [string[]]$ExemptObjectGUIDs
        ,
        [switch]$ReturnConflicts
        ,
        [parameter(Mandatory = $true, ParameterSetName = 'ExchangeSession')]
        [System.Management.Automation.Runspaces.PSSession]$ExchangeSession
        ,
        [parameter(Mandatory = $true, ParameterSetName = 'Hashtable')]
        [hashtable]$ProxyAddressHashtable
        ,
        [parameter()]
        [ValidateSet('SMTP', 'X500')]
        [string]$ProxyAddressType = 'SMTP'
    )
    if ($ProxyAddress -like "$($proxyaddresstype):*")
    {
        $ProxyAddress = $ProxyAddress.Split(':')[1]
    }
    #Test the ProxyAddress
    $ReturnedObjects = @(
        switch ($PSCmdlet.ParameterSetName)
        {
            'ExchangeSession'
            {
                try
                {
                    invoke-command -Session $ExchangeSession -ScriptBlock {Get-Recipient -identity $using:ProxyAddress -ErrorAction Stop} -ErrorAction Stop
                    Write-Verbose -Message "Existing object(s) Found for Alias $ProxyAddress"
                }
                catch
                {
                    if ($_.categoryinfo -like '*ManagementObjectNotFoundException*')
                    {
                        Write-Verbose -Message "No existing object(s) Found for Alias $ProxyAddress"
                    }
                    else
                    {
                        throw($_)
                    }
                }
            }
            'Hashtable'
            {
                if ($ProxyAddressHashtable.ContainsKey($ProxyAddress))
                {
                    $ProxyAddressHashtable.$ProxyAddress
                    Write-Verbose -Message "Existing object(s) Found for Alias $ProxyAddress"
                }
                else
                {
                    Write-Verbose -Message "No existing object(s) Found for Alias $ProxyAddress"
                }
            }
        }
    )
    if ($ReturnedObjects.Count -ge 1)
    {
        $ConflictingGUIDs = @($ReturnedObjects | ForEach-Object {$_.guid.guid} | Where-Object {$_ -notin $ExemptObjectGUIDs})
        if ($ConflictingGUIDs.count -gt 0)
        {
            if ($ReturnConflicts)
            {
                Return $ConflictingGUIDs
            }
            else
            {
                $false
            }
        }
        else
        {
            $true
        }
    }
    else
    {
        $true
    }

}
