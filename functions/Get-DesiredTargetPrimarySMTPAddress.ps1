Function Get-DesiredTargetPrimarySMTPAddress
{
    [cmdletbinding(DefaultParameterSetName = 'Hashtable')]
    param
    (
        [parameter(Mandatory)]
        $DesiredAlias
        ,
        [parameter(Mandatory)]
        [string]$TargetSMTPDomain
        ,
        [parameter(ParameterSetName = 'ExchangeSession',Mandatory)]
        [System.Management.Automation.Runspaces.PSSession]$TargetExchangeOrganizationSession
        ,
        [parameter(Mandatory, ParameterSetName = 'Hashtable')]
        [hashtable]$ProxyAddressHashtable
    )

    $DesiredPrimarySMTPAddress = $DesiredAlias + '@' + $TargetSMTPDomain
    Try
    {
        switch ($PSCmdlet.ParameterSetName)
        {
            'ExchangeSession'
            {
                if (Test-ExchangeProxyAddress -ProxyAddress $DesiredPrimarySMTPAddress -ExchangeSession $TargetExchangeOrganizationSession -ProxyAddressType SMTP)
                {
                    $DesiredPrimarySMTPAddress
                }
                else
                {
                    $(new-guid).guid + '@' + $TargetSMTPDomain
                }
        
            }
            'Hashtable'
            {
                if (-not $ProxyAddressHashtable.ContainsKey($DesiredPrimarySMTPAddress))
                {
                    $DesiredPrimarySMTPAddress
                }
                else
                {
                    $(new-guid).guid + '@' + $TargetSMTPDomain
                }
        
            }
        }
    }
    Catch
    {
        $(new-guid).guid + '@' + $TargetSMTPDomain
    }

}
