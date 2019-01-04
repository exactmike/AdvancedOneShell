    Function Get-DesiredTargetPrimarySMTPAddress {
        
        [cmdletbinding()]
        param
        (
        [parameter(ParameterSetName = 'Standard',Mandatory=$true)]
        $DesiredAlias
        ,
        [parameter(ParameterSetName = 'Standard',Mandatory)]
        [System.Management.Automation.Runspaces.PSSession]$TargetExchangeOrganizationSession
        ,
        [parameter(ParameterSetName = 'Standard',Mandatory=$true)]
        [string]$TargetSMTPDomain
        )
        $DesiredPrimarySMTPAddress = $DesiredAlias + '@' + $TargetSMTPDomain

        if (Test-ExchangeProxyAddress -ProxyAddress $DesiredPrimarySMTPAddress -ExchangeSession $TargetExchangeOrganizationSession -ProxyAddressType SMTP)
        {
            $DesiredPrimarySMTPAddress
        }
        else
        {
            throw "Desired Primary SMTP Address $DesiredPrimarySMTPAddress is not available."
        }
    
    }
