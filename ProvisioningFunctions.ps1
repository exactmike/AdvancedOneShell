function Get-DesiredTargetAlias
    {
        [cmdletbinding()]
        param
        (
            [parameter(ParameterSetName = 'NewPrefix',Mandatory=$true)]
            [parameter(ParameterSetName = 'Standard',Mandatory=$true)]
            [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
            $SourceAlias
            ,
            [parameter(ParameterSetName = 'NewPrefix',Mandatory=$true)]
            [parameter(ParameterSetName = 'Standard',Mandatory=$true)]
            [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
            $TargetExchangeOrganization
            ,
            [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
            [string]$ReplacementPrefix
            ,
            [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
            [string]$SourcePrefix
            ,
            [parameter(ParameterSetName = 'NewPrefix',Mandatory=$true)]
            [string]$NewPrefix
        )
        $Alias = $SourceAlias
        $Alias = $Alias -replace '\s|[^1-9a-zA-Z_-]',''
        switch ($PSCmdlet.ParameterSetName)
        {
            'ReplacePrefix'
            {
                $NewAlias = $Alias -replace "\b$($sourcePrefix)_",''
                $NewAlias = $NewAlias -replace "\b$($SourcePrefix)", ''
                $NewAlias = $NewAlias -replace "$($SourcePrefix)\b", ''
                $NewAlias = "$($ReplacementPrefix)_$($NewAlias)"
                $Alias = $NewAlias
            }
            'NewPrefix'
            {
                $Alias = $NewPrefix + '_' + $Alias
            }
            'Standard'
            {
                $Alias = $SourceAlias
            }
        }
        if (Test-ExchangeAlias -Alias $Alias -ExchangeOrganization $TargetExchangeOrganization) 
        {
            $Alias
        }
        else 
        {
            throw "Desired Alias $Alias, derived from Source Alias $SourceAlias is not available."
        }
    }
#end function Get-DesiredTargetAlias
function Get-DesiredTargetName
    {
        [cmdletbinding()]
        param
        (
        [parameter(ParameterSetName = 'NewPrefix',Mandatory=$true)]
        [parameter(ParameterSetName = 'Standard',Mandatory=$true)]
        [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
        $SourceName
        ,
        [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
        [string]$ReplacementPrefix
        ,
        [parameter(ParameterSetName = 'ReplacePrefix',Mandatory=$true)]
        [string]$SourcePrefix
        ,
        [parameter(ParameterSetName = 'NewPrefix',Mandatory=$true)]
        [string]$NewPrefix
        )
        $Name = $SourceName
        $Name = $Name -replace '|[^1-9a-zA-Z_-]',''
        switch ($PSCmdlet.ParameterSetName)
        {
            'ReplacePrefix'
            {
                $NewName = $Name -replace "\b$($sourcePrefix)_",''
                $NewName = $NewName -replace "\b$($SourcePrefix)", ''
                $NewName = $NewName -replace "$($SourcePrefix)\b", ''
                $NewName = "$($ReplacementPrefix)_$($NewName)"
                $Name = $NewName.Trim()
            }
            'NewPrefix'
            {
                $Name = $NewPrefix + '_' + $Name
            }
            'Standard'
            {
            }
        }
        Write-Output -InputObject $Name
    }
#end function Get-DesiredTargetName
function Get-DesiredTargetPrimarySMTPAddress
    {
        [cmdletbinding()]
        param
        (
        [parameter(ParameterSetName = 'Standard',Mandatory=$true)]
        $DesiredAlias
        ,
        [parameter(ParameterSetName = 'Standard',Mandatory=$true)]
        $TargetExchangeOrganization
        ,
        [parameter(ParameterSetName = 'Standard',Mandatory=$true)]
        [string]$TargetSMTPDomain
        )
        $DesiredPrimarySMTPAddress = $DesiredAlias + '@' + $TargetSMTPDomain

        if (Test-ExchangeProxyAddress -ProxyAddress $DesiredPrimarySMTPAddress -ExchangeOrganization $TargetExchangeOrganization -ProxyAddressType SMTP)
        {
            $DesiredPrimarySMTPAddress
        }
        else 
        {
            throw "Desired Primary SMTP Address $DesiredPrimarySMTPAddress is not available."
        }
    }
#end function Get-DesiredTargetPrimarySMTPAddress
function New-ResourceMailboxIntermediateObject
    {
        [cmdletbinding()]
        param
        (
            [parameter(Mandatory)]
            [psobject]$resource
            ,
            [parameter(Mandatory)]
            [System.Management.Automation.Runspaces.PSSession]$TargetExchangeOrganizationSession
            ,
            [parameter(Mandatory)]
            [string]$NewPrefix
            ,
            [parameter(Mandatory)]
            [string]$TargetSMTPDomain
            ,
            [parameter(Mandatory)]
            [string]$TargetDeliveryDomain
            ,
            [parameter()]
            [string[]]$DomainsToRemove
        )
        $IntermediateObjects = @(
        :nextResource foreach ($r in $resource)
        {
            $FriendlyIdentity = $r.PrimarySmtpAddress
            $message = "Get New Alias for $FriendlyIdentity $($r.ExchangeGUID)"
            try 
            {
                Write-Log -Message $message -EntryType Attempting
                $DesiredAlias = Get-DesiredTargetAlias -SourceAlias $r.alias -TargetExchangeOrganizationSession $TargetExchangeOrganizationSession -Verbose -NewPrefix $NewPrefix -ErrorAction Stop
                Write-Log -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            $message = "Get New SMTP Proxy Address for $FriendlyIdentity $($r.ExchangeGUID)"
            try 
            {
                Write-Log -Message $message -EntryType Attempting
                $DesiredNewProxyAddress = Get-DesiredTargetPrimarySMTPAddress -TargetExchangeOrganization $TargetExchangeOrganization -DesiredAlias $DesiredAlias -TargetSMTPDomain $TargetSMTPDomain -Verbose -ErrorAction Stop
                Write-Log -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            $message = "Get All Desired Proxy Addresses for $FriendlyIdentity $($r.ExchangeGUID)"
            try 
            {
                Write-Log -Message $message -EntryType Attempting
                $GetDesiredProxyAddressesParams = @{
                    CurrentProxyAddresses=$r.EmailAddresses
                    DesiredOrCurrentAlias=$DesiredAlias
                    DesiredPrimaryAddress=$r.PrimarySmtpAddress
                    LegacyExchangeDNs=$r.LegacyExchangeDN
                    TargetDeliveryDomain=$TargetDeliveryDomain
                    VerifyAddTargetAddress=$true
                    VerifySMTPAddressValidity=$true
                    AddressesToAdd="smtp:$DesiredNewProxyAddress"
                    ErrorAction = 'Stop'
                }
                if ($DomainsToRemove.Count -gt 0) {$GetDesiredProxyAddressesParams.DomainsToRemove = $DomainsToRemove}
                $DesiredAddresses = Get-DesiredProxyAddresses @GetDesiredProxyAddressesParams
                Write-Log -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            $message = "Check All Desired Proxy Addresses for $FriendlyIdentity $($r.ExchangeGUID) for conflicts in target Exchange Organization $TargetExchangeOrganization"
            try 
            {
                Write-Log -Message $message -EntryType Attempting
                $AnyConflicts = @(
                foreach ($a in $DesiredAddresses)
                {
                    $result = Test-ExchangeProxyAddress -ProxyAddress $a -ReturnConflicts -ExchangeOrganization $TargetExchangeOrganization -ErrorAction Stop
                    if ($result -ne $true)
                    {
                        $result
                    }
                })
                Write-Log -Message $message -EntryType Succeeded -Verbose
                if ($AnyConflicts.Count -gt 0)
                {
                    $conflictingGUIDsString = $AnyConflicts -join '|'
                    throw "$FriendlyIdentity $($r.ExchangeGUID) has conflicts in target Exchange Organization $TargetExchangeOrganization with the following objects:  $conflictingGUIDsString"
                }
            }
            catch
            {
                $myerror = $_
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            $message = "Get Desired Name for $FriendlyIdentity $($r.ExchangeGUID)"
            try 
            {
                $DesiredName = Get-DesiredTargetName -NewPrefix $NewPrefix -SourceName $r.Name -Verbose -ErrorAction Stop
                Write-Log -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            $message = "Check $FriendlyIdentity $($r.ExchangeGUID) RecipientTypeDetails $($r.RecipientTypeDetails) and Convert to SharedMailbox if needed"
            Write-Log -Message $message -EntryType Notification
            if ($r.RecipientTypeDetails -eq 'UserMailbox')
            {
                $message = "Convert $FriendlyIdentity $($r.ExchangeGUID) to SharedMailbox from $($r.RecipientTypeDetails)"
                Write-Log -Message $message -EntryType Notification
                $RTD = 'SharedMailbox'
            }
            else
            {
                $message = "Preserve $FriendlyIdentity $($r.ExchangeGUID) as $($r.RecipientTypeDetails)"
                Write-Log -Message $message -EntryType Notification
                $RTD = $r.RecipientTypeDetails
            }
            $message = "Get msExchRecipientDisplayType to use for $FriendlyIdentity $($r.ExchangeGUID)"
            try 
            {
                Write-Log -Message $message -EntryType Attempting
                $msExchRecipientDisplayType = Get-msExchRecipientDisplayTypeValue -RecipientTypeDetails $RTD -ErrorAction Stop
                Write-Log -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            $message = "Get msExchRecipientTypeDetails to use for $FriendlyIdentity $($r.ExchangeGUID)"
            try 
            {
                $msExchRecipientTypeDetails = Get-msExchRecipientTypeDetailsValue -RecipientTypeDetails $RTD -ErrorAction Stop
                Write-Log -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            $message = "Get msExchRemoteRecipientType to use for $FriendlyIdentity $($r.ExchangeGUID)"
            try 
            {
                $msExchRemoteRecipientType = Get-msExchRemoteRecipientTypeValue -RecipientTypeDetails $RTD -ErrorAction Stop
                Write-Log -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            $message = "Build Intermediate Object to use for creation of target object for $FriendlyIdentity $($r.ExchangeGUID)"
            Write-Log -Message $message -EntryType Notification
            $SAMLength = [math]::Min($desiredAlias.length,20)
            $IntermediateObject=[pscustomobject]@{
                msExchRecipientDisplayType = $msExchRecipientDisplayType
                msExchRecipientTypeDetails = $msExchRecipientTypeDetails
                msExchVersion = 44220983382016
                msExchUsageLocation = 'US'
                c = 'US'
                co = 'United States'
                extensionattribute5 = $r.Guid.Guid
                msExchPoliciesExcluded = '{26491cfc-9e50-4857-861b-0cb8df22b5d7}'
                msExchMailboxGUID = $($r.ExchangeGuid)
                Mail = $r.PrimarySmtpAddress
                TargetAddress = 'SMTP:' + $r.PrimarySmtpAddress
                mailNickName = $DesiredAlias
                SamAccountName = $DesiredAlias.substring(0,$SAMLength)
                proxyaddresses = [string[]]$DesiredAddresses
                Name = $DesiredName
                DisplayName = $DesiredName
                UserPrincipalName = $r.PrimarySmtpAddress
                employeeType = 'Shared Mailbox'
            }
            Write-Output -InputObject $IntermediateObject
        }
        )
        Write-Output -InputObject $IntermediateObjects
    }
#end function New-ResourceMailboxIntermediateObject
function Publish-ResourceObjects
    {
        :nextI foreach ($i in $IntermediateResourceObjects)
        {
            $message = "Create AD User Object for $($I.UserPrincipalName) $($I.msExchMailboxGUID.guid)"
            try 
            {
                Write-Log -Message $message -EntryType Attempting
                Push-Location
                Set-Location -Path $($targetActiveDirectory + ':\')
                $IHash = Convert-ObjectToHashTable -InputObject $I -NoEmpty -Exclude SAMAccountName -ErrorAction Stop
                $newADUser = New-ADUser -Path $targetUserOUDN -Server $targetDomain -Enabled:$false -OtherAttributes $IHash -Name $I.Name -ErrorAction Stop -SamAccountName $I.SamAccountName -PassThru #-WhatIf 
                Write-Log -Message $message -EntryType Succeeded -Verbose
                Pop-Location
            }
            catch
            {
                Pop-Location
                $myerror = $_
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextI
            }
            $message = "Add New Proxy Address and New Alias to Exchange Alias and Proxy Address Test tables"
            try 
            {
                Write-Log -Message $message -EntryType Attempting
                Add-ExchangeProxyAddressToTestExchangeProxyAddress -ProxyAddress $($i.mailNickName + '@' + $TargetSMTPDomain) -ObjectGUID $i.msExchMailboxGUID.Guid -ProxyAddressType SMTP
                Add-ExchangeAliasToTestExchangeAlias -Alias $i.mailNickName -ObjectGUID $i.msExchMailboxGUID.Guid 
                Write-Log -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextI
            }
            $message = "Add TargetDeliveryAddress $($i.mailNickName + "@$TargetDeliveryDomain") to Source Object $($i.UserPrincipalName) $($i.msExchMailboxGUID) "
            try 
            {
                Write-Log -Message $message -EntryType Attempting
                $AddEmailAddressParams = @{
                    ExchangeOrganization=$SourceExchangeOrganization
                    Identity=$i.msExchMailboxGUID
                    EmailAddresses=$($i.mailNickName + "@$TargetDeliveryDomain")
                    ErrorAction='Stop'
                }
                Add-EmailAddress @AddEmailAddressParams
                Write-Log -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-Log -Message $myerror.tostring() -ErrorLog -Verbose
            }
        }
    }
#end function Publish-ResourceObjects