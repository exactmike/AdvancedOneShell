    Function New-ResourceMailboxIntermediateObject {
        
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory)]
        [psobject[]]$Resource
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
        ,
        [parameter(Mandatory)]
        [string]$AliasFormula = '$_.mail.split("@")[0]'
        ,
        [parameter()]
        [switch]$PrefixOnlyIfNecessary
        ,
        [parameter()]
        [switch]$PreserveCurrentProxyAddresses
    )
    $IntermediateObjects = @(
        :nextResource foreach ($r in $Resource)
        {
            $FriendlyIdentity = $r.mail
            $message = "Get New Alias for $FriendlyIdentity"
            try
            {
                Write-OneShellLog -Message $message -EntryType Attempting
                $DesiredAlias = GetDesiredValueFromSourceObject -Formula $AliasFormula -InputObject $r
                $GetDesiredTargetAliasParams = @{
                    SourceAlias = $DesiredAlias
                    TargetExchangeOrganizationSession = $TargetExchangeOrganizationSession
                    NewPrefix = $NewPrefix
                    ErrorAction = 'Stop'
                    PrefixOnlyIfNecessary = $PrefixOnlyIfNecessary
                }
                $DesiredAlias = Get-DesiredTargetAlias @GetDesiredTargetAliasParams
                $Prefixed = $($DesiredAlias -like "$($NewPrefix + '_*')")
                Write-OneShellLog -Message $message -EntryType Succeeded -Verbose
                Write-OneShellLog -Message "New Alias for $FriendlyIdentity is $DesiredAlias" -EntryType Notification -Verbose
            }
            catch
            {
                $myerror = $_
                Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-OneShellLog -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            $message = "Get New SMTP Proxy Address for $FriendlyIdentity"
            try
            {
                Write-OneShellLog -Message $message -EntryType Attempting
                $DesiredNewProxyAddress = Get-DesiredTargetPrimarySMTPAddress -TargetExchangeOrganizationSession $TargetExchangeOrganizationSession -DesiredAlias $DesiredAlias -TargetSMTPDomain $TargetSMTPDomain -Verbose -ErrorAction Stop
                Write-OneShellLog -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-OneShellLog -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            $message = "Get All Desired Proxy Addresses for $FriendlyIdentity $($r.ExchangeGUID)"
            try
            {
                Write-OneShellLog -Message $message -EntryType Attempting
                $GetDesiredProxyAddressesParams = @{
                    DesiredOrCurrentAlias=$DesiredAlias
                    TargetDeliveryDomain=$TargetDeliveryDomain
                    VerifyAddTargetAddress=$true
                    VerifySMTPAddressValidity=$true
                    ErrorAction = 'Stop'
                }
                if ($PreserveCurrentProxyAddresses)
                {
                    $GetDesiredProxyAddressesParams.CurrentProxyAddresses=$r.EmailAddresses
                    $GetDesiredProxyAddressesParams.LegacyExchangeDNs=$r.LegacyExchangeDN
                    $GetDesiredProxyAddressesParams.AddressesToAdd="smtp:$DesiredNewProxyAddress"
                }
                else
                {
                    $GetDesiredProxyAddressesParams.DesiredPrimaryAddress=$DesiredNewProxyAddress
                }
                if ($DomainsToRemove.Count -gt 0) {$GetDesiredProxyAddressesParams.DomainsToRemove = $DomainsToRemove}
                $DesiredAddresses = Get-DesiredProxyAddresses @GetDesiredProxyAddressesParams
                $DesiredTargetAddress = $DesiredAddresses | Where-Object -FilterScript {$_ -like "*@$TargetDeliveryDomain"} | ForEach-Object {$_.split(':')[1]}
                Write-OneShellLog -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-OneShellLog -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            $message = "Check All Desired Proxy Addresses for $FriendlyIdentity for conflicts in target Exchange Organization"
            try
            {
                Write-OneShellLog -Message $message -EntryType Attempting
                $AnyConflicts = @(
                foreach ($a in $DesiredAddresses)
                {
                    $result = Test-ExchangeProxyAddress -ProxyAddress $a -ReturnConflicts -ExchangeSession $TargetExchangeOrganizationSession -ErrorAction Stop
                    if ($result -ne $true)
                    {
                        $result
                    }
                })
                Write-OneShellLog -Message $message -EntryType Succeeded -Verbose
                if ($AnyConflicts.Count -gt 0)
                {
                    $conflictingGUIDsString = $AnyConflicts -join '|'
                    throw "$FriendlyIdentity has conflicts in target Exchange Organization with the following objects:  $conflictingGUIDsString"
                }
            }
            catch
            {
                $myerror = $_
                Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-OneShellLog -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            $message = "Get Desired Name for $FriendlyIdentity"
            try
            {
                $GetDesiredTargetNameParams = @{
                    SourceName = $DesiredAlias
                    ErrorAction = 'Stop'
                }
                if ($PrefixOnlyIfNecessary -eq $false)
                {
                    $GetDesiredTargetNameParams.NewPrefix = $NewPrefix
                }
                $DesiredName = Get-DesiredTargetName @GetDesiredTargetNameParams
                Write-OneShellLog -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-OneShellLog -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            $message = "Get Desired Display Name for $FriendlyIdentity"
            try
            {
                $GetDesiredTargetNameParams = @{
                    SourceName = $r.DisplayName
                    ErrorAction = 'Stop'
                }
                if ($PrefixOnlyIfNecessary -eq $false)
                {
                    $GetDesiredTargetNameParams.NewPrefix = $NewPrefix
                }
                $DesiredDisplayName = Get-DesiredTargetName @GetDesiredTargetNameParams
                Write-OneShellLog -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-OneShellLog -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextResource
            }
            #need to update this code to propery specify and convert objects
            $message = "Check $FriendlyIdentity RecipientTypeDetails $($r.RecipientTypeDetails) and Convert to SharedMailbox if needed"
            Write-OneShellLog -Message $message -EntryType Notification
            $RecipientTypeDetails = Get-RecipientType -msExchRecipientTypeDetails $r.msExchRecipientTypeDetails
            if ($RecipientTypeDetails.Name -like '*User*')
            {
                $message = "Convert $FriendlyIdentity to SharedMailbox from $RecipienttypeDetails"
                Write-OneShellLog -Message $message -EntryType Notification
                $RTD = 'RemoteSharedMailbox'
            }
            else
            {
                $message = "Preserve $FriendlyIdentity as $RecipientTypeDetails"
                Write-OneShellLog -Message $message -EntryType Notification
                $RTD = $RecipientTypeDetails.Name
            }
            $message = "Build Intermediate Object to use for creation of target object for $FriendlyIdentity"
            Write-OneShellLog -Message $message -EntryType Notification
            $SAMLength = [math]::Min($desiredAlias.length,15)
            $IntermediateObject=[pscustomobject]@{
                #msExchPoliciesExcluded = '{26491cfc-9e50-4857-861b-0cb8df22b5d7}'
                #msExchMailboxGUID = $($r.ExchangeGuid)
                Mail = $DesiredNewProxyAddress
                TargetAddress = 'SMTP:' + $DesiredTargetAddress
                mailNickName = $DesiredAlias
                SamAccountName = $DesiredAlias.substring(0,$SAMLength) + $r.ObjectGUID.Guid.substring(0,5)
                proxyaddresses = [string[]]$DesiredAddresses
                Name = $DesiredName
                DisplayName = $DesiredDisplayName
                UserPrincipalName = $DesiredNewProxyAddress
                SourceIdentity = $FriendlyIdentity
                SourceObjectType = $r.ObjectClass
                employeeType = 'Resource'
                description = "Resource: $RTD"
                ResourceType = $RTD
                msExchRecipientDisplayType = $null
                msExchRecipientTypeDetails = $null
                msExchVersion = 88218628259840
                msExchUsageLocation = 'US'
                c = 'US'
                co = 'United States'
                extensionattribute5 = $r.ObjectGuid.Guid
                Prefixed = $Prefixed
            }
            Write-Output -InputObject $IntermediateObject
        }
        )
    Write-Output -InputObject $IntermediateObjects

    }
