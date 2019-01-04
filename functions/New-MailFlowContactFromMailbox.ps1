    Function New-MailFlowContactFromMailbox {
        
        [cmdletbinding()]
        param(
        [parameter(Mandatory)]
        [string]$Identity
        ,
        [parameter(Mandatory)]
        [string]$SourceExchangeOrganization
        ,
        [parameter(Mandatory)]
        [string]$TargetExchangeOrganization
        ,
        [parameter(Mandatory)]
        [string]$OrganizationalUnit
        ,
        [parameter()]
        [string[]]$DomainsToRemove
        ,
        [parameter()]
        [string[]]$AddressesToRemove
        ,
        [parameter()]
        [string[]]$AddressesToAdd
        ,
        [parameter()]
        [string]$TargetAddress
        ,
        [parameter()]
        [string]$TargetDeliveryDomain
        ,
        [string]$AliasPrefix
        ,
        [bool]$testOnly = $true
        ,
        [switch]$AddExternalEmailAddressToSourceObject
        )
        #region GetSourceObject
        $GetRecipientParams = @{
            cmdlet = 'Get-Recipient'
            ExchangeOrganization = $SourceExchangeOrganization
            ErrorAction = 'Stop'
            splat = @{
                Identity = $Identity
                ErrorAction = 'Stop'
            }
        }
        $message = "Get Recipient Object for Identity ($Identity) from Source Exchange Organization ($SourceExchangeOrganization)"
        Try
        {
            Write-OneShellLog -Message $message -EntryType Attempting
            $SourceRecipientObject = @(Invoke-ExchangeCommand @GetRecipientParams)
            Write-OneShellLog -Message $message -EntryType Succeeded
        }
        Catch
        {
            $MyError = $_
            Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
            Write-OneShellLog -Message $myerror.tostring() -ErrorLog
            Throw $MyError
        }
        $GetFullRecipientCmdlet = Get-RecipientCmdlet -Recipient $SourceRecipientObject -verb Get -ErrorAction Stop
        $GetFullRecipientCmdletParams = @{
            cmdlet = $GetFullRecipientCmdlet
            ExchangeOrganization = $SourceExchangeOrganization
            ErrorAction = 'Stop'
            splat = @{
                Identity = $Identity
                ErrorAction = 'Stop'
            }
        }
        Try
        {
            Write-OneShellLog -Message $message -EntryType Attempting
            $SourceRecipientObject = @(Invoke-ExchangeCommand @GetFullRecipientCmdletParams)
            Write-OneShellLog -Message $message -EntryType Succeeded
        }
        Catch
        {
            $MyError = $_
            Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
            Write-OneShellLog -Message $myerror.tostring() -ErrorLog
            Throw $MyError
        }
        #endregion
        #region CreateIntermediateObject
        $GetDesiredTargetAliasParams = @{
            SourceAlias = $aliasPrefix + $SourceRecipientObject.Alias
            TargetExchangeOrganization = $TargetExchangeOrganization
            ErrorAction = 'Stop'
        }
        $message = "Get Desired Alias for Identity ($Identity) from Target Exchange Organization ($TargetExchangeOrganization)"
        Try
        {
            Write-OneShellLog -Message $message -EntryType Attempting
            $DesiredAlias = Get-DesiredTargetAlias @GetDesiredTargetAliasParams
            Write-OneShellLog -Message $message -EntryType Succeeded
        }
        Catch
        {
            $MyError = $_
            Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
            Write-OneShellLog -Message $myerror.tostring() -ErrorLog
            Throw $MyError
        }
        $GetDesiredProxyAddressesParams = @{
            CurrentProxyAddresses = $SourceRecipientObject.emailAddresses
            DesiredOrCurrentAlias = $DesiredAlias
            LegacyExchangeDNs = $SourceRecipientObject.LegacyExchangeDN
            VerifyAddTargetAddress = $true
            TargetDeliveryDomain = $TargetDeliveryDomain
        }
        if ($DomainsToRemove.Count -ge 1)
        {
            $GetDesiredProxyAddressesParams.DomainsToRemove = $DomainsToRemove
        }
        if ($AddressesToAdd.Count -ge 1)
        {
            $GetDesiredProxyAddressesParams.AddressesToAdd = $AddressesToAdd
        }
        if ($AddressesToRemove.Count -ge 1)
        {
            $GetDesiredProxyAddressesParams.AddressesToAdd = $AddressesToRemove
        }
        $message = "Get Desired ProxyAddresses for Identity ($Identity)"
        Try
        {
            Write-OneShellLog -Message $message -EntryType Attempting
            $DesiredProxyAddresses = Get-DesiredProxyAddresses @GetDesiredProxyAddressesParams
            Write-OneShellLog -Message $message -EntryType Succeeded
        }
        Catch
        {
            $MyError = $_
            Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
            Write-OneShellLog -Message $myerror.tostring() -ErrorLog
            Throw $MyError
        }
        $IntermediateObject = [pscustomobject]@{
            Name = $SourceRecipientObject.Name
            Alias = $DesiredAlias
            EmailAddresses = $DesiredProxyAddresses
            ExternalEmailAddress = ($DesiredProxyAddresses | ? {$_ -like $('*@' + $targetDeliveryDomain)} | Select-Object -First 1).split(':')[1]
            PrimarySMTPAddress = ($DesiredProxyAddresses | ? {$_ -clike 'SMTP:*'} | Select-Object -first 1).split(':')[1]
            DisplayName = $SourceRecipientObject.DisplayName
            OrganizationalUnit = $OrganizationalUnit
            CustomAttribute5 = $SourceRecipientObject.guid.guid
            CustomAttribute6 = 'Temporary Mail Routing Contact'
            EmailAddressPolicyEnabled = $false
            AddExternalEmailAddressToSourceObject = $AddExternalEmailAddressToSourceObject
        }
        #endregion
        #region CreateTargetObject
        if ($testOnly -eq $true)
        {
            $IntermediateObject
        }
        else
        {
            #region UpdateSourceObject
            if ($AddExternalEmailAddressToSourceObject)
            {
                Add-EmailAddress -Identity $Identity -EmailAddresses $IntermediateObject.ExternalEmailAddress -ExchangeOrganization $SourceExchangeOrganization -ErrorAction Stop
            }
            #endregion
            $newMailContactParams = @{
                Cmdlet = 'New-MailContact'
                ExchangeOrganization = $TargetExchangeOrganization
                ErrorAction = 'Stop'
                Splat = @{
                    ErrorAction = 'Stop'
                    Name = $IntermediateObject.Name
                    Alias = $IntermediateObject.Alias
                    ExternalEmailaddress = $IntermediateObject.ExternalEmailAddress
                    PrimarySMTPAddress = $IntermediateObject.PrimarySMTPAddress
                    DisplayName = $IntermediateObject.DisplayName
                    OrganizationalUnit = $IntermediateObject.OrganizationalUnit
                }
            }
            $setMailContactParams = @{
                Cmdlet = 'Set-MailContact'
                ExchangeOrganization = $TargetExchangeOrganization
                ErrorAction = 'Stop'
                Splat = @{
                    Identity = $IntermediateObject.Alias
                    ErrorAction = 'Stop'
                    EmailAddresses = $IntermediateObject.EmailAddresses
                    CustomAttribute5 = $IntermediateObject.CustomAttribute5
                    CustomAttribute6 = $IntermediateObject.CustomAttribute6
                    EmailAddressPolicyEnabled = $IntermediateObject.EmailAddressPolicyEnabled
                }
            }
            $message = "Create New Mail Contact for Identity ($identity) in Target Exchange Organization ($TargetExchangeOrganization)"
            Try
            {
                Write-OneShellLog -Message $message -EntryType Attempting
                $NewMailContactOutput = Invoke-ExchangeCommand @newMailContactParams
                Write-OneShellLog -Message $message -EntryType Succeeded
            }
            Catch
            {
                $MyError = $_
                Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-OneShellLog -Message $myerror.tostring() -ErrorLog
                Throw $MyError
            }
            $message = "Get New Mail Contact for Identity ($identity) in Target Exchange Organization ($TargetExchangeOrganization)"
            Write-OneShellLog -Message $message -EntryType Attempting
            $FindAttemptCount = 0
            Do
            {
                Start-Sleep -Seconds 5
                $GetMailContactParams = @{
                    Cmdlet = 'Get-MailContact'
                    ExchangeOrganization = $TargetExchangeOrganization
                    ErrorAction = 'SilentlyContinue'
                    Splat = @{
                        ErrorAction = 'SilentlyContinue'
                        Identity = $IntermediateObject.Alias
                    }
                }
                $NewMailContact = @(Invoke-ExchangeCommand @GetMailContactParams)
                $FindAttemptCount++
                if ($FindAttemptCount -ge 15)
                {
                    Throw "Failed: $message"
                }
            }
            Until ($NewMailContact.Count -eq 1)
            $message  = "Set additional attributes for Mail Contact for Identity ($Identity) in Target Exchange Organization ($TargetExchangeOrganization)"
            Try
            {
                Write-OneShellLog -Message $message -EntryType Attempting
                Invoke-ExchangeCommand @setMailContactParams
                Write-OneShellLog -Message $message -EntryType Succeeded
            }
            Catch
            {
                $MyError = $_
                Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-OneShellLog -Message $myerror.tostring() -ErrorLog
                Throw $MyError
            }
        }
        #endregion
    
    }
