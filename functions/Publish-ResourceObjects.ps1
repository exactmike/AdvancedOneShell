    Function Publish-ResourceObjects {
        
        :nextI foreach ($i in $IntermediateResourceObjects)
        {
            $message = "Create AD User Object for $($I.UserPrincipalName) $($I.msExchMailboxGUID.guid)"
            try
            {
                Write-OneShellLog -Message $message -EntryType Attempting
                Push-Location
                Set-Location -Path $($targetActiveDirectory + ':\')
                $IHash = Convert-ObjectToHashTable -InputObject $I -NoEmpty -Exclude SAMAccountName -ErrorAction Stop
                $newADUser = New-ADUser -Path $targetUserOUDN -Server $targetDomain -Enabled:$false -OtherAttributes $IHash -Name $I.Name -ErrorAction Stop -SamAccountName $I.SamAccountName -PassThru #-WhatIf
                Write-OneShellLog -Message $message -EntryType Succeeded -Verbose
                Pop-Location
            }
            catch
            {
                Pop-Location
                $myerror = $_
                Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-OneShellLog -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextI
            }
            $message = "Add New Proxy Address and New Alias to Exchange Alias and Proxy Address Test tables"
            try
            {
                Write-OneShellLog -Message $message -EntryType Attempting
                Add-ExchangeProxyAddressToTestExchangeProxyAddress -ProxyAddress $($i.mailNickName + '@' + $TargetSMTPDomain) -ObjectGUID $i.msExchMailboxGUID.Guid -ProxyAddressType SMTP
                Add-ExchangeAliasToTestExchangeAlias -Alias $i.mailNickName -ObjectGUID $i.msExchMailboxGUID.Guid
                Write-OneShellLog -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-OneShellLog -Message $myerror.tostring() -ErrorLog -Verbose
                continue nextI
            }
            $message = "Add TargetDeliveryAddress $($i.mailNickName + "@$TargetDeliveryDomain") to Source Object $($i.UserPrincipalName) $($i.msExchMailboxGUID) "
            try
            {
                Write-OneShellLog -Message $message -EntryType Attempting
                $AddEmailAddressParams = @{
                    ExchangeOrganization=$SourceExchangeOrganization
                    Identity=$i.msExchMailboxGUID
                    EmailAddresses=$($i.mailNickName + "@$TargetDeliveryDomain")
                    ErrorAction='Stop'
                }
                Add-EmailAddress @AddEmailAddressParams
                Write-OneShellLog -Message $message -EntryType Succeeded -Verbose
            }
            catch
            {
                $myerror = $_
                Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
                Write-OneShellLog -Message $myerror.tostring() -ErrorLog -Verbose
            }
        }
    
    }
