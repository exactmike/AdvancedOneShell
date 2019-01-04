    Function Add-LicenseToMSOLUser {
        
        [cmdletbinding()]
        param(
        [parameter(Mandatory)]
        [ValidatePattern(".:.")]
        [string]$AccountSKUID
        ,
        [parameter()]
        [string[]]$DisabledPlans
        ,
        [Parameter(Mandatory,ParameterSetName='UserPrincipalName',ValueFromPipelineByPropertyName)]
        [string[]]$UserPrincipalName
        ,
        [Parameter(Mandatory,ParameterSetName='ObjectID',ValueFromPipeline)]
        [guid[]]$ObjectID
        ,
        [Parameter(Mandatory)]
        [bool]$CheckForExchangeOnlineRecipient = $true
        ,
        [parameter(Mandatory)]
        [string]$UsageLocation
        ,
        [parameter()]
        [string]$ExchangeOrganization
        )
        <#
        DynamicParam {
                $NewDynamicParameterParams=@{
                    Name = 'ExchangeOrganization'
                    ValidateSet = @(Get-OneShellVariableValue -name 'CurrentOrgAdminProfileSystems' | Where-Object SystemType -eq 'ExchangeOrganizations' | Select-Object -ExpandProperty Name)
                    Alias = @('Org','ExchangeOrg')
                    Position = 2
                    ParameterSetName = 'Organization'
                }
                New-DynamicParameter @NewDynamicParameterParams -Mandatory $false
            }#DynamicParam
        #>
        begin
        {
            if ($PSBoundParameters['CheckForExchangeOnlineRecipient'] -eq $true -and $PSBoundParameters.ContainsKey('ExchangeOrganization') -eq $false)
            {throw "ExchangeOrganization parameter required when -CheckForExchangeOnlineRecipient is True"}
            $newLicenseOptionsParams = @{
                AccountSkuID = $AccountSKUID
                ErrorAction = 'Stop'
            }
            if ($PSBoundParameters.ContainsKey('DisabledPlans'))
            {
                $newLicenseOptionsParams.DisabledPlans = $DisabledPlans
            }
            try
            {
                $message = "Build License Options Object"
                Write-OneShellLog -Message $message -EntryType Attempting
                $LicenseOptions = New-MsolLicenseOptions @newLicenseOptionsParams
                Write-OneShellLog -Message $message -EntryType Succeeded
            }
            catch
            {
                $myerror = $_
                Write-OneShellLog -Message $message -EntryType Failed -Verbose
                Write-OneShellLog -Message $_.tostring() -ErrorLog -Verbose
                throw("Failed:$message")
            }
            $IdentityParameter = $PSCmdlet.ParameterSetName
        }
        process
        {
            switch ($PSCmdlet.ParameterSetName)
            {
                'UserPrincipalName'
                {
                    $Identities = @($UserPrincipalName)
                    $GetMSOLUserParams = @{
                        UserPrincipalName = ''
                    }
                }

                'ObjectID'
                {
                    $Identities = @($ObjectID)
                    $GetMSOLUserParams = @{
                        ObjectID = ''
                    }
                }
            }
            $GetMSOLUserParams.ErrorAction = 'Stop'
            :nextID foreach ($ID in $Identities)
            {
                try
                {
                    $GetMSOLUserParams.$IdentityParameter = $ID.ToString()
                    $message = "Get MSOL User Object for $ID"
                    Write-OneShellLog -Message $message -EntryType Attempting
                    $MSOLUser = Get-MsolUser @GetMSOLUserParams
                    Write-OneShellLog -Message $message -EntryType Succeeded
                }
                catch
                {
                    $myerror = $_
                    Write-OneShellLog -Message $message -EntryType Failed -Verbose
                    Write-OneShellLog -Message $_.tostring() -ErrorLog -Verbose
                    continue nextID
                }
                if ($CheckForExchangeOnlineRecipient -eq $true)
                {
                    $message = "Lookup Exchange Online Recipient for Identity $ID"
                    $getRecipientParams = @{
                        Identity = $MSOLUser.objectID.guid
                        ErrorAction = 'Stop'
                    }
                    try
                    {
                        Write-OneShellLog -Message $message -EntryType Attempting
                        $EOLRecipient = Invoke-ExchangeCommand -cmdlet Get-Recipient -ExchangeOrganization $psboundparameters['exchangeOrganization'] -splat $getRecipientParams -ErrorAction Stop
                        Write-OneShellLog -Message $message -EntryType Succeeded
                    }
                    catch
                    {
                        $myerror = $_
                        Write-OneShellLog -Message $message -EntryType Failed -ErrorLog
                        Write-OneShellLog -Message $_.tostring() -ErrorLog -Verbose
                        continue nextID
                    }
                }
                $AssignedLicenseAccountSKUIDs = @($MSOLUser.licenses | Select-Object -ExpandProperty AccountSkuID)
                if ($AccountSKUID -notin $AssignedLicenseAccountSKUIDs)
                {
                    if ($MSOLUser.UsageLocation -eq $null)
                    {
                        $message = "UsageLocation for $ID is current NULL"
                        Write-OneShellLog -Message $message -EntryType Notification
                        $setMSOLUserParams = @{
                            ObjectID = $MSOLUser.ObjectID.guid
                            UsageLocation = $UsageLocation
                            ErrorAction = 'Stop'
                        }
                        $message = "Set UsageLocation for $ID to $UsageLocation"
                        try
                        {
                            Write-OneShellLog -Message $message -EntryType Attempting
                            Set-MsolUser @setMSOLUserParams
                            Write-OneShellLog -Message $message -EntryType Succeeded
                        }
                        catch
                        {
                            $myerror = $_
                            Write-OneShellLog -Message $message -EntryType Failed -ErrorLog
                            Write-OneShellLog -Message $_.tostring() -ErrorLog -Verbose
                            continue nextID
                        }
                    }#if usage location is null
                    $message = "Add $AccountSKUID license to MSOL User $ID"
                    $setMSOLUserLicenseParams = @{
                        ObjectID = $MSOLUser.ObjectID.guid
                        LicenseOptions = $LicenseOptions
                        AddLicenses = $AccountSKUID
                        ErrorAction = 'Stop'
                    }
                    try
                    {
                        Write-OneShellLog -Message $message -EntryType Attempting
                        Set-MsolUserLicense @setMSOLUserLicenseParams
                        Write-OneShellLog -Message $message -EntryType Succeeded
                    }
                    catch
                    {
                        $myerror = $_
                        Write-OneShellLog -Message $message -EntryType Failed -ErrorLog
                        Write-OneShellLog -Message $_.tostring() -ErrorLog -Verbose
                    }
                }
            }
        }
    
    }
