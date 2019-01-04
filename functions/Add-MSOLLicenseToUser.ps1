    Function Add-MSOLLicenseToUser {
        
        [cmdletbinding()]
        param(
            [parameter(Mandatory=$true,ParameterSetName = "Migration Wave Source Data")]
            [string]$Wave
            ,
            [parameter(ParameterSetName = "Migration Wave Source Data")]
            [ValidateSet('Full','Sub','All')]
            [string]$WaveType='Sub'
            ,
            [parameter(Mandatory=$true,ParameterSetName = "Custom Source Data")]
            $CustomSource
            ,
            [parameter(Mandatory=$true,ParameterSetName = "Single User",ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
            [string]$UserPrincipalName
            ,
            [switch]$AssignUsageLocation
            ,
            [string]$StaticUsagelocation
            ,
            [switch]$SelectServices
            ,
            [validateset('Exchange','Lync','SharePoint','OfficeWebApps','OfficeProPlus','LyncEV','AzureRMS')]
            [string[]]$DisabledServices
            ,
            [parameter(ParameterSetName = "Migration Wave Source Data")]
            [parameter(ParameterSetName = "Custom Source Data")]
            [switch]$DisabledServicesFromSourceColumn
            ,
            [parameter(Mandatory=$true,ParameterSetName = "Single User")]
            [validateset('E4','E3','E2','E1','K1')]
            [string]$LicenseTypeDesired
        )
        #initial setup
        #Build License Strings for Available License Plans
            [string]$E4AddLicenses = ($Global:TenantSubDomain + ':' + $Global:E4_SkuPartNumber)
            [string]$E3AddLicenses = ($Global:TenantSubDomain + ':' + $Global:E3_SkuPartNumber)
            [string]$E2AddLicenses = ($Global:TenantSubDomain + ':' + $Global:E2_SkuPartNumber)
            [string]$E1AddLicenses = ($Global:TenantSubDomain + ':' + $Global:E1_SkuPartNumber)
            [string]$K1AddLicenses = ($Global:TenantSubDomain + ':' + $Global:K1_SkuPartNumber)

        $InterchangeableLicenses = @($E4AddLicenses,$E3AddLicenses,$E2AddLicenses,$E1AddLicenses,$K1AddLicenses)

        #input and output files
        $stamp = Get-TimeStamp
        [string]$LogPath = $trackingfolder + $stamp + '-Apply-MSOLLicense-' + "$wave$BaseLogFilePath.log"
        [string]$ErrorLogPath = $trackingfolder + $stamp + '-ERRORS_Apply-MSOLLicense-' + "$wave$BaseLogFilePath.log"

        switch ($PSCmdlet.ParameterSetName) {
            'Migration Wave Source Data' {
                switch ($wavetype) {
                    'Full' {$WaveSourceData = @($SourceData | Where-Object {$_.wave -like "$wave*"})}
                    'Sub' {$WaveSourceData = @($SourceData | Where-Object {$_.wave -eq $wave})}
                    'All' {$WaveSourceData = $sourceData}
                }
            }
            'Custom Source Data' {
                #Validate Custom Source
                $CustomSourceColumns = $CustomSource | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
                Write-OneShellLog -verbose -message "Custom Source Columns are $CustomSourceColumns." -logpath $logpath
                $RequiredColumns = @('UserPrincipalName','LicenseTypeDesired')
                Write-OneShellLog -verbose -message "Required Columns are $RequiredColumns." -logpath $logpath
                $proceed = $true
                foreach ($reqcol in $RequiredColumns) {
                    if ($reqcol -notin $CustomSourceColumns) {
                        $Proceed = $false
                        Write-OneShellLog -errorlog -verbose -message "Required Column $reqcol is not found in the Custom Source data provided.  Processing cannot proceed." -logpath $logpath -errorlogpath $errorlogpath
                    }
                }
                if ($Proceed) {
                    $UsersToLicense = $CustomSource
                    Write-OneShellLog -verbose -message "Custom Source Data Columns Validated." -logpath $logpath
                }
                else {
                    $UsersToLicense = $null
                    Write-OneShellLog -errorlog -verbose -message "ERROR: Custom Source Data Colums failed validation.  Processing cannot proceed." -logpath $logpath -errorlogpath $errorlogpath
                }
            }
            'Single User' {
                    $UsersToLicense = @('' | Select-Object -Property @{n='UserPrincipalName';e={$UserPrincipalName}},@{n='DisableServices';e={$DisabledServices}},@{n='LicenseTypeDesired';e={[string]$LicenseTypeDesired}})
            }

        }
        #record record count to be processed
        $RecordCount = $UsersToLicense.count
        $b=0
        $WriteProgressParams = @{}
        $writeProgressParams.Activity = 'Processing User License and License Option Assignments'

        #main processing
        foreach ($Record in $UsersToLicense) {
            $b++
            #Modify the following lines if the organization has discrepancies between Azure UPN and AD UPN
            $CurrentADUPN = $Record.UserPrincipalName
            $CurrentAzureUPN = $Record.UserPrincipalName
            $WriteProgressParams.PercentComplete = ($b/$RecordCount*100)
            $writeProgressParams.Status = "Processing Record $b of $($RecordCount): $CurrentADUPN."
            $WriteProgressParams.CurrentOperation = "Reading Desired Licensing and Licensing Options."
            Write-Progress @WriteProgressParams
            #calculate any services to disable
            if ($SelectServices) {
                switch ($DisabledServicesFromSourceColumn) {
                    $true {
                        $DisableServices = if ([string]::IsNullOrWhiteSpace($record.DisableServices)) {$null} else {@($Record.DisableServices.split(';'))}
                    }
                    $false {
                        $DisableServices = $disabledServices
                    }
                }
            }
            #calculate license and license options to apply
            $msollicenseoptionsparams = @{}
            $msollicenseoptionsparams.ErrorAction = 'Stop'
            switch ($Record.LicensetypeDesired) {
                'E4' {
                    $DesiredLicense = $E4AddLicenses
                    $msollicenseoptionsparams.AccountSkuID = $DesiredLicense
                    if ($SelectServices) {
                        $DisabledPlans = @()
                        foreach ($service in $DisableServices) {
                            switch ($service) {
                                'Exchange' {$DisabledPlans += 'EXCHANGE_S_ENTERPRISE'}
                                'Lync' {$DisabledPlans += 'MCOSTANDARD'}
                                'SharePoint' {$DisabledPlans += 'SHAREPOINTENTERPRISE'}
                                'OfficeWebApps' {$DisabledPlans += 'SHAREPOINTWAC'}
                                'OfficeProPlus' {$DisabledPlans += 'OFFICESUBSCRIPTION'}
                                'LyncEV' {$DisabledPlans += 'MCOVOICECONF'}
                                'AzureRMS' {$DisabledPlans += 'RMS_S_ENTERPRISE'}
                            }
                        }
                        if ($DisabledPlans.Count -gt 0) {
                            Write-OneShellLog -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                            $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                        }
                    }
                    else {$msollicenseoptionsparams.DisabledPlans = $Null}
                    Write-OneShellLog -Message "Desired E4 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
                    #Create License Options Object
                    $LicenseOptions = New-MsolLicenseOptions @msollicenseoptionsparams
                    $Proceed = $true
                }
                'E3' {
                    $DesiredLicense = $E3AddLicenses
                    $msollicenseoptionsparams.AccountSkuID = $DesiredLicense
                    if ($SelectServices) {
                        $DisabledPlans = @()
                        foreach ($service in $DisabledServices) {
                            switch ($service) {
                                'Exchange' {$DisabledPlans += 'EXCHANGE_S_ENTERPRISE'}
                                'Lync' {$DisabledPlans += 'MCOSTANDARD'}
                                'SharePoint' {$DisabledPlans += 'SHAREPOINTENTERPRISE'}
                                'OfficeWebApps' {$DisabledPlans += 'SHAREPOINTWAC'}
                                'OfficeProPlus' {$DisabledPlans += 'OFFICESUBSCRIPTION'}
                                'AzureRMS' {$DisabledPlans += 'RMS_S_ENTERPRISE'}
                            }
                        }
                        if ($DisabledPlans.Count -gt 0) {
                            Write-OneShellLog -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                            $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                        }
                    }
                    else {$msollicenseoptionsparams.DisabledPlans = $Null}
                    Write-OneShellLog -Message "Desired E3 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
                    #Create License Options Object
                    $LicenseOptions = New-MsolLicenseOptions @msollicenseoptionsparams
                    $Proceed = $true
                }
                'E2' {
                    $DesiredLicense = $E2AddLicenses
                    $msollicenseoptionsparams.AccountSkuID = $DesiredLicense
                    if ($SelectServices) {
                        $DisabledPlans = @()
                        foreach ($service in $DisabledServices) {
                            switch ($service) {
                                'Exchange' {$DisabledPlans += 'EXCHANGE_S_STANDARD'}
                                'Lync' {$DisabledPlans += 'MCOSTANDARD'}
                                'SharePoint' {$DisabledPlans += 'SHAREPOINTSTANDARD'}
                                'OfficeWebApps' {$DisabledPlans += 'SHAREPOINTWAC'}
                            }
                        }
                        if ($DisabledPlans.Count -gt 0) {
                            Write-OneShellLog -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                            $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                        }
                    }
                    else {$msollicenseoptionsparams.DisabledPlans = $Null}
                    Write-OneShellLog -Message "Desired E2 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
                    #Create License Options Object
                    $LicenseOptions = New-MsolLicenseOptions @msollicenseoptionsparams

                    $Proceed = $true
                }
                'E1' {
                    $DesiredLicense = $E1AddLicenses
                    $msollicenseoptionsparams.AccountSkuID = $DesiredLicense
                    if ($SelectServices) {
                        $DisabledPlans = @()
                        foreach ($service in $DisabledServices) {
                            switch ($service) {
                            'Exchange' {$DisabledPlans += 'EXCHANGE_S_STANDARD'}
                            'Lync' {$DisabledPlans += 'MCOSTANDARD'}
                            'SharePoint' {$DisabledPlans += 'SHAREPOINTSTANDARD'}
                            }
                        }
                        if ($DisabledPlans.Count -gt 0) {
                            Write-OneShellLog -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                            $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                        }
                    }
                    else {$msollicenseoptionsparams.DisabledPlans = $Null}
                    Write-OneShellLog -Message "Desired E1 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
                    #Create License Options Object
                    $LicenseOptions = New-MsolLicenseOptions @msollicenseoptionsparams

                    $Proceed = $true
                }
                'K1' {
                    $DesiredLicense = $K1AddLicenses
                    $msollicenseoptionsparams.AccountSkuID = $DesiredLicense
                    if ($SelectServices) {
                        $DisabledPlans = @()
                        foreach ($service in $DisabledServices) {
                            switch ($service) {
                            'Exchange' {$DisabledPlans += 'EXCHANGE_S_DESKLESS'}
                            'SharePoint' {$DisabledPlans += 'SHAREPOINTDESKLESS'}
                            }
                        }
                        if ($DisabledPlans.Count -gt 0) {
                            Write-OneShellLog -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                            $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                        }
                    }
                    else {$msollicenseoptionsparams.DisabledPlans = $Null}
                    Write-OneShellLog -Message "Desired K1 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
                    #Create License Options Object
                    $LicenseOptions = New-MsolLicenseOptions @msollicenseoptionsparams
                    $Proceed = $true
                }
                Default {
                    $Proceed = $false
                    Write-OneShellLog -Message "No License Desired (non E4,E3,E2,E1,K1) Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
                }
            }
            #Lookup MSOL User Object
            if ($proceed) {
                $WriteProgressParams.CurrentOperation = "Looking up MSOL User Object."
                Write-Progress @WriteProgressParams
                Try {
                    Write-OneShellLog -Message "Looking up MSOL User Object $CurrentAzureUPN for AD User Object $CurrentADUPN" -Verbose -LogPath $LogPath
                    $CurrentMSOLUser = Get-MsolUser -UserPrincipalName $CurrentAzureUPN -ErrorAction Stop
                    Write-OneShellLog -Message "Found MSOL User for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                    $Proceed = $true
                }
                Catch {
                    $Proceed = $false
                    Write-OneShellLog -Message "ERROR: MSOL User for $CurrentAzureUPN not found." -Verbose -LogPath $LogPath
                    Write-OneShellLog -Message "ERROR: MSOL User for $CurrentAzureUPN not found." -LogPath $ErrorLogPath
                    Write-OneShellLog -Message $_.tostring() -LogPath $ErrorLogPath
                }
            }

            #Check Usage Location Assignment of User Object
            if ($Proceed) {

                switch -Wildcard ($CurrentMSOLUser.usagelocation) {
                    $null {
                        Try {
                            if ($AssignUsageLocation -and ($StaticUsagelocation -ne $null)) {Set-MsolUser -UserPrincipalName $CurrentAzureUPN -UsageLocation $StaticUsagelocation -ErrorAction Stop}
                            elseif ($AssignUsageLocation -and ($StaticUsagelocation -eq $null)) {Set-MsolUserUsageLocation -UserPrincipalName $CurrentAzureUPN -ErrorAction Stop}
                        }
                        Catch {
                            $Proceed = $false
                        }
                    }
                    '' {
                        Try {
                            if ($AssignUsageLocation -and ($StaticUsagelocation -ne $null)) {Set-MsolUser -UserPrincipalName $CurrentAzureUPN -UsageLocation $StaticUsagelocation -ErrorAction Stop}
                            elseif ($AssignUsageLocation -and ($StaticUsagelocation -eq $null)) {Set-MsolUserUsageLocation -UserPrincipalName $CurrentAzureUPN -ErrorAction Stop}
                        }
                        Catch {
                            $Proceed = $false
                        }
                    }
                    Default {
                        Write-OneShellLog -Message "Usage Location for MSOL User $CurrentAzureUPN is set to $($CurrentMSOLUser.UsageLocation)." -LogPath $LogPath -Verbose
                        $Proceed = $true
                    }

                }
            }
            #Determine License Operation Required
            if ($Proceed) {
            $WriteProgressParams.currentoperation = "Determining License Operation Required"
            Write-Progress @WriteProgressParams
            #Correct License Already Applied?
            if ($CurrentMSOLUser.IsLicensed) {$LicenseAssigned = $true} else {$LicenseAssigned = $false}
            Write-OneShellLog -Message "$CurrentADUPN license assignment status = $LicenseAssigned" -Verbose -LogPath $LogPath
            if ($CurrentMSOLUser.Licenses.AccountSkuId -contains $DesiredLicense) {$CorrectLicenseType = $True}
            else {
                $CorrectLicenseType = $false
                $LicenseToReplace = $CurrentMSOLUser.Licenses.AccountSkuID | where-object {$_ -in ($InterchangeableLicenses)}
            }
            Write-OneShellLog -Message "$CurrentADUPN correct license applied status = $CorrectLicenseType" -Verbose -LogPath $LogPath

            #Correct License Options Already Applied?
            if (-not $CorrectLicenseType) {$correctLicenseOptions = $false}
            else {
                #get current user's disabled plans
                $currentUserDisabledPlans = @($currentMSOLUser.Licenses.servicestatus | ? ProvisioningStatus -eq 'Disabled' | % {$_.servicePlan.ServiceName})
                #compare intended disabled plans to current user's disabled plans:
                $unintendedDisabledPlans = @($currentUserDisabledPlans | ?  {$_ -notin $DisabledPlans})
                $unintendedEnabledPlans = @($DisabledPlans | ? {$_ -notin $currentUserDisabledPlans})
                if ($unintendedDisabledPlans.count -gt 0 -or $unintendedEnabledPlans.Count -gt 0) {$correctLicenseOptions = $false}
                else {$correctLicenseOptions = $true}
            }
            Write-OneShellLog -Message "$CurrentADUPN correct license options applied status = $correctLicenseOptions" -Verbose -LogPath $LogPath
            #Set Operation To Process on User
            $MSOLUserLicenseParams = @{}
            $MSOLUserLicenseParams.ErrorAction = 'Stop'
            $MSOLUserLicenseParams.UserPrincipalName = $CurrentAzureUPN
            if (-not $LicenseAssigned) {$LicenseOperation = 'Assign'}
            if ($licenseAssigned -and $CorrectLicenseType -and $correctLicenseOptions) {$LicenseOperation = 'None'}
            if ($licenseAssigned -and $CorrectLicenseType -and -not $correctLicenseOptions) {$LicenseOperation = 'Options'}
            if ($LicenseAssigned -and -not $CorrectLicenseType) {$LicenseOperation = 'Replace'}


            Write-OneShellLog -Message "$CurrentADUPN license operation selected = $LicenseOperation" -Verbose -LogPath $LogPath

            #Process License Operation
            switch ($LicenseOperation) {
                'None' {Write-OneShellLog -Message "$CurrentAzureUPN is already correctly licensed." -Verbose -LogPath $LogPath
                }
                'Assign'{
                    Try {
                        $MSOLUserLicenseParams.AddLicenses = $DesiredLicense
                        $MSOLUserLicenseParams.LicenseOptions = $LicenseOptions
                        Write-OneShellLog -Message "Setting User License for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                        Set-MsolUserLicense @MSOLUserLicenseParams
                        Write-OneShellLog -Message "Success: Assigned User License for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                    }
                    Catch {
                        Write-OneShellLog -Message "ERROR: License could not be assigned for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                        Write-OneShellLog -Message "ERROR: License could not be assigned for $CurrentAzureUPN" -LogPath $ErrorLogPath
                        Write-OneShellLog -Message $_.tostring() -Verbose -errorlogpath $ErrorLogPath
                    }
                }
                'Replace' {
                    Try {
                        $MSOLUserLicenseParams.AddLicenses = $DesiredLicense
                        $MSOLUserLicenseParams.LicenseOptions = $LicenseOptions
                        $MSOLUserLicenseParams.RemoveLicenses = $LicenseToReplace
                        Write-OneShellLog -Message "Replacing User License for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                        Set-MsolUserLicense @MSOLUserLicenseParams
                        Write-OneShellLog -Message "Success: Replaced User License for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                    }
                    Catch {
                        Write-OneShellLog -Message "ERROR: License could not be replaced for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                        Write-OneShellLog -Message "ERROR: License could not be replaced for $CurrentAzureUPN" -LogPath $ErrorLogPath
                        Write-OneShellLog -Message $_.tostring() -Verbose -LogPath $ErrorLogPath
                    }
                }
                'Options' {
                    Try {
                        #$MSOLUserLicenseParams.AddLicenses = $DesiredLicense
                        $MSOLUserLicenseParams.LicenseOptions = $LicenseOptions
                        Write-OneShellLog -Message "Setting User License Options for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                        Set-MsolUserLicense @MSOLUserLicenseParams
                        Write-OneShellLog -Message "Success: Set User License Options for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                    }
                    Catch {
                        Write-OneShellLog -Message "ERROR: License options could not be set for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                        Write-OneShellLog -Message "ERROR: License options could not be set for $CurrentAzureUPN" -LogPath $ErrorLogPath
                        Write-OneShellLog -Message $_.tostring() -Verbose -LogPath $ErrorLogPath
                    }
                }
            }

        }
        }
    
    }
