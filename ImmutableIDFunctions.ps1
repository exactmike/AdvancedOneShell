function Set-ImmutableIDAttributeValue
    {
        [cmdletbinding(DefaultParameterSetName='Single',SupportsShouldProcess=$true)]
        param
        (
            [parameter(ParameterSetName = 'EntireForest')]
            [switch]$EntireForest
            ,
            [parameter(ParameterSetName = 'EntireForest',Mandatory)]
            [string]$ForestFQDN
            ,            
            [parameter(ParameterSetName = 'SearchBase',Mandatory)]
            [ValidateSet('Base','OneLevel','SubTree')]
            [string]$SearchScope = 'SubTree'
            ,
            #Should be a valid Distinguished Name
            [parameter(ParameterSetName = 'SearchBase')]
            [ValidateScript({Test-Path $_})]
            [string]$SearchBase
            ,
            [parameter(ParameterSetName = 'Single',ValueFromPipelineByPropertyName = $true)]
            [Alias('DistinguishedName','SamAccountName','ObjectGUID')]
            [string]$Identity
            ,
            [string]$ImmutableIDAttribute = 'mS-DS-ConsistencyGuid'
            ,
            [string]$ImmutableIDAttributeSource = 'ObjectGUID'
            ,
            [switch]$OnlyReport
            ,
            [bool]$ExportResults = $true
            ,
            [switch]$OnlyUpdateNull
        )
        Begin
        {
            #Check Current PSDrive Location: Should be AD, Should be GC, Should be Root of the PSDrive
            $Location = Get-Location
            $PSDriveTests = @{
                ProviderIsActiveDirectory = $($Location.Provider.ToString() -like '*ActiveDirectory*')
                LocationIsRootOfDrive = ($Location.Path.ToString() -eq $($Location.Drive.ToString() + ':\'))
                ProviderPathIsRootDSE = ($Location.ProviderPath.ToString() -eq '//RootDSE/')
            }#PSDriveTests
            if ($PSDriveTests.Values -contains $false)
            {
                Write-Log -ErrorLog -Verbose -Message "Set-ImmutableIDAttributeValue may not continue for the following reason(s) related to the command prompt location:"
                Write-Log -ErrorLog -Verbose -Message $($PSDriveTests.GetEnumerator() | Where-Object -filter {$_.Value -eq $False} | Select-Object @{n='TestName';e={$_.Key}},Value | ConvertTo-Json -Compress)
                Write-Error -Message "Set-ImmutableIDAttributeValue may not continue due to the command prompt location.  Review Error Log for details." -ErrorAction Stop
            }#If
            #Setup operational parameters for Get-ADObject based on Parameter Set
            $GetADObjectParams = @{
                Properties = @('CanonicalName',$ImmutableIDAttributeSource,$ImmutableIDAttribute)
                ErrorAction = 'Stop'
            }#GetADObjectParams
            switch ($PSCmdlet.ParameterSetName)
            {
                'EntireForest'
                {
                    $GetADObjectParams.ResultSetSize = $null
                    $GetADObjectParams.Filter = {objectCategory -eq 'Person' -or objectCategory -eq 'Group'}
                    Try
                    {
                        $message = "Find AD Forest $ForestFQDN"
                        Write-Log -Message $message -EntryType Attempting
                        $Forest = Get-ADForest -Server $ForestFQDN -ErrorAction Stop
                        $message = $message + "with domains $($forest.Domains -join ', ')"
                        Write-Log -Message $message -EntryType Succeeded
                        Write-Verbose -Message "Forest Found: $($forest.name)"
                    }
                    catch
                    {
                        Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
                        throw "Failed to get AD Forest $ForestFQDN"
                    }                    
                }#EntireForest
                'Single'
                {
                    #$GetADObjectParams.ResultSetSize = 1
                }#Single
                'SearchBase'
                {
                    $GetADObjectParams.ResultSetSize = $null
                    $GetADObjectParams.Filter = {objectCategory -eq 'Person' -or objectCategory -eq 'Group'}
                    $GetADObjectParams.SearchBase = $SearchBase
                    $GetADObjectParams.SearchScope = $SearchScope
                }#SearchBase
            }#Switch
            #Setup Export Files if $ExportResults is $true
            if ($ExportResults)
            {
                $ADObjectGetSuccesses = @()
                $ADObjectGetFailures = @()
                $Successes = @()
                $Failures = @()
                $ExportName = "SetImmutableIDAttributeValueResults"
            }#if
        }#Begin
        Process
        {
            $message = $PSCmdlet.MyInvocation.InvocationName + ': Get AD Objects with the Get-ADObject cmdlet.'
            Write-Log -Message $message -Verbose -EntryType Attempting            
            switch ($PSCmdlet.ParameterSetName)
            {
                'single'
                {
                    Try
                    {
                        $GetADObjectParams.Identity = $Identity
                        $adobjects = @(Get-ADObject @GetADObjectParams | Select-Object -ExcludeProperty Item,Property* -Property *,@{n='Domain';e={Get-AdObjectDomain -adobject $_ -ErrorAction Stop}})
                        $message = $PSCmdlet.MyInvocation.InvocationName + ": Get $($adObjects.Count) AD Objects with the Get-ADObject cmdlet."
                        $ADObjectGetSuccesses += $Identity
                        Write-Log -Message $message -Verbose -EntryType Succeeded
                        if ($OnlyUpdateNull -eq $true)
                        {
                            $adobjects = $adobjects | Where-Object -FilterScript {$null -eq $_.$($ImmutableIDAttribute)}
                        }                        
                    }#Try
                    catch
                    {
                        Write-Log -Message $message -Verbose -EntryType Failed
                        Write-Log -Message $_.tostring() -ErrorLog
                        $ADObjectGetFailures += $Identity | Select-Object @{n='Identity';e={$Identity}},@{n='TimeStamp';e={Get-TimeStamp}},@{n='Status';e={'Failed'}},@{n='ErrorString';e={$_.tostring()}}
                    }
                }
                'SearchBase'
                {
                    $adobjects = @()
                }
                'EntireForest'
                {
                    $ADObjects = @(
                        foreach ($d in $Forest.domains)
                        {
                            $GetADObjectParams.Server = $d
                            Get-ADObject @GetADObjectParams | Select-Object -ExcludeProperty Item,Property* -Property *,@{n='Domain';e={Get-AdObjectDomain -adobject $_ -ErrorAction Stop}}
                        }
                    )
                    if ($OnlyUpdateNull -eq $true)
                    {
                        $adobjects = $adobjects | Where-Object -FilterScript {$null -eq $_.$($ImmutableIDAttribute)}
                    }
                }
            }#end switch
            if ($OnlyReport -eq $true)
            {
                Export-Data -DataToExportTitle TargetADObjectsForSetImmutableID -DataToExport $adobjects -DataType csv
            }
            else
            {
                #Modify the objects that need modifying
                $O = 0 #Current Object Counter
                $ObjectCount = $adobjects.Count
                $AllResults = @(
                    $adobjects | ForEach-Object {
                        $CurrentObject = $_
                        $O++ #Current Object Counter Incremented
                        $LogString = "Set-ImmutableIDAttributeValue: Set Immutable ID Attribute $ImmutableIDAttribute for Object $($CurrentObject.ObjectGUID.tostring()) with the Set-ADObject cmdlet."
                        Write-Progress -Activity "Setting Immutable ID Attribute for $ObjectCount AD Object(s)" -PercentComplete $($O/$ObjectCount*100) -CurrentOperation $LogString
                        Try
                        {
                            if ($PSCmdlet.ShouldProcess($CurrentObject.ObjectGUID,"Set-ADObject $ImmutableIDAttribute with value $ImmutableIDAttributeSource"))
                            {
                                Write-Log -Message $LogString -EntryType Attempting
                                Set-ADObject -Identity $CurrentObject.ObjectGUID -Add @{$ImmutableIDAttribute=$($CurrentObject.$($ImmutableIDAttributeSource))} -Server $CurrentObject.Domain -ErrorAction Stop -confirm:$false #-WhatIf
                                Write-Log -Message $LogString -EntryType Succeeded
                                if ($ExportResults)
                                {
                                    $attributeset = @('ObjectGUID','Domain','ObjectClass','DistinguishedName',@{n='TimeStamp';e={Get-TimeStamp}},@{n='Status';e={'Succeeded'}},@{n='ErrorString';e={'None'}},@{n='SourceAttribute';e={$ImmutableIDAttributeSource}},@{n='TargetAttribute';e={$ImmutableIDAttribute}})
                                    Write-Output -InputObject ($CurrentObject | Select-Object -Property $attributeset)
                                }#if
                            }#if
                        }#try
                        Catch
                        {
                            Write-Log -Message $LogString -EntryType Failed -ErrorLog -Verbose
                            Write-Log -Message $_.ToString() -ErrorLog
                            if ($ExportResults)
                            {
                                $attributeset = @('ObjectGUID','Domain','ObjectClass','DistinguishedName',@{n='TimeStamp';e={Get-TimeStamp}},@{n='Status';e={'Succeeded'}},@{n='ErrorString';e={'None'}},@{n='SourceAttribute';e={$ImmutableIDAttributeSource}},@{n='TargetAttribute';e={$ImmutableIDAttribute}})                    
                                Write-Output -inputObject ($CurrentObject | Select-Object -Property $attributeset)
                            }#if
                        }#Catch
                    }#ForEach-Object
                )
                Write-Progress -Activity "Setting Immutable ID Attribute for $ObjectCount AD Object(s)" -Completed
            }#end else
        }
        End
        {
            If ($ExportResults)
            {
                if ($PSCmdlet.ParameterSetName -eq 'Single')
                {
                    $AllLookupAttempts = $ADObjectGetSuccesses.Count + $ADObjectGetFailures.Count
                    Write-Log -Message "Set-ImmutableIDAttributeValue Get AD Object Results: Total Attempts: $AllLookupAttempts; Successes: $($ADObjectGetSuccesses.Count); Failures: $($ADObjectGetFailures.count)" -Verbose
                    Export-Data -DataToExportTitle 'ImmutableIDSingleUpdateGetFailures' -DataToExport $ADObjectGetFailures -DataType csv 
                }
                Write-Log -message "Set-ImmutableIDAttributeValue Set AD Object Results: Total Attempts: $($AllResults.Count); Successes: $($Successes.Count); Failures: $($Failures.Count)." -Verbose
                Export-Data -DataToExportTitle $ExportName -DataToExport $AllResults -DataType csv
            }
            Write-Log -Message "Set-ImmutableIDAttributeValue Operations Completed." -Verbose
        }
    }
function Join-ADObjectByImmutableID
    {
        [cmdletbinding(SupportsShouldProcess)]
        param
        (
            $SourceForestDrive #Source ADForest PSDriveName Without any path/punctuation
            ,
            $SourceObjectGUID
            ,
            $SourceImmutableIDAttribute = 'mS-DS-ConsistencyGUID'
            ,
            $TargetForestDrive #Target ADForest PSDriveName Without any path/punctuation
            ,
            $TargetObjectGUID
            ,
            $TargetImmutableIDAttribute = 'mS-DS-ConsistencyGUID'
        )
        Push-Location
        try
        {
            Set-Location $($SourceForestDrive + ':\') -ErrorAction Stop
            $SourceObjectFromGlobalCatalog = Get-AdObject -Identity $SourceObjectGUID -Property CanonicalName -ErrorAction Stop
            $SourceObjectDomain = Get-AdObjectDomain -adobject $SourceObjectFromGlobalCatalog -ErrorAction Stop
            $SourceObject = Get-AdObject -Identity $SourceObjectGUID -Server $SourceObjectDomain -Property CanonicalName,$SourceImmutableIDAttribute -ErrorAction Stop
            if ($null -eq $($SourceObject.$($SourceImmutableIDAttribute)))
            {
                Throw "Source Object $SourceObjectGUID's source Immutable ID attribute $SourceImmutableIDAttribute is NULL"
            }
        }
        catch
        {
            Pop-Location
            $_
            Throw "Source Object $sourceObjectGUID Failure for Source Forest PSDrive $sourceForestDrive"
        }
        try
        {
            Set-Location $($TargetForestDrive + ':\') -ErrorAction Stop
            $TargetObjectFromGlobalCatalog = Get-AdObject -Identity $TargetObjectGUID -Property CanonicalName -ErrorAction Stop
            $TargetObjectDomain = Get-AdObjectDomain -adobject $TargetObjectFromGlobalCatalog -ErrorAction Stop
            $TargetObject = Get-AdObject -Identity $TargetObjectGUID -Server $TargetObjectDomain -Property CanonicalName,$TargetImmutableIDAttribute -ErrorAction Stop
            if ($null -ne $($TargetObject.$($TargetImmutableIDAttribute)))
            {
                Throw "Target Object $TargetObjectGUID's target Immutable ID attribute $targetImmutableIDAttribute is NOT currently NULL"
            }
            if ($PSCmdlet.ShouldProcess($TargetObjectGUID,"Set-ADObject $TargetObjectGUID attribute $TargetImmutableIDAttribute with value $($SourceObject.$($SourceImmutableIDAttribute))"))
            {
                Set-ADObject -Identity $TargetObjectGUID -Add @{$TargetImmutableIDAttribute=$($SourceObject.$($SourceImmutableIDAttribute))} -Server $TargetObjectDomain -ErrorAction Stop -confirm:$false
            }
        }
        catch
        {
            Pop-Location
            $_
            Throw "Target Object $TargetObjectGUID Failure for Target Forest PSDrive $TargetForestDrive"
        }
        Pop-Location
    }
