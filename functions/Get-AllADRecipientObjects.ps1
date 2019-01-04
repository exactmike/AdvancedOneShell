    Function Get-AllADRecipientObjects {
        
        [cmdletbinding()]
        param
        (
            [Parameter()]
            [AllowNull()]
            [int]$ResultSetSize = $null
            ,
            [bool]$Passthru = $true
            ,
            [switch]$ExportData
            ,
            [parameter(Mandatory)]
            [System.Management.Automation.Runspaces.PSSession]$ADSession
        )
        $ADUserAttributes = Get-OneShellVariableValue -Name ADUserAttributes
        $ADGroupAttributesWMembership = Get-OneShellVariableValue -Name ADGroupAttributesWMembership
        $ADContactAttributes = Get-OneShellVariableValue -Name ADContactAttributes
        $ADPublicFolderAttributes = Get-OneShellVariableValue -Name ADPublicFolderAttributes

        Write-Verbose -Message "ADUserAttributes being returned: $($ADUserAttributes -join ',')"
        Write-Verbose -Message "ADGroupAttributes being returned: $($ADGroupAttributesWMembership -join ',')"
        Write-Verbose -Message "ADContactAttributes being returned: $($ADContactAttributes -join ',')"
        Write-Verbose -Message "ADPublicFolderAttributes being returned: $($ADPublicFolderAttributes -join ',')"

        #Start Job to Get Groups
        Write-Verbose -Message "Starting Get AD Groups Job"
        $AllGroupsJob = Invoke-Command -session $ADSession -scriptblock {Get-ADGroup -ResultSetSize $using:ResultSetSize -Properties $using:ADGroupAttributesWMembership -Filter * | Select-Object -Property * -ExcludeProperty Property*,Item} -AsJob
        Write-Verbose -Message "Receiving Get AD Groups Job"
        $AllGroups = Receive-Job -Job $AllGroupsJob -Wait
        Write-Verbose -Message "Get AD Groups Job Returned $($AllGroups.count) Groups."

        #Start Job to Get Contacts
        Write-Verbose -Message "Starting Get AD Contacts Job"
        $AllContactsJob = Invoke-Command -Session $ADSession -ScriptBlock {Get-ADObject -Filter {objectclass -eq 'contact'} -Properties $using:ADContactAttributes -ResultSetSize $using:ResultSetSize | Select-Object -Property * -ExcludeProperty Property*,Item} -AsJob
        
        #Process Groups
        Write-Verbose -Message "Processing Groups to find Mail Enabled Groups"
        $AllMailEnabledGroups = $AllGroups | Where-Object -FilterScript {$_.legacyExchangeDN -ne $NULL -or $_.mailNickname -ne $NULL -or $_.proxyAddresses -ne $NULL}
        Write-Verbose -Message "Found $($AllMailEnabledGroups.count) Mail Enabled Groups"

        #Wait on Contacts Job if needed
        Write-Verbose -Message "Receiving Get AD Contacts Job"
        $AllContacts = Receive-Job -Job $AllContactsJob -Wait
        Write-Verbose -Message "Get AD Contacts Job Returned $($AllContacts.count) Contacts."

        #Start Job to Get Users
        Write-Verbose -Message "Starting Get AD Users Job"
        $AllUsersJob = Invoke-Command -Session $ADSession -ScriptBlock {Get-ADUser -ResultSetSize $using:ResultSetSize -Filter * -Properties $using:ADUserAttributes | Select-Object -Property * -ExcludeProperty Property*,Item} -AsJob
        
        #Process Contacts
        Write-Verbose -Message "Processing Contacts to find Mail Enabled Contacts"
        $AllMailEnabledContacts = $AllContacts | Where-Object -FilterScript {$_.legacyExchangeDN -ne $NULL -or $_.mailNickname -ne $NULL -or $_.proxyAddresses -ne $NULL}
        Write-Verbose -Message "Found $($AllMailEnabledContacts.count) Mail Enabled Contacts"

        #Wait on Users Job if needed
        Write-Verbose -Message "Receiving Get AD Users Job"
        $AllUsers = Receive-Job -Job $AllUsersJob -Wait
        Write-Verbose -Message "Get AD Users Job Returned $($AllUsers.count) Users."

        #Start Job to Get Public Folders
        Write-Verbose -Message "Starting Get AD PublicFolders Job"
        $AllPublicFoldersJob = Invoke-Command -Session $ADSession -ScriptBlock {Get-ADObject -Filter {objectclass -eq 'publicFolder'} -ResultSetSize $using:ResultSetSize -Properties $using:ADPublicFolderAttributes | Select-Object -Property * -ExcludeProperty Property*,Item} -AsJob

        #Process Users
        Write-Verbose -Message "Processing Users to find Mail Enabled Users"
        $AllMailEnabledUsers = $AllUsers  | Where-Object -FilterScript {$_.legacyExchangeDN -ne $NULL -or $_.mailNickname -ne $NULL -or $_.proxyAddresses -ne $NULL}
        Write-Verbose -Message "Found $($AllMailEnabledUsers.count) Mail Enabled Users"
        
        #Wait on Public Folders Job if needed
        Write-Verbose -Message "Receiving Get AD PublicFolders Job"
        $AllPublicFolders = Receive-Job -Job $AllPublicFoldersJob -Wait
        Write-Verbose -Message "Get AD PublicFolders Job Returned $($AllPublicFolders.count) PublicFolders."

        #Process Public Folders
        Write-Verbose -Message "Processing PublicFolders to find Mail Enabled PublicFolders"
        $AllMailEnabledPublicFolders = $AllPublicFolders  | Where-Object -FilterScript {$_.legacyExchangeDN -ne $NULL -or $_.mailNickname -ne $NULL -or $_.proxyAddresses -ne $NULL}
        Write-Verbose -Message "Found $($AllMailEnabledPublicFolders.count) Mail Enabled PublicFolders"

        #Combine objects for output
        Write-Verbose -Message "Combining Objects for Output"
        $AllMailEnabledADObjects = @($AllMailEnabledGroups; $AllMailEnabledContacts; $AllMailEnabledUsers; $AllMailEnabledPublicFolders)
        Write-Verbose -Message "Found $($AllMailEnabledADObjects.count) AD Objects that are Mail Enabled Recipients"

        #output
        if ($Passthru) {$AllMailEnabledADObjects}
        if ($ExportData) {Export-OneShellData -DataToExport $AllMailEnabledADObjects -DataToExportTitle 'AllADRecipientObjects' -Depth 3 -DataType clixml}
    
    }
