function Export-ExchangeRecipient
{
    [cmdletbinding()]
    param(
        [parameter(Mandatory)]
        [ValidateScript( { Test-Path -type Container -Path $_ })]
        [string]$OutputFolderPath
        ,
        [string]$RecipientFilter = 'CustomAttribute5 -notlike "ACS*"'
        ,
        [parameter(Mandatory)]
        [ValidateSet('Mailbox', 'CASMailbox', 'RemoteMailbox', 'ResourceCalendarProcessing', 'PublicFolderMailbox', 'ArbitrationMailbox', 'MailboxStatistics', 'PublicFolder', 'PublicFolderStatistics', 'MailPublicFolder', 'Contact', 'DistributionGroup', 'DistributionGroupMember', 'UnifiedGroup', 'UnifiedGroupMember', 'MailUser')]
        [string[]]$Operation #Specify the types of recipients or recipient data to include in the export.  For DistributionGroupMember and UnifiedGroupMember you must include DistributionGroup or UnifiedGroup respectively in order to successfully export membership.
        ,
        [parameter()]
        [switch]$CompressOutput
    )


    #Write-Verbose -Verbose -Message 'Export-ExchangeRecipient Version x.x'

    $ErrorActionPreference = 'Continue'

    #For commands that support Recipient Filtering
    $GetRParams = @{
        ResultSize    = 'Unlimited'
        WarningAction = 'SilentlyContinue'
    }
    if (-not [string]::IsNullOrWhiteSpace($RecipientFilter))
    {
        $GetRParams.Filter = $RecipientFilter
    }

    #For commands that do not support recipient filtering
    $GetNonFRParams = @{
        ResultSize    = 'Unlimited'
        WarningAction = 'SilentlyContinue'
    }

    $ExchangeRecipients = [PSCustomObject]@{
        OrganizationConfig = Get-OrganizationConfig
    }

    $AMParams = @{
        MemberType = 'NoteProperty'
    }

    $OpCount = 0
    $OpTotalCount = $Operation.Count

    Switch ($Operation | Sort-Object)
    {
        'Mailbox'
        {

            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(Get-Mailbox @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'CASMailbox'
        {
            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(Get-CASMailbox @GetNonFRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'RemoteMailbox'
        {
            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(Get-RemoteMailbox @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'ResourceCalendarProcessing'
        {
            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(
                Get-Mailbox -Filter "RecipientTypeDetails -eq 'RoomMailbox' -or RecipientTypeDetails -eq 'EquipmentMailbox'" -ResultSize Unlimited).ForEach( {
                    $mb = $_
                    Get-CalendarProcessing -Identity $mb.exchangeguid.guid |
                    Select-Object -Property *, @{n = 'ExchangeGUID'; e = { $mb.exchangeguid.guid } }
                })
            $ExchangeRecipients | Add-Member @AMParams
        }
        'PublicFolderMailbox'
        {
            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(Get-Mailbox -PublicFolder @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'ArbitrationMailbox'
        {
            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(Get-Mailbox -Arbitration @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'MailboxStatistics'
        {
            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(Get-Mailbox @GetRParams | ForEach-Object { Get-MailboxStatistics -identity $_.ExchangeGUID.guid -WarningAction 'SilentlyContinue' })
            $ExchangeRecipients | Add-Member @AMParams
        }
        'PublicFolder'
        {
            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(Get-PublicFolder -recurse @GetNonFRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'PublicFolderStatistics'
        {
            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(Get-PublicFolderStatistics @GetNonFRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'MailPublicFolder'
        {
            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(Get-MailPublicFolder @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'Contact'
        {
            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(Get-MailContact @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'DistributionGroup'
        {
            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(Get-DistributionGroup @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'DistributionGroupMember'
        {
            $getDGMemberParams = @{
                Identity   = ''
                ResultSize = 'Unlimited'
            }

            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $dgCount = 0
            $dgTotalCount = $ExchangeRecipients.DistributionGroup.Count
            $AMParams.Value = @(

                foreach ($dg in $ExchangeRecipients.DistributionGroup)
                {
                    $dgCount++
                    Write-Progress -Activity 'Exporting Distribution Group Membership' -CurrentOperation $dg.DisplayName -Status "Group $dgCount of $dgTotalCount" -Id 1 -ParentId 0
                    $getDGMemberParams.Identity = $dg.guid.guid
                    Get-DistributionGroupMember @getDGMemberParams |
                    Select-Object -Property DisplayName,Alias,PrimarySMTPAddress,RecipientTypeDetails,GUID,ExchangeGUID,ExternalDirectoryObjectID,@{name='MemberOfGUID';e= {$dg.guid.guid}}, @{n='MemberOfPrimarySMTPAddress';e= {$dg.PrimarySMTPAddress}},@{n='MemberOfDisplayName';e= {$dg.DisplayName}}
                }
                Write-Progress -Id 1 -Completed -Activity 'Exporting Distribution Group Membership'
            )
            $ExchangeRecipients | Add-Member @AMParams
        }
        'UnifiedGroup'
        {
            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(Get-UnifiedGroup @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'UnifiedGroupMember'
        {
            $getUGMemberParams = @{
                Identity   = ''
                ResultSize = 'Unlimited'
                LinkType   = 'Member'
            }

            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(

                foreach ($ug in $ExchangeRecipients.UnifiedGroup)
                {
                    $getUGMemberParams.Identity = $ug.guid.guid
                    Get-UnifiedGroupLinks @getUGMemberParams |
                    Select-Object -Property DisplayName,Alias,PrimarySMTPAddress,RecipientTypeDetails,GUID,ExchangeGUID,ExternalDirectoryObjectID,@{name='MemberOfGUID';e= {$ug.guid.guid}}, @{n='MemberOfPrimarySMTPAddress';e= {$ug.PrimarySMTPAddress}},@{n='MemberOfDisplayName';e= {$ug.DisplayName}}
                }

            )
            $ExchangeRecipients | Add-Member @AMParams
        }
        'MailUser'
        {
            $AMParams.Name = $_
            $OpCount++
            Write-Progress -Activity 'Exporting Exchange Recipients' -CurrentOperation $AMParams.Name -Status "Operation $OpCount of $OpTotalCount" -Id 0
            $AMParams.Value = @(Get-MailUser @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
    }


    $DateString = Get-Date -Format yyyyMMddHHmmss

    $OutputFileName = $($ExchangeRecipients.OrganizationConfig.Name.split('.')[0]) + 'ExchangeRecipientsAsOf' + $DateString
    $OutputFilePath = Join-Path $OutputFolderPath $($OutputFileName + '.xml')

    $ExchangeRecipients | Export-Clixml -Path $OutputFilePath -Encoding utf8

    if ($CompressOutput)
    {
        $ArchivePath = Join-Path -Path $OutputFolderPath -ChildPath $($OutputFileName + '.zip')

        Compress-Archive -Path $OutputFilePath -DestinationPath $ArchivePath

        Remove-Item -Path $OutputFilePath -Confirm:$false

    }

}