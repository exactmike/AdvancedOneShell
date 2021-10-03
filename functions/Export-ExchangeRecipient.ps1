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
        [ValidateSet('Mailbox', 'CASMailbox', 'RemoteMailbox', 'ResourceCalendarProcessing', 'PublicFolderMailbox', 'ArbitrationMailbox', 'MailboxStatistics', 'PublicFolder', 'PublicFolderStatistics', 'MailPublicFolder', 'Contact', 'DistributionGroup', 'MailUser')]
        [string[]]$Operation
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

    Switch ($Operation)
    {
        'Mailbox'
        {
            $AMParams.Name = $_
            $AMParams.Value = @(Get-Mailbox @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'CASMailbox'
        {
            $AMParams.Name = $_
            $AMParams.Value = @(Get-CASMailbox @GetNonFRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'RemoteMailbox'
        {
            $AMParams.Name = $_
            $AMParams.Value = @(Get-RemoteMailbox @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'ResourceCalendarProcessing'
        {
            $AMParams.Name = $_
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
            $AMParams.Value = @(Get-Mailbox -PublicFolder @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'ArbitrationMailbox'
        {
            $AMParams.Name = $_
            $AMParams.Value = @(Get-Mailbox -Arbitration @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'MailboxStatistics'
        {
            $AMParams.Name = $_
            $AMParams.Value = @(Get-Mailbox @GetRParams | ForEach-Object { Get-MailboxStatistics -identity $_.ExchangeGUID.guid -WarningAction 'SilentlyContinue' })
            $ExchangeRecipients | Add-Member @AMParams
        }
        'PublicFolder'
        {
            $AMParams.Name = $_
            $AMParams.Value = @(Get-PublicFolder -recurse @GetNonFRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'PublicFolderStatistics'
        {
            $AMParams.Name = $_
            $AMParams.Value = @(Get-PublicFolderStatistics @GetNonFRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'MailPublicFolder'
        {
            $AMParams.Name = $_
            $AMParams.Value = @(Get-MailPublicFolder @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'Contact'
        {
            $AMParams.Name = $_
            $AMParams.Value = @(Get-MailContact @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'DistributionGroup'
        {
            $AMParams.Name = $_
            $AMParams.Value = @(Get-DistributionGroup @GetRParams)
            $ExchangeRecipients | Add-Member @AMParams
        }
        'MailUser'
        {
            $AMParams.Name = $_
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