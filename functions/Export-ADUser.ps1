function Export-ADUser
{
    [cmdletbinding()]
    param(
        [parameter(Mandatory)]
        [string]$OutputFolderPath
        ,
        [parameter(Mandatory)]
        [string]$Domain
        ,
        [parameter()]
        [bool]$Exchange = $true
        ,
        [parameter()]
        [string[]]$CustomProperty
        ,
        [parameter()]
        [switch]$CompressOutput
    )


    if ($null -eq $(Get-Module -Name ActiveDirectory))
    {
        Import-Module ActiveDirectory -ErrorAction Stop
    }

    switch ($Exchange)
    {
        $true
        {
            $Properties = 'DisplayName', 'Mail', 'proxyAddresses', 'SamAccountName', 'UserPrincipalName', 'Company', 'department', 'objectSID', 'msExchMasterAccountSid', 'DistinguishedName', 'ObjectGUID', 'mailnickname', 'mS-DS-ConsistencyGUID', 'msExchMailboxGuid', 'physicalDeliveryOfficeName'
        }
        $false
        {
            $Properties = 'DisplayName', 'Mail', 'proxyAddresses', 'SamAccountName', 'UserPrincipalName', 'Company', 'department', 'objectSID', 'DistinguishedName', 'ObjectGUID', 'mS-DS-ConsistencyGUID', 'physicalDeliveryOfficeName'
        }
    }

    if ($CustomProperty.Count -ge 1)
    {
        $Properties = $Properties += $CustomProperty
    }

    $DateString = Get-Date -Format yyyyMMddhhmmss

    $OutputFileName = $DateString + $Domain + 'Users'
    $OutputFilePath = Join-Path -Path $OutputFolderPath -ChildPath $($OutputFileName + '.xml')

    $ADUsers = Get-ADUser -Properties $Properties -filter * #| Sort-Object -Property $Properties -Descending

    $ADUsers | Export-Clixml -Path $outputFilePath

    if ($CompressOutput)
    {
        $ArchivePath = Join-Path -Path $OutputFolderPath -ChildPath $($OutputFileName + '.zip')

        Compress-Archive -Path $OutputFilePath -DestinationPath $ArchivePath

        Remove-Item -Path $OutputFilePath -Confirm:$false

    }
}