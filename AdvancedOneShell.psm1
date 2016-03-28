function Get-ExistingProxyAddressTypes {
param(
[object[]]$proxyAddresses
)
$ProxyAddresses | ForEach-Object -Process {$_.split(':')[0]} | Sort-Object | Select-Object -Unique
}
function Get-DesiredProxyAddresses {
param(
    [parameter(Mandatory=$true)]
    [string[]]$CurrentProxyAddresses
    ,
    [string]$DesiredPrimaryAddress
    ,
    [string]$DesiredOrCurrentAlias
    ,
    [string[]]$LegacyExchangeDNs
    ,
    [psobject[]]$Recipients
    ,
    [parameter()]
    [switch]$VerifyAddTargetAddress
    ,
    [string]$TargetDeliveryDomain = $global:TargetDeliveryDomain
)
$DesiredProxyAddresses = $CurrentProxyAddresses.Clone()
if($DesiredPrimaryAddress) {
    if (($currentPrimary = $CurrentProxyAddresses | Where-Object {$_ -clike 'SMTP:*'} | foreach {$_.split(':')[1]}).count -eq 1) {
        if ($currentPrimary -ceq $DesiredPrimaryAddress) {
        }#if
        else {
            $DesiredProxyAddresses = @($DesiredProxyAddresses | where-object {$_ -notlike "smtp:$DesiredPrimaryAddress"})
            $DesiredProxyAddresses = @($DesiredProxyAddresses | where-object {$_ -notlike "SMTP:$currentPrimary"})
            $DesiredProxyAddresses += $("smtp:$currentPrimary")
            $DesiredProxyAddresses += $("SMTP:$DesiredPrimaryAddress")
        }#else
    }#if
}#if
if ($LegacyExchangeDNs.Count -ge 1) {
    foreach ($LED in $LegacyExchangeDNs) {
        $existingProxyAddressTypes = Get-ExistingProxyAddressTypes -proxyAddresses $DesiredProxyAddresses
        $type = 'X500'
        if ($existingProxyAddressTypes -ccontains $type) {
            $type = $type.ToLower()
        }
        $newX500 = "$type`:$LED"
        if ($newX500 -in $DesiredProxyAddresses) {
        }
        else {
            $DesiredProxyAddresses += $newX500
        }
    }
}
if ($VerifyAddTargetAddress) {
    if ($DesiredOrCurrentAlias -and $TargetDeliveryDomain) {
        $DesiredTargetAddress = "smtp:$DesiredOrCurrentAlias@$TargetDeliveryDomain"
        if (($DesiredProxyAddresses | Where-Object {$_ -eq $DesiredTargetAddress}).count -lt 1) {
            $DesiredProxyAddresses += $DesiredTargetAddress
        }#if
    }#if
    else {
        Write-Log -Message 'ERROR: VerifyAddTargetAddress was specified but DesiredOrCurrentAlias or TargetDeliveryDomain were not specified.'
        throw('ERROR: VerifyAddTargetAddress was specified but DesiredOrCurrentAlias or TargetDeliveryDomain were not specified.')
    }#else
}#if
if ($Recipients.Count -ge 1) {
    $RecipientProxyAddresses = @()
    foreach ($recipient in $Recipients) {
        $paProperty = if (Test-Member -InputObject $recipient -Name emailaddresses) {'EmailAddresses'} elseif (Test-Member -InputObject $recipient -Name proxyaddresses ) {'proxyAddresses'} else {$null}
        if ($paProperty) {
        $existingProxyAddressTypes = Get-ExistingProxyAddressTypes -proxyAddresses $DesiredProxyAddresses
            $rpa = @($recipient.$paProperty)
            foreach ($a in $rpa) {
                $type = $a.split(':')[0]
                $address = $a.split(':')[1]
                if ($existingProxyAddressTypes -ccontains $type) {
                    $la = $type.tolower() + ':' +  $address
                }
                else {
                    $la = $a
                }
                $RecipientProxyAddresses += $la
            }#foreach
        }#if
    }#foreach
    if ($RecipientProxyAddresses.count -ge 1) {
        $add = @($RecipientProxyAddresses | Where-Object {$DesiredProxyAddresses -inotcontains $_})
        $DesiredProxyAddresses += @($add)
    }#if
}#if
Return $DesiredProxyAddresses
}#function get-desiredproxyaddresses
function Get-RecipientType {
[cmdletbinding()]
param
(
[parameter(ParameterSetName = 'DisplayType')]
[string]$msExchRecipientDisplayType
,
[parameter(ParameterSetName = 'TypeDetails')]
[string]$msExchRecipientTypeDetails
)
$DisplayTypes = @(
            [pscustomobject]@{Value=1;Type='Universal Distribution Group';Name='DistributionGroup'}
            [pscustomobject]@{Value=1073741833;Type='Universal Security Group';Name='SecurityDistributionGroup'}
            [pscustomobject]@{Value=3;Type='Dynamic Distribution Group';Name='DynamicDistributionGroup'}
            [pscustomobject]@{Value=1073741824;Type='User Mailbox (User, Shared, or Linked)';Name='UserMailbox'}
            [pscustomobject]@{Value=7;Type='Room Mailbox';Name='RoomMailbox'}
            [pscustomobject]@{Value=8;Type='Equipment Mailbox';Name='EquipmentMailbox'}            
            [pscustomobject]@{Value=6;Type='Mail User, Mail Contact';Name='RemoteMailUser'}
            [pscustomobject]@{Value=2;Type='Public Folder';Name='PublicFolder'}
            [pscustomobject]@{Value=4;Type='Outlook Only:Organization';Name='Organization'}
            [pscustomobject]@{Value=5;Type='Outlook Only:Private Distribution List';Name='PrivateDistributionList'}
            [pscustomobject]@{Value=-2147483642;Type='Remote User Mailbox';Name='RemoteUserMailbox'}
            [pscustomobject]@{Value=-2147481594;Type='Remote Equipment Mailbox';Name='RemoteEquipmentMailbox'}
            [pscustomobject]@{Value=-2147483642;Type='Remote Shared Mailbox';Name='RemoteSharedMailbox'}
            [pscustomobject]@{Value=-2147481850;Type='Remote Room Mailbox';Name='RemoteRoomMailbox'}
)
$RecipientTypeDetailsTypes = @(
            [pscustomobject]@{Value=1;Type='User Mailbox';Name='UserMailbox'}
            [pscustomobject]@{Value=2;Type='Linked Mailbox';Name='LinkedMailbox'}
            [pscustomobject]@{Value=4;Type='Shared Mailbox';Name='SharedMailbox'}
            [pscustomobject]@{Value=8;Type='Legacy Mailbox';Name='LegacyMailbox'}
            [pscustomobject]@{Value=16;Type='Room Mailbox';Name='RoomMailbox'}
            [pscustomobject]@{Value=32;Type='Equipment Mailbox';Name='EquipmentMailbox'}
            [pscustomobject]@{Value=64;Type='Mail Contact';Name='MailContact'}
            [pscustomobject]@{Value=128;Type='Mail User';Name='MailUser'}
            [pscustomobject]@{Value=256;Type='Mail Enabled Universal Distribution Group';Name='MailUniversalDistributionGroup'}
            [pscustomobject]@{Value=512;Type='Mail Enabled Non-Universal Distribution Group';Name='MailNonUniversalDistributionGroup'}
            [pscustomobject]@{Value=1024;Type='Mail Enabled Universal Security Group';Name='MailUniversalSecurityGroup'}
            [pscustomobject]@{Value=2048;Type='Dynamic Distribution Group';Name='DynamicDistributionGroup'}
            [pscustomobject]@{Value=4096;Type='Public Folder';Name='PublicFolder'}
            [pscustomobject]@{Value=8192;Type='System Attendant Mailbox';Name='SystemAttendantMailbox'}
            [pscustomobject]@{Value=16384;Type='System Mailbox';Name='SystemMailbox'}
            [pscustomobject]@{Value=32768;Type='Cross Forest Mail Contact';Name='MailForestContact'}
            [pscustomobject]@{Value=65536;Type='User';Name='User'}
            [pscustomobject]@{Value=131072;Type='Contact';Name='Contact'}
            [pscustomobject]@{Value=262144;Type='Universal Distribution Group';Name='UniversalDistributionGroup'}
            [pscustomobject]@{Value=524288;Type='Universal Security Group';Name='UniversalSecurityGroup'}
            [pscustomobject]@{Value=1048576;Type='Non Universal Group';Name='NonUniversalGroup'}
            [pscustomobject]@{Value=2097152;Type='Disabled User';Name='DisabledUser'}
            [pscustomobject]@{Value=4194304;Type='Microsoft Exchange';Name='MicrosoftExchange'}
            [pscustomobject]@{Value=8388608;Type='Arbitration Mailbox';Name='ArbitrationMailbox'}
            [pscustomobject]@{Value=16777216;Type='Mailbox Plan';Name='MailboxPlan'}
            [pscustomobject]@{Value=33554432;Type='Linked User';Name='LinkedUser'}
            [pscustomobject]@{Value=268435456;Type='Room List';Name='RoomList'}
            [pscustomobject]@{Value=536870912;Type='Discovery Mailbox';Name='DiscoveryMailbox'}
            [pscustomobject]@{Value=1073741824;Type='Role Group';Name='RoleGroup'}
            [pscustomobject]@{Value=2147483648;Type='Remote Mailbox';Name='RemoteMailbox'}
            [pscustomobject]@{Value=137438953472;Type='Team Mailbox';Name='TeamMailbox'}
)
switch ($PSCmdlet.ParameterSetName) 
{
    'DisplayType'
    {
    $DisplayTypes | Where-Object -FilterScript {$_.Value -eq $msExchRecipientDisplayType}
    }
    'TypeDetails'
    {
    $RecipientTypeDetailsTypes | Where-object -FilterScript {$_.Value -eq $msExchRecipientTypeDetails}
    }
}
}#function Get-RecipientType
Function Export-FailureRecord {
[cmdletbinding()]
param(
[string]$Identity
,
[string]$ExceptionCode
,
[string]$FailureGroup
,
[string]$ExceptionDetails
,
[string]$RelatedObjectIdentifier
,
[string]$RelatedObjectIdentifierType
)#Param
    $Exception=[ordered]@{
        Identity = $Identity
        ExceptionCode = $ExceptionCode
        ExceptionDetails = $ExceptionDetails
        FailureGroup = $FailureGroup
        RelatedObjectIdentifier = $RelatedObjectIdentifier
        RelatedObjectIdentifierType = $RelatedObjectIdentifierType
        TimeStamp = Get-TimeStamp
    }
    try {
    $ExceptionObject = $Exception | Convert-HashTableToObject
    Export-Data -DataToExportTitle $FailureGroup -DataToExport $ExceptionObject -Append -DataType csv -ErrorAction Stop
    $Global:SEATO_Exceptions += $ExceptionObject
    }
    catch {
    Write-Log -Message "FAILED: to write Exception Record for $identity with Exception Code $ExceptionCode and Failure Group $FailureGroup" -ErrorLog
    }
}#Function Export-FailureRecord
function Move-StagedAccountToOperationalOU {
param(
[parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
[string[]]$SAMAccountName
,
[string]$DestinationOU
)
begin {}
process {
    foreach ($S in $SAMAccountName) {
        try {
            Write-Log -Message "Attempting: Get-ADUser -Identity $S" -Verbose
            $aduser = Get-ADUser -Identity $S -ErrorAction Stop
            Write-Log -Message "Succeeded: Get-ADUser -Identity $S" -Verbose
        }#try
        catch {
            Write-Log -Message "FAILED: Get-ADUser -Identity $S" -Verbose -ErrorLog
            Write-Log -Message $_.tostring()
        }#catch
        try {
            Write-Log -Message "Attempting: Move-ADObject -TargetPath $DestinationOU" -Verbose
            $aduser | Move-ADObject -TargetPath $DestinationOU -ErrorAction Stop
            Write-Log -Message "Succeeded: Move-ADObject -TargetPath $DestinationOU" -Verbose
        }#try
        catch {
            Write-Log -Message "FAILED: Move-ADObject -TargetPath $DestinationOU" -Verbose -ErrorLog
            Write-Log -Message $_.tostring()
        }#catch
    }#foreach
}
end{}
}#function Move-StagedAccountToOperationalOU
function Update-PostMigrationMailboxUser {
[cmdletbinding()]
param(
[parameter(
    Mandatory=$true,
    ParameterSetName = 'Individual'
    )]
[string[]]$Identity
,
[parameter(
    Mandatory=$true,
    ParameterSetName = 'InputList'
    )]
[array]$InputList
,
[parameter(Mandatory=$true)]
[string[]]$SourceAD
,
[parameter(Mandatory=$true)]
[string]$TargetAD
,
[parameter(Mandatory=$true)]
[string]$TargetExchangeOrg
,
[parameter(Mandatory = $true)]
[validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid')]
[string]$TargetLookupAttribute
,
[parameter(Mandatory = $true)]
[validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid')]
[string]$SourceLookupAttribute
,
[parameter(Mandatory = $true)]
[validateset('UserMailbox','LinkedMailbox','LegacyMailbox','MailUser','RemoteUserMailbox','RemoteSharedMailbox','RemoteRoomMailbox','RemoteEquipmentMailbox')]
[string]$TargetRecipientTypeDetails
,
[parameter(Mandatory = $true)]
[string]$TargetDeliveryDomain
)
begin {
    switch ($PSCmdlet.ParameterSetName) {
        'Individual' {
            $recordcount = $Identity.Count
        }#Individual
        'InputList' {
            if ($TargetLookupAttribute -in ($InputList | get-member -MemberType Properties | select-object -ExpandProperty Name)) {
                $Identity = @($InputList | Select-Object -ExpandProperty $TargetLookupAttribute)
                $RecordCount = $Identity.count
            }#if
            else {
                Write-Log -Message "FAILED: InputList does not contain the Target Lookup Attribute $TargetLookupAttribute." -Verbose -ErrorLog
                throw("FAILED: InputList does not contain the Target Lookup Attribute $TargetLookupAttribute.")
            }#else
        }#InputList
    }#switch
    $Global:Exceptions = @()
    $Global:ProcessedUsers = @()
    $msExchRecipientTypeDetails = switch ($TargetRecipientTypeDetails) {'UserMailbox' {1} 'LinkedMailbox' {2} 'LegacyMailbox'{8} 'MailUser' {128} 'RemoteUserMailbox' {2147483648} 'RemoteEquipmentMailbox' {17179869184} 'RemoteSharedMailbox' {34359738368} 'RemoteRoomMailbox' {8589934592}}
}#begin
process{
    $cr = 0
    foreach ($ID in $Identity) {
        try {
            $cr++
            $writeProgressParams = @{
            Activity = "Update-PostMailboxMigrationUser: TargetForest $TargetAD"
            CurrentOperation = "Processing Record $cr of $recordcount : $ID"
            Status = "Lookup User with $ID by $TargetLookupAttribute in Target AD"
            PercentComplete = $cr/$RecordCount*100
            }
            Write-Progress @writeProgressParams
           ################################################################################################################################################################
           #lookup users in source Active Directory Environments
           ################################################################################################################################################################
           #Lookup Target AD User
           ################################################################################################################################################################
            try {
                Write-Log -Message "Attempting: Find AD User $ID in Target AD Forest $TargetAD" -Verbose
                $TADU = @(Find-Aduser -Identity $ID -IdentityType $TargetLookupAttribute -ADInstance $TargetAD -ErrorAction Stop)
                Write-Log -Message "Succeeded: Find AD User $ID in Target AD Forest $TargetAD" -Verbose
            }#try
            catch {
                Write-Log -Message "FAILED: Find AD User $ID in Target AD Forest $TargetAD" -Verbose -ErrorLog
                Write-Log -Message $_.tostring() -ErrorLog
                $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'TargetADUserNotFound'}}
                Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                throw("User Object for value $ID in Attribute $TargetLookupAttribute in Target AD $TargetAD not found.")
            }#catch
            if ($TADU.count -gt 1) {#check for ambiguous results
                Write-Log -Message "FAILED: Find AD User $ID in Target AD Forest $TargetAD returned multiple objects/ambiguous results." -Verbose -ErrorLog
                $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'TargetADUserAmbiguous'}}
                Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                throw("User Object for value $ID in Attribute $TargetLookupAttribute in Target AD $TargetAD was ambiguous.")
            }#if
            else {
                $TADU = $TADU[0]
                $TADUGUID = $TADU.objectguid
                Write-Log -Message "NOTE: Target AD User in $TargetAD Identified with ObjectGUID: $TADUGUID" -Verbose
            }
            ################################################################################################################################################################
            #Lookup Matching Source AD User
            ################################################################################################################################################################
            $writeProgressParams.status = "Lookup User with $($TADU.$SourceLookupAttribute) by $SourceLookupAttribute in Source AD"
            Write-Progress @writeProgressParams
            $SADU = @()
            foreach ($ad in $SourceAD) {
                try {
                    Write-Log -message "Attempting: Find Matching User for $ID in Source AD $ad by Lookup Attribute $SourceLookupAttribute" -Verbose
                    $SADU += Find-Aduser -Identity $($TADU.$SourceLookupAttribute) -IdentityType $SourceLookupAttribute -ADInstance $ad -ErrorAction Stop
                    Write-Log -message "Succeeded: Find Matching User for $ID in Source AD $ad by Lookup Attribute $SourceLookupAttribute" -Verbose
                }#try
                catch {
                    Write-Log -message "FAILED: Find Matching User for $ID in Source AD $ad by Lookup Attribute $SourceLookupAttribute" -Verbose -ErrorLog
                    Write-Log -Message $_.tostring() -ErrorLog
                }
            }#foreach
            #check for no results or ambiguous results
            switch ($SADU.count) {
                1 {
                    Write-Log -message "Succeeded: Found exactly 1 Matching User for $ID in $($SourceAD -join ' & ') by Lookup Attribute $SourceLookupAttribute" -Verbose
                    $SADU = $SADU[0]
                    $SADUGUID = $SADU.objectguid
                    Write-Log -Message "NOTE: Source AD User Identified in with ObjectGUID: $SADUGUID" -Verbose
                }#1
                0 {
                    Write-Log -message "FAILED: Found 0 Matching User for $ID in Source AD $($SourceAD -join ' & ') by Lookup Attribute $SourceLookupAttribute" -Verbose
                    $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'SourceADUserNotFound'}}
                    Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                    throw("User Object for value $ID in Attribute $SourceLookupAttribute in Source AD $($SourceAD -join ' & ') not found.")
                }#0
                Default {
                    Write-Log -message "FAILED: Found multiple ambiguous matching User for $ID in Source AD $($SourceAD -join ' & ') by Lookup Attribute $SourceLookupAttribute" -Verbose
                    $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'SourceADUserAmbiguous'}}
                    Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                    throw("User Object for value $ID in Attribute $SourceLookupAttribute in Source AD $($SourceAD -join ' & ') was ambiguous.")
                }#Default
            }#switch $SADU.count
            ################################################################################################################################################################
            #Calculate Address Changes
            ################################################################################################################################################################
            $writeProgressParams.status = "Calculate Proxy Address and Target Address Changes"
            Write-Progress @writeProgressParams
            try {
                Write-Log -Message "Attempting: Find Current proxy $TargetDeliveryDomain SMTP Address for Target AD User $TADUGUID" -Verbose
                $TargetDeliveryDomainAddress = ($TADU.proxyaddresses | Where-Object {$_ -like "smtp:*@$TargetDeliveryDomain"} | Select-Object -First 1).split(':')[1]
                Write-Log -Message "Succeeded: Find Current proxy $TargetDeliveryDomain SMTP Address for Target AD User $TADUGUID : $TargetDeliveryDomainAddress" -Verbose
            }#try
            catch {
                Write-Log -Message "FAILED: Find Current proxy $TargetDeliveryDomain SMTP Address for Target AD User $TADUGUID" -Verbose -ErrorLog
                Write-Log -Message $_.tostring() -ErrorLog
                Write-Log -Message "NOTE: $TargetDeliveryDomain SMTP Proxy Address for Target AD User $TADUGUID will be added." -Verbose -ErrorLog
                $AddTargetDeliveryProxyAddress = $true
            }#catch
            #setup for get-desiredproxyaddresses function to calculate updated addresses
            $GetDesiredProxyAddressesParams = @{
                CurrentProxyAddresses=$TADU.proxyAddresses
                LegacyExchangeDNs=@($SADU.legacyExchangeDN)
                Recipients = $SADU
                DesiredOrCurrentAlias = $TADU.mailNickName
            }
            if ($AddTargetDeliveryProxyAddress) {$GetDesiredProxyAddressesParams.VerifyAddTargetAddress = $true}
            $DesiredProxyAddresses = Get-DesiredProxyAddresses @GetDesiredProxyAddressesParams
            if ($AddTargetDeliveryProxyAddress) {$TargetDeliveryDomainAddress = ($DesiredProxyAddresses | Where-Object {$_ -like "smtp:*@$TargetDeliveryDomain"} | Select-Object -First 1).split(':')[1]}
            #preparation activities complete, time to write changes to Target AD
            Write-Log -message "Using AD Cmdlets to set attributes for $TADUGUID in $TargetAD" -Verbose
            $writeProgressParams.status = "Updating Attributes for $TADUGUID in $TargetAD using AD Cmdlets"
            Write-Progress @writeProgressParams
            #ClearTargetAttributes
            $setaduserparams1 = @{
                Identity=$TADUGUID
                clear='proxyaddresses','targetaddress','msExchRecipientDisplayType','msExchRecipientTypeDetails','msExchUsageLocation'
                Server=$TADU.CanonicalName.split('/')[0] #get's the Target AD Users Domain to use as the server value with set-aduser
                ErrorAction = 'Stop'
            }#setaduserparams1
            try {
                Write-Log -message "Attempting: Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD" -Verbose
                set-aduser @setaduserparams1
                Write-Log -message "Succeeded: Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD" -Verbose
            }#try
            catch {
                Write-Log -message "FAILED: Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD" -Verbose -ErrorLog
                Write-Log -Message $_.tostring() -ErrorLog
                $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'FailedToClearTargetAttributes'}}
                Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                throw("Failed to clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD")
            }#catch
            #SetNewValuesOnTargetAttributes
            $setaduserparams2 = @{
                identity=$TADUGUID
                add=@{
                    targetaddress = "SMTP:$TargetDeliveryDomainAddress"
                    proxyaddresses = [string[]]$DesiredProxyAddresses
                    msExchRecipientDisplayType = -2147483642
                    msExchRecipientTypeDetails = 2147483648
                    msExchRemoteRecipientType = 4
                }
                Server=$TADU.CanonicalName.split('/')[0]
                ErrorAction = 'Stop'
            }#setaduserparams2
            if ($TADU.c) {$setaduserparams1.msExchangeUsageLocation = $TADU.c}
            try {
                Write-Log -message "Attempting: SET target attributes $($setaduserparams2.'Add'.keys -join ';') for $TADUGUID in $TargetAD" -Verbose
                set-aduser @setaduserparams2
                Write-Log -message "Succeeded: SET target attributes $($setaduserparams2.'Add'.keys -join ';') for $TADUGUID in $TargetAD" -Verbose
            }#try
            catch {
                Write-Log -message "FAILED: SET target attributes $($setaduserparams2.'Add'.keys -join ';')  for $ID in $TargetAD" -Verbose -ErrorLog
                Write-Log -Message $_.tostring() -ErrorLog
                $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'FailedToSetTargetAttributes'}}
                Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                throw("Failed to set target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD")
            }#catch
            #have exchange clean up and set version/legacyexchangedn
            #wait until Exchange "sees" the new attributes in the Global Catalog
            $writeProgressParams.status = "Enabling ADUser $TADUGUID in $TargetAD"
            Write-Progress @writeProgressParams
            $EnableADAccountParams = @{
                identity=$TADUGUID
                Server=$TADU.CanonicalName.split('/')[0]
                ErrorAction = 'Stop'
            }#EnableADAccountParams
            try {
                Write-Log -message "Attempting: Enable-ADAccount $TADUGUID in $TargetAD" -Verbose
                Enable-ADAccount @EnableADAccountParams
                Write-Log -message "Succeeded: Enable-ADAccount $TADUGUID in $TargetAD" -Verbose
            }#try
            catch {
                Write-Log -message "FAILED: Enable-ADAccount $TADUGUID in $TargetAD" -Verbose -ErrorLog
                Write-Log -Message $_.tostring() -ErrorLog
                $Global:Exceptions += $ID | Select-Object *,@{n='Exception';e={'FailedToEnableAccount'}}
                Export-Data -DataToExportTitle PostMailboxMigrationExceptionUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                throw("Failed to set target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD")
            }#catch

            $writeProgressParams.status = "Updating Recipient $TADUGUID in $TargetExchangeOrg"
            Write-Progress @writeProgressParams
            do {
                $count++
                start-sleep -Seconds 3
                Connect-Exchange -ExchangeOrganization $TargetExchangeOrg
                $Recipient = @(Invoke-ExchangeCommand -cmdlet get-recipient -ExchangeOrganization $TargetExchangeOrg -string "-Identity $TADUGUID -ErrorAction SilentlyContinue")
            }
            until ($Recipient.count -ge 1 -or $count -ge 15)
            #now that we found the object as a recipient, go ahead and run Update-Recipient against the object
            try {
                Write-Log -message "Attempting: Update Recipient $DesiredUPNAndPrimarySMTPAddress in $TargetExchangeOrg" -Verbose
                $Global:ErrorActionPreference = 'Stop'
                Connect-Exchange -ExchangeOrganization $TargetExchangeOrg
                Invoke-ExchangeCommand -cmdlet Update-Recipient -ExchangeOrganization $TargetExchangeOrg -string "-Identity $TADUGUID -ErrorAction Stop"
                $Global:ErrorActionPreference = 'Continue'
                Write-Log -message "Succeeded: Update Recipient $DesiredUPNAndPrimarySMTPAddress in $TargetExchangeOrg" -Verbose
            }
            catch {
                $Global:ErrorActionPreference = 'Continue'
                Write-Log -message "FAILED: Update Recipient $DesiredUPNAndPrimarySMTPAddress in $TargetExchangeOrg" -Verbose -ErrorLog
                Write-Log -message $_.tostring() -ErrorLog
                $Global:Exceptions += $DesiredUPNAndPrimarySMTPAddress | Select-Object *,@{n='Exception';e={'FailedToUpdateRecipient'}}
                Export-Data -DataToExportTitle TargetForestExceptionsUsers -DataToExport $Global:Exceptions[-1] -DataType csv -Append
                throw("Failed to Update Recipient for $TADUGUID in $TargetExchangeOrg")
            }
            $ProcessedUser = $TADU | Select-Object -Property SAMAccountName,DistinguishedName,@{n='UserPrincipalname';e={$DesiredUPNAndPrimarySMTPAddress}},@{n='ObjectGUID';e={$TADUGUID}}
            $Global:ProcessedUsers += $ProcessedUser
            Write-Log -Message "NOTE: Processing for $DesiredUPNAndPrimarySMTPAddress with GUID $TADUGUID in $TargetAD and $TargetExchangeOrg has completed successfully." -Verbose
        Export-Data -DataToExportTitle PostMailboxMigrationProcessedUsers -DataToExport $ProcessedUser -DataType csv -Append
    }#try
    catch {
        $_
    }
    }#foreach
}#process
end{
if ($Global:ProcessedUsers.count -ge 1) {
    Write-Log -Message "Successfully Processed $($Global:ProcessedUsers.count) Users." -Verbose
}
if ($Global:Exceptions.count -ge 1) {
    Write-Log -Message "Processed $($Global:Exceptions.count) Users with Exceptions." -Verbose
}
}#end
}#function
function Add-MSOLLicenseToUser {
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
        Write-Log -verbose -message "Custom Source Columns are $CustomSourceColumns." -logpath $logpath
        $RequiredColumns = @('UserPrincipalName','LicenseTypeDesired')
        Write-Log -verbose -message "Required Columns are $RequiredColumns." -logpath $logpath
        $proceed = $true
        foreach ($reqcol in $RequiredColumns) {
            if ($reqcol -notin $CustomSourceColumns) {
                $Proceed = $false
                Write-Log -errorlog -verbose -message "Required Column $reqcol is not found in the Custom Source data provided.  Processing cannot proceed." -logpath $logpath -errorlogpath $errorlogpath
            }
        }
        if ($Proceed) {
            $UsersToLicense = $CustomSource
            Write-Log -verbose -message "Custom Source Data Columns Validated." -logpath $logpath
        }
        else {
            $UsersToLicense = $null
            Write-Log -errorlog -verbose -message "ERROR: Custom Source Data Colums failed validation.  Processing cannot proceed." -logpath $logpath -errorlogpath $errorlogpath
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
                    Write-Log -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                    $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                }
            }
            else {$msollicenseoptionsparams.DisabledPlans = $Null}
            Write-Log -Message "Desired E4 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
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
                    Write-Log -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                    $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                }
            }
            else {$msollicenseoptionsparams.DisabledPlans = $Null}
            Write-Log -Message "Desired E3 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
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
                    Write-Log -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                    $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                }
            }
            else {$msollicenseoptionsparams.DisabledPlans = $Null}
            Write-Log -Message "Desired E2 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
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
                    Write-Log -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                    $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                }
            }
            else {$msollicenseoptionsparams.DisabledPlans = $Null}
            Write-Log -Message "Desired E1 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
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
                    Write-Log -verbose -message "Desired Disabled Plans have been calculated as follows: $DisabledPlans" -LogPath $LogPath
                    $msollicenseoptionsparams.DisabledPlans = $DisabledPlans
                }
            }
            else {$msollicenseoptionsparams.DisabledPlans = $Null}
            Write-Log -Message "Desired K1 License and License Options Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
            #Create License Options Object
            $LicenseOptions = New-MsolLicenseOptions @msollicenseoptionsparams
            $Proceed = $true
        }
        Default {
            $Proceed = $false
            Write-Log -Message "No License Desired (non E4,E3,E2,E1,K1) Determined for $CurrentADUPN." -Verbose -LogPath $LogPath
        }
    }
    #Lookup MSOL User Object
    if ($proceed) {
        $WriteProgressParams.CurrentOperation = "Looking up MSOL User Object."
        Write-Progress @WriteProgressParams
        Try {
            Write-Log -Message "Looking up MSOL User Object $CurrentAzureUPN for AD User Object $CurrentADUPN" -Verbose -LogPath $LogPath
            $CurrentMSOLUser = Get-MsolUser -UserPrincipalName $CurrentAzureUPN -ErrorAction Stop
            Write-Log -Message "Found MSOL User for $CurrentAzureUPN" -Verbose -LogPath $LogPath
            $Proceed = $true
        }
        Catch {
            $Proceed = $false
            Write-Log -Message "ERROR: MSOL User for $CurrentAzureUPN not found." -Verbose -LogPath $LogPath
            Write-Log -Message "ERROR: MSOL User for $CurrentAzureUPN not found." -LogPath $ErrorLogPath
            Write-Log -Message $_.tostring() -LogPath $ErrorLogPath
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
                Write-Log -Message "Usage Location for MSOL User $CurrentAzureUPN is set to $($CurrentMSOLUser.UsageLocation)." -LogPath $LogPath -Verbose
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
    Write-Log -Message "$CurrentADUPN license assignment status = $LicenseAssigned" -Verbose -LogPath $LogPath
    if ($CurrentMSOLUser.Licenses.AccountSkuId -contains $DesiredLicense) {$CorrectLicenseType = $True}
    else {
        $CorrectLicenseType = $false
        $LicenseToReplace = $CurrentMSOLUser.Licenses.AccountSkuID | where-object {$_ -in ($InterchangeableLicenses)}
    }
    Write-Log -Message "$CurrentADUPN correct license applied status = $CorrectLicenseType" -Verbose -LogPath $LogPath

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
    Write-Log -Message "$CurrentADUPN correct license options applied status = $correctLicenseOptions" -Verbose -LogPath $LogPath
    #Set Operation To Process on User
    $MSOLUserLicenseParams = @{}
    $MSOLUserLicenseParams.ErrorAction = 'Stop'
    $MSOLUserLicenseParams.UserPrincipalName = $CurrentAzureUPN
    if (-not $LicenseAssigned) {$LicenseOperation = 'Assign'}
    if ($licenseAssigned -and $CorrectLicenseType -and $correctLicenseOptions) {$LicenseOperation = 'None'}
    if ($licenseAssigned -and $CorrectLicenseType -and -not $correctLicenseOptions) {$LicenseOperation = 'Options'}
    if ($LicenseAssigned -and -not $CorrectLicenseType) {$LicenseOperation = 'Replace'}


    Write-Log -Message "$CurrentADUPN license operation selected = $LicenseOperation" -Verbose -LogPath $LogPath

    #Process License Operation
    switch ($LicenseOperation) {
        'None' {Write-Log -Message "$CurrentAzureUPN is already correctly licensed." -Verbose -LogPath $LogPath
        }
        'Assign'{
            Try {
                $MSOLUserLicenseParams.AddLicenses = $DesiredLicense
                $MSOLUserLicenseParams.LicenseOptions = $LicenseOptions
                Write-Log -Message "Setting User License for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                Set-MsolUserLicense @MSOLUserLicenseParams
                Write-Log -Message "Success: Assigned User License for $CurrentAzureUPN" -Verbose -LogPath $LogPath
            }
            Catch {
                Write-Log -Message "ERROR: License could not be assigned for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                Write-Log -Message "ERROR: License could not be assigned for $CurrentAzureUPN" -LogPath $ErrorLogPath
                Write-Log -Message $_.tostring() -Verbose -errorlogpath $ErrorLogPath
            }
        }
        'Replace' {
            Try {
                $MSOLUserLicenseParams.AddLicenses = $DesiredLicense
                $MSOLUserLicenseParams.LicenseOptions = $LicenseOptions
                $MSOLUserLicenseParams.RemoveLicenses = $LicenseToReplace
                Write-Log -Message "Replacing User License for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                Set-MsolUserLicense @MSOLUserLicenseParams
                Write-Log -Message "Success: Replaced User License for $CurrentAzureUPN" -Verbose -LogPath $LogPath
            }
            Catch {
                Write-Log -Message "ERROR: License could not be replaced for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                Write-Log -Message "ERROR: License could not be replaced for $CurrentAzureUPN" -LogPath $ErrorLogPath
                Write-Log -Message $_.tostring() -Verbose -LogPath $ErrorLogPath
            }
        }
        'Options' {
            Try {
                #$MSOLUserLicenseParams.AddLicenses = $DesiredLicense
                $MSOLUserLicenseParams.LicenseOptions = $LicenseOptions
                Write-Log -Message "Setting User License Options for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                Set-MsolUserLicense @MSOLUserLicenseParams
                Write-Log -Message "Success: Set User License Options for $CurrentAzureUPN" -Verbose -LogPath $LogPath
            }
            Catch {
                Write-Log -Message "ERROR: License options could not be set for $CurrentAzureUPN" -Verbose -LogPath $LogPath
                Write-Log -Message "ERROR: License options could not be set for $CurrentAzureUPN" -LogPath $ErrorLogPath
                Write-Log -Message $_.tostring() -Verbose -LogPath $ErrorLogPath
            }
        }
    }

}
}
}
function Set-ImmutableIDAttributeValue {
[cmdletbinding(
    DefaultParameterSetName='Single'
    ,
    SupportsShouldProcess=$true
)]
param(
[parameter(ParameterSetName = 'EntireForest')]
[switch]$EntireForest
,
[parameter(ParameterSetName = 'SearchBase',Mandatory = $true)]
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
[bool]$ExportResults = $true
)
Begin {
    #Check Current PSDrive Location: Should be AD, Should be GC, Should be Root of the PSDrive
    $Location = Get-Location
    $PSDriveTests = @{
        ProviderIsActiveDirectory = $($Location.Provider.ToString() -like '*ActiveDirectory*')
        LocationIsRootOfDrive = ($Location.Path.ToString() -eq $($Location.Drive.ToString() + ':\'))
        ProviderPathIsRootDSE = ($Location.ProviderPath.ToString() -eq '//RootDSE/')
    }#PSDriveTests
    if ($PSDriveTests.Values -contains $false) {
        Write-Log -ErrorLog -Verbose -Message "Set-ImmutableIDAttributeValue may not continue for the following reason(s) related to the command prompt location:"
        Write-Log -ErrorLog -Verbose -Message $($PSDriveTests.GetEnumerator() | Where-Object -filter {$_.Value -eq $False} | Select-Object @{n='TestName';e={$_.Key}},Value | ConvertTo-Json -Compress)
        Write-Error -Message "Set-ImmutableIDAttributeValue may not continue due to the command prompt location.  Review Error Log for details." -ErrorAction Stop
    }#If
    #Setup operational parameters for Get-ADObject based on Parameter Set
    $GetADObjectParams = @{
        Properties = @('CanonicalName',$ImmutableIDAttributeSource)
        ErrorAction = 'Stop'
    }#GetADObjectParams
    switch ($PSCmdlet.ParameterSetName) {
        'EntireForest' {
            $GetADObjectParams.ResultSetSize = $null
            $GetADObjectParams.Filter = {objectCategory -eq 'Person' -or objectCategory -eq 'Group'}
        }#EntireForest
        'Single' {
            #$GetADObjectParams.ResultSetSize = 1
        }#Single
        'SearchBase' {
            $GetADObjectParams.ResultSetSize = $null
            $GetADObjectParams.Filter = {objectCategory -eq 'Person' -or objectCategory -eq 'Group'}
            $GetADObjectParams.SearchBase = $SearchBase
            $GetADObjectParams.SearchScope = $SearchScope
        }#SearchBase
    }#Switch
    #Setup Export Files if $ExportResults is $true
    if ($ExportResults) {
        $ADObjectGetSuccesses = @()
        $ADObjectGetFailures = @()
        $Successes = @()
        $Failures = @()
        $ExportName = "SetImmutableIDAttributeValueResults"
    }#if
}#Begin
Process {
    if ($PSCmdlet.ParameterSetName -eq 'Single') {
        $GetADObjectParams.Identity = $Identity
    }#if
    Try {
        $logstring = $PSCmdlet.MyInvocation.InvocationName + ': Get AD Objects with the Get-ADObject cmdlet.'
        Write-Log -Message $logstring -Verbose -EntryType Attempting
        $adobjects = @(Get-ADObject @GetADObjectParams | Select-Object -ExcludeProperty Item,Property* -Property *,@{n='Domain';e={Get-AdObjectDomain -adobject $_}})
        $ObjectCount = $adobjects.Count
        $logstring = $PSCmdlet.MyInvocation.InvocationName + ": Get $ObjectCount AD Objects with the Get-ADObject cmdlet."
        if ($PSCmdlet.ParameterSetName -eq 'Single') {
            $ADObjectGetSuccesses += $Identity
        }
        Write-Log -Message $logstring -Verbose -EntryType Succeeded
        }#Try
    catch {
        Write-Log -Message $logstring -Verbose -EntryType Failed
        Write-Log -Message $_.tostring() -ErrorLog
        if ($ExportResults -and $PSCmdlet.ParameterSetName -eq 'Single') {
            $ADObjectGetFailures += $Identity | Select-Object @{n='Identity';e={$Identity}},@{n='TimeStamp';e={Get-TimeStamp}},@{n='Status';e={'Failed'}},@{n='ErrorString';e={$_.tostring()}}
        }
    }
    $O = 0 #Current Object Counter
    $adobjects | ForEach-Object {
        $CurrentObject = $_
        $O++ #Current Object Counter Incremented
        $LogString = "Set-ImmutableIDAttributeValue: Set Immutable ID Attribute $ImmutableIDAttribute for Object $($CurrentObject.ObjectGUID.tostring()) with the Set-ADObject cmdlet."
        Write-Progress -Activity "Setting Immutable ID Attribute for $ObjectCount AD Object(s)" -PercentComplete $($O/$ObjectCount*100) -CurrentOperation $LogString
        Try {
            if ($PSCmdlet.ShouldProcess($CurrentObject.ObjectGUID,"Set-ADObject $ImmutableIDAttribute with value $ImmutableIDAttributeSource")) {
                Write-Log -Message $LogString -EntryType Attempting
                Set-ADObject -Identity $CurrentObject.ObjectGUID -Add @{$ImmutableIDAttribute=$($CurrentObject.$($ImmutableIDAttributeSource))} -Server $CurrentObject.Domain -ErrorAction Stop -confirm:$false #-WhatIf
                Write-Log -Message $LogString -EntryType Succeeded
                if ($ExportResults) {
                    $Successes += $CurrentObject | Select-Object *,@{n='TimeStamp';e={Get-TimeStamp}},@{n='Status';e={'Succeeded'}},@{n='ErrorString';e={'None'}}
                }#if
            }#if
        }#try
        Catch {
            Write-Log -Message $LogString -EntryType Failed -ErrorLog -Verbose
            Write-Log -Message $_.ToString() -ErrorLog
            if ($ExportResults) {
                $Failures += $CurrentObject | Select-Object *,@{n='TimeStamp';e={Get-TimeStamp}},@{n='Status';e={'Failed'}},@{n='ErrorString';e={$_.tostring()}}
            }#if
        }#Catch
    }#ForEach-Object
    Write-Progress -Activity "Setting Immutable ID Attribute for $ObjectCount AD Object(s)" -Completed
}
End {
    If ($ExportResults) {
        if ($PSCmdlet.ParameterSetName -eq 'Single') {
            $AllLookupAttempts = $ADObjectGetSuccesses.Count + $ADObjectGetFailures.Count
            Write-Log -Message "Set-ImmutableIDAttributeValue Get AD Object Results: Total Attempts: $AllLookupAttempts; Successes: $($ADObjectGetSuccesses.Count); Failures: $($ADObjectGetFailures.count)" -Verbose
        }
        $AllResults = $Failures + $Successes
        Write-Log -message "Set-ImmutableIDAttributeValue Set AD Object Results: Total Attempts: $($AllResults.Count); Successes: $($Successes.Count); Failures: $($Failures.Count)." -Verbose
        Export-Data -DataToExportTitle $ExportName -DataToExport $AllResults -DataType csv
    }
    Write-Log -Message "Set-ImmutableIDAttributeValue Operations Completed." -Verbose
}
}
function Set-ExchangeAttributesOnTargetObject {
[cmdletbinding()]
param
(
[Parameter(ParameterSetName='SourceLookup')]
[switch]$SourceLookup
,
[Parameter(ParameterSetName='SourceDataProvided')]
[switch]$SourceDataProvided
,
[parameter(Mandatory=$true,ParameterSetName='SourceDataProvided')]
$SourceData
,
[parameter(Mandatory=$true,ParameterSetName='SourceLookup')]
[ValidateScript({$_ -in $(Get-PSDrive -PSProvider ActiveDirectory | Select-Object -ExpandProperty Name)})]
[string]$SourceAD 
,
[parameter(Mandatory = $true,ParameterSetName='SourceLookup')]
[validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','employeeNumber','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid','SID','msExchMasterAccountSID')]
[string]$SourceLookupAttribute
,
[parameter(Mandatory = $true,ParameterSetName='SourceLookup',ValueFromPipeline=$true)]
[string[]]$SourceLookupValue
,
<#
[parameter(ParameterSetName='SourceLookup')]
[validateset('UserMailbox','LinkedMailbox','LegacyMailbox','MailUser','RemoteUserMailbox')]
[string]$SourceRecipientTypeDetails
,
#>
[parameter(Mandatory=$true)]
[ValidateScript({$_ -in $(Get-PSDrive -PSProvider ActiveDirectory | Select-Object -ExpandProperty Name)})]
[string]$TargetAD
,
[parameter(Mandatory = $true)]
[validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','employeeNumber','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid','SID','msExchMasterAccountSID','GivenNameSurname')]
[string]$TargetLookupPrimaryAttribute
,
[parameter()]
[validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','employeeNumber','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid','SID','msExchMasterAccountSID','GivenNameSurname')]
[string]$TargetLookupSecondaryAttribute
,
[parameter(Mandatory = $true)]
[validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','employeeNumber','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid','SID','msExchMasterAccountSID','GivenNameSurname')]
[string]$TargetLookupPrimaryValue
,
[parameter(Mandatory = $true)]
[validateset('SAMAccountName','UserPrincipalName','ProxyAddress','Mail','employeeNumber','extensionattribute5','extensionattribute11','DistinguishedName','CanonicalName','ObjectGUID','mS-DS-ConsistencyGuid','SID','msExchMasterAccountSID','GivenNameSurname')]
[string]$TargetLookupSecondaryValue
,
[parameter(Mandatory = $true)]
[string]$TargetDeliveryDomain = $CurrentOrgProfile.office365tenants[0].TargetDomain
,
[parameter()]
[string]$ForceTargetPrimarySMTPDomain
,
[boolean]$DeleteContact = $true
,
[boolean]$DeleteSourceObject = $true
,
[boolean]$DisableEmailAddressPolicyInTarget = $true
,
[boolean]$ReplaceUPN = $false
,
[parameter(Mandatory=$false)]
[ValidateScript({$_ -in $(Get-PSSession | Where-Object ConfigurationName -eq 'Microsoft.Exchange' | Select-Object -ExpandProperty Name | ForEach-Object {$($_ -split '-')[0]})})]
[string]$TargetExchangeOrganization
,
[switch]$UpdateTargetRecipient
,
[string[]]$TargetAttributestoClear = 
    $(@(
        'Mail'
        'mailNickName'
        'msExchArchiveGUID'
        'msExchArchiveName'
        'msExchMailboxGUID'
        'msExchPoliciesExcluded'
        'msExchRecipientDisplayType'
        'msExchRecipientTypeDetails'
        'msExchRemoteRecipientType'
        'msExchUserCulture'
        'msExchVersion'
        'msExchUsageLocation'
        'proxyaddresses'
        'targetaddress'
        'extensionattribute5'
        'c'
        'co'
        #'countrycode'
    ))
,
[int]$ADSyncDelayInSeconds = 90
, 
[switch]$ClearGlobalTrackingVariables
,
[switch]$TestOnly
)#Param
begin 
{
    #Set up the global tracking/reporting variables if needed and/or clear them if requested
    $GlobalTrackingVariables = 
        @(
            'SEATO_Exceptions'
            ,'SEATO_ProcessedUsers'
            ,'SEATO_MailContactsFound'
            ,'SEATO_MailContactDeletionFailures'
            ,'SEATO_OriginalTargetUsers'
            ,'SEATO_OriginalSourceUsers'
            ,'SEATO_MembershipAddFailures'
        )
    if ($ClearGlobalTrackingVariables) 
    {
        foreach ($var in $GlobalTrackingVariables) 
        {
            Set-Variable -Scope Global -Name $var -Value @()
        }
    }
    else {
        foreach ($var in $GlobalTrackingVariables) {
            if (-not (Test-Path -Path "variable:global:$var")) {
                Set-Variable -Scope Global -Name $var -Value @()
            }
        }
    }
}#begin
process
{
#Process processes all incoming Source Lookup Values or Source Data and outputs Source Object Data
#setup for Source Operation specified
    switch ($PSCmdlet.ParameterSetName) 
    {
        'SourceLookup' 
        {
            #Populate SourceData
            #############################################################
            #lookup user in the Source AD
            #############################################################
            $SourceData = 
            @(
                $recordcount = $SourceLookupValue.Count
                $cr = 0
                foreach ($value in $SourceLookupValue) 
                {
                    $cr++
                    $writeProgressParams = 
                        @{
                            Activity = "Find Source Object"
                            Status = "Processing Record $cr of $recordcount : $value"
                            PercentComplete = $cr/$RecordCount*100
                        }#writeProgressParams
                    $writeProgressParams.currentoperation = "Find User with value $value in $SourceLookupAttribute in Source Object Forest $SourceAD"
                    Write-Progress @writeProgressParams
                    $TrialSADU = $null
                    $TrialSADU = 
                    @(
                        try {
                            Write-Log -message $writeProgressParams.currentoperation -EntryType Attempting
                            Find-ADUser -Identity $value -IdentityType $sourceLookupAttribute -AmbiguousAllowed -ActiveDirectoryInstance $SourceAD -ErrorAction Stop
                            Write-Log -message $writeProgressParams.currentoperation -EntryType Succeeded
                        }#try
                        catch {
                            Write-Log -message $writeProgressParams.currentoperation -Verbose -EntryType Failed -ErrorLog
                            Write-Log -Message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity $ID -ExceptionCode 'SourceADUserNotFound' -FailureGroup NotProcessed -RelatedObjectIdentifier $ID -RelatedObjectIdentifierType $SourceLookupAttribute
                        }
                    )#TrialSADU
                    #Determine action based on the results of the lookup attempt in the target AD
                    switch ($TrialSADU.count) {
                        1 {
                            Write-Log -message "Succeeded: Found exactly 1 Matching User with value $value in $SourceLookupAttribute in Source Object Forest $SourceAD"
                            #output the object into $SourceData
                            $TrialSADU[0]
                            Write-Log -Message "Source AD User Identified in with ObjectGUID: $($TrialSADU[0].objectguid)" -EntryType Notification
                        }#1
                        0 {
                            Write-Log -message "FAILED: Found 0 Matching Users with value $value in $SourceLookupAttribute in Source Object Forest $SourceAD" -Verbose
                            Export-FailureRecord -Identity $ID -ExceptionCode 'SourceADUserNotFound' -FailureGroup NotProcessed -RelatedObjectIdentifier $ID -RelatedObjectIdentifierType $SourceLookupAttribute
                        }#0
                        Default {
                            Write-Log -message "FAILED: Found multiple ambiguous Matching Users with value $value in $SourceLookupAttribute in Source Object Forest $SourceAD" -Verbose
                            Export-FailureRecord -Identity $ID -ExceptionCode 'SourceADUserAmbiguous' -FailureGroup NotProcessed -RelatedObjectIdentifier $ID -RelatedObjectIdentifierType $SourceLookupAttribute  
                        }#Default
                    }#switch $SADU.count
                }#foreach $value in $SourceLookupValue
            )#SourceData Array
         }#SourceLookup
        'SourceDataProvided' 
        {
            #validate attributes in SourceData?
        }#SourceDataProvided
    }#switch $PSCmdlet.ParameterSetName
    Write-Log -Message "Completed Source Object Lookup/Validation Operations" -EntryType Notification -Verbose
}#process
end
{
#End processes the Source Object Data, finds target and related object (such as contacts) and makes determinations about processing in the Intermediate Object stage
#End then processes the objects, adding/adjusting attribute values on the target objects and deleting source and related objects if specified
    $IntermediateObjects = 
    @(
        $recordcount = $SourceData.Count
        $cr = 0
        :nextID foreach ($SADU in $SourceData) 
        {
            #region Preparing To Generate Intermediate Object
            #region FindTADU
            $cr++
            $SADUGUID = $SADU.objectguid.guid
            #value is the value to use in lookup.  attribute is the attribute to lookup against.
            $ID = $SADU.$TargetLookupPrimaryValue
            if ($TargetLookupSecondaryAttribute) 
            {
                $trySecondary = $true
                $SecondaryID = $SADU.$targetLookupSecondaryValue
            }#if
            else 
            {
                $trySecondary = $false
            }#else
            $writeProgressParams = 
            @{
                Activity = "Preparing for User Updates"
                CurrentOperation = "Find Target Object with $ID in $TargetLookupPrimaryAttribute in $targetAD"
                Status = "Processing Record $cr of $recordcount : $ID"
                PercentComplete = $cr/$RecordCount*100
            }#writeProgressParams
            Write-Progress @writeProgressParams
            try 
            {
                Write-Log -Message $writeProgressParams.CurrentOperation -EntryType Attempting
                $TrialTADU = @(Find-Aduser -Identity $ID -IdentityType $TargetLookupPrimaryAttribute -ADInstance $TargetAD -ErrorAction Stop -AmbiguousAllowed)
                $TrialTADU = @($TrialTADU | Where-Object {$_.ObjectGUID -ne $SADUGUID})
                if ($TrialTADU.Count -eq 0 -and $trySecondary) 
                {
                    if ($TargetLookupSecondaryAttribute -eq 'GivenNameSurname') 
                    {
                        $GivenName = $SADU.GivenName
                        $SurName = $SADU.Surname
                        Write-log -Message "Attempting Secondary Attribute Lookup using GivenName: $givenName Surname: $SurName" -EntryType Notification
                        $TrialTADU = @(Find-ADUser -GivenName $GivenName -SurName $SurName -IdentityType GivenNameSurname -AmbiguousAllowed -ADInstance $TargetAD -ErrorAction Stop)
                    }
                    else 
                    {
                        Write-log -Message "Attempting Secondary Attribute Lookup using $secondaryID in $TargetLookupSecondaryAttribute" -EntryType Notification
                        $TrialTADU = @(Find-Aduser -Identity $SecondaryID -IdentityType $TargetLookupSecondaryAttribute -ADInstance $TargetAD -ErrorAction Stop -AmbiguousAllowed)
                    }
                }#if
                Write-Log -Message $writeProgressParams.CurrentOperation -EntryType Succeeded
            }#try
            catch 
            {
                Write-Log -Message $writeProgressParams.CurrentOperation -EntryType Failed -Verbose -ErrorLog
                Write-Log -Message $_.tostring() -ErrorLog
                Export-FailureRecord -Identity $ID -ExceptionCode 'TargetADUserNotFound' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType 'ObjectGUID' 
                continue nextID
            }#catch
            #Determine action based on the results of the lookup attempt in the target AD
            #filter SADU out of TADU results
            $TrialTADU = @($TrialTADU | Where-Object {$_.ObjectGUID -ne $SADUGUID})
            switch ($TrialTADU.count) 
            {
                1 
                {
                    Write-Log -message "Succeeded: Found exactly 1 Matching User" -Verbose
                    $TADU = $TrialTADU[0]
                    $TADUGUID = $TADU.objectguid.guid
                    Write-Log -Message "Target AD User Identified in $TargetAD with ObjectGUID: $TADUGUID" -Verbose -EntryType Notification
                }#1
                0 
                {
                    Write-Log -message "Found 0 Matching Users for User $ID" -Verbose -EntryType Failed
                    Export-FailureRecord -Identity $ID -ExceptionCode 'TargetADUserNotFound' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType 'ObjectGUID' 
                    continue nextID
                }#0
                Default 
                {#check for ambiguous results
                    Write-Log -Message "Find AD User returned multiple objects/ambiguous results for User $ID." -Verbose -ErrorLog -EntryType Failed
                    Export-FailureRecord -Identity $ID -ExceptionCode 'TargetADUserAmbiguous' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType 'ObjectGUID' 
                    continue nextID
                }
            }#switch
            #endregion FindTADU
            #region FindSADUExchangeDetails
            $writeProgressParams.CurrentOperation = "Get Source Object $ID Exchange Details"
            Write-Progress @writeProgressParams
            #Determine Source Object Exchange Recipient Status
            $SourceUserObjectIsExchangeRecipient = $(
                if ($SADU.msExchRecipientTypeDetails -ne $null -or $TADU.msExchRecipientDisplayType -ne $null) 
                {
                    $true
                } 
                else 
                {
                    $false
                }
            )
            if ($SourceUserObjectIsExchangeRecipient) {
                try 
                {
                    $SADUCurrentPrimarySmtpAddress = Find-PrimarySMTPAddress -ProxyAddresses $SADU.proxyaddresses -Identity $ID -ErrorAction Stop
                }#try
                catch 
                {
                    Write-Log -Message $_.tostring() -ErrorLog
                    Export-FailureRecord -Identity $ID -ExceptionCode 'SourceADUserPrimarySMTPNotFound' -FailureGroup NotProcessed
                    continue nextID
                }#catch
                $SADUUserObjectExchangeRecipientType = Get-RecipientType -msExchRecipientTypeDetails $SADU.msExchRecipientTypeDetails | Select-Object -ExpandProperty Name
                }
            #endregion FindSADUExchangeDetails
            #region FindTADUExchangeDetails
            #Determine Target Object Exchange Recipient Status
            $TargetUserObjectIsExchangeRecipient = $(
                if ($TADU.msExchRecipientTypeDetails -ne $null -or $TADU.msExchRecipientDisplayType -ne $null) 
                {
                    $true
                } 
                else 
                {
                    $false
                }
            )
            if ($TargetUserObjectIsExchangeRecipient) {
                try 
                {
                    $TADUCurrentPrimarySmtpAddress = Find-PrimarySMTPAddress -ProxyAddresses $TADU.proxyaddresses -Identity $ID -ErrorAction Stop
                }#try
                catch 
                {
                    Write-Log -Message $_.tostring() -ErrorLog
                    #Export-FailureRecord -Identity $ID -ExceptionCode 'TargetADUserPrimarySMTPNotFound' -FailureGroup NotProcessed
                    #continue nextID
                }#catch
                $TargetUserObjectExchangeRecipientType = Get-RecipientType -msExchRecipientTypeDetails $TADU.msExchRecipientTypeDetails | Select-Object -ExpandProperty Name
            }
            #endregion FindTADUExchangeDetails
            #region FindContacts
            #lookup mail contacts in the Target AD (using Source AD Proxy addresses, target address, and altRecipient)
            $writeProgressParams.currentOperation = "Get any mail contacts for $ID in target AD $TargetAD"
            Write-Progress @writeProgressParams
            $MailContacts = @()
            #look for contacts via proxy addresses
            :nextAddr foreach ($addr in $SADU.proxyaddresses) 
            {
                try 
                {
                    #Write-Log -message "Find Mail Contact for $addr in $TargetAD" -EntryType Attempting
                    $MailContact = @(Find-ADContact -Identity $addr -IdentityType ProxyAddress -AmbiguousAllowed -ActiveDirectoryInstance $TargetAD -ErrorAction Stop)
                    #Write-Log -message "No Errors: Find Mail Contact for $addr in $TargetAD" -EntryType Succeeded
                }#try
                catch 
                {
                    Write-Log -message "Unexpected Error: Find Mail Contact for $addr in $TargetAD" -EntryType Failed -Verbose -ErrorLog
                    Write-Log -message $_.tostring() -ErrorLog
                    Export-FailureRecord -Identity "$ID`:$addr" -ExceptionCode 'UnexpectedFailureDuringMailContactLookup' -FailureGroup ContactLookupFailure -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                    continue nextAddr
                }#catch
                If ($MailContact.count -ge 1) 
                {
                    Write-Log -Message "NOTE: A mail contact was found for $addr in $TargetAD." -Verbose
                    if ($MailContacts.distinguishedname -notcontains $MailContact.Distinguishedname) 
                    {
                        $MailContacts += $MailContact
                        $Global:SEATO_MailContactsFound += $MailContact
                    }
                }#if
            }#foreach
            #look for contacts via target address if target was not in proxy addresses
            if ($SADU.TargetAddress -ne $null -and $SADU.proxyaddresses -notcontains $SADU.TargetAddress) 
            {
                $addr = $SADU.targetaddress
                try 
                {
                    #Write-Log -message "Find Mail Contact for $addr in $TargetAD" -EntryType Attempting
                    $MailContact = @(Find-ADContact -Identity $addr -IdentityType ProxyAddress -AmbiguousAllowed -ActiveDirectoryInstance $TargetAD -ErrorAction Stop)
                    #Write-Log -message "No Errors: Find Mail Contact for $addr in $TargetAD" -EntryType Succeeded
                }#try
                catch 
                {
                    Write-Log -message "Unexpected Error: Find Mail Contact for $addr in $TargetAD" -EntryType Failed -Verbose -ErrorLog
                    Write-Log -message $_.tostring() -ErrorLog
                    Export-FailureRecord -Identity "$ID`:$addr" -ExceptionCode 'UnexpectedFailureDuringMailContactLookup' -FailureGroup ContactLookupFailure -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                }#catch
                If ($MailContact.count -ge 1) 
                {
                    Write-Log -Message "NOTE: A mail contact was found for $addr in $TargetAD." -Verbose
                    if ($MailContacts.distinguishedname -notcontains $MailContact.Distinguishedname) {
                        $MailContacts += $MailContact
                        $Global:SEATO_MailContactsFound += $MailContact
                    }
                }#if
            }
            #look for contacts via altRecipient if altRecipient has a value and it is not equal to the TADU DN
            if (-not [string]::IsNullOrWhiteSpace($SADU.altRecipient) -and $SADU.altRecipent -ne $TADU.distinguishedName)
            {
                $addr = $SADU.altRecipient
                try 
                {
                    #Write-Log -message "Find Mail Contact for $addr in $TargetAD" -EntryType Attempting
                    $MailContact = @(Find-ADContact -Identity $addr -IdentityType DistinguishedName -ActiveDirectoryInstance $TargetAD -ErrorAction Stop)
                    Write-Log -message "No Errors: Find Mail Contact for $addr in $TargetAD" -EntryType Succeeded
                }#try
                catch 
                {
                    Write-Log -message "Unexpected Error: Find Mail Contact for $addr in $TargetAD" -EntryType Failed -Verbose -ErrorLog
                    Write-Log -message $_.tostring() -ErrorLog
                    Export-FailureRecord -Identity "$ID`:$addr" -ExceptionCode 'UnexpectedFailureDuringMailContactLookup' -FailureGroup ContactLookupFailure -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                }#catch
                If ($MailContact.count -ge 1) 
                {
                    Write-Log -Message "NOTE: A mail contact was found for $addr in $TargetAD." -Verbose
                    if ($MailContacts.distinguishedname -notcontains $MailContact.Distinguishedname) 
                    {
                        $MailContacts += $MailContact
                        $Global:SEATO_MailContactsFound += $MailContact
                    }
                }#if
            }
            $addr = $null
            Write-Log -Message "A total of $($MailContacts.count) mail contacts were found for $ID in $TargetAD" -Verbose -EntryType Notification
            #endregion FindContacts
            #region BuildDesiredProxyAddresses
            #First, check desired Alias and desired PrimarySMTPAddress for conflicts
            $AliasAndPrimarySMTPAttemptCount = 0
            $ExemptObjectGUIDs = @($SADUGUID,$TADUGUID)
            Do {
                $AliasPass = $false
                $PrimarySMTPPass = $false
                $AliasAndPrimarySMTPAttemptCount++
                switch ($AliasAndPrimarySMTPAttemptCount) {
                    1 
                    {
                        $DesiredAlias = $SADU.givenname.substring(0,1) + $SADU.surname
                    }
                    2 
                    {
                        $DesiredAlias = $SADU.givenname.substring(0,2) + $SADU.surname
                    }
                    3 
                    {
                        $DesiredAlias = $SADU.givenname.substring(0,3) + $SADU.surname
                    }
                }
                $DesiredUPNAndPrimarySMTPAddress = $DesiredAlias + '@' + $ForceTargetPrimarySMTPDomain
                if (Test-ExchangeAlias -Alias $DesiredAlias -ExemptObjectGUIDs $ExemptObjectGUIDs -ExchangeOrganization $TargetExchangeOrganization) {
                    $AliasPass = $true
                }
                if (Test-ExchangeProxyAddress -ProxyAddress $DesiredUPNAndPrimarySMTPAddress -ProxyAddressType SMTP -ExemptObjectGUIDs $ExemptObjectGUIDs -ExchangeOrganization $TargetExchangeOrganization) {
                    $PrimarySMTPPass = $true
                }
            }
            Until (($AliasPass -and $PrimarySMTPPass) -or $AliasAndPrimarySMTPAttemptCount -gt 3)
            if ($AliasAndPrimarySMTPAttemptCount -gt 3) {
                    Write-Log -message "Was not able to find a valid alias and/or PrimarySMTPAddress to Assign to the target: $ID" -Verbose -EntryType Failed
                    Export-FailureRecord -Identity $ID -ExceptionCode 'InvalidAliasOrPrimarySMTPAddress' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                    continue nextID
            }
            #Build Proxy Address Array to use in Target AD
            $writeProgressParams.currentOperation = "Build Proxy Addresses Array for $ID"
            Write-Progress @writeProgressParams
            #setup for get-desiredproxyaddresses function to calculate updated addresses
            $GetDesiredProxyAddressesParams = @{
                CurrentProxyAddresses=$SADU.proxyAddresses
                LegacyExchangeDNs=@($SADU.legacyExchangeDN)
                Recipients = @($TADU)
                VerifyAddTargetAddress = $true
                DesiredOrCurrentAlias = $DesiredAlias
                TargetDeliveryDomain = $TargetDeliveryDomain
            }
            if (-not [string]::IsNullOrWhiteSpace($ForceTargetPrimarySMTPDomain)) {
                $GetDesiredProxyAddressesParams.DesiredPrimaryAddress=$DesiredUPNAndPrimarySMTPAddress
            }
            #include contacts legacyexchangedn as x500 and proxy addresses if contacts were found
            if ($MailContacts.Count -ge 1) {$GetDesiredProxyAddressesParams.legacyExchangeDNs += $MailContacts.LegacyExchangeDN; $GetDesiredProxyAddressesParams.Recipients += $MailContacts}
            $DesiredProxyAddresses = Get-DesiredProxyAddresses @GetDesiredProxyAddressesParams
            #endregion BuildDesiredProxyAddresses
            #endregion Preparing To Generate Intermediate Object
            #region IntermediateObjectGeneration
            $IntermediateObject = [pscustomobject]@{
                SourceUserObjectCN = $SADU.CanonicalName
                SourceUserObjectGUID = $SADUGUID
                SourceUserObjectEnabled = $SADU.Enabled
                SourceUserGivenName = $SADU.GivenName
                SourceUserSurName = $SADU.Surname
                SourceUserMail = $SADU.Mail
                SourceUserAlias = $SADU.mailNickname
                SourceUserObjectIsExchangeRecipient = $SourceUserObjectIsExchangeRecipient
                SourceUserObjectExchangeRecipientType = $SADUUserObjectExchangeRecipientType
                SourceUserPrimarySMTPAddress = $SADUCurrentPrimarySmtpAddress
                SourceUserObject = $SADU
                TargetUserObjectCN = $TADU.CanonicalName
                TargetUserObjectGUID = $TADUGUID
                TargetUserObjectEnabled = $TADU.Enabled
                TargetUserGivenName = $TADU.GivenName
                TargetUserSurName = $TADU.Surname
                TargetUserMail = $TADU.Mail
                TargetUserAlias = $TADU.mailNickname
                TargetUserObjectIsExchangeRecipient = $TargetUserObjectIsExchangeRecipient
                TargetUserObjectExchangeRecipientType = $TargetUserObjectExchangeRecipientType
                TargetUserObjectPrimarySMTPAddress = $TADUCurrentPrimarySmtpAddress
                TargetUserObject = $TADU
                MatchingContactObject = $MailContacts
                MatchedContactCount = $MailContacts.Count
                DesiredProxyAddresses = $DesiredProxyAddresses
                DesiredUPNAndPrimarySMTPAddress = $DesiredUPNAndPrimarySMTPAddress
                DesiredAlias = $DesiredAlias
                DesiredCoexistenceRoutingAddress = "O365_$($DesiredAlias)@$($SADUCurrentPrimarySmtpAddress.split('@')[1])"
                DesiredTargetAddress = "SMTP:$($DesiredAlias)@$($TargetDeliveryDomain)"
                ShouldUpdateUPN = if ($DesiredUPNAndPrimarySMTPAddress -ne $TADU.UserPrincipalName) {$true} else {$false}
            }
            #Below disabled due to new logic for determining TargetOperation below in #region write
            #$shouldProcessObject = if ($IntermediateObject.SourceUserObjectEnabled -or $IntermediateObject.TargetUserObjectIsExchangeRecipient) {$false} else {$true}
            #Write-Verbose "Should Process Object is $shouldProcessObject" -Verbose
            #$IntermediateObject | Add-Member -MemberType NoteProperty -Name ShouldProcessObject -Value $shouldProcessObject
            #output Intermediate Object to IntermediateObjects
            $IntermediateObject
            #endregion IntermediateObjectGeneration
        }#foreach $SADU in $SourceData
        $writeProgressParams.currentOperation = "Completing Intermediate Objects Operations"
        Write-Progress @writeProgressParams -Completed
    )#IntermediateObjects
    Write-Log -Message "$($IntermediateObjects.count) Object(s) Processed (Lookup of Source and Target Objects and Attribute Calculations)." -EntryType Notification
    #region CYABackup
    #depth must be 2 or greater to capture and restore MV attributes like proxy addresses correctly
    Export-Data -DataToExport $IntermediateObjects -DataToExportTitle IntermediateObjects -Depth 3 -DataType json
    #endregion CYABackup
    #############################################################
    #preparation activities complete, time to write changes to Target AD Objects
    #############################################################
    #region write
    if ($TestOnly) 
    {
        $IntermediateObjects
    }
    else {
        $recordcount = $IntermediateObjects.Count
        $cr = 0
        #Write to Target Objects, Delete Source Objects, Update Exchange Recipient for Target Objects
        $ProcessedObjects = 
        @(
            :nextIntObj 
            foreach ($IntObj in $IntermediateObjects) {
                #region PrepareForTargetOperation
                $cr++
                $writeProgressParams = @{
                    Activity = "Update Target Object"
                    Status = "Processing Record $cr of $recordcount : $value"
                    PercentComplete = $cr/$RecordCount*100
                }
                $TADUGUID = $IntObj.TargetUserObjectGUID
                $SADUGUID = $IntObj.SourceUserObjectGUID
                $TADU = $IntObj.TargetUserObject
                $SADU = $IntObj.SourceUserObject
                $TargetDomain = Get-ADObjectDomain -adobject $TADU
                $writeProgressParams.currentoperation = "Updating Attributes for $TADUGUID in $TargetAD using AD Cmdlets"
                Write-Progress @writeProgressParams
                #endregion PrepareForTargetOperation
                #region DetermineTargetOperation
                if ($intobj.TargetUserObjectIsExchangeRecipient) 
                {
                    switch -Wildcard ($IntObj.TargetUserObjectExchangeRecipientType) 
                    {
                        *Mailbox*
                        {
                            $TargetOperation = 'UpdateAndMigrateOnPremisesMailbox'
                        }
                        Default
                        {
                            $message = "Source Object $SADUGUID and/or Target Object $TADUGUID should not be processed because of dual enabled user accounts or dual recipient objects."
                            Write-Log -message $message -EntryType Failed -Verbose -ErrorLog
                            Export-FailureRecord -Identity $TADUGUID -ExceptionCode 'FailedDueToDualUserOrRecipient' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                            Continue nextIntObj
                        }
                    }
                }
                else
                {$TargetOperation = 'EnableRemoteMailbox'}
                #endregion DetermineTargetOperation
                #region PerformTargetAttributeUpdate
                switch ($TargetOperation) 
                {
                    'EnableRemoteMailbox' 
                    {
                        #############################################################
                        #clear the target attributes in target AD
                        #############################################################
                        #ClearTargetAttributes
                        $setaduserparams1 = @{
                            Identity=$TADUGUID
                            clear=$TargetAttributestoClear
                            Server=$TargetDomain
                            ErrorAction = 'Stop'
                        }#setaduserparams1
                        #add UPN to clear list if UPN needs to be replaced
                        if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                            $setaduserparams1.'Clear' += 'UserPrincipalName'
                        }
                        $message = "Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD"
                        try {
                            Write-Log -message $message -EntryType Attempting
                            set-aduser @setaduserparams1
                            Write-Log -message $message -EntryType Succeeded
                        }#try
                        catch {
                            Write-Log -message $message -EntryType Failed -Verbose -ErrorLog
                            Write-Log -Message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity $TADUGUID -ExceptionCode 'FailedToClearTargetAttributes' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                            Continue nextIntObj
                        }#catch
                        #############################################################
                        #set new values on the target attributes in target AD
                        #############################################################
                        $setaduserparams2 = @{
                            identity=$TADUGUID
                            add=@{
                                #Adjust this section to use parameters depending on the recipient type which should be created.  Following is currently set for Remote Mailbox.
                                msExchRecipientDisplayType = -2147483642 #RemoteUserMailbox
                                msExchRecipientTypeDetails = 2147483648 #RemoteUserMailbox
                                msExchRemoteRecipientType = 1 #ProvisionMailbox
                                msExchVersion = 44220983382016
                                targetaddress = $intobj.DesiredTargetAddress
                                #hard coding these for a customer for now
                                msExchUsageLocation = 'US'
                                c = 'US'
                                co = 'United States'
                                #countrycode = 840
                                extensionattribute5 = $($SADUGUID.tostring())
                            }
                            Server=$TargetDomain
                            ErrorAction = 'Stop'
                        }#setaduserparams2
                        #Disable Email Address Policy if admin user specified the parameter, otherwise leave status quo of source object
                        if ($DisableEmailAddressPolicyInTarget) {
                            $setaduserparams2.'add'.msExchPoliciesExcluded = '{26491cfc-9e50-4857-861b-0cb8df22b5d7}'
                        }
                        elseif (-not [string]::IsNullOrWhiteSpace($SADU.msExchPoliciesExcluded)) {
                            $setaduserparams2.'add'.msExchPoliciesExcluded = $SADU.msExchPoliciesExcluded
                        }
                        <#region NotUsing
                                if(-not [string]::IsNullOrWhiteSpace($sadu.displayname)) {
                                $setaduserparams2.'add'.DisplayName = [string]$($SADU.displayName)
                                }
                                if(-not [string]::IsNullOrWhiteSpace($sadu.department)) {
                                $setaduserparams2.'add'.department = [string]$($SADU.department)
                                }
                                if(-not [string]::IsNullOrWhiteSpace($SADU.msExchMailboxGUID)) {
                                $setaduserparams2.'add'.msExchMailboxGUID = [byte[]]$($SADU.msExchMailboxGUID)
                                }
                                if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveGUID)) {
                                $setaduserparams2.'add'.msExchArchiveGUID = [byte[]]$($SADU.msExchArchiveGUID)
                                }
                                if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveName)) {
                                $setaduserparams2.'add'.msExchArchiveName = [byte[]]$($SADU.msExchArchiveName)
                                }
                                if (-not [string]::IsNullOrWhiteSpace($TADU.c)) {
                                    $setaduserparams2.'add'.msExchangeUsageLocation = $TADU.c
                                }
                        endRegion NotUsing#>
                        if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredUPNAndPrimarySMTPAddress)) {
                            $setaduserparams2.'add'.Mail = $intobj.DesiredUPNAndPrimarySMTPAddress
                        }
                        if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredAlias)) {
                            $setaduserparams2.'add'.mailNickName = [string]$intobj.DesiredAlias
                        }
                        if(-not [string]::IsNullOrWhiteSpace($SADU.msExchangeUserCulture)) {
                            $setaduserparams2.'add'.msExchangeUserCulture = [string]$SADU.msExchangeUserCulture
                        }
                        if(-not [string]::IsNullOrWhiteSpace($IntObj.DesiredProxyAddresses)) {
                            $setaduserparams2.'add'.proxyaddresses = [string[]]$($IntObj.DesiredProxyAddresses)
                        }
                        if(-not [string]::IsNullOrWhiteSpace($sadu.msExchMasterAccountSID)) {
                            $setaduserparams2.'add'.msExchMasterAccountSID = [string]$($SADU.msExchMasterAccountSID)
                        }
                        if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                            $setaduserparams2.'add'.UserPrincipalName = $intobj.DesiredUPNAndPrimarySMTPAddress
                        }
                        try {
                            $message = "SET target attributes $($setaduserparams2.'Add'.keys -join ';') for $TADUGUID in $TargetAD"
                            Write-Log -message $message -EntryType Attempting
                            set-aduser @setaduserparams2
                            Write-Log -message $message -EntryType Succeeded
                        }#try
                        catch {
                            Write-Log -message "FAILED: SET target attributes $($setaduserparams2.'Add'.keys -join ';')  for $TADUGUID in $TargetAD" -Verbose -ErrorLog
                            Write-Log -Message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity $id -ExceptionCode 'FailedToSetTargetAttributes' -FailureGroup PartiallyProcessed -RelatedObjectIdentifier $TADUGUID -RelatedObjectIdentifierType ObjectGUID
                            continue nextIntObj
                        }#catch
                    }#EnableRemoteMailbox
                    'UpdateAndMigrateOnPremisesMailbox' 
                    {
                        #############################################################
                        #clear the target attributes in target AD
                        #############################################################
                        #ClearTargetAttributes
                        #Since this is an existing mailbox, we won't clear all the mail attributes
                        $ExemptTargetAttributesForExistingMailbox = 
                        @(
	                        'msExchArchiveGUID'
                            'msExchArchiveName'
                            'msExchMailboxGUID'
	                        'msExchRecipientDisplayType'
	                        'msExchRecipientTypeDetails'
	                        'msExchRemoteRecipientType'
	                        'msExchUserCulture'
                            'msExchVersion'
                        )
                        $TargetAttributestoClearForExistingMailbox = $TargetAttributestoClear | Where-Object {$_ -notin $ExemptTargetAttributesForExistingMailbox}
                        $setaduserparams1 = @{
                            Identity=$TADUGUID
                            clear=$TargetAttributestoClearForExistingMailbox
                            Server=$TargetDomain
                            ErrorAction = 'Stop'
                        }#setaduserparams1
                        #add UPN to clear list if UPN needs to be replaced
                        if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                            $setaduserparams1.'Clear' += 'UserPrincipalName'
                        }
                        $message = "Clear target attributes $($setaduserparams1.clear -join ',') for $TADUGUID in $TargetAD"
                        try {
                            Write-Log -message $message -EntryType Attempting
                            set-aduser @setaduserparams1
                            Write-Log -message $message -EntryType Succeeded
                        }#try
                        catch {
                            Write-Log -message $message -EntryType Failed -Verbose -ErrorLog
                            Write-Log -Message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity $TADUGUID -ExceptionCode 'FailedToClearTargetAttributes' -FailureGroup NotProcessed -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                            Continue nextIntObj
                        }#catch
                        #############################################################
                        #set new values on the target attributes in target AD
                        #############################################################
                        $setaduserparams2 = @{
                            identity=$TADUGUID
                            add=@{
                                #Adjust this section to use parameters depending on the recipient type which should be created.  Following is currently set for Existing Mailbox.
                                #msExchRecipientDisplayType = -2147483642 #RemoteUserMailbox
                                #msExchRecipientTypeDetails = 2147483648 #RemoteUserMailbox
                                #msExchRemoteRecipientType = 1 #ProvisionMailbox
                                #msExchVersion = 44220983382016
                                targetaddress = $intobj.DesiredTargetAddress
                                #hard coding these for a customer for now
                                msExchUsageLocation = 'US'
                                c = 'US'
                                co = 'United States'
                                #countrycode = 840
                                extensionattribute5 = $($SADUGUID.tostring())
                            }
                            Server=$TargetDomain
                            ErrorAction = 'Stop'
                        }#setaduserparams2
                        #Disable Email Address Policy if admin user specified the parameter, otherwise leave status quo of source object
                        if ($DisableEmailAddressPolicyInTarget) {
                            $setaduserparams2.'add'.msExchPoliciesExcluded = '{26491cfc-9e50-4857-861b-0cb8df22b5d7}'
                        }
                        elseif (-not [string]::IsNullOrWhiteSpace($SADU.msExchPoliciesExcluded)) {
                            $setaduserparams2.'add'.msExchPoliciesExcluded = $SADU.msExchPoliciesExcluded
                        }
                        <#region NotUsing
                                if(-not [string]::IsNullOrWhiteSpace($sadu.displayname)) {
                                $setaduserparams2.'add'.DisplayName = [string]$($SADU.displayName)
                                }
                                if(-not [string]::IsNullOrWhiteSpace($sadu.department)) {
                                $setaduserparams2.'add'.department = [string]$($SADU.department)
                                }
                                if(-not [string]::IsNullOrWhiteSpace($SADU.msExchMailboxGUID)) {
                                $setaduserparams2.'add'.msExchMailboxGUID = [byte[]]$($SADU.msExchMailboxGUID)
                                }
                                if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveGUID)) {
                                $setaduserparams2.'add'.msExchArchiveGUID = [byte[]]$($SADU.msExchArchiveGUID)
                                }
                                if(-not [string]::IsNullOrWhiteSpace($SADU.msExchArchiveName)) {
                                $setaduserparams2.'add'.msExchArchiveName = [byte[]]$($SADU.msExchArchiveName)
                                }
                                if (-not [string]::IsNullOrWhiteSpace($TADU.c)) {
                                    $setaduserparams2.'add'.msExchangeUsageLocation = $TADU.c
                                }
                                if(-not [string]::IsNullOrWhiteSpace($SADU.msExchangeUserCulture)) {
                                    $setaduserparams2.'add'.msExchangeUserCulture = [string]$SADU.msExchangeUserCulture
                                }
                                if(-not [string]::IsNullOrWhiteSpace($sadu.msExchMasterAccountSID)) {
                                    $setaduserparams2.'add'.msExchMasterAccountSID = [string]$($SADU.msExchMasterAccountSID)
                                }
                        endRegion NotUsing#>
                        if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredUPNAndPrimarySMTPAddress)) {
                            $setaduserparams2.'add'.Mail = $intobj.DesiredUPNAndPrimarySMTPAddress
                        }
                        if(-not [string]::IsNullOrWhiteSpace($intobj.DesiredAlias)) {
                            $setaduserparams2.'add'.mailNickName = [string]$intobj.DesiredAlias
                        }
                        if(-not [string]::IsNullOrWhiteSpace($IntObj.DesiredProxyAddresses)) {
                            $setaduserparams2.'add'.proxyaddresses = [string[]]$($IntObj.DesiredProxyAddresses)
                        }
                        if ($ReplaceUPN -and (-not [string]::IsNullOrWhiteSpace($IntObj.DesiredUPNAndPrimarySMTPAddress)) -and $IntObj.ShouldUpdateUPN) {
                            $setaduserparams2.'add'.UserPrincipalName = $intobj.DesiredUPNAndPrimarySMTPAddress
                        }
                        try {
                            $message = "SET target attributes $($setaduserparams2.'Add'.keys -join ';') for $TADUGUID in $TargetAD"
                            Write-Log -message $message -EntryType Attempting
                            set-aduser @setaduserparams2
                            Write-Log -message $message -EntryType Succeeded
                        }#try
                        catch {
                            Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                            Write-Log -Message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity $id -ExceptionCode 'FailedToSetTargetAttributes' -FailureGroup PartiallyProcessed -RelatedObjectIdentifier $TADUGUID -RelatedObjectIdentifierType ObjectGUID
                            continue nextIntObj
                        }#catch
                    }#UpdateAndMigrateOnPremisesMailbox
                }#Switch $TargetOperation
                #endregion PerformTargetAttributeUpdate
                #region GroupMemberships
                #############################################################
                #add TADU to memberof groups from SADU
                #############################################################
                $writeProgressParams = @{
                    Activity = "Update Target Object Group Memberships"
                    Status = "Processing Record $cr of $recordcount : $value"
                    PercentComplete = $cr/$RecordCount*100
                }
                $writeProgressParams.currentoperation = "Adding $TADUGUID to Groups using AD Cmdlets"
                Write-Progress @writeProgressParams
                if ($SADU.memberof.count -ge 1) {
                    foreach ($groupDN in $SADU.memberof) {
                        try {
                            $message = "Add $TADUGUID to group $groupDN"
                            $GroupObject = Get-ADGroup -Identity $groupDN -Properties CanonicalName
                            $Domain = Get-ADObjectDomain -adobject $GroupObject
                            Write-Log -message $message -EntryType Attempting
                            Add-ADGroupMember -Identity $groupDN -Members $TADUGUID -ErrorAction Stop -Confirm:$false -Server $Domain
                            Write-Log -message $message -EntryType Succeeded
                        }
                        catch {
                            Write-Log -message $message -EntryType Failed -Verbose -ErrorLog
                            Write-Log -Message $_.tostring() -ErrorLog
                            Export-FailureRecord -Identity $TADUGUID -ExceptionCode "GroupMembershipFailure:$groupDN" -FailureGroup GroupMembership -RelatedObjectIdentifier $GroupDN -RelatedObjectIdentifierType DistinguishedName
                        }
                    }
                }
                #endregion GroupMemberships
                #region ContactDeletionAndAttributeCopy
                #############################################################
                #delete contact objects found in the Target AD
                #############################################################
                $writeProgressParams = @{
                    Activity = "Process Target Object Related Contacts"
                    Status = "Processing Record $cr of $recordcount : $value"
                    PercentComplete = $cr/$RecordCount*100
                }
                $writeProgressParams.currentoperation = "Processing Contacts for $TADUGUID"
                Write-Progress @writeProgressParams
                if ($deletecontact -and $intobj.MatchingContactObject.count -ge 1) {
                    Write-Log -message "Attempting: Delete $($intobj.MatchingContactObject.count) Mail Contact(s) from $TargetAD" -Verbose
                    foreach ($c in $intobj.MatchinContactObject) {
                        try {
                            Write-Log -message "Attempting: Delete $($c.distinguishedname) Mail Contact from $TargetAD" -Verbose
                            Push-Location -StackName DeleteADObject
                            Set-Location $("$TargetAD" + ":")
                            $splat = @{Identity = $c.distinguishedname;Confirm=$false;ErrorAction='Stop'}
                            Remove-ADObject @splat
                            Pop-Location -StackName DeleteADObject
                            Write-Log -message "Succeeded: Delete $($c.distinguishedname) Mail Contact from $TargetAD" -Verbose
                        }#try
                        catch {
                            Pop-Location -StackName DeleteADObject
                            #$Global:ErrorActionPreference = 'Continue'
                            Write-Log -message "FAILED: Delete $($c.distinguishedname) Mail Contact from $TargetAD" -Verbose -ErrorLog
                            Write-Log -Message $_.tostring() -ErrorLog
                            $Global:SEATO_MailContactDeletionFailures+=$c
                        }#catch
                    }#foreach
                }#if
                #############################################################
                #copy contact object memberships to Target AD User
                #############################################################
                if ($deletecontact -and $intobj.MatchingContactObject.count -ge 1) {
                    Write-Log -message "Attempting: Add $TADUGUID to Contacts' Distribution Groups in $TargetAD" -Verbose
                    $Groups = @($intobj.MatchingContactObject.memberof)
                    foreach ($group in $Groups) {
                        if ($group) {
                            try {
                                $message = "Add $TADUGUID as member to group $group"
                                Write-Log -message $message -EntryType Attempting
                                Push-Location -StackName DeleteADObject
                                Set-Location $("$TargetAD" + ":")
                                $splat = @{Identity = $group;Confirm=$false;ErrorAction='Stop';Members=$TADUGUID}
                                Add-ADGroupMember @splat
                                Pop-Location -StackName DeleteADObject
                                Write-Log -message $message -EntryType Succeeded
                            }#try
                            catch {
                                Pop-Location -StackName DeleteADObject
                                Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                                Write-Log -Message $_.tostring() -ErrorLog
                                Export-FailureRecord -Identity $TADUGUID -ExceptionCode "GroupMembershipFailure:$group" -FailureGroup GroupMembership -RelatedObjectIdentifier $GroupDN -RelatedObjectIdentifierType DistinguishedName
                            }#catch
                        }#IF
                    }#foreach
                }#if
                #endregion ContactDeletionAndAttributeCopy
                #region DeleteSourceObject
                #############################################################
                #Delete SADU from AD
                #############################################################
                if ($DeleteSourceObject) {
                    try {
                        $message = "Remove Object $SADUGUID from AD $SourceAD"
                        Write-Log -message $message -EntryType Attempting
                        $Splat = @{
                            Identity = $SADUGUID
                            ErrorAction = 'Stop'
                            Confirm = $false
                            Server = Get-ADObjectDomain -adobject $SADU
                        }
                        Remove-ADObject @splat
                        Write-Log -message $message -EntryType Succeeded
                    }
                    catch {
                        Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                        Write-Log -Message $_.tostring() -ErrorLog
                        Export-FailureRecord -Identity $SADUGUID -ExceptionCode "SourceObjectRemovalFailure:$SADUGUID" -FailureGroup SourceObjectRemoval -RelatedObjectIdentifier $SADUGUID -RelatedObjectIdentifierType ObjectGUID
                    }
                }
                #endregion DeleteSourceObject
                #region RefreshTargetObjectRecipient
                #############################################################
                #Refresh Exchange recipient object for TADU
                #############################################################
                if ($UpdateTargetRecipient) {
                    $RecipientFound = $false
                    do {
                        Connect-Exchange -ExchangeOrganization $TargetExchangeOrganization
                        $RecipientFound = Invoke-ExchangeCommand -cmdlet 'Get-Recipient' -ExchangeOrganization $TargetExchangeOrganization -string "-Identity $TADUGUID -ErrorAction SilentlyContinue" -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 1
                    }
                    until ($RecipientFound)
                    try {
                        $message = "Update-Recipient $TADUGUID in Exchange Organization $TargetExchangeOrganization"
                        Write-Log -message $message -EntryType Attempting
                        $Splat = @{
                            Identity = $TADUGUID
                            ErrorAction = 'Stop'
                        }
                        $ErrorActionPreference = 'Stop'
                        Connect-Exchange -ExchangeOrganization $TargetExchangeOrganization
                        Invoke-ExchangeCommand -cmdlet 'Update-Recipient' -splat $Splat -ExchangeOrganization $TargetExchangeOrganization -ErrorAction Stop
                        Write-Log -message $message -EntryType Succeeded
                        $ErrorActionPreference = 'Continue'
                    }
                    catch {
                        $ErrorActionPreference = 'Continue'
                        Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                        Write-Log -Message $_.tostring() -ErrorLog
                        Export-FailureRecord -Identity $TADUGUID -ExceptionCode "UpdateRecipientFailure:$TADUGUID" -FailureGroup UpdateRecipient -RelatedObjectIdentifier $TADUGUID -RelatedObjectIdentifierType ObjectGUID
                    }
                }
                #endregion RefreshTargetObjectRecipient
                $IntObj
            }#foreach
        )#ProcessedObjects
#endregion write
        if ($ProcessedObjects.Count -ge 1) {
            #Start a Directory Synchronization to Azure AD Tenant 
            #Wait first for AD replication
            Write-Log -Message "Waiting for $ADSyncDelayInSeconds seconds for AD Synchronization before starting an Azure AD Directory Synchronization." -Verbose -EntryType Notification
            New-Timer -units Seconds -length $ADSyncDelayInSeconds -showprogress -Frequency 5 -voice
            #Write-Log -Message "Starting an Azure AD Directory Synchronization." -Verbose -EntryType Notification
            #Start-DirectorySynchronization
        }
        foreach ($IntObj in $ProcessedObjects) {
            $SADUGUID = $IntObj.SourceUserObjectGUID
            $TADUGUID = $IntObj.TargetUserObjectGUID
            $TADU = Find-ADUser -Identity $IntObj.TargetUserObjectGUID -IdentityType ObjectGUID -ActiveDirectoryInstance $TargetAD
            #region WaitforDirectorySynchronization
            #############################################################
            #Request Directory Synchronization and Wait for Completion to Set Forwarding
            #############################################################
            $GUIDMATCH = $false
            $TestDirectorySynchronizationParams = @{
                Identity = $IntObj.DesiredUPNAndPrimarySMTPAddress
                MaxSyncWaitMinutes = 10
                DeltaSyncExpectedMinutes = 2
                SyncCheckInterval = 15
                ExchangeOrganization = 'OL'
                RecipientAttributeToCheck = 'CustomAttribute5'
                RecipientAttributeValue = $SADUGUID
                InitiateSynchronization = $true
            }
            $DirSyncTest = Test-DirectorySynchronization @TestDirectorySynchronizationParams
            #endregion WaitforDirectorySynchronization
            #region SetMailboxForwarding
            if ($DirSyncTest) {
                try {
                    $message = "Set Exchange Online Mailbox $($IntObj.DesiredUPNAndPrimarySMTPAddress) for forwarding to $($IntObj.DesiredCoexistenceRoutingAddress)."
                    Connect-Exchange -ExchangeOrganization OL
                    $ErrorActionPreference = 'Stop'
                    Write-Log -message $message -EntryType Attempting
                    Invoke-ExchangeCommand -cmdlet 'Set-Mailbox' -ExchangeOrganization OL -string "-Identity $($IntObj.DesiredUPNAndPrimarySMTPAddress) -ForwardingSmtpAddress $($IntObj.DesiredCoexistenceRoutingAddress)" -ErrorAction Stop
                    Write-Log -message $message -EntryType Succeeded
                    $ErrorActionPreference = 'Continue'
                }
                catch {
                    Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                    Write-Log -Message $_.tostring() -ErrorLog
                    Export-FailureRecord -Identity $($IntObj.DesiredUPNAndPrimarySMTPAddress) -ExceptionCode "SetCoexistenceForwardingFailure:$($IntObj.DesiredUPNAndPrimarySMTPAddress)" -FailureGroup SetCoexistenceForwarding
                    $ErrorActionPreference = 'Continue'
                }
            }
            else {
                $message = "Set Exchange Online Mailbox $($IntObj.DesiredUPNAndPrimarySMTPAddress) for forwarding to $($IntObj.DesiredCoexistenceRoutingAddress). Sync Related Failure."
                Write-Log -message $message -Verbose -ErrorLog -EntryType Failed
                Export-FailureRecord -Identity $($IntObj.DesiredUPNAndPrimarySMTPAddress) -ExceptionCode "SetCoexistenceForwardingFailure:$($IntObj.DesiredUPNAndPrimarySMTPAddress)" -FailureGroup SetCoexistenceForwarding
            }
            #endregion SetMailboxForwarding
            #############################################################
            #Processing Complete: Report Results
            #############################################################
            $ProcessedUser = $TADU | Select-Object -Property SAMAccountName,DistinguishedName,UserPrincipalname,@{n='OriginalPrimarySMTPAddress';e={$IntObj.SourceUserMail}},@{n='CoexistenceForwardingAddress';e={$IntObj.DesiredCoexistenceRoutingAddress}},@{n='ObjectGUID';e={$_.ObjectGUID.GUID}},@{n='TimeStamp';e={Get-TimeStamp}}
            $Global:SEATO_ProcessedUsers += $ProcessedUser
            Write-Log -Message "NOTE: Processing for $($TADU.UserPrincipalName) with GUID $TADUGUID in $TargetAD has completed successfully." -Verbose
        }#foreach
        #region ReportAllResults
        if ($Global:SEATO_ProcessedUsers.count -ge 1) {
            Write-Log -Message "Successfully Processed $($Global:SEATO_ProcessedUsers.count) Users."
            Export-Data -DataToExportTitle TargetForestProcessedUsers -DataToExport $Global:SEATO_ProcessedUsers -DataType csv #-Append
        }
        if ($Global:SEATO_Exceptions.count -ge 1) {
            Write-Log -Message "Processed $($Global:SEATO_Exceptions.count) Users with Exceptions."
        }
        if ($Global:SEATO_MailContactsFound.count -ge 1) {
            Write-Log -Message "$($Global:SEATO_MailContactsFound.count) Contacts were found and are being exported."
            Export-Data -DataToExportTitle FoundMailContacts -DataToExport $Global:SEATO_MailContactsFound -Depth 2 -DataType xml
        }
        if ($Global:SEATO_OriginalTargetUsers.count -ge 1) {
            Write-Log -Message "$($Global:SEATO_OriginalTargetUsers.count) Original Target Users were attempted for processing and are being exported."
            Export-Data -DataToExportTitle OriginalTargetUsers -DataToExport $Global:SEATO_OriginalTargetUsers -Depth 2 -DataType xml
        }
        if ($Global:SEATO_MailContactDeletionFailures.Count -ge 1) {
            Write-Log -Message "$($Global:SEATO_MailContactDeletionFailures.Count) Mail Contact(s) NOT successfully deleted.  Exporting them for review."
            Export-Data -DataToExportTitle MailContactsNOTDeleted -DataToExport $Global:SEATO_MailContactDeletionFailures -DataType csv
        }
        #endregion ReportAllResults
    }#else
}#end
}#function
function Add-EmailAddress {
[cmdletbinding()]
param
(
[string]$Identity
, 
[string[]]$EmailAddresses
, 
[string]$ExchangeOrganization
)
    #Get the Recipient Object for the specified Identity
    try 
    {
        $message = "Get Recipient for Identity $Identity"
        #Write-Log -Message $message -EntryType Attempting -Verbose
        $Recipient = Invoke-ExchangeCommand -cmdlet Get-Recipient -string "-Identity $Identity -ErrorAction 'Stop'" -ErrorAction Stop
        #Write-Log -Message $message -EntryType Succeeded -Verbose
    }
    catch 
    {
        Write-Log -Message $message -EntryType Failed -Verbose -ErrorLog
        Write-Log -Message $_.tostring() -ErrorLog
        Return
    }
    #Determine the Set cmdlet to use based on the Recipient Object
    $cmdlet = Get-RecipientCmdlet -Recipient $Recipient -verb Set -ErrorAction Stop
    try 
    {
        $message = "Add Email Address $($EmailAddresses -join ',') to recipient $Identity"
        Write-Log -Message $message -EntryType Attempting -Verbose
        $splat = @{
            Identity = $Identity
            EmailAddresses = @{Add = $EmailAddresses}
            ErrorAction = 'Stop'
        }
        Invoke-ExchangeCommand -cmdlet $cmdlet -splat $splat -ExchangeOrganization $ExchangeOrganization -ErrorAction Stop
        Write-Log -Message $message -EntryType Succeeded -Verbose
    }
    catch 
    {
        Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
        Write-Log -Message $_.tostring() -ErrorLog
    }
}
function Remove-EmailAddress {
[cmdletbinding()]
param
(
[string]$Identity
, 
[string[]]$EmailAddresses
, 
[string]$ExchangeOrganization
)
    #Get the Recipient Object for the specified Identity
    try 
    {
        $message = "Get Recipient for Identity $Identity"
        #Write-Log -Message $message -EntryType Attempting -Verbose
        $Recipient = Invoke-ExchangeCommand -cmdlet Get-Recipient -string "-Identity $Identity -ErrorAction 'Stop'" -ErrorAction Stop
        #Write-Log -Message $message -EntryType Succeeded -Verbose
    }
    catch 
    {
        Write-Log -Message $message -EntryType Failed -Verbose -ErrorLog
        Write-Log -Message $_.tostring() -ErrorLog
        Return
    }
    #Determine the Set cmdlet to use based on the Recipient Object
    $cmdlet = Get-RecipientCmdlet -Recipient $Recipient -verb Set -ErrorAction Stop
    try 
    {
        $message = "Remove Email Address $($EmailAddresses -join ',') to recipient $Identity"
        Write-Log -Message $message -EntryType Attempting -Verbose
        $splat = @{
            Identity = $Identity
            EmailAddresses = @{Remove = $EmailAddresses}
            ErrorAction = 'Stop'
        }
        Invoke-ExchangeCommand -cmdlet $cmdlet -splat $splat -ExchangeOrganization $ExchangeOrganization -ErrorAction Stop
        Write-Log -Message $message -EntryType Succeeded -Verbose
    }
    catch 
    {
        Write-Log -Message $message -EntryType Failed -ErrorLog -Verbose
        Write-Log -Message $_.tostring() -ErrorLog
    }
}