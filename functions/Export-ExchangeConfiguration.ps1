function Export-ExchangeConfiguration
{
    [cmdletbinding()]
    param(
        [parameter(Mandatory)]
        [ValidateScript( { Test-Path -type Container -Path $_ })]
        [string]$OutputFolderPath
        ,
        [parameter(Mandatory)]
        [ValidateSet('Online', 'Exchange2013+')]
        [string]$OrganizationType
    )

    $ErrorActionPreference = 'Continue'

    $ExchangeConfiguration = @{
        OrganizationConfig              = Get-OrganizationConfig
        AdminAuditLogConfig             = Get-AdminAuditLogConfig
        PartnerApplication              = Get-PartnerApplication
        AuthServer                      = Get-AuthServer
        FederatedOrganizationIdentifier = Get-FederatedOrganizationIdentifier
        FederationTrust                 = Get-FederationTrust
        AvailabilityAddressSpace        = Get-AvailabilityAddressSpace
        AvailabilityConfig              = Get-AvailabilityConfig
        OrganizationRelationship        = Get-OrganizationRelationship
        SharingPolicy                   = Get-SharingPolicy
        MigrationConfig                 = Get-MigrationConfig
        MigrationEndpoint               = Get-MigrationEndpoint
        AcceptedDomain                  = Get-AcceptedDomain
        RemoteDomain                    = Get-RemoteDomain
        TransportConfig                 = Get-TransportConfig
        TransportRule                   = Get-TransportRule
        SMIMEConfig                     = Get-SmimeConfig
        EmailAddressPolicy              = Get-EmailAddressPolicy
        AddressBookPolicy               = Get-AddressBookPolicy
        OWAMailboxPolicy                = Get-OWAMailboxPolicy
        MobileDeviceMailboxPolicy       = Get-MobileDeviceMailboxPolicy
        ActiveSyncDeviceClass           = Get-ActiveSyncDeviceClass | Select-Object -Property DeviceModel, DeviceType -Unique
        ActiveSyncDeviceAccessRule      = Get-ActiveSyncDeviceAccessRule
        RetentionPolicy                 = Get-RetentionPolicy
        RetentionTag                    = Get-RetentionPolicyTag
        IntraOrganizationConnector      = @(Get-IntraOrganizationConnector)
    }



    switch ($OrganizationType)
    {

        'Exchange2013+'
        {
            $Premises = @{
                ExchangeServer                 = @(Get-ExchangeServer)
                SettingOverride                = Get-SettingOverride
                AuthConfig                     = Get-AuthConfig
                CmdletExtensionAgent           = Get-CmdletExtensionAgent
                HybridConfiguration            = Get-HybridConfiguration
                PendingFederatedDomain         = Get-PendingFederatedDomain
                DatabaseAvailabilityGroup      = Get-DatabaseAvailabilityGroup
                MailboxDatabase                = Get-MailboxDatabase
                MailboxServer                  = Get-MailboxServer
                DeliveryAgentConnector         = Get-DeliveryAgentConnector
                ForeignConnector               = Get-ForeignConnector
                FrontEndTransportService       = Get-FrontendTransportService
                MailboxTransportService        = Get-MailboxTransportService
                ReceiveConnector               = Get-ReceiveConnector
                SendConnector                  = Get-SendConnector
                TransportAgent                 = Get-TransportAgent
                TransportPipeline              = Get-TransportPipeline
                UPNSuffix                      = Get-UserPrincipalNamesSuffix
                ClientAccessServer             = Get-ClientAccessServer
                ClientAccessArray              = Get-ClientAccessArray
                PowershellVirtualDirectory     = Get-PowershellVirtualDirectory -adpropertiesonly
                ActiveSyncVirtualDirectory     = Get-ActiveSyncVirtualDirectory -adpropertiesonly
                OABVirtualDirectory            = Get-OabVirtualDirectory -adpropertiesonly
                OWAVirtualDirectory            = Get-OwaVirtualDirectory -adpropertiesonly
                ECPVirtualDirectory            = Get-EcpVirtualDirectory -adpropertiesonly
                WebServicesVirtualDirectory    = Get-WebServicesVirtualDirectory -adpropertiesonly
                MAPIVirtualDirectory           = Get-MapiVirtualDirectory -adpropertiesonly
                OutlookProvider                = Get-OutlookProvider
                OutlookAnywhere                = Get-OutlookAnywhere
                RPCClientAccess                = Get-RPCClientAccess
                AddressList                    = Get-AddressList
                GlobalAddressList              = Get-GlobalAddressList
                OfflineAddressBook             = Get-OfflineAddressBook
                IntraOrganizationConfiguration = Get-IntraOrganizationConfiguration
            }

            $Premises.NetworkConnectionInfo = @(
                foreach ($s in $Premises.ExchangeServer) { Get-NetworkConnectionInfo -Identity $s.fqdn }
            )

            $Premises.ExchangeCertificate = @(
                foreach ($s in $Premises.ExchangeServer) { Get-ExchangeCertificate -Server $s.fqdn }
            )

            $Online = @{ }
        }

        'Online'
        {
            $Online = @{
                OutboundConnector = Get-OutboundConnector
                InboundConnector  = Get-InboundConnector
            }
            $Online.IntraOrganizationConfiguration = @(
                foreach ($c in $ExchangeConfiguration.IntraOrganizationConnector)
                {
                    $null = $c.id -match "[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}"
                    Get-IntraOrganizationConfiguration -OrganizationGuid $matches[0]
                }
            )
            $Premises = @{ }
        }
    }

    $ExchangeConfigObject = New-Object -Property $($ExchangeConfiguration + $Premises + $Online) -TypeName PSCustomObject

    $DateString = Get-Date -Format yyyyMMddHHmmss

    $OutputFileName = $($ExchangeConfigObject.OrganizationConfig.Name.split('.')[0]) + 'ExchangeConfigOn' + $DateString + '.xml'
    $OutputFilePath = Join-Path $OutputFolderPath $OutputFileName

    $ExchangeConfigObject | Export-Clixml -Path $OutputFilePath -Encoding utf8

}