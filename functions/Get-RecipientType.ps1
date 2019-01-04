    Function Get-RecipientType {
        
        [cmdletbinding()]
        param
        (
        [parameter(ParameterSetName = 'msExchRecipientDisplayType')]
        [string]$msExchRecipientDisplayType
        ,
        [parameter(ParameterSetName = 'msExchRecipientTypeDetails')]
        [string]$msExchRecipientTypeDetails
        ,
        [parameter(ParameterSetName = 'msExchRemoteRecipientType')]
        [string]$msExchRemoteRecipientType
        )
        $msExchRecipientDisplayTypes = @(
            [pscustomobject]@{Value=0;Type='Shared Mailbox';Name='SharedMailbox'}
            [pscustomobject]@{Value=1;Type='Universal Distribution Group';Name='DistributionGroup'}
            [pscustomobject]@{Value=10;Type='Arbitration Mailbox';Name='ArbitrationMailbox'}
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
            [pscustomobject]@{Value=-2147481850;Type='Remote Room Mailbox';Name='RemoteRoomMailbox'}
        )
        $msExchRecipientTypeDetailsTypes = @(
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
            [pscustomobject]@{Value=2147483648;Type='Remote Mailbox';Name='RemoteUserMailbox'}
            [pscustomobject]@{Value=137438953472;Type='Team Mailbox';Name='TeamMailbox'}
            [pscustomobject]@{Value=4294967296;Type='Computer';Name='Computer'}
            [pscustomobject]@{Value=8589934592;Type='Remote Room Mailbox';Name='RemoteRoomMailbox'}
            [pscustomobject]@{Value=17179869184;Type='Remote Equipment Mailbox';Name='RemoteEquipmentMailbox'}
            [pscustomobject]@{Value=34359738368;Type='Remote Shared Mailbox';Name='RemoteSharedMailbox'}
            [pscustomobject]@{Value=68719476736;Type='Public Folder Mailbox';Name='PublicFolderMailbox'}
            [pscustomobject]@{Value=274877906944;Type='Remote Team Mailbox';Name='RemoteTeamMailbox'}
            [pscustomobject]@{Value=549755813888;Type='Monitoring Mailbox';Name='MonitoringMailbox'}
            [pscustomobject]@{Value=1099511627776;Type='Group Mailbox';Name='GroupMailbox'}
            [pscustomobject]@{Value=2199023255552;Type='Linked Room Mailbox';Name='LinkedRoomMailbox'}
            [pscustomobject]@{Value=4398046511104;Type='Audit Log Mailbox';Name='AuditLogMailbox'}
            [pscustomobject]@{Value=8796093022208;Type='Remote Group Mailbox';Name='RemoteGroupMailbox'}
            [pscustomobject]@{Value=17592186044416;Type='Scheduling Mailbox';Name='SchedulingMailbox'}
            [pscustomobject]@{Value=35184372088832;Type='Guest Mail User';Name='GuestMailUser'}
            [pscustomobject]@{Value=70368744177664;Type='Aux Audit Log Mailbox';Name='AuxAuditLogMailbox'}
            [pscustomobject]@{Value=140737488355328;Type='Supervisory Review Policy Mailbox';Name='SupervisoryReviewPolicyMailbox'}
        )
        $msExchRemoteRecipientTypeTypes = @(
            [PSCustomObject]@{Type ='ProvisionMailbox';Value = 1;Migrated=$false;Archive=$false}
            [PSCustomObject]@{Type ='ProvisionArchive';Value = 2;Migrated=$false;Archive=$false}
            [PSCustomObject]@{Type ='ProvisionMailbox,ProvisionArchive';Value = 3;Migrated=$false;Archive=$true}
            [PSCustomObject]@{Type ='Migrated';Value = 4;Migrated=$true;Archive=$false}
            [PSCustomObject]@{Type ='RoomMailbox';Value = 33;Migrated=$false}
            [PSCustomObject]@{Type ='RoomMailbox';Value = 36;Migrated=$true}
            [PSCustomObject]@{Type ='SharedMailbox';Value = 96;Migrated=$false}
            [PSCustomObject]@{Type ='SharedMailbox';Value = 100;Migrated=$true}
            [PSCustomObject]@{Type ='LinkedMailbox';Value = $null;Migrated=$null}
            [PSCustomObject]@{Type ='RemoteEquipmentMailbox';Value = 68}
            [PSCustomObject]@{Type ='RoomMailbox';Value = 32}
            [PSCustomObject]@{Type ='DiscoveryMailbox';Value = $null}
            [PSCustomObject]@{Type ='ArbitrationMailbox';Value = $null}
            [PSCustomObject]@{Type ='LegacyMailbox';Value = $null}
            [PSCustomObject]@{Type ='EquipmentMailbox';Value = 64}
            [PSCustomObject]@{Type ='MailContact';Value = $null}
            [PSCustomObject]@{Type ='MailForestContact';Value = $null}
            [PSCustomObject]@{Type ='MailUser';Value = $null}
            [PSCustomObject]@{Type ='MailUniversalDistributionGroup';Value = $null}
            [PSCustomObject]@{Type ='MailUniversalSecurityGroup';Value = $null}
            [PSCustomObject]@{Type ='DynamicDistributionGroup';Value = $null}
            [PSCustomObject]@{Type ='PublicFolder';Value = $null}
        )
        switch ($PSCmdlet.ParameterSetName)
        {
            'msExchRecipientDisplayType'
            {
                $msExchRecipientDisplayTypes | Where-Object -FilterScript {$_.Value -eq $msExchRecipientDisplayType}
            }
            'msExchRecipientTypeDetails'
            {
                $msExchRecipientTypeDetailsTypes | Where-object -FilterScript {$_.Value -eq $msExchRecipientTypeDetails}
            }
            'msExchRemoteRecipientType'
            {
                $msExchRemoteRecipientTypeTypes | Where-object -FilterScript {$_.Value -eq $msExchRemoteRecipientType}
            }
        }
    
    }
