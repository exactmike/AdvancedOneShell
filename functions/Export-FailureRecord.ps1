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
        Export-OneShellData -DataToExportTitle $FailureGroup -DataToExport $ExceptionObject -Append -DataType csv -ErrorAction Stop
        $Global:SEATO_Exceptions += $ExceptionObject
        }
        catch {
        Write-OneShellLog -Message "FAILED: to write Exception Record for $identity with Exception Code $ExceptionCode and Failure Group $FailureGroup" -ErrorLog
        }
    
    }
