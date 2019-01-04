    Function Set-UsageLocationForMSOLUser {
        
        [cmdletbinding()]
        param(
        [parameter(Mandatory)]
        [string]$UsageLocation
        ,
        [Parameter(Mandatory,ParameterSetName='UPN')]
        [string]$UserPrincipalName
        ,
        [Parameter(Mandatory,ParameterSetName='ObjectID')]
        [string]$ObjectID
        )

    
    }
