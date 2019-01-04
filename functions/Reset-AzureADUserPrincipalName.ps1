    Function Reset-AzureADUserPrincipalName {
        
        [cmdletbinding()]
        param
        (
            [Parameter(Mandatory,ParameterSetName='ExistingUPN')]
            [string]$UserPrincipalName
            ,
            [Parameter(Mandatory,ParameterSetName='ObjectID')]
            [string]$ObjectID
            ,
            [parameter(Mandatory)]
            [string]$TenantName
            ,
            [Parameter(Mandatory)]
            [string]$DesiredUserPrincipalName
            ,
            [switch]$Verify
        )
        if ((Connect-MSOnlineTenant -Tenant $TenantName) -ne $true)
        {throw {"Could not connect to MSOnline Tenant $($TenantName)"}}
        $message = "Get Azure AD User using $($PSCmdlet.ParameterSetName) "
        $splat = @{
            ErrorAction = 'Stop'
        }#splat
        switch ($PSCmdlet.ParameterSetName)
        {
            'ExistingUPN'
            {
                $message = $message + $UserPrincipalName
                $splat.UserPrincipalName = $UserPrincipalName
            }#ExistingUPN
            'ObjectID'
            {
                $message = $message + $ObjectID
                $splat.ObjectID = $ObjectID
            }#objectID
        }#switch
        try
        {
            Write-OneShellLog -Message $message -EntryType Attempting
            $OriginalAzureADUser = Get-MsolUser @splat
            Write-OneShellLog -Message $message -EntryType Succeeded
        }#try
        catch
        {
            $myerror = $_
            Write-OneShellLog -Message $message -EntryType Failed -ErrorLog
            Write-OneShellLog -Message $myerror.tostring() -ErrorLog
            throw {$myerror}
        }#catch
        $message = "Get Tenant domain to use for temporary UPN value"
        try
        {
            Write-OneShellLog -Message $message -EntryType Attempting
            $TenantDomain = Get-MsolDomain -ErrorAction Stop | Where-Object -FilterScript {$_.Name -like '*.onmicrosoft.com' -and $_.name -notlike '*.mail.onmicrosoft.com'} | Select-Object -ExpandProperty Name
            Write-OneShellLog -Message $message -EntryType Succeeded
        }#try
        catch
        {
            $myerror = $_
            Write-OneShellLog -Message $message -EntryType Failed -ErrorLog
            Write-OneShellLog -Message $myerror.tostring() -ErrorLog
            throw {$myerror}
        }#catch
        $temporaryUPN = $OriginalAzureADUser.ObjectID.guid + '@' + $TenantDomain
        $message = "Set Azure AD User $($OriginalAzureADUser.ObjectID.guid) UserPrincipalName to temporary value $temporaryUPN"
        $splat = @{
            ObjectID = $OriginalAzureADUser.objectID.guid
            NewUserPrincipalName = $temporaryUPN
            ErrorAction = 'Stop'
        }#splat
        try
        {
            Write-OneShellLog -Message $message -EntryType Attempting
            Set-MsolUserPrincipalName @splat | Out-Null #temporary password output thrown away
            Write-OneShellLog -Message $message -EntryType Succeeded
        }#try
        catch
        {
            $myerror = $_
            Write-OneShellLog -Message $message -EntryType Failed -ErrorLog
            Write-OneShellLog -Message $myerror.tostring() -ErrorLog
            throw {$myerror}
        }#catch
        $message = "Set Azure AD User $($OriginalAzureADUser.ObjectID.guid) UserPrincipalName to Desired value $DesiredUserPrincipalName"
        $splat = @{
            ObjectID = $OriginalAzureADUser.objectID.guid
            NewUserPrincipalName = $DesiredUserPrincipalName
            ErrorAction = 'Stop'
        }#splat
        try
        {
            Write-OneShellLog -Message $message -EntryType Attempting
            Set-MsolUserPrincipalName @splat
            Write-OneShellLog -Message $message -EntryType Succeeded
        }#try
        catch
        {
            $myerror = $_
            Write-OneShellLog -Message $message -EntryType Failed -ErrorLog
            Write-OneShellLog -Message $myerror.tostring() -ErrorLog
            throw {$myerror}
        }#catch
        if ($PSBoundParameters.ContainsKey('Verify'))
        {
            $splat = @{
                ObjectID = $OriginalAzureADUser.objectID.guid
                ErrorAction = 'Stop'
            }#splat
            Get-MsolUser @splat
        }
    
    }
