    Function Add-EmailAddress {
        
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
            #Write-OneShellLog -Message $message -EntryType Attempting -Verbose
            $Splat = @{
                Identity = $Identity
                ErrorAction = 'Stop'
            }
            $Recipient = Invoke-ExchangeCommand -cmdlet Get-Recipient -splat $Splat -ErrorAction Stop -ExchangeOrganization $ExchangeOrganization
            #Write-OneShellLog -Message $message -EntryType Succeeded -Verbose
        }
        catch
        {
            Write-OneShellLog -Message $message -EntryType Failed -Verbose -ErrorLog
            Write-OneShellLog -Message $_.tostring() -ErrorLog
            Return
        }
        #Determine the Set cmdlet to use based on the Recipient Object
        $cmdlet = Get-RecipientCmdlet -Recipient $Recipient -verb Set -ErrorAction Stop
        try
        {
            $message = "Add Email Address $($EmailAddresses -join ',') to recipient $Identity"
            Write-OneShellLog -Message $message -EntryType Attempting -Verbose
            $splat = @{
                Identity = $Identity
                EmailAddresses = @{Add = $EmailAddresses}
                ErrorAction = 'Stop'
            }
            Invoke-ExchangeCommand -cmdlet $cmdlet -splat $splat -ExchangeOrganization $ExchangeOrganization -ErrorAction Stop
            Write-OneShellLog -Message $message -EntryType Succeeded -Verbose
        }
        catch
        {
            Write-OneShellLog -Message $message -EntryType Failed -ErrorLog -Verbose
            Write-OneShellLog -Message $_.tostring() -ErrorLog
        }
    
    }
