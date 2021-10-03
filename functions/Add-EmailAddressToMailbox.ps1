function Add-EmailAddressToMailbox
{
    [cmdletbinding()]
    param(
        [parameter(Mandatory)]
        [string]$Identity
        ,
        [parameter(Mandatory)]
        [string[]]$EmailAddress
    )

    $EmailAddresses = @{
        Add = @()
    }

    foreach ($e in $EmailAddress)
    {
        $EmailAddresses.Add += $e
    }

    $SetMailboxParams = @{
        Identity       = $Identity
        EmailAddresses = $EmailAddresses
    }

    Set-Mailbox @SetMailboxParams

}
