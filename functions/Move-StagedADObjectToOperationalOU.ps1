    Function Move-StagedADObjectToOperationalOU {
        
        param(
        [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [string[]]$Identity
        ,
        [string]$DestinationOU
        )
        begin {}
        process {
            foreach ($I in $Identity) {
                try {
                    $message = "Find AD Object: $I"
                    Write-OneShellLog -Message $message -EntryType Attempting
                    $aduser = Get-ADObject -Identity $I -ErrorAction Stop
                    Write-OneShellLog -Message $message -EntryType Succeeded
                }#try
                catch {
                    Write-OneShellLog -Message $message -Verbose -EntryType Failed -ErrorLog
                    Write-OneShellLog -Message $_.tostring() -ErrorLog
                }#catch
                try {
                    $message = "Move-ADObject -Identity $I -TargetPath $DestinationOU"
                    Write-OneShellLog -Message $message -EntryType Attempting
                    $aduser | Move-ADObject -TargetPath $DestinationOU -ErrorAction Stop
                    Write-OneShellLog -Message $message -EntryType Succeeded
                }#try
                catch {
                    Write-OneShellLog -Message $message -Verbose -ErrorLog -EntryType Failed
                    Write-OneShellLog -Message $_.tostring() -ErrorLog
                }#catch
            }#foreach
        }
        end{}
    
    }
