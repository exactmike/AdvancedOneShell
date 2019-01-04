    Function Get-AOSVariable {
        
        param
        (
        [string]$Name
        )
        Get-Variable -Scope Script -Name $name
    
    }
