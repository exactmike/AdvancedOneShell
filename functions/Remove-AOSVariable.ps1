    Function Remove-AOSVariable {
        
        param
        (
        [string]$Name
        )
        Remove-Variable -Scope Script -Name $name
    
    }
