    Function Get-AOSVariableValue {
        
        param
        (
        [string]$Name
        )
        Get-Variable -Scope Script -Name $name -ValueOnly
    
    }
