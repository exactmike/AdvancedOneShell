    Function New-AOSVariable {
        
        param
        (
        [string]$Name
        ,
        $Value
        )
        New-Variable -Scope Script -Name $name -Value $Value
    
    }
