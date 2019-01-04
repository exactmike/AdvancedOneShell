    Function GetDesiredValueFromSourceObject {
        
        param
        (
            [string]$Formula
            ,
            [psobject]$InputObject
        )
        $ScriptBlock = [scriptblock]::Create($Formula)
        $InputObject | ForEach-Object -Process $ScriptBlock
    
    }
