    Function New-NestingOrderedGroupArray
    {  
        [cmdletbinding()]
        param
        (
            $Groups
            ,
            [hashtable]$GroupsDNHash
        )
        $OutputGroups = @{}
        $NestingLevel = 0
        $NoMoreNests = $false
        $groups | Add-Member -memberType NoteProperty -Name BadMemberOf -Value @()
        Do {
            Write-Verbose -Message "Begin Nesting Level $NestingLevel"
            foreach ($group in $Groups)
            {
                if (-not $OutputGroups.ContainsKey($($group.DistinguishedName)))
                {
                    Write-Verbose -message "Nesting Level $NestingLevel; Processing Group $($group.DistinguishedName)"
                    if ($NestingLevel -eq 0 -and $group.memberof.Count -eq 0)
                    {
                        #these groups have no memberships in other groups and can only be containting groups so we create/populate them last
                        $Group | Add-Member -MemberType NoteProperty -Name NestingLevel -Value $NestingLevel
                        $OutputGroups.$($Group.DistinguishedName) = $Group
                        Write-Verbose -Message "Nesting Level $NestingLevel; added Group $($Group.DistinguishedName) to Output"
                    }
                    elseif ($NestingLevel -eq 0 -and $group.MemberOf.Count -ge 1)
                    {
                        #these groups have memberships in other groups but they may not be mail enabled so we might treat them like they don't
                        $Tests = @(
                            foreach ($mo in $group.MemberOf)
                            {
                                $GroupsDNHash.ContainsKey($mo)
                                #if none of these MemberOf are in other groups in our set then we don't care about them for nesting purposes and threat this group as nesting 0
                            }
                        )
                        if ($Tests -notcontains $true)
                        {
                            $Group | Add-Member -MemberType NoteProperty -Name NestingLevel -Value $NestingLevel
                            $OutputGroups.$($Group.DistinguishedName) = $Group
                            Write-Verbose -Message "Nesting Level $NestingLevel; added Group $($Group.DistinguishedName) to Output"
                        }
                    }
                    elseif ($NestingLevel -ge 1 -and $Group.memberof.Count -ge 1)
                    {
                        $Tests = @(
                            foreach ($mo in $group.MemberOf)
                            {
                                switch ($GroupsDNHash.ContainsKey($mo))
                                {
                                    $true
                                    {
                                        $InitialResult = ($OutputGroups.ContainsKey($mo) -and $OutputGroups.$($mo).NestingLevel -lt $NestingLevel)
                                        if ($false -eq $InitialResult)
                                        {
                                            if  ($group.DistinguishedName -in $GroupsDNHash.$mo.MemberOf)
                                            {
                                                Write-Warning -Message "Found Mutually Nested groups $($group.DistinguishedName) and $mo"
                                                $group.MemberOf.Remove($mo)
                                                $group.BadMemberOf += $mo
                                                $true    
                                            }
                                            else 
                                            {
                                                $false
                                            }
                                        }
                                        else
                                        {
                                            $InitialResult
                                        }
                                    }
                                    $false
                                    {
                                        $true
                                    }
                                }
                            }    
                        )
                        if ($Tests -notcontains $false)
                        {
                            #add the group to the output at the current nesting level
                            $Group | Add-Member -MemberType NoteProperty -Name NestingLevel -Value $NestingLevel
                            $OutputGroups.$($group.DistinguishedName) = $Group
                            Write-Verbose -Message "Nesting Level $NestingLevel; added Group $($Group.DistinguishedName) to Output"
                        }
                    }
                }
            }
            Write-Verbose -Message "End Nesting Level $NestingLevel"
            if ($OutputGroups.Keys.Count -eq $Groups.count)
            {
                Write-Verbose -Message "$($OutputGroups.Keys.Count) of $($Groups.Count) completed"
                Write-Verbose -Message "No More Nests Required"
                $NoMoreNests = $true
            }
            else
            {
                Write-Verbose -Message "$($OutputGroups.Keys.Count) of $($Groups.Count) completed."
            }            
            $NestingLevel++
        }
        Until
        ($NoMoreNests)
        $OrderedGroups = $OutputGroups.Values | Sort-Object -Property NestingLevel -Descending
        $OrderedGroups
    }
