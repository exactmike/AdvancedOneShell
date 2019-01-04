    Function New-NestingOrderedGroupArray {
        
        [cmdletbinding()]
        param(
        $Groups
        )
        $GroupsDNHash = @{}
        $groups | Select-Object -ExpandProperty DistinguishedName | ForEach-Object {$GroupsDNHash.$($_) = $true}
        $OutputGroups = @{}
        $NestingLevel = 0
        Do {
            foreach ($group in $Groups)
            {
                if ($NestingLevel -eq 0 -and $group.memberof.Count -eq 0)
                {
                    #these groups have no memberships in other groups and can only be containting groups so we create/populate them last
                    $Group | Add-Member -MemberType NoteProperty -Name NestingLevel -Value $NestingLevel
                    $OutputGroups.$($Group.DistinguishedName) = $Group
                    Write-Verbose -Message "added Group $($Group.DistinguishedName) to Output at Nesting Level $NestingLevel"
                }
                elseif ($NestingLevel -ge 1 -and $Group.memberof.Count -ge 1 -and (-not $OutputGroups.ContainsKey($($Group.DistinguishedName))))
                {
                    $testGroupMemberships = @{}
                    foreach ($membership in $group.memberof)
                    {
                        #if the member of is not in the Groups array then we ignore it
                        if ($GroupsDNHash.ContainsKey($membership))
                        {
                            #if the member of is in the groups array then we make sure that the group we would be a member of is created and populated after the member group
                            $testGroupMemberships.$($membership) = ($OutputGroups.ContainsKey($membership) -and $OutputGroups.$($membership).NestingLevel -lt $NestingLevel)
                        }
                    }
                    if ($testGroupMemberships.ContainsValue($false))
                    #do nothing yet - wait until no $false values appear
                    {} else
                    {
                        #add the group to the output at the current nesting level
                        $Group | Add-Member -MemberType NoteProperty -Name NestingLevel -Value $NestingLevel
                        $OutputGroups.$($Group.DistinguishedName) = $Group
                        Write-Verbose -Message "added Group $($Group.DistinguishedName) to Output at Nesting Level $NestingLevel"
                    }
                }
            }
            if ($OutputGroups.Keys.Count -eq $Groups.count)
            {
                Write-Verbose -Message "No More Nests Required"
                $NoMoreNests = $true
            }
            $NestingLevel++
        }
        Until
        ($NoMoreNests)
        $OrderedGroups = $OutputGroups.Values | Sort-Object -Property NestingLevel -Descending
        Write-Output $OrderedGroups
    
    }
