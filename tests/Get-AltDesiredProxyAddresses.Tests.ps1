$CommandName = $MyInvocation.MyCommand.Name.Replace(".Tests.ps1", "")
. Join-Path (Join-Path $PSScriptRoot 'Functions') $($CommandName + '.ps1')


Describe "$CommandName Unit Tests" -Tag 'UnitTests' {
    Context "Validate parameters" {
        $paramCount = 19
        $defaultParamCount = 11
        [object[]]$params = (Get-ChildItem Function:\Get-AltDesiredProxyAddresses).Parameters.Keys
        $knownParameters = @("CurrentProxyAddresses","DesiredPrimarySMTPAddress","DesiredOrCurrentAlias","LegacyExchangeDNs","Recipients","AddPrimarySMTPAddressForAlias","AddTargetSMTPAddress","VerifyTargetSMTPAddress","TargetDeliverySMTPDomain","PrimarySMTPDomain","AddressesToRemove","AddressesToAdd","DomainsToAdd","ForceDomainsToAdd","DomainsToRemove","VerifySMTPAddressValidity","ExchangeSession","TestAddressAvailabilityInExchangeSession","TestAddressAvailabilityExemptGUID")
        It "Should contain specific parameters" {
            ( @(Compare-Object -ReferenceObject $knownParameters -DifferenceObject $params -IncludeEqual | Where-Object SideIndicator -eq "==").Count ) | Should Be $paramCount
        }
        It "Should contain $paramCount parameters" {
            $params.Count - $defaultParamCount | Should Be $paramCount
        }
    }
}

Describe "$CommandName Integration Tests" -Tags "IntegrationTests" {
    Context "Command preserves values from CurrentProxyAddresses" {
        $results = Get-AltDesiredProxyAddresses -CurrentProxyAddresses 'smtp:user@contoso.com','SMTP:user.name@contoso.com','smtp:user@fabrikam.com'
        It "Should be exactly the right values" {
            $results[0] | Should BeExactly 'smtp:user@contoso.com'
            $results[1] | Should BeExactly 'SMTP:user.name@contoso.com'
            $results[2] | Should BeExactly 'smtp:user@fabrikam.com'
        }
    }
}
Describe "$CommandName Integration Tests" -Tags "IntegrationTests" {
    Context "Command updates the PrimarySMTPAddress based on PrimarySMTPDomain parameter" {
        $results = Get-AltDesiredProxyAddresses -CurrentProxyAddresses 'smtp:user@contoso.com','SMTP:user.name@contoso.com','smtp:user@fabrikam.com' -PrimarySMTPDomain 'fabrikam.com'
        It "Should be exactly the right values" {
            $results[0] | Should BeExactly 'smtp:user@contoso.com'
            $results[1] | Should BeExactly 'smtp:user.name@contoso.com'
            $results[2] | Should BeExactly 'SMTP:user@fabrikam.com'
        }
        It "Should contain the right number of results" {
            $results.count | Should Be 3
        }
    }
}
Describe "$CommandName Integration Tests" -Tags "IntegrationTests" {
    Context "Command preserves the PrimarySMTPAddress of CurrentProxyAddresses demoting a primary value from AddressesToAdd" {
        $results = Get-AltDesiredProxyAddresses -CurrentProxyAddresses 'smtp:user@contoso.com','SMTP:user.name@contoso.com','smtp:user@fabrikam.com' -AddressesToAdd 'SMTP:user@northwindtraders.com'
        It "Should be exactly the right values" {
            $results[0] | Should BeExactly 'smtp:user@contoso.com'
            $results[1] | Should BeExactly 'smtp:user@fabrikam.com'
            $results[2] | Should BeExactly 'smtp:user@northwindtraders.com'
            $results[3] | Should BeExactly 'SMTP:user.name@contoso.com'
        }
        It "Should contain the right number of results" {
            $results.count | Should Be 4
        }
    }
}
Describe "$CommandName Integration Tests" -Tags "IntegrationTests" {
    Context "Command updates the PrimarySMTPAddress based on PrimarySMTPDomain parameter and demotes a primary value from AddressesToAdd" {
        $results = Get-AltDesiredProxyAddresses -CurrentProxyAddresses 'smtp:user@contoso.com','SMTP:user.name@contoso.com','smtp:user@fabrikam.com' -PrimarySMTPDomain 'fabrikam.com' -AddressesToAdd 'SMTP:user@northwindtraders.com'
        It "Should be exactly the right values" {
            $results[0] | Should BeExactly 'smtp:user@contoso.com'
            $results[1] | Should BeExactly 'smtp:user.name@contoso.com'
            $results[2] | Should BeExactly 'smtp:user@northwindtraders.com'
            $results[3] | Should BeExactly 'SMTP:user@fabrikam.com'
        }
        It "Should contain the right number of results" {
            $results.count | Should Be 4
        }
    }
}
Describe "$CommandName Integration Tests" -Tags "IntegrationTests" {
    Context "Command updates the PrimarySMTPAddress based on DesiredPrimarySMTPAddress parameter and demotes a primary value from AddressesToAdd" {
        $results = Get-AltDesiredProxyAddresses -CurrentProxyAddresses 'smtp:user@contoso.com','SMTP:user.name@contoso.com','smtp:user@fabrikam.com' -AddressesToAdd 'SMTP:user@northwindtraders.com' -DesiredPrimarySMTPAddress 'SMTP:user.name@northwindtraders.com'
        It "Should be exactly the right values" {
            $results[0] | Should BeExactly 'smtp:user@contoso.com'
            $results[1] | Should BeExactly 'smtp:user@fabrikam.com'
            $results[2] | Should BeExactly 'smtp:user.name@contoso.com'
            $results[3] | Should BeExactly 'smtp:user@northwindtraders.com'
            $results[4] | Should BeExactly 'SMTP:user.name@northwindtraders.com'
        }
        It "Should contain the right number of results" {
            $results.count | Should Be 5
        }
    }
}
Describe "$CommandName Integration Tests" -Tags "IntegrationTests" {
    Context "Command updates the PrimarySMTPAddress based on DesiredPrimarySMTPAddress parameter and demotes a primary value from AddressesToAdd" {
        $results = Get-AltDesiredProxyAddresses -CurrentProxyAddresses 'smtp:user@contoso.com','SMTP:user.name@contoso.com','smtp:user@fabrikam.com' -AddressesToAdd 'SMTP:user@northwindtraders.com' -DesiredOrCurrentAlias 'Me.User' -AddPrimarySMTPAddressForAlias -PrimarySMTPDomain 'fabrikam.com'
        It "Should be exactly the right values" {
            $results[0] | Should BeExactly 'smtp:user@contoso.com'
            $results[1] | Should BeExactly 'smtp:user.name@contoso.com'
            $results[2] | Should BeExactly 'smtp:user@northwindtraders.com'
            $results[3] | Should BeExactly 'smtp:user@fabrikam.com'
            $results[4] | Should BeExactly 'SMTP:Me.User@fabrikam.com'
        }
        It "Should contain the right number of results" {
            $results.count | Should Be 5
        }
    }
}