#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$SourceDirectory = (Join-Path (Split-Path $PSScriptRoot) 'active-directory')
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (Get-Module ActiveDirectory) { throw 'Run fixture tests in a fresh PowerShell process.' }

function Assert-Equal {
    param($Actual, $Expected, [string]$Label)
    if ($Actual -cne $Expected) {
        throw "${Label}: expected [$Expected], got [$Actual]."
    }
}

function Invoke-Report {
    param([string]$Name, [hashtable]$Arguments = @{})
    $reportPath = Join-Path $SourceDirectory ('Export-SwpAD' + $Name + '.ps1')
    $base = @{ SearchBase = 'OU=Lab,DC=example,DC=test'; OutputDirectory = $output }
    $result = & $reportPath @base @Arguments
    if ($result -isnot [System.IO.FileInfo]) { throw "Expected one CSV file for $Name." }
    if ($result.Name -notmatch '-\d{8}T\d{9}Z\.csv$') { throw "Missing UTC timestamp: $($result.Name)." }
    $script:lastFile = $result.FullName
    Import-Csv -LiteralPath $result.FullName
}

function Assert-Bucket {
    param([object[]]$Rows, [string]$Column, [string]$Name, [int]$Count)
    $match = @($Rows | Where-Object { $_.$Column -ceq $Name })
    Assert-Equal $match.Count 1 "$Column/$Name row count"
    Assert-Equal $match[0].Count $Count "$Column/$Name count"
}

$previousModulePath = $env:PSModulePath
$fixture = Join-Path ([System.IO.Path]::GetTempPath()) ('swp-ad-fixtures-' + [guid]::NewGuid().ToString('N'))
$output = Join-Path $fixture 'reports'
$fakeModule = Join-Path $fixture 'ActiveDirectory'
try {
    $null = New-Item -ItemType Directory -Path $output, $fakeModule
    @'
function Get-ADUser {
    [CmdletBinding()]
    param($Filter, $SearchBase, $SearchScope, $ResultPageSize, $Properties, $Server)
    $global:SwpAdLastQuery = $PSBoundParameters
    if ($global:SwpAdQueryFailure) {
        $global:SwpAdUsers[0]
        throw 'Synthetic directory query failure'
    }
    $global:SwpAdUsers
}
function Get-ADGroup {
    [CmdletBinding()]
    param($Filter, $SearchBase, $SearchScope, $ResultPageSize, $Properties, $Server)
    $global:SwpAdLastQuery = $PSBoundParameters
    $global:SwpAdGroups
}
Export-ModuleMember -Function Get-ADUser, Get-ADGroup
'@ | Set-Content -LiteralPath (Join-Path $fakeModule 'ActiveDirectory.psm1') -Encoding ASCII
    $env:PSModulePath = $fixture + [System.IO.Path]::PathSeparator + $previousModulePath

    $utcNow = [datetime]::UtcNow
    $global:SwpAdQueryFailure = $false
    $global:SwpAdLastQuery = $null
    $global:SwpAdUsers = @(
        [pscustomobject]@{
            Enabled = $true; whenCreated = $utcNow
            AccountExpirationDate = $utcNow.AddDays(-1); accountExpires = 12345L
            Manager = 'CN=Never Export'; Department = ' Lab North '; Office = '=1+1'; Title = 'Tech'; Company = 'Lab'
            mail = ' hidden@example.test '; EmployeeID = '-secret1'; employeeType = 'Employee'
        }
        [pscustomobject]@{
            Enabled = $false; whenCreated = $utcNow.AddYears(-2)
            AccountExpirationDate = $utcNow.AddDays(2); accountExpires = 12346L
            Department = 'lab north'; Office = 'Annex'; Title = ' '; Company = 'Lab'
            mail = 'HIDDEN@example.test'; EmployeeID = ' -SECRET1 '; employeeType = 'Temp'
        }
        [pscustomobject]@{
            Enabled = $true; whenCreated = $utcNow.AddDays(70)
            AccountExpirationDate = $utcNow.AddDays(70); accountExpires = 12347L
            Department = 'Another'; Office = '=1+1'; Title = 'Temp'; Company = ' '
            mail = 'else@example.test'; EmployeeID = '001'; employeeType = ' contractor '
        }
        [pscustomobject]@{ Enabled = $null; accountExpires = [int64]::MaxValue; Department = ' ' }
        [pscustomobject]@{ accountExpires = 8000L }
    )
    $global:SwpAdGroups = @(
        [pscustomobject]@{ Description = '=do not write'; ManagedBy = 'CN=Never Export' }
        [pscustomobject]@{ Description = ' ' }
    )

    $expected = @{
        AccountCreationTrend = 6; AccountExpiryBuckets = 5; AccountStateTotals = 3
        DepartmentDistribution = 4; DuplicateEmployeeIdCounts = 1; DuplicateMailCounts = 1
        EmployeeTypeCoverage = 3; GroupDocumentationCoverage = 2; OfficeDistribution = 4
        ProfileCompleteness = 4; ManagerReferenceCoverage = 1
    }
    foreach ($name in ($expected.Keys | Sort-Object)) {
        $fileName = Join-Path $SourceDirectory ('Export-SwpAD' + $name + '.ps1')
        $tokens = $null
        $parseErrors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($fileName, [ref]$tokens, [ref]$parseErrors)
        Assert-Equal @($parseErrors).Count 0 "$name parse errors"
        $directoryCommands = @($ast.FindAll({
            param($node)
            $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -match '-AD'
        }, $true) | ForEach-Object { $_.GetCommandName() })
        $expectedCommand = if ($name -eq 'GroupDocumentationCoverage') { 'Get-ADGroup' } else { 'Get-ADUser' }
        Assert-Equal ($directoryCommands -join ',') $expectedCommand "$name directory commands"
        $arguments = @{}
        switch ($name) {
            'AccountCreationTrend' { $arguments.Months = 4 }
            'AccountExpiryBuckets' { $arguments.DaysAhead = 30 }
            'DepartmentDistribution' { $arguments.Labels = @('Lab North','Lab South') }
            'OfficeDistribution' { $arguments.Labels = @('=1+1','Annex') }
            'EmployeeTypeCoverage' { $arguments.ApprovedValues = @('Employee','Contractor') }
            'GroupDocumentationCoverage' {
                $arguments.SearchScope = 'Subtree'
                $arguments.Server = 'dc1.example.test'
            }
        }

        $rows = @(Invoke-Report $name $arguments)
        Assert-Equal $rows.Count $expected[$name] "$name CSV rows"
        Assert-Equal $global:SwpAdLastQuery['SearchBase'] 'OU=Lab,DC=example,DC=test' "$name search base"
        Assert-Equal $global:SwpAdLastQuery['ResultPageSize'] 500 "$name page size"
        Assert-Equal $global:SwpAdLastQuery['Filter'] '*' "$name scoped filter"
        if ($name -eq 'GroupDocumentationCoverage') {
            Assert-Equal $global:SwpAdLastQuery['SearchScope'] 'Subtree' 'Explicit subtree'
            Assert-Equal $global:SwpAdLastQuery['Server'] 'dc1.example.test' 'Explicit server'
        }
        else {
            Assert-Equal $global:SwpAdLastQuery['SearchScope'] 'OneLevel' "$name default scope"
            Assert-Equal $global:SwpAdLastQuery.ContainsKey('Server') $false "$name default server"
        }
        switch ($name) {
            'AccountCreationTrend' {
                Assert-Bucket $rows PeriodUtc ($utcNow.ToString('yyyy-MM', [Globalization.CultureInfo]::InvariantCulture)) 1
                Assert-Bucket $rows PeriodUtc Earlier 1
                Assert-Bucket $rows PeriodUtc Unknown 3
            }
            'AccountExpiryBuckets' {
                foreach ($bucket in @('Expired','DueWithinWindow','Later','NoExpiryOrNotVisible','Unknown')) {
                    Assert-Bucket $rows Bucket $bucket 1
                }
            }
            'AccountStateTotals' {
                Assert-Bucket $rows State Enabled 2
                Assert-Bucket $rows State Disabled 1
                Assert-Bucket $rows State Unknown 2
            }
            'DepartmentDistribution' {
                Assert-Bucket $rows Label 'Lab North' 2
                Assert-Bucket $rows Label 'Lab South' 0
                Assert-Bucket $rows Bucket Other 1
                Assert-Bucket $rows Bucket MissingOrNotVisible 2
            }
            'OfficeDistribution' {
                Assert-Bucket $rows Label "'=1+1" 2
                Assert-Bucket $rows Label 'Annex' 1
                Assert-Bucket $rows Bucket MissingOrNotVisible 2
            }
            { $_ -in 'DuplicateEmployeeIdCounts','DuplicateMailCounts' } {
                Assert-Equal $rows[0].PopulatedObjects 3 "$name populated"
                Assert-Equal $rows[0].MissingOrNotVisible 2 "$name missing"
                Assert-Equal $rows[0].DuplicateValues 1 "$name groups"
                Assert-Equal $rows[0].ObjectsWithDuplicateValue 2 "$name affected"
            }
            'EmployeeTypeCoverage' {
                Assert-Bucket $rows Bucket Recognized 2
                Assert-Bucket $rows Bucket Other 1
                Assert-Bucket $rows Bucket MissingOrNotVisible 2
            }
            'GroupDocumentationCoverage' {
                foreach ($row in $rows) {
                    Assert-Equal $row.Present 1 'Group field present'
                    Assert-Equal $row.MissingOrNotVisible 1 'Group field missing'
                    Assert-Equal $row.Total 2 'Group population'
                }
            }
            'ManagerReferenceCoverage' {
                Assert-Equal $rows[0].Field Manager 'Manager field'
                Assert-Equal $rows[0].Present 1 'Manager present'
                Assert-Equal $rows[0].MissingOrNotVisible 4 'Manager missing'
                Assert-Equal $rows[0].Total 5 'Manager population'
            }
            'ProfileCompleteness' {
                $department = $rows | Where-Object Field -eq Department
                Assert-Equal $department.Present 3 'Department present'
                Assert-Equal $department.MissingOrNotVisible 2 'Department missing'
                Assert-Equal $department.Total 5 'Department population'
            }
        }

        $csvText = Get-Content -LiteralPath $script:lastFile -Raw
        if ($csvText -match 'hidden@example|secret1|CN=Never Export|=do not write|"Another"') {
            throw "Source attribute leaked in $name."
        }
        Write-Host "PASS $name"
    }

    $before = @(Get-ChildItem -LiteralPath $output).Count
    $global:SwpAdQueryFailure = $true
    $didThrow = $false
    try { $null = Invoke-Report AccountStateTotals }
    catch {
        if ($_.Exception.Message -ne 'Synthetic directory query failure') { throw }
        $didThrow = $true
    }
    Assert-Equal $didThrow $true 'Query failure surfaced'
    Assert-Equal @(Get-ChildItem -LiteralPath $output).Count $before 'No partial output on query failure'

    $global:SwpAdQueryFailure = $false
    $global:SwpAdUsers = @()
    $empty = @(Invoke-Report AccountStateTotals)
    Assert-Bucket $empty State Enabled 0
    Assert-Bucket $empty State Disabled 0
    Assert-Bucket $empty State Unknown 0
    Assert-Equal @(Get-ChildItem -LiteralPath $output).Count ($before + 1) 'Earlier CSV kept on rerun'

    $didThrow = $false
    $file = Join-Path $SourceDirectory 'Export-SwpADAccountStateTotals.ps1'
    try { $null = & $file -SearchBase 'DC=example,DC=test' -OutputDirectory $output }
    catch [System.Management.Automation.ParameterBindingException] { $didThrow = $true }
    Assert-Equal $didThrow $true 'Domain-root scope rejected'
    Write-Host 'PASS scope, no-match, no partial file, and rerun checks'
}
finally {
    Remove-Module ActiveDirectory -ErrorAction SilentlyContinue
    $env:PSModulePath = $previousModulePath
    Remove-Variable -Name SwpAdUsers, SwpAdGroups, SwpAdLastQuery, SwpAdQueryFailure -Scope Global -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $fixture) { Remove-Item -LiteralPath $fixture -Recurse -Force }
}
