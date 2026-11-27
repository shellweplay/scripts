$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$temp = Join-Path ([IO.Path]::GetTempPath()) ('swp-utilities-' + [guid]::NewGuid())
New-Item -Path $temp -ItemType Directory | Out-Null

function Assert-True($Value, [string] $Message) {
    if (!$Value) { throw $Message }
}

# Offline fixture only: shadow the module import and directory cmdlets.
function Import-Module { param([string] $Name) Assert-True ($Name -eq 'ActiveDirectory') 'Unexpected module import.' }
function Get-ADUser {
    param($LDAPFilter, $Server, $SearchBase, $SearchScope, $ResultPageSize, $Properties)
    Assert-True ($Server -eq 'dc01.example.test') 'Unexpected server.'
    if ($LDAPFilter.Contains('broken')) { throw 'Simulated directory read failure.' }
    if ($LDAPFilter.Contains('demo1')) { [pscustomobject]@{ SamAccountName = 'demo1'; Enabled = $true; SID = 'S-1-5-21-10-20-30-1101'; adminCount = 1 } }
    if ($LDAPFilter.Contains('demo2')) { [pscustomobject]@{ SamAccountName = 'demo2'; Enabled = $false; SID = 'S-1-5-21-10-20-30-1102'; adminCount = 1 } }
    if ($LDAPFilter -like '*literal*') { Assert-True ($LDAPFilter.Contains('literal\2a\28x\29')) 'Roster did not escape filter metacharacters.' }
}
function Get-ADGroupMember {
    param($Identity, [switch] $Recursive, $Server)
    Assert-True $Recursive 'Expected recursive membership.'
    if ($Identity.EndsWith('1202')) { throw 'Simulated group failure.' }
    [pscustomobject]@{ ObjectClass = 'user'; SID = 'S-1-5-21-10-20-30-1101' }
}
function Get-ADGroup {
    param($LDAPFilter, $Filter, $Server, $SearchBase, $SearchScope, $ResultPageSize, $Properties)
    Assert-True ($Server -eq 'dc01.example.test' -and $SearchBase) 'Expected one scoped server.'
    if ($LDAPFilter) {
        Assert-True ($LDAPFilter -eq '(info=*note\2a\28x\29*)') 'Notes search did not escape the literal text.'
    }
    [pscustomobject]@{
        ObjectGUID = [guid]'82ab1625-325d-4989-9407-dd0b85a0828b'; SamAccountName = 'lab-group'
        whenCreated = [datetime]'2025-01-01'; whenChanged = [datetime]'2025-02-01'; Description = '=fixture'; Mail = 'group@example.test'; Info = 'Synthetic note'
    }
}
try {
    $server = 'dc01.example.test'
    $base = 'OU=People,DC=example,DC=test'
    $inputPath = Join-Path $temp 'accounts.txt'
    $output = Join-Path $temp 'roster.csv'
    @('demo1', 'demo2', 'missing', 'broken', 'literal*(x)') | Set-Content -LiteralPath $inputPath
    & "$root/active-directory/Export-SwpADUserRosterStatus.ps1" -Server $server -SearchBase $base -InputPath $inputPath -OutputPath $output -BatchSize 1
    $rows = @(Import-Csv -LiteralPath $output)
    Assert-True (($rows.Status -join ',') -eq 'Enabled,Disabled,NotFound,LookupFailed,NotFound') 'Incorrect roster status handling.'
    $rejected = $false
    try { & "$root/active-directory/Export-SwpADUserRosterStatus.ps1" -Server $server -SearchBase $base -InputPath $inputPath -OutputPath $output } catch { $rejected = $true }
    Assert-True $rejected 'Roster accepted overwriting an existing report.'

    $inputPath = Join-Path $temp 'accounts.csv'
    $output = Join-Path $temp 'review.csv'
    @('SamAccountName', 'demo1', 'demo2', 'broken') | Set-Content -LiteralPath $inputPath
    & "$root/active-directory/Export-SwpADAdminCountReview.ps1" -Server $server -InputPath $inputPath -ApprovedGroup 'S-1-5-21-10-20-30-1201','S-1-5-21-10-20-30-1202' -OutputPath $output
    $rows = @(Import-Csv -LiteralPath $output)
    Assert-True (($rows.ReviewStatus -join ',') -eq 'ApprovedGroupMatch,MembershipUnknown,LookupFailed') 'Incorrect partial group handling.'

    $output = Join-Path $temp 'notes.csv'
    & "$root/active-directory/Find-SwpADGroupsByNote.ps1" -Server $server -SearchBase $base -Text 'note*(x)' -OutputPath $output
    $rows = @(Import-Csv -LiteralPath $output)
    Assert-True ($rows.Count -eq 1 -and $rows[0].SamAccountName -eq 'lab-group' -and $rows[0].PSObject.Properties.Name -notcontains 'Info') 'Notes results disclose text or failed to return the group.'

    $output = Join-Path $temp 'attributes.csv'
    & "$root/active-directory/Export-SwpADGroupAttributes.ps1" -Server $server -SearchBase $base,'OU=Child,OU=People,DC=example,DC=test' -OutputPath $output
    $rows = @(Import-Csv -LiteralPath $output)
    Assert-True ($rows.Count -eq 1 -and $rows[0].Description -eq "'=fixture" -and $rows[0].PSObject.Properties.Name -notcontains 'Notes') 'Attributes export did not deduplicate, escape formulas, or omit notes.'

    $inputs = Join-Path $temp 'samples'
    New-Item -ItemType Directory -Path $inputs | Out-Null
    @('Status,Name', 'Active,=example', 'Archived,other') | Set-Content -LiteralPath (Join-Path $inputs 'sample-a.csv')
    @('Name,Status', 'entry,active') | Set-Content -LiteralPath (Join-Path $inputs 'sample-b.csv')
    $output = Join-Path $temp 'matched.csv'
    & "$root/powershell/Export-SwpMatchingCsvRows.ps1" -InputDirectory $inputs -Column Status -Equals Active -OutputPath $output
    $rows = @(Import-Csv -LiteralPath $output)
    Assert-True ($rows.Count -eq 2 -and $rows[0].Name -eq "'=example" -and $rows[1].SourceFile -eq 'sample-b.csv') 'CSV utility failed matching, escaping, or tracking source.'
    @('Status,Other', 'Active,test') | Set-Content -LiteralPath (Join-Path $inputs 'sample-c.csv')
    $badOutput = Join-Path $temp 'bad.csv'
    $rejected = $false
    try { & "$root/powershell/Export-SwpMatchingCsvRows.ps1" -InputDirectory $inputs -Column Status -Equals Active -OutputPath $badOutput } catch { $rejected = $true }
    Assert-True ($rejected -and !(Test-Path -LiteralPath $badOutput)) 'Schema mismatch should not write a partial report.'
    Write-Host 'Five offline PowerShell utility fixtures passed.'
} finally {
    Remove-Item -LiteralPath $temp -Recurse -Force
}
