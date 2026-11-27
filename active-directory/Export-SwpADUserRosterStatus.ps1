<#
.SYNOPSIS
Reports AD lookup and account state for a supplied list of sAMAccountNames.
.DESCRIPTION
Queries one explicit domain controller and OU with bounded LDAP batches.
Writes a new local CSV. NotFound means no visible match in the selected OU.
A failed batch is LookupFailed, never NotFound. Treat output as sensitive.
Version 1.0.0.
.EXAMPLE
./Export-SwpADUserRosterStatus.ps1 -Server dc01.example.test -SearchBase 'OU=People,DC=example,DC=test' -InputPath ./accounts.txt -OutputPath ./roster.csv
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Server,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $SearchBase,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $InputPath,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $OutputPath,
    [ValidateRange(1, 500)] [int] $BatchSize = 100
)
$ErrorActionPreference = 'Stop'
Import-Module ActiveDirectory
function Protect-Cell([string] $Value) {
    if ($Value -match '^[\s\x00-\x1f]*[=+@-]|^[\t\r\n]') { return "'" + $Value }
    return $Value
}

$names = @(Get-Content -LiteralPath $InputPath | ForEach-Object { $_.Trim() } |
    Where-Object { $_ } | Select-Object -Unique)
if (!$names.Count) { throw 'Provide at least one sAMAccountName, one per line.' }
if (Test-Path -LiteralPath $OutputPath) { throw 'Choose a new output file.' }
$results = for ($offset = 0; $offset -lt $names.Count; $offset += $BatchSize) {
    $batch = @($names[$offset..([Math]::Min($offset + $BatchSize, $names.Count) - 1)])
    $parts = foreach ($name in $batch) {
        $escaped = $name.Replace('\', '\5c').Replace('*', '\2a').Replace('(', '\28').Replace(')', '\29').Replace([string][char]0, '\00')
        '(sAMAccountName=' + $escaped + ')'
    }
    $found = @{}
    $failed = $false
    try {
        $filter = '(|' + ($parts -join '') + ')'
        $users = @(Get-ADUser -LDAPFilter $filter -Server $Server -SearchBase $SearchBase -SearchScope Subtree -ResultPageSize 200 -Properties whenCreated, whenChanged)
        foreach ($user in $users) { $found[$user.SamAccountName] = $user }
    } catch {
        $failed = $true
        Write-Warning ('AD lookup failed for a batch of {0} account(s). Check connectivity and permissions.' -f $batch.Count)
    }
    foreach ($name in $batch) {
        $user = $found[$name]
        $status = if ($failed) { 'LookupFailed' } elseif (!$user) { 'NotFound' } elseif ($null -eq $user.Enabled) { 'Unknown' } elseif ($user.Enabled) { 'Enabled' } else { 'Disabled' }
        [pscustomobject]@{
            SamAccountName = Protect-Cell $name
            Status = $status
            CreatedUtc = if (!$failed -and $null -ne $user.whenCreated) { $user.whenCreated.ToUniversalTime().ToString('o') }
            ChangedUtc = if (!$failed -and $null -ne $user.whenChanged) { $user.whenChanged.ToUniversalTime().ToString('o') }
        }
    }
}
$results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8 -NoClobber
Write-Host ('Exported {0} roster result(s).' -f @($results).Count)
