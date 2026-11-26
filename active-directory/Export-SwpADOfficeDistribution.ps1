#requires -Version 5.1
<#
.SYNOPSIS
    AD office distribution.
.DESCRIPTION
    Count visible users mapped to administrator-supplied office labels; all other values stay aggregated.
    Makes a paged, read-only query against one explicit organizational unit.
    Writes aggregate counts only; review and protect the CSV before sharing.
    Pass approved office labels only; values are trimmed and matched without regard to case.
    Configured labels with a formula marker are prefixed with an apostrophe for spreadsheet safety. No unlisted text or account identities are exported.
.PARAMETER SearchBase
    Distinguished name of the authorized OU to query. No default domain scope.
.PARAMETER SearchScope
    OneLevel (default) or Subtree (includes nested OUs).
.PARAMETER Server
    Optional approved AD DS domain controller. Otherwise use normal module selection.
.PARAMETER OutputDirectory
    Existing local filesystem folder. Defaults to the current directory.
.PARAMETER Labels
    Up to 100 approved office labels to count separately.
.EXAMPLE
    .\Export-SwpADOfficeDistribution.ps1 -SearchBase "OU=Lab,DC=example,DC=test" -Labels 'Lab North','Lab South'
.NOTES
    ShellWePlay version 1.0.0. Offline fixture checks do not replace live lab validation.
.LINK
    https://learn.microsoft.com/powershell/module/activedirectory/get-aduser
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('(?i)^(OU=).+')]
    [string]$SearchBase,

    [ValidateSet('OneLevel', 'Subtree')]
    [string]$SearchScope = 'OneLevel',

    [ValidateNotNullOrEmpty()]
    [string]$Server,

    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory = '.'
,

    [Parameter(Mandatory)]
    [ValidateCount(1, 100)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Labels
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Resolve a literal directory before querying; no folders are created automatically.
$destination = Get-Item -LiteralPath $OutputDirectory
if ($destination -isnot [System.IO.DirectoryInfo]) {
    throw 'OutputDirectory must be an existing filesystem directory.'
}

Import-Module ActiveDirectory -ErrorAction Stop
$query = @{
    Filter = '*'
    SearchBase = $SearchBase
    SearchScope = $SearchScope
    ResultPageSize = 500
    ErrorAction = 'Stop'
    Properties = @('Office')
}
if ($PSBoundParameters.ContainsKey('Server')) {
    $query.Server = $Server
}

# Do not emit a partial CSV if the directory query fails partway through.
# Compare trimmed values case-insensitively; never export unrecognized AD values.
$counts = [System.Collections.Generic.Dictionary[string,long]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($label in $Labels) {
    $key = $label.Trim()
    if ($key.Length -eq 0 -or $key -match '[\p{Cc}\p{Cf}\p{Zl}\p{Zp}]') {
        throw 'Labels must be nonblank and cannot contain control characters.'
    }
    if ($counts.ContainsKey($key)) {
        throw 'Labels must be distinct after trimming and ignoring case.'
    }
    $counts.Add($key, 0L)
}
$missing = 0L
$other = 0L
Get-ADUser @query | ForEach-Object {
    $property = $_.PSObject.Properties['Office']
    $key = if ($null -eq $property) { '' } else { ([string]$property.Value).Trim() }
    if ($key.Length -eq 0) { $missing++ }
    elseif ($counts.ContainsKey($key)) { $counts[$key]++ }
    else { $other++ }
}
$rows = foreach ($label in $Labels) {
    $key = $label.Trim()
    $csvLabel = if ($key -match '^[=+@-]') { "'" + $key } else { $key }
    [pscustomobject]@{ Bucket = 'Configured'; Label = $csvLabel; Count = $counts[$key] }
}
$rows += @(
    [pscustomobject]@{ Bucket = 'Other'; Label = ''; Count = $other }
    [pscustomobject]@{ Bucket = 'MissingOrNotVisible'; Label = ''; Count = $missing }
)

# Text from the directory is never written to these reports. The first row is a
# fixed header even if this OU contains no visible objects. CreateNew is atomic
# and never overwrites an existing report. Acquire it before managing cleanup.
$columns = @('Bucket', 'Label', 'Count')
$header = ($columns | ForEach-Object { '"{0}"' -f $_ }) -join ','
$stamp = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ', [System.Globalization.CultureInfo]::InvariantCulture)
$path = Join-Path -Path $destination.FullName -ChildPath ('Export-SwpADOfficeDistribution-{0}.csv' -f $stamp)
$file = [System.IO.File]::Open($path, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
try {
    try {
        $writer = [System.IO.StreamWriter]::new($file, [System.Text.UTF8Encoding]::new($true))
        try {
            $writer.WriteLine($header)
            $rows | Select-Object -Property $columns | ConvertTo-Csv -NoTypeInformation |
                Select-Object -Skip 1 | ForEach-Object { $writer.WriteLine($_) }
        }
        finally {
            $writer.Dispose()
        }
    }
    finally {
        $file.Dispose()
    }
}
catch {
    try {
        [System.IO.File]::Delete($path)
    }
    catch {
        Write-Warning 'Could not remove the incomplete local report.'
    }
    throw
}
Get-Item -LiteralPath $path
