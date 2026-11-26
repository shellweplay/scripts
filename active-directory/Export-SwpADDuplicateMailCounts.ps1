#requires -Version 5.1
<#
.SYNOPSIS
    AD duplicate mail counts.
.DESCRIPTION
    Count duplicates in the visible mail field without revealing any values or account identities.
    Makes a paged, read-only query against one explicit organizational unit.
    Writes aggregate counts only; review and protect the CSV before sharing.
    Trims surrounding whitespace and compares values case-insensitively; blanks never count as duplicates.
    A duplicate is a review signal, not evidence of an error. Only the selected OU scope is assessed.
.PARAMETER SearchBase
    Distinguished name of the authorized OU to query. No default domain scope.
.PARAMETER SearchScope
    OneLevel (default) or Subtree (includes nested OUs).
.PARAMETER Server
    Optional approved AD DS domain controller. Otherwise use normal module selection.
.PARAMETER OutputDirectory
    Existing local filesystem folder. Defaults to the current directory.
.EXAMPLE
    .\Export-SwpADDuplicateMailCounts.ps1 -SearchBase "OU=Lab,DC=example,DC=test"
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
    Properties = @('mail')
}
if ($PSBoundParameters.ContainsKey('Server')) {
    $query.Server = $Server
}

# Do not emit a partial CSV if the directory query fails partway through.
# Dictionary keys are held in memory only; do not log or export attribute values.
$counts = [System.Collections.Generic.Dictionary[string,long]]::new([System.StringComparer]::OrdinalIgnoreCase)
$missing = 0L
$populated = 0L
Get-ADUser @query | ForEach-Object {
    $property = $_.PSObject.Properties['mail']
    $key = if ($null -eq $property) { '' } else { ([string]$property.Value).Trim() }
    if ($key.Length -eq 0) { $missing++ }
    else {
        $populated++
        if ($counts.ContainsKey($key)) { $counts[$key]++ }
        else { $counts.Add($key, 1L) }
    }
}
$duplicates = 0L
$affected = 0L
foreach ($count in $counts.Values) {
    if ($count -gt 1) {
        $duplicates++
        $affected += $count
    }
}
$rows = [pscustomobject]@{
    PopulatedObjects = $populated
    MissingOrNotVisible = $missing
    DuplicateValues = $duplicates
    ObjectsWithDuplicateValue = $affected
}

# Text from the directory is never written to these reports. The first row is a
# fixed header even if this OU contains no visible objects. CreateNew is atomic
# and never overwrites an existing report. Acquire it before managing cleanup.
$columns = @('PopulatedObjects', 'MissingOrNotVisible', 'DuplicateValues', 'ObjectsWithDuplicateValue')
$header = ($columns | ForEach-Object { '"{0}"' -f $_ }) -join ','
$stamp = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ', [System.Globalization.CultureInfo]::InvariantCulture)
$path = Join-Path -Path $destination.FullName -ChildPath ('Export-SwpADDuplicateMailCounts-{0}.csv' -f $stamp)
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
