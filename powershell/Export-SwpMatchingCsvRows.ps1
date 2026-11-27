<#
.SYNOPSIS
Combines rows from same-schema CSV files where a named field has an exact value.
.DESCRIPTION
Reads only CSVs in one selected folder, without recursion. Column order can
differ between files; the set of column names must match. SourceFile is reserved.
Matching is case-insensitive by default. Local only, and creates a new file
only when rows match. Version 1.0.0.
.EXAMPLE
./Export-SwpMatchingCsvRows.ps1 -InputDirectory ./samples -Column Status -Equals Active -OutputPath ./active-rows.csv
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $InputDirectory,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Column,
    [Parameter(Mandatory)] [AllowEmptyString()] [string] $Equals,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $OutputPath,
    [switch] $CaseSensitive
)
$ErrorActionPreference = 'Stop'
function Protect-Cell([string] $Value) {
    if ($Value -match '^[\s\x00-\x1f]*[=+@-]|^[\t\r\n]') { return "'" + $Value }
    return $Value
}
if (!(Test-Path -LiteralPath $InputDirectory -PathType Container)) { throw 'Choose an existing input directory.' }
if (Test-Path -LiteralPath $OutputPath) { throw 'Choose a new output file.' }
$files = @(Get-ChildItem -LiteralPath $InputDirectory -Filter '*.csv' -File | Sort-Object Name)
if (!$files.Count) { throw 'No CSV input files were found.' }
$schema = $null
$matchedRows = @(foreach ($file in $files) {
    $rows = @(Import-Csv -LiteralPath $file.FullName)
    # Header-only CSVs have no data to contribute and are skipped.
    if (!$rows.Count) { continue }
    $fields = @($rows[0].PSObject.Properties.Name)
    if (@($fields | Where-Object { $_ -notmatch '^[\p{L}_]' }).Count) {
        throw 'Input column names must start with a letter or underscore.'
    }
    if ($fields -notcontains $Column -or $fields -contains 'SourceFile') {
        throw 'Every non-empty input CSV must include the selected column and must not use the reserved SourceFile heading.'
    }
    if ($null -eq $schema) { $schema = $fields }
    if ($schema.Count -ne $fields.Count -or @($fields | Where-Object { $schema -notcontains $_ }).Count) {
        throw 'Input CSVs with data must have the same column names.'
    }
    foreach ($item in $rows) {
        $value = [string]$item.PSObject.Properties[$Column].Value
        $equal = if ($CaseSensitive) { $value -ceq $Equals } else { $value -ieq $Equals }
        if ($equal) {
            $record = [ordered]@{}
            foreach ($field in $schema) { $record[$field] = Protect-Cell $item.PSObject.Properties[$field].Value }
            $record['SourceFile'] = Protect-Cell $file.Name
            [pscustomobject]$record
        }
    }
})
if (!$matchedRows.Count) { Write-Host 'No matching rows. No CSV written.'; return }
$matchedRows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8 -NoClobber
Write-Host ('Exported {0} matching row(s).' -f $matchedRows.Count)
