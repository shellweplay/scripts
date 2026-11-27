<#
.SYNOPSIS
Exports standard group attributes under explicitly selected OUs.
.DESCRIPTION
Uses one specified DC. Overlapping search bases are deduplicated by GUID.
Uses only standard AD attributes. Add -IncludeNotes only if the local
report is approved to contain Notes text. No directory writes. A query
failure aborts before export; a zero-result run creates no file.
Version 1.0.0.
.EXAMPLE
./Export-SwpADGroupAttributes.ps1 -Server dc01.example.test -SearchBase 'OU=Groups,DC=example,DC=test' -OutputPath ./group-attributes.csv
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Server,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string[]] $SearchBase,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $OutputPath,
    [switch] $IncludeNotes
)
$ErrorActionPreference = 'Stop'
Import-Module ActiveDirectory
function Protect-Cell([string] $Value) {
    if ($Value -match '^[\s\x00-\x1f]*[=+@-]|^[\t\r\n]') { return "'" + $Value }
    return $Value
}
if (Test-Path -LiteralPath $OutputPath) { throw 'Choose a new output file.' }
$fields = @('whenCreated', 'whenChanged', 'description', 'mail')
if ($IncludeNotes) { $fields += 'info' }
$seen = [System.Collections.Generic.HashSet[guid]]::new()
$rows = @(foreach ($base in ($SearchBase | Select-Object -Unique)) {
    if ([string]::IsNullOrWhiteSpace($base)) { throw 'Search bases cannot be blank.' }
    Get-ADGroup -Filter * -SearchBase $base -SearchScope Subtree -Server $Server -ResultPageSize 200 -Properties $fields | ForEach-Object {
        if (!$_.ObjectGUID) { throw 'A returned group is missing its object GUID.' }
        if ($seen.Add($_.ObjectGUID)) {
            $row = [ordered]@{
                SamAccountName = Protect-Cell $_.SamAccountName
                CreatedUtc = if ($null -ne $_.whenCreated) { $_.whenCreated.ToUniversalTime().ToString('o') }
                ChangedUtc = if ($null -ne $_.whenChanged) { $_.whenChanged.ToUniversalTime().ToString('o') }
                Description = Protect-Cell $_.Description
                Mail = Protect-Cell $_.Mail
            }
            if ($IncludeNotes) { $row['Notes'] = Protect-Cell $_.Info }
            [pscustomobject]$row
        }
    }
})
if (!$rows.Count) { Write-Host 'No visible groups. No CSV written.'; return }
$rows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8 -NoClobber
Write-Host ('Exported {0} unique group(s).' -f $rows.Count)
