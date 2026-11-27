<#
.SYNOPSIS
Finds group objects whose Notes (info) field contains literal text.
.DESCRIPTION
Uses a server-side substring search under one explicit OU. Writes group
names and timestamps, not the potentially sensitive Notes text. Read only.
Writes a new local CSV; a zero-match run creates no file. Version 1.0.0.
.EXAMPLE
./Find-SwpADGroupsByNote.ps1 -Server dc01.example.test -SearchBase 'OU=Groups,DC=example,DC=test' -Text 'review cycle' -OutputPath ./note-matches.csv
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Server,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $SearchBase,
    [Parameter(Mandatory)] [ValidateLength(3, 128)] [string] $Text,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $OutputPath
)
$ErrorActionPreference = 'Stop'
Import-Module ActiveDirectory
if ($Text.Trim().Length -lt 3) { throw 'Enter at least three non-padding characters.' }
if (Test-Path -LiteralPath $OutputPath) { throw 'Choose a new output file.' }

$escaped = $Text.Trim().Replace('\', '\5c').Replace('*', '\2a').Replace('(', '\28').Replace(')', '\29').Replace([string][char]0, '\00')
$groups = @(Get-ADGroup -LDAPFilter ('(info=*' + $escaped + '*)') -SearchBase $SearchBase -SearchScope Subtree -Server $Server -ResultPageSize 200 -Properties whenCreated, whenChanged)
if (!$groups.Count) { Write-Host 'No visible matches. No CSV written.'; return }
$groups | ForEach-Object {
    $name = [string]$_.SamAccountName
    if ($name -match '^[\s\x00-\x1f]*[=+@-]|^[\t\r\n]') { $name = "'" + $name }
    [pscustomobject]@{
        SamAccountName = $name
        CreatedUtc = if ($null -ne $_.whenCreated) { $_.whenCreated.ToUniversalTime().ToString('o') }
        ChangedUtc = if ($null -ne $_.whenChanged) { $_.whenChanged.ToUniversalTime().ToString('o') }
    }
} | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8 -NoClobber
Write-Host ('Exported {0} group match(es).' -f $groups.Count)
