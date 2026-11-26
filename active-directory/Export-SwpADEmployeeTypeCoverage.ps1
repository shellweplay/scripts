#requires -Version 5.1
<#
.SYNOPSIS
    AD employee type coverage.
.DESCRIPTION
    Measure how many user objects have a known, other, or unspecified employee classification against an approved list.
    Makes a paged, read-only query against one explicit organizational unit.
    Writes aggregate counts only; review and protect the CSV before sharing.
    Approved values come from the administrator, not a guessed HR policy.
    The CSV contains neither the allowlist nor user classification values; service accounts in the selected scope are not automatically excluded.
.PARAMETER SearchBase
    Distinguished name of the authorized OU to query. No default domain scope.
.PARAMETER SearchScope
    OneLevel (default) or Subtree (includes nested OUs).
.PARAMETER Server
    Optional approved AD DS domain controller. Otherwise use normal module selection.
.PARAMETER OutputDirectory
    Existing local filesystem folder. Defaults to the current directory.
.PARAMETER ApprovedValues
    Up to 100 approved employee types, matched case-insensitively.
.EXAMPLE
    .\Export-SwpADEmployeeTypeCoverage.ps1 -SearchBase "OU=Lab,DC=example,DC=test" -ApprovedValues 'Employee','Contractor'
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
    [string[]]$ApprovedValues
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
    Properties = @('employeeType')
}
if ($PSBoundParameters.ContainsKey('Server')) {
    $query.Server = $Server
}

# Do not emit a partial CSV if the directory query fails partway through.
$known = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($value in $ApprovedValues) {
    $key = $value.Trim()
    if ($key.Length -eq 0 -or -not $known.Add($key)) {
        throw 'ApprovedValues must be nonblank and distinct after trimming and ignoring case.'
    }
}
$counts = [ordered]@{ Recognized = 0L; Other = 0L; MissingOrNotVisible = 0L }
Get-ADUser @query | ForEach-Object {
    $property = $_.PSObject.Properties['employeeType']
    $type = if ($null -eq $property) { '' } else { ([string]$property.Value).Trim() }
    if ($type.Length -eq 0) { $counts.MissingOrNotVisible++ }
    elseif ($known.Contains($type)) { $counts.Recognized++ }
    else { $counts.Other++ }
}
$rows = foreach ($key in $counts.Keys) {
    [pscustomobject]@{ Bucket = $key; Count = $counts[$key] }
}

# Text from the directory is never written to these reports. The first row is a
# fixed header even if this OU contains no visible objects. CreateNew is atomic
# and never overwrites an existing report. Acquire it before managing cleanup.
$columns = @('Bucket', 'Count')
$header = ($columns | ForEach-Object { '"{0}"' -f $_ }) -join ','
$stamp = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ', [System.Globalization.CultureInfo]::InvariantCulture)
$path = Join-Path -Path $destination.FullName -ChildPath ('Export-SwpADEmployeeTypeCoverage-{0}.csv' -f $stamp)
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
