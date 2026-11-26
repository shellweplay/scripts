#requires -Version 5.1
<#
.SYNOPSIS
    AD account state totals.
.DESCRIPTION
    Summarize enabled, disabled, and unknown states for visible user objects in an OU.
    Makes a paged, read-only query against one explicit organizational unit.
    Writes aggregate counts only; review and protect the CSV before sharing.
    Enabled status is a configuration flag, not proof of current access or recent activity.
    Unknown means the state was not visible; no account identities are exported.
.PARAMETER SearchBase
    Distinguished name of the authorized OU to query. No default domain scope.
.PARAMETER SearchScope
    OneLevel (default) or Subtree (includes nested OUs).
.PARAMETER Server
    Optional approved AD DS domain controller. Otherwise use normal module selection.
.PARAMETER OutputDirectory
    Existing local filesystem folder. Defaults to the current directory.
.EXAMPLE
    .\Export-SwpADAccountStateTotals.ps1 -SearchBase "OU=Lab,DC=example,DC=test"
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
}
if ($PSBoundParameters.ContainsKey('Server')) {
    $query.Server = $Server
}

# Do not emit a partial CSV if the directory query fails partway through.
$counts = @{ Enabled = 0L; Disabled = 0L; Unknown = 0L }
Get-ADUser @query | ForEach-Object {
    $property = $_.PSObject.Properties['Enabled']
    if ($null -eq $property -or $property.Value -isnot [bool]) { $counts.Unknown++ }
    elseif ($property.Value) { $counts.Enabled++ }
    else { $counts.Disabled++ }
}
$rows = foreach ($state in @('Enabled', 'Disabled', 'Unknown')) {
    [pscustomobject]@{ State = $state; Count = $counts[$state] }
}

# Text from the directory is never written to these reports. The first row is a
# fixed header even if this OU contains no visible objects. CreateNew is atomic
# and never overwrites an existing report. Acquire it before managing cleanup.
$columns = @('State', 'Count')
$header = ($columns | ForEach-Object { '"{0}"' -f $_ }) -join ','
$stamp = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ', [System.Globalization.CultureInfo]::InvariantCulture)
$path = Join-Path -Path $destination.FullName -ChildPath ('Export-SwpADAccountStateTotals-{0}.csv' -f $stamp)
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
