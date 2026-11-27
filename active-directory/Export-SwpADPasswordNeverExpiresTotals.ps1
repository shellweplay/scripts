#requires -Version 5.1
<#
.SYNOPSIS
    AD password-never-expires flag totals.
.DESCRIPTION
    Count the PasswordNeverExpires setting for visible user objects, including disabled accounts.
    Makes a paged, read-only query against one explicit organizational unit.
    Writes aggregate counts only; review and protect the CSV before sharing.
    The flag does not establish effective password age, policy compliance, or actual account expiry.
    Unknown means the property was missing or not Boolean; no account identities are exported.
    CSV: State (NeverExpires, FlagNotSet, Unknown), Count (visible user objects).
    Requires RSAT ActiveDirectory and read access to the chosen OU and account flags.
.PARAMETER SearchBase
    Distinguished name of the authorized OU to query. No default domain scope.
.PARAMETER SearchScope
    OneLevel (default) or Subtree (includes nested OUs).
.PARAMETER Server
    Optional approved AD DS domain controller. Otherwise use normal module selection.
.PARAMETER OutputDirectory
    Existing local filesystem folder. Defaults to the current directory.
.EXAMPLE
    .\Export-SwpADPasswordNeverExpiresTotals.ps1 -SearchBase "OU=Lab,DC=mylab,DC=dev"
.NOTES
    ShellWePlay version 1.0.0 (2026-11-27): initial release. Offline fixture checks do not replace live lab validation.
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
    Properties = @('PasswordNeverExpires')
    ErrorAction = 'Stop'
}
if ($PSBoundParameters.ContainsKey('Server')) {
    $query.Server = $Server
}

# Do not emit a partial CSV if the directory query fails partway through.
$counts = @{ NeverExpires = 0L; FlagNotSet = 0L; Unknown = 0L }
Get-ADUser @query | ForEach-Object {
    $property = $_.PSObject.Properties['PasswordNeverExpires']
    if ($null -eq $property -or $property.Value -isnot [bool]) { $counts.Unknown++ }
    elseif ($property.Value) { $counts.NeverExpires++ }
    else { $counts.FlagNotSet++ }
}
$rows = foreach ($state in @('NeverExpires', 'FlagNotSet', 'Unknown')) {
    [pscustomobject]@{ State = $state; Count = $counts[$state] }
}

# Text from the directory is never written to these reports. The first row is a
# fixed header even if this OU contains no visible objects. CreateNew is atomic
# and never overwrites an existing report. Acquire it before managing cleanup.
$columns = @('State', 'Count')
$header = ($columns | ForEach-Object { '"{0}"' -f $_ }) -join ','
$stamp = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ', [System.Globalization.CultureInfo]::InvariantCulture)
$path = Join-Path -Path $destination.FullName -ChildPath ('Export-SwpADPasswordNeverExpiresTotals-{0}.csv' -f $stamp)
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
Write-Host ('Visible users: {0}; never-expires flag: {1}; unknown: {2}; saved: {3}' -f ($counts.NeverExpires + $counts.FlagNotSet + $counts.Unknown), $counts.NeverExpires, $counts.Unknown, $path)
Get-Item -LiteralPath $path
