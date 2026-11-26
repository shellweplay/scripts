#requires -Version 5.1
<#
.SYNOPSIS
    AD account expiration buckets.
.DESCRIPTION
    Bucket visible user objects by configured AD account-expiration time, without exposing identities.
    Makes a paged, read-only query against one explicit organizational unit.
    Writes aggregate counts only; review and protect the CSV before sharing.
    Uses one UTC reference instant; DueWithinWindow includes its final timestamp.
    Account expiration is not password expiration. Unset and unreadable dates cannot always be distinguished.
.PARAMETER SearchBase
    Distinguished name of the authorized OU to query. No default domain scope.
.PARAMETER SearchScope
    OneLevel (default) or Subtree (includes nested OUs).
.PARAMETER Server
    Optional approved AD DS domain controller. Otherwise use normal module selection.
.PARAMETER OutputDirectory
    Existing local filesystem folder. Defaults to the current directory.
.PARAMETER DaysAhead
    Future horizon measured in days (1 to 365).
.EXAMPLE
    .\Export-SwpADAccountExpiryBuckets.ps1 -SearchBase "OU=Lab,DC=example,DC=test" -DaysAhead 30
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

    [ValidateRange(1, 365)]
    [int]$DaysAhead = 30
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
    Properties = @('accountExpires', 'AccountExpirationDate')
}
if ($PSBoundParameters.ContainsKey('Server')) {
    $query.Server = $Server
}

# Do not emit a partial CSV if the directory query fails partway through.
$now = [DateTime]::UtcNow
$end = $now.AddDays($DaysAhead)
$counts = [ordered]@{ Expired = 0L; DueWithinWindow = 0L; Later = 0L; NoExpiryOrNotVisible = 0L; Unknown = 0L }
Get-ADUser @query | ForEach-Object {
    $dateProperty = $_.PSObject.Properties['AccountExpirationDate']
    $rawProperty = $_.PSObject.Properties['accountExpires']
    $date = if ($null -ne $dateProperty) { $dateProperty.Value }
    $raw = if ($null -ne $rawProperty) { $rawProperty.Value }
    if ($date -is [datetime]) {
        $value = $date.ToUniversalTime()
        if ($value -le $now) { $counts.Expired++ }
        elseif ($value -le $end) { $counts.DueWithinWindow++ }
        else { $counts.Later++ }
    }
    elseif ($null -ne $date) { $counts.Unknown++ }
    elseif ($null -eq $raw -or $raw -eq 0 -or $raw -eq [Int64]::MaxValue) {
        $counts.NoExpiryOrNotVisible++
    }
    else { $counts.Unknown++ }
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
$path = Join-Path -Path $destination.FullName -ChildPath ('Export-SwpADAccountExpiryBuckets-{0}.csv' -f $stamp)
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
