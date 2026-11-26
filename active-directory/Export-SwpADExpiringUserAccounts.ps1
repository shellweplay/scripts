#requires -Version 5.1
<#
.SYNOPSIS
Report AD user accounts with account expiration in a future UTC window.
.DESCRIPTION
Reads user objects from one explicit DC and OU/container. Filters raw accountExpires
FILETIME on the server; 0 and Int64.MaxValue mean never expires. Excludes already
expired accounts and does not measure password expiration or actual sign-in ability.
Enabled and disabled users are included unless -EnabledOnly is supplied.
.PARAMETER Server
FQDN of one domain controller (not a domain name or global catalog).
.PARAMETER SearchBase
Distinguished name of an OU/container in that controller's domain.
.PARAMETER DaysAhead
Positive whole days ahead of the frozen UTC report time. Default 30.
.PARAMETER SearchScope
Subtree (default) or OneLevel.
.PARAMETER EnabledOnly
Include only accounts observed as enabled.
.PARAMETER CsvDirectory
Existing protected filesystem directory. Defaults to current directory.
.PARAMETER MaxRecords
Abort rather than export partial data if more than this many candidates return.
.PARAMETER PageSize
LDAP result page size. Default 500.
.EXAMPLE
.\Export-SwpADExpiringUserAccounts.ps1 -Server 'dc01.mylab.dev' -SearchBase 'OU=Lab,DC=mylab,DC=dev' -DaysAhead 30 -CsvDirectory 'C:\Reports'
.NOTES
Version 1.0.0 - 2026-11-26. Lab validation pending.
Windows PowerShell 5.1 with RSAT ActiveDirectory module; current Windows identity
must have read access to users and accountExpires. No directory writes.
Fields: SamAccountName, DisplayName, UserPrincipalName, ObjectGuid, Enabled,
DistinguishedName, AccountExpiresUtc, DaysUntilExpiration (fractional elapsed
days rounded to two decimals), ReportAsOfUtc, WindowEndUtc, Server, SearchBase.
The UTC date names the UTF-8 CSV. Existing files are never overwritten. Empty
results produce a header-only CSV. Missing/invalid values on returned candidates
abort the report. Unreadable values may be excluded by LDAP; not a complete
forest inventory or snapshot. Protect local exports.
References:
https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-aduser
https://learn.microsoft.com/en-us/windows/win32/adschema/a-accountexpires
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Server,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SearchBase,
    [ValidateRange(1,3650)][int]$DaysAhead = 30,
    [ValidateSet('Subtree','OneLevel')][string]$SearchScope = 'Subtree',
    [switch]$EnabledOnly,
    [string]$CsvDirectory = '.',
    [ValidateRange(1,1000000)][int]$MaxRecords = 100000,
    [ValidateRange(1,1000)][int]$PageSize = 500
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function ConvertTo-SwpCsvCell {
    param($Value)
    if ($null -eq $Value) { return '[null]' }
    $text = [string]$Value
    if ($text -match '^[\s\x00-\x1f]*[=+\-@]' -or $text -match '^[\t\r\n]') { return "'" + $text }
    return $text
}

function ConvertFrom-SwpAccountExpires {
    param($Value)
    if ($null -eq $Value) { return $null }
    $raw = 0L
    if (-not [long]::TryParse([string]$Value, [ref]$raw)) { throw 'Invalid accountExpires returned.' }
    if ($raw -eq 0 -or $raw -eq [long]::MaxValue) { return $null }
    if ($raw -lt 0) { throw 'Negative accountExpires returned.' }
    return [datetime]::FromFileTimeUtc($raw)
}

function Write-SwpCsv {
    param([string]$Path, [string[]]$Lines)
    $stream = $null
    $writer = $null
    $created = $false
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $created = $true
        $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
        foreach ($line in $Lines) { $writer.WriteLine($line) }
        $writer.Dispose()
        $writer = $null
        $stream = $null
    } catch {
        $failure = $_
        try { if ($null -ne $writer) { $writer.Dispose() } elseif ($null -ne $stream) { $stream.Dispose() } } catch {}
        if ($created) {
            try { [IO.File]::Delete($Path) }
            catch { throw "CSV export and cleanup failed at '$Path'; remove the partial file explicitly." }
        }
        throw $failure
    }
}

if ($Server -notmatch '^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$') { throw 'Specify a single DC FQDN, for example dc01.mylab.dev.' }
if ([string]::IsNullOrWhiteSpace($SearchBase)) { throw 'SearchBase must not be blank.' }
$directory = Get-Item -LiteralPath $CsvDirectory
if (-not $directory.PSIsContainer -or $directory.PSProvider.Name -ne 'FileSystem') { throw 'CsvDirectory must be an existing filesystem directory.' }
if (-not (Get-Command Get-ADUser -ErrorAction SilentlyContinue)) { Import-Module ActiveDirectory -ErrorAction Stop }
$asOf = [datetime]::UtcNow
$end = $asOf.AddDays($DaysAhead)
# Strictly future: +1 FILETIME tick, inclusive upper bound. Never-expire
# sentinel values fall outside the numeric range. No raw user input in filter.
$lower = ($asOf.ToFileTimeUtc() + 1).ToString([Globalization.CultureInfo]::InvariantCulture)
$upper = $end.ToFileTimeUtc().ToString([Globalization.CultureInfo]::InvariantCulture)
$filter = "(&(accountExpires>=$lower)(accountExpires<=$upper))"
Write-Host "UTC expiration window: after $($asOf.ToString('o')) through $($end.ToString('o'))"
# Fetch one extra candidate to distinguish an exact limit from a truncated set.
$users = @(Get-ADUser -Server $Server -SearchBase $SearchBase -SearchScope $SearchScope -LDAPFilter $filter -Properties accountExpires,displayName -ResultPageSize $PageSize -ResultSetSize ($MaxRecords + 1) -ErrorAction Stop)
if ($users.Count -gt $MaxRecords) { throw 'Candidate limit exceeded; no CSV exported. Narrow the OU or window.' }
$rows = @(foreach ($user in $users) {
    $expires = ConvertFrom-SwpAccountExpires $user.accountExpires
    if ($null -eq $expires) { throw 'A matching candidate returned no finite expiration; rerun and verify read visibility.' }
    # Recheck in case a value changed while the directory query was running.
    if ($expires -le $asOf -or $expires -gt $end) { throw 'A candidate expiration changed outside the report window; rerun for consistent results.' }
    if ($null -eq $user.Enabled -or -not $user.ObjectGuid) { throw 'A candidate lacks required status or identity data.' }
    if ($EnabledOnly -and -not $user.Enabled) { continue }
    [pscustomobject][ordered]@{
        SamAccountName = $user.SamAccountName
        DisplayName = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        ObjectGuid = $user.ObjectGuid
        Enabled = $user.Enabled
        DistinguishedName = $user.DistinguishedName
        AccountExpiresUtc = $expires.ToString('o')
        DaysUntilExpiration = [Math]::Round(($expires - $asOf).TotalDays, 2)
        ReportAsOfUtc = $asOf.ToString('o')
        WindowEndUtc = $end.ToString('o')
        Server = $Server
        SearchBase = $SearchBase
    }
})
$columns = @('SamAccountName','DisplayName','UserPrincipalName','ObjectGuid','Enabled','DistinguishedName','AccountExpiresUtc','DaysUntilExpiration','ReportAsOfUtc','WindowEndUtc','Server','SearchBase')
$lines = @('"' + ($columns -join '","') + '"')
if ($rows.Count -gt 0) {
    $safeRows = foreach ($row in ($rows | Sort-Object AccountExpiresUtc,SamAccountName)) {
        $safe = [ordered]@{}
        foreach ($column in $columns) { $safe[$column] = ConvertTo-SwpCsvCell $row.$column }
        [pscustomobject]$safe
    }
    $lines = @($safeRows | ConvertTo-Csv -NoTypeInformation)
}
$path = Join-Path $directory.FullName ('AD_ExpiringUserAccounts_' + $asOf.ToString('yyyy-MM-dd') + '.csv')
Write-SwpCsv -Path $path -Lines $lines
Write-Host "Candidates: $($users.Count); exported: $($rows.Count); saved: $path"
