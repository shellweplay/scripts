#requires -Version 7.0
<#
.SYNOPSIS
Check the policy-reported expiry of selected Microsoft 365 groups against a configurable review window.
.DESCRIPTION
Default review window: the next 30 elapsed days, inclusive of its upper endpoint; -ReviewWithinDays adjusts it.
Only Unified groups are evaluated. Null dates do not establish policy exemption. A past date is not proof of deletion or actual lifecycle outcome.
Runs one Graph GET per unique approved ID and writes a local CSV only after all reads succeed.
No login, consent, Graph writes, or tenant-wide discovery. SDK may retry transient failures.
.PARAMETER TenantId
The intended tenant GUID and a matching Global, delegated, Process-scoped Graph context.
.PARAMETER ObjectId
An explicit list of 1-250 approved group object GUIDs. No UPN lookups or wildcard filters. Duplicates are ignored.
.PARAMETER CsvDirectory
Existing protected filesystem directory. Default: current directory. The report will not replace an existing file.
.PARAMETER ReviewWithinDays
Whole days ahead from report start; upper endpoint is inclusive.
Default: 30.
.EXAMPLE
# Create ./Reports first and replace both fictional GUIDs.
Connect-MgGraph -TenantId '11111111-1111-4111-8111-111111111111' -Scopes 'Group.Read.All' -ContextScope Process -Environment Global
./Export-SwpEntraGroupExpirationReview.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -ObjectId '22222222-2222-4222-8222-222222222222' -CsvDirectory './Reports'
.NOTES
Version 1.0.0 (2026-11-26). Lab validation pending.
Requires PowerShell 7 and Microsoft.Graph.Authentication, with approved delegated Group.Read.All.
Read scope approval and the signed-in user's directory access still apply. Use a dedicated session.
Reports can identify directory objects by ID; protect local CSVs and never commit them.
GroupKind: Microsoft 365, Not Microsoft 365, or Unknown.
ExpiresUtc: Policy-reported Microsoft 365 group expiry timestamp; blank when unavailable or not applicable.
ExpiryReview: Within window, Beyond window, At or past expiry, Not reported, Not applicable, or Unknown group type.
References: https://learn.microsoft.com/en-us/graph/api/group-get?view=graph-rest-1.0
https://learn.microsoft.com/en-us/graph/api/resources/group?view=graph-rest-1.0
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][guid]$TenantId,
    [Parameter(Mandatory)][ValidateCount(1,250)][guid[]]$ObjectId,
    [ValidateRange(1,3650)][int]$ReviewWithinDays = 30,
    [ValidateNotNullOrEmpty()][string]$CsvDirectory = '.'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Get-SwpField {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return ,($Object[$Name]) }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property) { return ,($property.Value) }
    return $null
}

function Convert-SwpTimestamp {
    param($Value)
    if ($Value -is [datetimeoffset]) { return $Value.ToUniversalTime() }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return $null }
        return ([datetimeoffset]$Value).ToUniversalTime()
    }
    if ($Value -isnot [string] -or $Value -cnotmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})\z') {
        return $null
    }
    $parsed = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$parsed)) { return $parsed }
    return $null
}

function Get-SwpVerifiedContext {
    param([guid]$ExpectedTenant, [string]$Scope)
    $current = Get-MgContext
    if ($null -eq $current -or [string](Get-SwpField $current 'TenantId') -ine $ExpectedTenant.ToString() -or
        (Get-SwpField $current 'AuthType') -ine 'Delegated' -or
        (Get-SwpField $current 'ContextScope') -ine 'Process' -or
        (Get-SwpField $current 'Environment') -ine 'Global' -or
        [string]::IsNullOrWhiteSpace([string](Get-SwpField $current 'Account'))) {
        throw 'Connect to the supplied tenant in a delegated, Global, Process-scoped Graph session first. No query was made with the unverified context.'
    }
    if ($Scope -notin (Get-SwpField $current 'Scopes')) {
        throw "The current context lacks the required delegated scope: $Scope. This report never grants or requests it."
    }
    return $current
}

function Protect-SwpCsvCell {
    param($Value)
    # Neutralize formula-like text without coercing real numeric report values.
    if ($Value -is [string] -and $Value -match '^[\s\p{Cc}\p{Cf}]*[=+@-]') { return "'" + $Value }
    return $Value
}

if ($TenantId -eq [guid]::Empty -or [guid]::Empty -in $ObjectId) {
    throw 'Tenant and object identifiers must be nonzero GUIDs.'
}
$directory = Get-Item -LiteralPath $CsvDirectory -ErrorAction Stop
if ($directory.PSProvider.Name -ne 'FileSystem' -or -not $directory.PSIsContainer) {
    throw 'CsvDirectory must be an existing filesystem directory.'
}
if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue) -or
    -not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
    throw 'Microsoft.Graph.Authentication is required; review and install it separately. This report installs nothing.'
}

$requiredScope = 'Group.Read.All'
$context = Get-SwpVerifiedContext $TenantId $requiredScope
$account = [string](Get-SwpField $context 'Account')
$clientId = [string](Get-SwpField $context 'ClientId')
$asOf = [DateTimeOffset]::UtcNow
$seen = [System.Collections.Generic.HashSet[guid]]::new()
$rows = [System.Collections.Generic.List[object]]::new()
foreach ($id in $ObjectId) {
    # One GET for each unique approved ID; no tenant enumeration or pagination.
    if (-not $seen.Add($id)) { continue }
    Write-Progress -Activity 'Entra Microsoft 365 group expiry review' -Status "Read $($rows.Count) of $($ObjectId.Count) submitted entries"
    $verified = Get-SwpVerifiedContext $TenantId $requiredScope
    if ([string](Get-SwpField $verified 'Account') -ine $account -or
        [string](Get-SwpField $verified 'ClientId') -ine $clientId) {
        throw 'The Graph account or client changed. No partial report was written.'
    }
    $uri = 'https://graph.microsoft.com/v1.0/groups/{0}?$select=id,groupTypes,expirationDateTime' -f $id.ToString()
    try {
        $record = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject -ErrorAction Stop
    }
    catch {
        throw "Graph could not read selected object $id. Verify the approved roster and access, and retry later if throttled. No partial report was written. Graph error: $($_.Exception.Message)"
    }
    if ([string](Get-SwpField $record 'id') -ine $id.ToString()) {
        throw "Graph did not return the expected ID for object $id. No partial report was written."
    }

    $data = [ordered]@{ ObjectId = $id.ToString() }
    $types = Get-SwpField $record 'groupTypes'
    $data['GroupKind'] = 'Unknown'
    $data['ExpiresUtc'] = ''
    $data['ExpiryReview'] = 'Unknown group type'
    if ($types -is [System.Collections.IList]) {
        if ('Unified' -in $types) {
            $data['GroupKind'] = 'Microsoft 365'
            $data['ExpiryReview'] = 'Not reported'
            $expiry = Convert-SwpTimestamp (Get-SwpField $record 'expirationDateTime')
            if ($null -ne $expiry) {
                $data['ExpiresUtc'] = $expiry.ToString('o')
                $data['ExpiryReview'] = if ($expiry -le $asOf) { 'At or past expiry' }
                    elseif ($expiry -le $asOf.AddDays($ReviewWithinDays)) { 'Within window' }
                    else { 'Beyond window' }
            }
        }
        else {
            $data['GroupKind'] = 'Not Microsoft 365'
            $data['ExpiryReview'] = 'Not applicable'
        }
    }
    $data['ReportAsOfUtc'] = $asOf.ToString('o')
    foreach ($key in @($data.Keys)) { $data[$key] = Protect-SwpCsvCell $data[$key] }
    $rows.Add([pscustomobject]$data)
}
Write-Progress -Activity 'Entra Microsoft 365 group expiry review' -Completed

$stamp = $asOf.ToString('yyyyMMddTHHmmssfffZ', [Globalization.CultureInfo]::InvariantCulture)
$csvPath = Join-Path $directory.FullName "GroupExpirationReview-${stamp}.csv"
$stream = $null
$writer = $null
$created = $false
try {
    # CreateNew atomically refuses an existing filename, even with concurrent runs.
    $stream = [IO.File]::Open($csvPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    $created = $true
    $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
    foreach ($line in ($rows.ToArray() | ConvertTo-Csv -NoTypeInformation)) { $writer.WriteLine($line) }
    $writer.Dispose()
}
catch {
    try {
        if ($null -ne $writer) { $writer.Dispose() }
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
        if ($created) { [IO.File]::Delete($csvPath) }
    }
    throw
}
finally {
    if ($null -ne $stream) { $stream.Dispose() }
}
[pscustomobject]@{ Rows = $rows.Count; CsvPath = $csvPath }
