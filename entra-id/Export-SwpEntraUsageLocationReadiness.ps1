#requires -Version 7.0
<#
.SYNOPSIS
Check whether selected users have a usage location in the expected two-letter form before a licensing review.
.DESCRIPTION
This checks format, not membership in an ISO code list, correctness for the person, or service eligibility.
No usage locations or license assignments are changed.
Runs one Graph GET per unique approved ID and writes a local CSV only after all reads succeed.
No login, consent, Graph writes, or tenant-wide discovery. SDK may retry transient failures.
.PARAMETER TenantId
The intended tenant GUID and a matching Global, delegated, Process-scoped Graph context.
.PARAMETER ObjectId
An explicit list of 1-250 approved user object GUIDs. No UPN lookups or wildcard filters. Duplicates are ignored.
.PARAMETER CsvDirectory
Existing protected filesystem directory. Default: current directory. The report will not replace an existing file.
.EXAMPLE
# Create ./Reports first and replace both fictional GUIDs.
Connect-MgGraph -TenantId '11111111-1111-4111-8111-111111111111' -Scopes 'User.Read.All' -ContextScope Process -Environment Global
./Export-SwpEntraUsageLocationReadiness.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -ObjectId '22222222-2222-4222-8222-222222222222' -CsvDirectory './Reports'
.NOTES
Version 1.0.0 (2026-11-26). Lab validation pending.
Requires PowerShell 7 and Microsoft.Graph.Authentication, with approved delegated User.Read.All.
Read scope approval and the signed-in user's directory access still apply. Use a dedicated session.
Reports can identify directory objects by ID; protect local CSVs and never commit them.
UsageLocation: Two-letter value, normalized to uppercase, or blank if absent or unexpected.
UsageLocationReview: Two-letter format, Missing or unreturned, or Unexpected format.
References: https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0
https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][guid]$TenantId,
    [Parameter(Mandatory)][ValidateCount(1,250)][guid[]]$ObjectId,
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

$requiredScope = 'User.Read.All'
$context = Get-SwpVerifiedContext $TenantId $requiredScope
$account = [string](Get-SwpField $context 'Account')
$clientId = [string](Get-SwpField $context 'ClientId')
$asOf = [DateTimeOffset]::UtcNow
$seen = [System.Collections.Generic.HashSet[guid]]::new()
$rows = [System.Collections.Generic.List[object]]::new()
foreach ($id in $ObjectId) {
    # One GET for each unique approved ID; no tenant enumeration or pagination.
    if (-not $seen.Add($id)) { continue }
    Write-Progress -Activity 'Entra usage location readiness' -Status "Read $($rows.Count) of $($ObjectId.Count) submitted entries"
    $verified = Get-SwpVerifiedContext $TenantId $requiredScope
    if ([string](Get-SwpField $verified 'Account') -ine $account -or
        [string](Get-SwpField $verified 'ClientId') -ine $clientId) {
        throw 'The Graph account or client changed. No partial report was written.'
    }
    $uri = 'https://graph.microsoft.com/v1.0/users/{0}?$select=id,usageLocation' -f $id.ToString()
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
    $code = [string](Get-SwpField $record 'usageLocation')
    $data['UsageLocation'] = ''
    if ([string]::IsNullOrWhiteSpace($code)) { $data['UsageLocationReview'] = 'Missing or unreturned' }
    elseif ($code -cmatch '^[A-Za-z]{2}\z') {
        $data['UsageLocation'] = $code.ToUpperInvariant()
        $data['UsageLocationReview'] = 'Two-letter format'
    }
    else { $data['UsageLocationReview'] = 'Unexpected format' }
    $data['ReportAsOfUtc'] = $asOf.ToString('o')
    foreach ($key in @($data.Keys)) { $data[$key] = Protect-SwpCsvCell $data[$key] }
    $rows.Add([pscustomobject]$data)
}
Write-Progress -Activity 'Entra usage location readiness' -Completed

$stamp = $asOf.ToString('yyyyMMddTHHmmssfffZ', [Globalization.CultureInfo]::InvariantCulture)
$csvPath = Join-Path $directory.FullName "UsageLocationReadiness-${stamp}.csv"
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
