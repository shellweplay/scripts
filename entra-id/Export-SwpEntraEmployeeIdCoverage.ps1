#requires -Version 7.0
<#
.SYNOPSIS
Check whether the employeeId field is populated for an approved user roster.
.DESCRIPTION
A populated value does not prove HR accuracy, uniqueness, employment or access eligibility.
The employee ID itself is never exported; use this for a narrow metadata completeness review.
Runs one Graph GET per unique approved ID and writes a local CSV only after all reads succeed.
No login, consent, Graph writes, or tenant-wide discovery. SDK may retry transient failures.
.PARAMETER TenantId
The intended tenant GUID and a matching Global, delegated, Process-scoped Graph context.
.PARAMETER ObjectId
An explicit list of 1-250 approved user object GUIDs. No UPN lookups or wildcard filters. Duplicates are ignored.
.PARAMETER CsvDirectory
Existing protected filesystem directory. Default: current directory. The report will not replace an existing file.
.EXAMPLE
# Fictional mylab.dev tenant and user IDs. Create ./Reports first and replace both GUIDs.
Connect-MgGraph -TenantId '11111111-1111-4111-8111-111111111111' -Scopes 'User.Read.All' -ContextScope Process -Environment Global
./Export-SwpEntraEmployeeIdCoverage.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -ObjectId '22222222-2222-4222-8222-222222222222' -CsvDirectory './Reports'
.NOTES
Version 1.0.0 (2026-11-27): initial release. Lab validation pending.
Requires PowerShell 7 and Microsoft.Graph.Authentication, with approved delegated User.Read.All.
Read scope approval and the signed-in user's directory access still apply. Use a dedicated session.
Reports can identify directory objects by ID; protect local CSVs and never commit them.
ObjectId: Approved user object GUID.
EmployeeIdReview: Present, Missing or unreturned, or Unexpected type.
ReportAsOfUtc: Frozen UTC timestamp for the report.
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
    Write-Progress -Activity 'Entra employee ID coverage' -Status "Read $($rows.Count) of $($ObjectId.Count) submitted entries"
    $verified = Get-SwpVerifiedContext $TenantId $requiredScope
    if ([string](Get-SwpField $verified 'Account') -ine $account -or
        [string](Get-SwpField $verified 'ClientId') -ine $clientId) {
        throw 'The Graph account or client changed. No partial report was written.'
    }
    $uri = 'https://graph.microsoft.com/v1.0/users/{0}?$select=id,employeeId' -f $id.ToString()
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
    # Inspect only presence; never put the underlying employee ID into the CSV.
    $employeeId = Get-SwpField $record 'employeeId'
    $data['EmployeeIdReview'] = 'Missing or unreturned'
    if ($null -ne $employeeId -and $employeeId -isnot [string]) {
        $data['EmployeeIdReview'] = 'Unexpected type'
    }
    elseif (-not [string]::IsNullOrWhiteSpace($employeeId)) {
        $data['EmployeeIdReview'] = 'Present'
    }
    $data['ReportAsOfUtc'] = $asOf.ToString('o')
    foreach ($key in @($data.Keys)) { $data[$key] = Protect-SwpCsvCell $data[$key] }
    $rows.Add([pscustomobject]$data)
}
Write-Progress -Activity 'Entra employee ID coverage' -Completed

$stamp = $asOf.ToString('yyyyMMddTHHmmssfffZ', [Globalization.CultureInfo]::InvariantCulture)
$csvPath = Join-Path $directory.FullName "EmployeeIdCoverage-${stamp}.csv"
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
Write-Host ('Reviewed {0} unique users; saved: {1}' -f $rows.Count, $csvPath)
[pscustomobject]@{ Rows = $rows.Count; CsvPath = $csvPath }
