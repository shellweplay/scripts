#requires -Version 7.0
[CmdletBinding()]
param([string]$SourceDirectory = (Join-Path $PSScriptRoot '../entra-id'))
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$tenant = [guid]'11111111-1111-4111-8111-111111111111'
$id = [guid]'22222222-2222-4222-8222-222222222222'
$id2 = [guid]'33333333-3333-4333-8333-333333333333'
$out = Join-Path ([IO.Path]::GetTempPath()) ("swp-entra-mock-" + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$global:SwpBatchTestContext = [ordered]@{
    TenantId = $tenant.ToString(); AuthType = 'Delegated'; ContextScope = 'Process'; Environment = 'Global'
    Scopes = @('User.Read.All', 'User.ReadBasic.All', 'GroupMember.Read.All', 'Group.Read.All'); Account = 'mock@example.invalid'; ClientId = $tenant.ToString()
}
$global:SwpBatchTestRequests = [System.Collections.Generic.List[string]]::new()
$global:SwpBatchTestFixture = @{}
$global:SwpBatchTestFailId = ''
$global:SwpBatchTestWrongId = ''
function Get-MgContext { return $global:SwpBatchTestContext }
function Invoke-MgGraphRequest {
    [CmdletBinding()]param([string]$Method, [string]$Uri, [string]$OutputType)
    if ($Method -cne 'GET' -or $OutputType -cne 'PSObject' -or $Uri -cnotmatch '^https://graph\.microsoft\.com/v1\.0/(users|groups)/([0-9a-f-]{36})\?\$select=([a-zA-Z,]+)\z') {
        throw 'Unexpected Graph request from mock.'
    }
    $requestId = $Matches[2]
    $selected = $Matches[3].Split(',')
    $global:SwpBatchTestRequests.Add($Uri)
    if ($requestId -eq $global:SwpBatchTestFailId) { throw 'Simulated provider failure' }
    $response = [ordered]@{ id = $requestId }
    if ($global:SwpBatchTestWrongId) { $response['id'] = $global:SwpBatchTestWrongId }
    foreach ($key in $global:SwpBatchTestFixture.Keys) {
        if ($key -cnotin $selected) { throw "Missing Graph property in request: $key" }
        $response[$key] = $global:SwpBatchTestFixture[$key]
    }
    return [pscustomobject]$response
}
function Assert-SwpEqual {
    param($Actual, $Expected, [string]$Why)
    if ([string]$Actual -cne [string]$Expected) { throw "$Why expected <$Expected>, got <$Actual>." }
}
function Invoke-SwpCase {
    param([string]$Report, [string]$Resource, [hashtable]$Fields, [hashtable]$Expected, [guid[]]$Ids = @($id, $id))
    $global:SwpBatchTestFixture = $Fields
    $global:SwpBatchTestRequests.Clear()
    $path = Join-Path $SourceDirectory "Export-SwpEntra$Report.ps1"
    $summary = & $path -TenantId $tenant -ObjectId $Ids -CsvDirectory $out
    $rows = @(Import-Csv -LiteralPath $summary.CsvPath)
    Assert-SwpEqual $rows.Count 1 "$Report cases"
    Assert-SwpEqual $summary.Rows 1 "$Report summary"
    Assert-SwpEqual $global:SwpBatchTestRequests.Count 1 "$Report should deduplicate identical input IDs"
    if (-not $global:SwpBatchTestRequests[0].Contains("/v1.0/$Resource/")) { throw "Wrong resource in $Report" }
    Assert-SwpEqual $rows[0].ObjectId $id.ToString() "$Report object ID"
    foreach ($key in $Expected.Keys) { Assert-SwpEqual ($rows[0].$key) $Expected[$key] "$Report / $key" }
    if ($rows[0].ReportAsOfUtc -notmatch '\+00:00$') { throw 'Expected UTC report timestamp.' }
    Remove-Item -LiteralPath $summary.CsvPath
}
function Assert-SwpRejected {
    param([scriptblock]$Operation, [string]$ErrorPattern, [int]$ExpectedCalls)
    $global:SwpBatchTestRequests.Clear()
    try { & $Operation; throw 'Operation was not rejected.' }
    catch {
        if ($_.Exception.Message -notlike $ErrorPattern) { throw }
    }
    Assert-SwpEqual $global:SwpBatchTestRequests.Count $ExpectedCalls 'Failure path Graph call count'
    Assert-SwpEqual @(Get-ChildItem -LiteralPath $out).Count 0 'Failure path must not leave a CSV'
}
try {
    $scripts = @(Get-ChildItem -LiteralPath $SourceDirectory -Filter 'Export-SwpEntra*.ps1' | Where-Object Name -notin @(
        'Export-SwpEntraRecentAppRegistrations.ps1', 'Export-SwpEntraPendingGuestInvitations.ps1'))
    Assert-SwpEqual $scripts.Count 10 'Number of new standalone script files'
    foreach ($file in $scripts) {
        $tokens = $null; $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
        if ($errors.Count -gt 0) { throw "PowerShell parser error: $($file.Name) / $($errors[0])" }
    }

    Invoke-SwpCase RosterAccountState users @{ accountEnabled = $false; userType = 'Member' } @{ AccountState = 'Disabled'; UserType = 'Member' }
    Invoke-SwpCase RosterAccountState users @{ accountEnabled = 'false'; userType = 'Other' } @{ AccountState = 'Unknown'; UserType = 'Unknown' }
    Invoke-SwpCase OrganizationProfileGaps users @{ department = 'Sales'; jobTitle = '  ' } @{
        MissingFields = 'jobTitle;officeLocation'; MissingCount = 2 }
    Invoke-SwpCase PrimaryMailReview users @{ mail = 'alternate@example.invalid'; userPrincipalName = 'signin@example.invalid' } @{
        MailReview = 'Different values' }
    Invoke-SwpCase UsageLocationReadiness users @{ usageLocation = 'us' } @{ UsageLocation = 'US'; UsageLocationReview = 'Two-letter format' }
    Invoke-SwpCase UsageLocationReadiness users @{ usageLocation = 'USA' } @{ UsageLocation = ''; UsageLocationReview = 'Unexpected format' }
    Invoke-SwpCase LicenseAssignmentPresence users @{ assignedLicenses = @(@{ skuId = $id2.ToString() }, @{ skuId = $id2.ToString() }) } @{
        AssignedSkuCount = '1'; LicenseReview = 'Assigned' }
    Invoke-SwpCase LicenseAssignmentPresence users @{ assignedLicenses = @() } @{ AssignedSkuCount = '0'; LicenseReview = 'None returned' }
    Invoke-SwpCase LicenseAssignmentPresence users @{ assignedLicenses = @(@{ skuId = 'invalid' }) } @{
        AssignedSkuCount = ''; LicenseReview = 'Unknown' }
    Invoke-SwpCase HybridSyncFreshness users @{ onPremisesSyncEnabled = $true; onPremisesLastSyncDateTime = '2020-01-01T02:00:00+02:00' } @{
        SyncState = 'Currently synced'; LastSyncUtc = '2020-01-01T00:00:00.0000000+00:00'; SyncReview = 'Review age' }
    Invoke-SwpCase HybridSyncFreshness users @{ onPremisesSyncEnabled = $true; onPremisesLastSyncDateTime = 'not-a-date' } @{
        LastSyncUtc = ''; ElapsedDays = ''; SyncReview = 'Timestamp unavailable' }
    Invoke-SwpCase PreferredLanguageReadiness users @{ preferredLanguage = [string]([char]9) + '=2+2' } @{
        PreferredLanguage = "'" + [char]9 + '=2+2'; LanguageReview = 'Present' }
    Invoke-SwpCase GroupDocumentationGaps groups @{ displayName = 'Example'; description = '' } @{
        NameField = 'Present'; DescriptionField = 'Blank or unreturned' }
    Invoke-SwpCase GroupMailReadiness groups @{ mailEnabled = $true; mail = ' ' } @{ MailState = 'Mail-enabled'; MailReview = 'Review mail field' }
    Invoke-SwpCase GroupExpirationReview groups @{ groupTypes = @('Unified'); expirationDateTime = [datetimeoffset]::UtcNow.AddDays(5).ToString('o') } @{
        GroupKind = 'Microsoft 365'; ExpiryReview = 'Within window' }
    Invoke-SwpCase GroupExpirationReview groups @{ groupTypes = @('Unified') } @{
        GroupKind = 'Microsoft 365'; ExpiresUtc = ''; ExpiryReview = 'Not reported' }
    Invoke-SwpCase GroupExpirationReview groups @{ groupTypes = @() } @{ GroupKind = 'Not Microsoft 365'; ExpiryReview = 'Not applicable' }
    Invoke-SwpCase GroupExpirationReview groups @{} @{ GroupKind = 'Unknown'; ExpiryReview = 'Unknown group type' }

    $roster = Join-Path $SourceDirectory 'Export-SwpEntraRosterAccountState.ps1'
    $global:SwpBatchTestFixture = @{ accountEnabled = $true; userType = 'Member' }
    Assert-SwpRejected { & $roster -TenantId $id2 -ObjectId $id -CsvDirectory $out } '*unverified context*' 0
    Assert-SwpRejected { & $roster -TenantId $tenant -ObjectId ([guid]::Empty) -CsvDirectory $out } '*nonzero GUIDs*' 0
    $global:SwpBatchTestContext['AuthType'] = 'AppOnly'
    Assert-SwpRejected { & $roster -TenantId $tenant -ObjectId $id -CsvDirectory $out } '*unverified context*' 0
    $global:SwpBatchTestContext['AuthType'] = 'Delegated'
    $global:SwpBatchTestContext['Scopes'] = @('Group.Read.All')
    Assert-SwpRejected { & $roster -TenantId $tenant -ObjectId $id -CsvDirectory $out } '*lacks the required delegated scope*' 0
    $global:SwpBatchTestContext['Scopes'] = @('User.Read.All', 'User.ReadBasic.All', 'GroupMember.Read.All', 'Group.Read.All')
    $global:SwpBatchTestFailId = $id2.ToString()
    Assert-SwpRejected { & $roster -TenantId $tenant -ObjectId @($id, $id2) -CsvDirectory $out } '*No partial report was written*' 2
    $global:SwpBatchTestFailId = ''
    $global:SwpBatchTestWrongId = $id2.ToString()
    Assert-SwpRejected { & $roster -TenantId $tenant -ObjectId $id -CsvDirectory $out } '*did not return the expected ID*' 1

    Write-Host 'Entra identity batch: 10 parser checks, 18 report cases, and 6 guarded-failure cases passed; no network.'
}
finally {
    Remove-Item -LiteralPath $out -Recurse -Force
}
