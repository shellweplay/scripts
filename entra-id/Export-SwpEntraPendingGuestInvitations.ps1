#requires -Version 7.0
<#
.SYNOPSIS
Export current Entra Guest accounts whose invitation state is PendingAcceptance.
.DESCRIPTION
Reads Graph v1.0 users filtered to Guest, selects nondefault properties explicitly,
and checks current externalUserState locally. No sign-in or audit-log reads.
CreatedDateTimeUtc is account creation; InvitationStateChangedUtc is the latest
externalUserState transition, not the invitation-email-sent time. PendingAcceptance
does not prove that the person never previously accepted. Null/invalid dates
and future ages remain unavailable, not zero. Whole elapsed days use a frozen UTC
timestamp. Includes name, UPN, mail, ID, enabled status, state, both dates/ages,
and ReportAsOfUtc. User.Read.All delegated is needed for the selected full profile.
.PARAMETER MinimumAccountAgeDays
Optional minimum whole days since ACCOUNT CREATION, not invitation age. Defaults
to zero (all pending guests including missing creation dates). Positive thresholds
exclude unavailable ages and report their count.
.EXAMPLE
# Fictional mylab.dev tenant; replace GUID and sign in with an authorized account.
Connect-MgGraph -TenantId '11111111-1111-4111-8111-111111111111' -Scopes 'User.Read.All' -ContextScope Process
./Export-SwpEntraPendingGuestInvitations.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -CsvDirectory './Reports' -MinimumAccountAgeDays 30

.PARAMETER TenantId
Required intended tenant GUID. Global cloud, delegated, Process-scoped context.
.PARAMETER CsvDirectory
Existing protected filesystem directory; defaults to the current directory.
Dated UTF-8 CSV, atomic no-overwrite open; [null] indicates unavailable values.
Formula-like cells are neutralized. Zero matches produce a header-only CSV.
.PARAMETER MaxPages
Maximum pages per collection (default 10000). A collection exceeding the bound fails.
.PARAMETER MaxRecords
Maximum records per collection (default 1000000). Review memory/query load.
.PARAMETER MaxRetries
Maximum additional attempts for HTTP 429/503 (default 3); honors Retry-After
up to 60 seconds, otherwise stops. The Graph SDK can also retry internally.
.NOTES
Version 1.0.0 - 2026-11-25. Lab validation pending.
PowerShell 7 and Microsoft.Graph.Authentication (install separately).
Use approved permissions and a dedicated session; no module install, consent,
authentication, or directory write is performed by this report. No snapshot
isolation; objects may change during paging. Protect local directory data.
References: https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0
https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0
https://learn.microsoft.com/en-us/graph/permissions-reference
https://learn.microsoft.com/en-us/graph/paging
#>
[CmdletBinding()]
param(
    [ValidateRange(0,365000)][int]$MinimumAccountAgeDays = 0,
    [Parameter(Mandatory)][guid]$TenantId,
    [string]$CsvDirectory = '.',
    [ValidateRange(1,100000)][int]$MaxPages = 10000,
    [ValidateRange(1,10000000)][int]$MaxRecords = 1000000,
    [ValidateRange(0,5)][int]$MaxRetries = 3
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
function Get-SwpValue {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return ,($Object[$Name]) }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property) { return ,($property.Value) }
    return $null
}
function Assert-SwpContext {
    param([string[]]$RequiredScopes)
    if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) {
        throw 'Microsoft.Graph.Authentication is required. Review and install it separately; this script never installs modules.'
    }
    $context = Get-MgContext
    if ($null -eq $context) { throw 'No Graph context. Connect explicitly with the scopes documented in this script help.' }
    if ([string](Get-SwpValue $context 'TenantId') -ine $TenantId.ToString()) { throw 'Wrong tenant context. No Graph request was made. Disconnect explicitly and reconnect to the intended tenant.' }
    if ([string](Get-SwpValue $context 'AuthType') -ine 'Delegated' -or [string]::IsNullOrWhiteSpace([string](Get-SwpValue $context 'Account'))) { throw 'A delegated signed-in account is required; application-only access is outside this script.' }
    if ([string](Get-SwpValue $context 'Environment') -ine 'Global') { throw 'Only the Global Microsoft Graph cloud is supported by this artifact.' }
    if ([string](Get-SwpValue $context 'ContextScope') -ine 'Process') { throw 'Use a Process-scoped Graph connection; persistent CurrentUser context is not accepted.' }
    $actualScopes = Get-SwpValue $context 'Scopes'
    foreach ($scope in $RequiredScopes) { if ($scope -notin $actualScopes) { throw "Missing delegated scope $scope. Reconnect explicitly for the required purpose; permissions are not broadened automatically." } }
    return $context
}

function Assert-SwpGraphUri {
    param([string]$Address, [string]$ExpectedPath = '')
    $uri = $null
    if (-not [uri]::TryCreate($Address, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -cne 'https' -or $uri.Host -ine 'graph.microsoft.com' -or $uri.Port -ne 443 -or $uri.UserInfo -or $uri.Fragment -or -not $uri.AbsolutePath.StartsWith('/v1.0/', [StringComparison]::Ordinal)) {
        throw 'Rejected Graph URL: only HTTPS graph.microsoft.com/v1.0 with no credentials or fragment is permitted.'
    }
    if ($ExpectedPath -and $uri.AbsolutePath -cne $ExpectedPath) { throw 'Rejected pagination URL: the collection path changed.' }
}
function Invoke-SwpGet {
    param([string]$Address)
    Assert-SwpGraphUri $Address
    for ($attempt = 0; $attempt -le $MaxRetries; $attempt++) {
        # Recheck before every call: another script in this process must not switch tenants unnoticed.
        $verifiedContext = Assert-SwpContext $script:SwpRequiredScopes
        if ([string](Get-SwpValue $verifiedContext 'Account') -ine $script:SwpInitialAccount -or [string](Get-SwpValue $verifiedContext 'ClientId') -ine $script:SwpInitialClientId) { throw 'The Graph account or client changed during this report. No mixed-identity report was produced.' }
        try { return Invoke-MgGraphRequest -Method GET -Uri $Address -OutputType PSObject -ErrorAction Stop }
        catch {
            $response = Get-SwpValue $_.Exception 'Response'
            $rawStatus = Get-SwpValue $response 'StatusCode'
            if ($null -eq $rawStatus) { $rawStatus = Get-SwpValue $_.Exception 'ResponseStatusCode' }
            $status = 0
            if ($null -ne $rawStatus) { try { $status = [int]$rawStatus } catch { $status = 0 } }
            if ($status -notin @(429, 503) -or $attempt -ge $MaxRetries) { throw }
            $delay = [int][Math]::Pow(2, $attempt + 1)
            $headers = Get-SwpValue $response 'Headers'
            $retry = Get-SwpValue $headers 'Retry-After'
            if ($retry -is [array]) { $retry = $retry[0] }
            $typedRetry = Get-SwpValue $headers 'RetryAfter'
            if ($null -ne $retry) {
                $seconds = 0
                $retryDate = [DateTimeOffset]::MinValue
                if ([int]::TryParse([string]$retry, [ref]$seconds)) { $delay = [Math]::Max(0, $seconds) }
                elseif ([DateTimeOffset]::TryParse([string]$retry, [ref]$retryDate)) { $delay = [int][Math]::Ceiling([Math]::Max(0, ($retryDate - [DateTimeOffset]::UtcNow).TotalSeconds)) }
                else { throw 'Unrecognized Retry-After value. Stopping rather than guessing a retry deadline.' }
            } elseif ($null -ne $typedRetry) {
                $delta = Get-SwpValue $typedRetry 'Delta'
                $date = Get-SwpValue $typedRetry 'Date'
                if ($null -ne $delta) { $delay = [int][Math]::Ceiling($delta.TotalSeconds) }
                elseif ($null -ne $date) { $delay = [int][Math]::Ceiling([Math]::Max(0, ($date - [DateTimeOffset]::UtcNow).TotalSeconds)) }
            }
            if ($delay -gt 60) { throw 'Retry-After exceeds the 60-second per-retry budget. No partial report was produced; retry later.' }
            Start-Sleep -Seconds ([Math]::Max(0, $delay))
        }
    }
}
function Get-SwpPages {
    param([string]$Address, [System.Collections.Generic.List[object]]$Items = [System.Collections.Generic.List[object]]::new())
    $expectedPath = ([uri]$Address).AbsolutePath
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $next = $Address
    $pageCount = 0
    while (-not [string]::IsNullOrEmpty($next)) {
        Assert-SwpGraphUri $next $expectedPath
        if (-not $seen.Add($next)) { throw 'Repeated pagination URL: refusing an incomplete looped report.' }
        if ($pageCount -ge $MaxPages) { throw 'Page limit exceeded. No partial report/CSV was produced. Narrow the query or explicitly review a higher bound.' }
        $page = Invoke-SwpGet $next
        $values = Get-SwpValue $page 'value'
        if ($null -eq $values -or $values -is [string] -or $values -isnot [System.Collections.IEnumerable]) { throw 'Malformed Graph collection response: no value array.' }
        foreach ($item in $values) {
            if ([string]::IsNullOrWhiteSpace([string](Get-SwpValue $item 'id'))) { throw 'Graph returned an object without an ID; report aborted rather than dropping it.' }
            if ($items.Count -ge $MaxRecords) { throw 'Record limit exceeded. No partial report/CSV was produced.' }
            [void]$items.Add($item)
        }
        $pageCount++
        $next = [string](Get-SwpValue $page '@odata.nextLink')
    }
    return ,($items.ToArray())
}
function ConvertTo-SwpCsvCell {
    param($Value)
    if ($null -eq $Value) { return '[null]' }
    $text = [string]$Value
    # Neutralize formula prefixes even behind spaces/control characters. Raw output objects remain unchanged.
    if ($text -match '^[\s\x00-\x1f]*[=+\-@]' -or $text -match '^[\t\r\n]') { return "'" + $text }
    return $text
}
function Export-SwpCsv {
    param([object[]]$Rows, [string]$Prefix, [string[]]$Columns)
    $directory = Get-Item -LiteralPath $CsvDirectory -ErrorAction Stop
    if (-not $directory.PSIsContainer -or $directory.PSProvider.Name -ne 'FileSystem') { throw 'CsvDirectory must be an existing filesystem directory.' }
    $path = Join-Path $directory.FullName ($Prefix + '_' + $asOf.ToString('yyyy-MM-dd') + '.csv')
    # Even a zero-match report has a header. Preserve nulls using the documented marker.
    $safeRows = foreach ($row in $Rows) {
        $safe = [ordered]@{}
        foreach ($name in $Columns) { $safe[$name] = ConvertTo-SwpCsvCell (Get-SwpValue $row $name) }
        [pscustomobject]$safe
    }
    if ($Rows.Count -eq 0) { $lines = @(('"' + ($Columns -join '","') + '"')) }
    else { $lines = @($safeRows | ConvertTo-Csv -NoTypeInformation) }
    Write-SwpCsvFile -Path $path -Lines $lines
}
function Write-SwpCsvFile {
    param([string]$Path, [string[]]$Lines)
    $stream = $null
    $writer = $null
    $created = $false
    try {
        # CreateNew is the overwrite guarantee, not a vulnerable Test-Path-then-write check.
        $stream = [IO.File]::Open($path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $created = $true
        $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
        foreach ($line in $lines) { $writer.WriteLine($line) }
        $writer.Flush()
        $writer.Dispose(); $writer = $null; $stream = $null
        return $path
    } catch {
        $failure = $_
        # Dispose may itself fail after a disk/flush error; still attempt partial-file cleanup.
        try { if ($null -ne $writer) { $writer.Dispose() } elseif ($null -ne $stream) { $stream.Dispose() } } catch { }
        try { if ($created -and [IO.File]::Exists($path)) { [IO.File]::Delete($path) } }
        catch { throw "CSV export failed and partial-file cleanup also failed at '$path'. Remove the partial file explicitly. Original failure: $($failure.Exception.Message)" }
        throw $failure
    }
}


function ConvertTo-SwpUtc {
    param($Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    if ($Value -is [DateTimeOffset]) { return $Value.ToUniversalTime() }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return $null }
        return ([DateTimeOffset]$Value).ToUniversalTime()
    }
    $parsed = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) { return $parsed.ToUniversalTime() }
    return $null
}
function Format-SwpUtc {
    param($Date)
    if ($null -ne $Date) { return $Date.ToString('o') }
    return $null
}
function Get-SwpAge {
    param($Date)
    if ($null -eq $Date -or $Date -gt $asOf) { return $null }
    return [Math]::Floor(($asOf - $Date).TotalDays)
}
if ($TenantId -eq [guid]::Empty) { throw 'TenantId cannot be empty.' }
$script:SwpRequiredScopes = @('User.Read.All')
$context = Assert-SwpContext $script:SwpRequiredScopes
$script:SwpInitialAccount = [string](Get-SwpValue $context 'Account')
$script:SwpInitialClientId = [string](Get-SwpValue $context 'ClientId')
if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) { throw 'Microsoft.Graph.Authentication is required.' }
$asOf = [DateTimeOffset]::UtcNow

$filter = [uri]::EscapeDataString("userType eq 'Guest'")
$address = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,mail,accountEnabled,createdDateTime,userType,externalUserState,externalUserStateChangeDateTime&$filter=' + $filter
$guests = Get-SwpPages $address
$pending = 0
$unknownAge = 0
$rows = @(foreach ($guest in $guests) {
    if ((Get-SwpValue $guest 'userType') -ne 'Guest' -or (Get-SwpValue $guest 'externalUserState') -ne 'PendingAcceptance') { continue }
    $pending++
    $created = ConvertTo-SwpUtc (Get-SwpValue $guest 'createdDateTime')
    $changed = ConvertTo-SwpUtc (Get-SwpValue $guest 'externalUserStateChangeDateTime')
    $age = Get-SwpAge $created
    if ($null -eq $age) { $unknownAge++ }
    if ($MinimumAccountAgeDays -gt 0 -and ($null -eq $age -or $age -lt $MinimumAccountAgeDays)) { continue }
    [pscustomobject][ordered]@{
        DisplayName = Get-SwpValue $guest 'displayName'
        UserPrincipalName = Get-SwpValue $guest 'userPrincipalName'
        Mail = Get-SwpValue $guest 'mail'
        UserObjectId = Get-SwpValue $guest 'id'
        AccountEnabled = Get-SwpValue $guest 'accountEnabled'
        CreatedDateTimeUtc = Format-SwpUtc $created
        ExternalUserState = Get-SwpValue $guest 'externalUserState'
        InvitationStateChangedUtc = Format-SwpUtc $changed
        DaysSinceAccountCreated = $age
        DaysSinceStateChange = Get-SwpAge $changed
        ReportAsOfUtc = Format-SwpUtc $asOf
    }
})
$columns = @('DisplayName','UserPrincipalName','Mail','UserObjectId','AccountEnabled','CreatedDateTimeUtc','ExternalUserState','InvitationStateChangedUtc','DaysSinceAccountCreated','DaysSinceStateChange','ReportAsOfUtc')
$path = Export-SwpCsv -Rows $rows -Prefix 'Entra_PendingGuestInvitations' -Columns $columns
Write-Host "UTC observation: $($asOf.ToString('o')); guests scanned: $($guests.Count); pending: $pending; unknown account ages: $unknownAge; exported: $($rows.Count)"
Write-Host "Saved: $path"
