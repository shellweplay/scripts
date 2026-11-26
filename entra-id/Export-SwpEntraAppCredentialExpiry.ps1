#requires -Version 7.0
<#
.SYNOPSIS
Report application registration credential expirations across current Entra apps.
.DESCRIPTION
Enumerates Graph v1.0 applications without a creation-date filter. Returns
certificate, client-secret and other key credential metadata whose validity
end is in the next DaysAhead days, with optional already-expired credentials.
Includes unknown/invalid dates and explicit coverage-warning rows for incomplete
collections so missing data is not interpreted as zero risk. This is not the
recently created application report and does not enumerate service principals.
No owner lookup, audit or sign-in data, secret values, private keys, or raw
certificate key exports. Graph may return secret hints in credential metadata;
these are ignored. The list endpoint does not return raw certificate key data
per the keyCredential resource documentation. Only GET reads are performed.
.PARAMETER TenantId
Required intended tenant GUID; delegated, Global cloud, Process-scoped context.
.PARAMETER DaysAhead
Number of positive whole days in the future expiry window (default 30).
.PARAMETER IncludeExpired
Also include credentials whose valid end timestamp is on or before report time.
.PARAMETER CsvDirectory
Existing protected filesystem directory; default current directory.
Dated UTF-8 CSV; no overwrite; zero rows produce header-only CSV.
.PARAMETER MaxPages
Maximum application pages, default 10000. Exceeding the limit aborts.
.PARAMETER MaxRecords
Maximum returned applications, default 1000000. Exceeding the limit aborts.
.PARAMETER MaxRetries
Additional attempts on HTTP 429/503, default 3; Retry-After up to 60 seconds.
SDK may also perform retries. No partial main collection is exported on failure.
.EXAMPLE
# Fictional mylab.dev tenant; replace GUID with your authorized test tenant.
Connect-MgGraph -TenantId '11111111-1111-4111-8111-111111111111' -Scopes 'Application.Read.All' -ContextScope Process
./Export-SwpEntraAppCredentialExpiry.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -DaysAhead 30 -IncludeExpired -CsvDirectory './Reports'
.NOTES
Version 1.0.0 - 2026-11-26. Lab validation pending.
PowerShell 7, Microsoft.Graph.Authentication installed separately; delegated
Application.Read.All and signed-in account rights. No consent or install by
the report. The keyCredentials selection has a 150 requests/minute tenant limit.
One report timestamp, not an atomic directory snapshot. Scope is visible
application objects (including derived types), not enterprise apps/service
principals, federated identity credentials or deleted apps. Credential validity
does not prove actual usage or an application outage. Validity start is not a
creation/upload date. CustomKeyIdentifierHex is not a verified thumbprint.
Output: RowKind (Credential or CoverageWarning), AppName, ApplicationClientId,
ApplicationObjectId, ApplicationObjectType, CredentialKind, CredentialDescription,
CredentialId, StartDateTimeUtc, EndDateTimeUtc, DaysUntilExpiry (signed fractional
days), ValidityStatus, KeyType, KeyUsage, CustomKeyIdentifierHex, IdentifierStatus,
CredentialMetadataStatus, ReportAsOfUtc, WindowEndUtc. [null] means unavailable.
UnknownDates/InvalidDateRange remain visible regardless of window; CoverageWarning
rows have no credential and identify incomplete/unavailable application metadata.
Formula-like cells are neutralized. Treat exports as confidential.
References:
https://learn.microsoft.com/en-us/graph/api/application-list?view=graph-rest-1.0
https://learn.microsoft.com/en-us/graph/api/resources/keycredential?view=graph-rest-1.0
https://learn.microsoft.com/en-us/graph/api/resources/passwordcredential?view=graph-rest-1.0
https://learn.microsoft.com/en-us/graph/paging
https://learn.microsoft.com/en-us/graph/permissions-reference
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][guid]$TenantId,
    [ValidateRange(1,3650)][int]$DaysAhead = 30,
    [switch]$IncludeExpired,
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
function Get-SwpCredentialCollection {
    param($Application, [string]$Name)
    $raw = Get-SwpValue $Application $Name
    $items = [System.Collections.Generic.List[object]]::new()
    if ($null -eq $raw) { return [pscustomobject]@{ Items=$items.ToArray(); Status='Unavailable' } }
    # A missing/null property, a scalar and a returned [] are different evidence.
    if ($raw -is [string] -or $raw -is [System.Collections.IDictionary] -or $raw -isnot [System.Collections.IEnumerable]) {
        return [pscustomobject]@{ Items=$items.ToArray(); Status='Malformed' }
    }
    $status = 'Complete'
    foreach ($item in $raw) {
        if ($null -eq $item -or ($item -isnot [System.Collections.IDictionary] -and $item -isnot [pscustomobject])) {
            $status = 'Malformed'
            continue
        }
        [void]$items.Add($item)
    }
    return [pscustomobject]@{ Items=$items.ToArray(); Status=$status }
}
function ConvertTo-SwpCredentialDate {
    param($Value)
    if ($Value -is [DateTimeOffset] -or $Value -is [datetime]) { return ConvertTo-SwpUtc $Value }
    # Graph DateTimeOffset values carry a full ISO timestamp and offset. Never
    # guess a date from an arbitrary string or stringify an unexpected object.
    if ($Value -is [string] -and $Value -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})$') {
        return ConvertTo-SwpUtc $Value
    }
    return $null
}
function Get-SwpCredentialValidity {
    param($Start, $End)
    if ($null -eq $Start -or $null -eq $End) { return 'UnknownDates' }
    if ($Start -ge $End) { return 'InvalidDateRange' }
    if ($End -le $asOf) { return 'Expired' }
    if ($Start -gt $asOf) { return 'NotYetValid' }
    return 'WithinValidityPeriod'
}
function ConvertTo-SwpCustomIdentifier {
    param($Value)
    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) {
        return [pscustomobject]@{ Hex=$null; Status='Absent' }
    }
    $bytes = $null
    try {
        if ($Value -is [string]) {
            # Graph's JSON binary encoding is base64, even when text looks like
            # hexadecimal. Never guess the encoding from the characters alone.
            $bytes = [Convert]::FromBase64String($Value)
        } elseif ($Value -is [byte[]]) {
            $bytes = $Value
        } elseif ($Value -is [array]) {
            $buffer = [System.Collections.Generic.List[byte]]::new()
            foreach ($part in $Value) {
                if (($part -isnot [byte] -and $part -isnot [int] -and $part -isnot [long]) -or $part -lt 0 -or $part -gt 255) { throw 'Invalid binary identifier.' }
                [void]$buffer.Add([byte]$part)
            }
            $bytes = $buffer.ToArray()
        } else { throw 'Unsupported identifier representation.' }
        if ($bytes.Length -eq 0) { return [pscustomobject]@{ Hex=$null; Status='Absent' } }
        return [pscustomobject]@{ Hex=[BitConverter]::ToString($bytes).Replace('-',''); Status='Decoded' }
    } catch {
        # Never echo the supplied value or exception: malformed input can contain
        # unexpected data. This status is enough to show the loss of metadata.
        return [pscustomobject]@{ Hex=$null; Status='Invalid' }
    }
}
function Get-SwpCredentialMetadata {
    param($Application)
    $keys = Get-SwpCredentialCollection $Application 'keyCredentials'
    $secrets = Get-SwpCredentialCollection $Application 'passwordCredentials'
    $details = [System.Collections.Generic.List[object]]::new()
    $limited = $false
    $unknownKeyType = $false
    $unknownValidity = $false
    $certificateCount = 0
    $otherKeyCount = 0
    $expiredCount = 0
    foreach ($source in @(@{ IsKey=$true; Collection=$keys }, @{ IsKey=$false; Collection=$secrets })) {
        foreach ($credential in $source.Collection.Items) {
            $kind = 'ClientSecret'
            $keyType = $null
            $keyUsage = $null
            if ($source.IsKey) {
                $rawType = Get-SwpValue $credential 'type'
                if ($rawType -is [string] -and -not [string]::IsNullOrWhiteSpace($rawType)) { $keyType = $rawType }
                else { $unknownKeyType = $true; $limited = $true }
                $rawUsage = Get-SwpValue $credential 'usage'
                if ($rawUsage -is [string] -and -not [string]::IsNullOrWhiteSpace($rawUsage)) { $keyUsage = $rawUsage }
                else { $limited = $true }
                if ($keyType -in @('AsymmetricX509Cert','X509Cert')) { $kind = 'Certificate'; $certificateCount++ }
                else { $kind = 'OtherKey'; $otherKeyCount++ }
            }
            $rawName = Get-SwpValue $credential 'displayName'
            $name = $null
            if ($rawName -is [string]) { $name = $rawName }
            elseif ($null -ne $rawName) { $limited = $true }
            $rawId = Get-SwpValue $credential 'keyId'
            $credentialId = $null
            $idStatus = 'Missing'
            if ($null -ne $rawId -and -not ($rawId -is [string] -and [string]::IsNullOrWhiteSpace($rawId))) {
                $parsedId = [guid]::Empty
                if (($rawId -is [string] -or $rawId -is [guid]) -and [guid]::TryParse([string]$rawId, [ref]$parsedId) -and $parsedId -ne [guid]::Empty) {
                    $credentialId = $parsedId.ToString()
                    $idStatus = 'Valid'
                } else { $idStatus = 'Invalid' }
            }
            # Microsoft marks passwordCredential.customKeyIdentifier "Do not use".
            # Ignore it even if returned; only key credentials have an identifier here.
            $identifier = [pscustomobject]@{ Hex=$null; Status='NotApplicable' }
            if ($source.IsKey) { $identifier = ConvertTo-SwpCustomIdentifier (Get-SwpValue $credential 'customKeyIdentifier') }
            if ($idStatus -ne 'Valid' -or $identifier.Status -eq 'Invalid') { $limited = $true }
            $validFrom = ConvertTo-SwpCredentialDate (Get-SwpValue $credential 'startDateTime')
            $validUntil = ConvertTo-SwpCredentialDate (Get-SwpValue $credential 'endDateTime')
            $validity = Get-SwpCredentialValidity -Start $validFrom -End $validUntil
            $daysUntil = $null
            if ($validity -in @('UnknownDates','InvalidDateRange')) { $unknownValidity = $true; $limited = $true }
            else {
                $daysUntil = [Math]::Round(($validUntil - $asOf).TotalDays, 2, [MidpointRounding]::AwayFromZero)
                if ($validity -eq 'Expired') { $expiredCount++ }
            }
            # Explicit fields only. Do not serialize the credential object, key,
            # secretText or hint, even when a test/response unexpectedly supplies it.
            [void]$details.Add([pscustomobject][ordered]@{
                AppName = Get-SwpValue $Application 'displayName'
                ApplicationClientId = Get-SwpValue $Application 'appId'
                ApplicationObjectId = Get-SwpValue $Application 'id'
                CredentialKind = $kind
                CredentialDescription = $name
                CredentialId = $credentialId
                StartDateTimeUtc = Format-SwpUtc $validFrom
                EndDateTimeUtc = Format-SwpUtc $validUntil
                DaysUntilExpiry = $daysUntil
                ValidityStatus = $validity
                KeyType = $keyType
                KeyUsage = $keyUsage
                CustomKeyIdentifierHex = $identifier.Hex
                IdentifierStatus = "KeyId=$idStatus;CustomKeyIdentifier=$($identifier.Status)"
                ReportAsOfUtc = Format-SwpUtc $asOf
            })
        }
    }
    $metadataStatus = 'Complete'
    if ($keys.Status -eq 'Unavailable' -and $secrets.Status -eq 'Unavailable') { $metadataStatus = 'Unavailable' }
    elseif ($keys.Status -ne 'Complete' -or $secrets.Status -ne 'Complete') { $metadataStatus = 'Incomplete' }
    elseif ($limited) { $metadataStatus = 'LimitedDetails' }
    $secretCount = $null
    if ($secrets.Status -eq 'Complete') { $secretCount = $secrets.Items.Count }
    if ($keys.Status -ne 'Complete' -or $unknownKeyType) { $certificateCount = $null; $otherKeyCount = $null }
    if ($keys.Status -ne 'Complete' -or $secrets.Status -ne 'Complete' -or $unknownValidity) { $expiredCount = $null }
    return [pscustomobject]@{
        Rows=$details.ToArray(); CertificateCount=$certificateCount; ClientSecretCount=$secretCount
        OtherKeyCredentialCount=$otherKeyCount; ExpiredCredentialCount=$expiredCount; CredentialMetadataStatus=$metadataStatus
    }
}

if ($TenantId -eq [guid]::Empty) { throw 'TenantId cannot be empty.' }
$script:SwpRequiredScopes = @('Application.Read.All')
$context = Assert-SwpContext $script:SwpRequiredScopes
$script:SwpInitialAccount = [string](Get-SwpValue $context 'Account')
$script:SwpInitialClientId = [string](Get-SwpValue $context 'ClientId')
if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) { throw 'Microsoft.Graph.Authentication is required.' }
$asOf = [DateTimeOffset]::UtcNow
$end = $asOf.AddDays($DaysAhead)
Write-Host "UTC expiry window: after $($asOf.ToString('o')) through $($end.ToString('o')); include expired: $IncludeExpired"
# Use the collection endpoint only; no single-object key material retrieval.
$apps = Get-SwpPages 'https://graph.microsoft.com/v1.0/applications?$select=id,appId,displayName,keyCredentials,passwordCredentials'
$columns = @('RowKind','AppName','ApplicationClientId','ApplicationObjectId','ApplicationObjectType','CredentialKind','CredentialDescription','CredentialId','StartDateTimeUtc','EndDateTimeUtc','DaysUntilExpiry','ValidityStatus','KeyType','KeyUsage','CustomKeyIdentifierHex','IdentifierStatus','CredentialMetadataStatus','ReportAsOfUtc','WindowEndUtc')
$rows = [System.Collections.Generic.List[object]]::new()
$warnings = 0
$matched = 0
$unknown = 0
foreach ($app in $apps) {
    $type = Get-SwpValue $app '@odata.type'
    if (-not $type) { $type = '#microsoft.graph.application' }
    $metadata = Get-SwpCredentialMetadata $app
    if ($metadata.CredentialMetadataStatus -ne 'Complete') {
        $warnings++
        $warningRow = [ordered]@{}
        foreach ($column in $columns) { $warningRow[$column] = $null }
        $warningRow.RowKind = 'CoverageWarning'
        $warningRow.AppName = Get-SwpValue $app 'displayName'
        $warningRow.ApplicationClientId = Get-SwpValue $app 'appId'
        $warningRow.ApplicationObjectId = Get-SwpValue $app 'id'
        $warningRow.ApplicationObjectType = $type
        $warningRow.CredentialMetadataStatus = $metadata.CredentialMetadataStatus
        $warningRow.ReportAsOfUtc = Format-SwpUtc $asOf
        $warningRow.WindowEndUtc = Format-SwpUtc $end
        [void]$rows.Add([pscustomobject]$warningRow)
    }
    foreach ($credential in $metadata.Rows) {
        # Compare full timestamps rather than rounded DaysUntilExpiry.
        $unclassifiable = $credential.ValidityStatus -in @('UnknownDates','InvalidDateRange')
        if ($unclassifiable) { $unknown++ }
        else {
            $expires = ConvertTo-SwpUtc $credential.EndDateTimeUtc
            if ($expires -le $asOf) {
                if (-not $IncludeExpired) { continue }
            } elseif ($expires -gt $end) { continue }
            $matched++
        }
        $row = [ordered]@{}
        foreach ($column in $columns) { $row[$column] = Get-SwpValue $credential $column }
        $row.RowKind = 'Credential'
        $row.ApplicationObjectType = $type
        $row.CredentialMetadataStatus = $metadata.CredentialMetadataStatus
        $row.WindowEndUtc = Format-SwpUtc $end
        [void]$rows.Add([pscustomobject]$row)
    }
}
$path = Export-SwpCsv -Rows $rows.ToArray() -Prefix 'Entra_AppCredentialExpiry' -Columns $columns
Write-Host "Applications scanned: $($apps.Count); matching credentials: $matched; unknown/invalid date rows: $unknown; application metadata warnings: $warnings"
Write-Host "Saved: $path"
