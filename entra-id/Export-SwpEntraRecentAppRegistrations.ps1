#requires -Version 7.0
<#
.SYNOPSIS
Export current application registrations created within a backward-looking UTC window.
.DESCRIPTION
Reads Graph v1.0 applications and their current owners. Locally filters createdDateTime
because application creation time filtering is not assumed supported. Deleted apps
and unusable creation timestamps are excluded. Derived types remain identified.
OwnerSummary is a JSON array with each owner's ID, type, name, UPN and mail together.
OwnerCount is null for incomplete enumeration, zero only for a successful empty list.
OwnerLookupStatus distinguishes complete, limited details, incomplete and failed.
Current owners are not evidence of who created an application. AppName, client/object
IDs, creation time, whole elapsed AgeDays, audience, publisher domain, description,
type and ReportAsOfUtc are exported. Requires Application.Read.All and
User.ReadBasic.All delegated for application and basic human owner metadata.
Version 1.1.0 preserves one application row and adds credential counts/status.
A companion CSV contains one row per returned certificate, client secret or other
key credential. CredentialDescription is displayName; validity start/end dates
are not creation/upload dates. DaysBack still filters application creation time.
Only allowlisted metadata is exported: never secretText, hint or raw key material.
CustomKeyIdentifierHex is a decoded custom identifier, not a verified thumbprint.
ValidityStatus describes dates only, never application health or actual use.
Missing/malformed collections produce unknown counts, not an empty inventory.
.PARAMETER DaysBack
Positive whole number of days before one frozen UTC report timestamp; prompts if omitted.
.EXAMPLE
# Fictional mylab.dev tenant; replace GUID and sign in with an authorized account.
Connect-MgGraph -TenantId '11111111-1111-4111-8111-111111111111' -Scopes 'Application.Read.All','User.ReadBasic.All' -ContextScope Process
./Export-SwpEntraRecentAppRegistrations.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -DaysBack 30 -CsvDirectory './Reports'

.PARAMETER TenantId
Required intended tenant GUID. Global cloud, delegated, Process-scoped context.
.PARAMETER CsvDirectory
Existing protected filesystem directory; defaults to the current directory.
Two dated UTF-8 CSVs, atomic no-overwrite opens; [null] means unavailable.
Formula-like cells are neutralized, including signed negative day text. Zero
matches produce two header-only files. A companion-write failure removes newly
created files where possible and throws; an incomplete pair is never success.
.PARAMETER MaxPages
Maximum pages per collection (default 10000). A collection exceeding the bound fails.
.PARAMETER MaxRecords
Maximum records per collection (default 1000000). Review memory/query load.
.PARAMETER MaxRetries
Maximum additional attempts for HTTP 429/503 (default 3); honors Retry-After
up to 60 seconds, otherwise stops. The Graph SDK can also retry internally.
.NOTES
Version 1.1.0 - 2026-11-25. Adds credential metadata; lab validation pending.
PowerShell 7 and Microsoft.Graph.Authentication (install separately).
Use approved permissions and a dedicated session; no module install, consent,
authentication, or directory write is performed by this report. No snapshot
isolation; objects may change during paging. Protect local directory data.
Both files use the same ReportAsOfUtc. DaysUntilExpiry is signed fractional days
rounded to two decimals; missing/invalid validity endpoints leave it unavailable.
ValidityStatus is Expired (end <= report time), NotYetValid (start > report time),
WithinValidityPeriod, UnknownDates or InvalidDateRange (start >= end).
CredentialMetadataStatus is Complete, LimitedDetails, Incomplete or Unavailable.
Optional absent display names/custom identifiers do not imply incomplete data.
IdentifierStatus reports KeyId=Valid|Missing|Invalid and
CustomKeyIdentifier=Decoded|Absent|Invalid|NotApplicable independently, separated
by a semicolon. Password custom identifiers are not used and are NotApplicable.
Known X509 key types are certificates; other key types remain OtherKey. Missing
key types leave certificate/other-key counts unknown. An unavailable collection,
unknown validity dates or invalid range leaves ExpiredCredentialCount unknown.
The selected keyCredentials property has a tenant-level Graph throttle (currently
150 requests/minute); existing bounded GET retries honor Retry-After. No extra
write scopes, credential retrieval calls, audit logs or sign-in queries are used.
References: https://learn.microsoft.com/en-us/graph/api/application-list?view=graph-rest-1.0
https://learn.microsoft.com/en-us/graph/api/application-list-owners?view=graph-rest-1.0
https://learn.microsoft.com/en-us/graph/api/resources/application?view=graph-rest-1.0
https://learn.microsoft.com/en-us/graph/api/resources/keycredential?view=graph-rest-1.0
https://learn.microsoft.com/en-us/graph/api/resources/passwordcredential?view=graph-rest-1.0
https://learn.microsoft.com/en-us/graph/permissions-reference
https://learn.microsoft.com/en-us/graph/paging
#>
[CmdletBinding()]
param(
    [ValidatePattern('^[0-9]{1,6}$')][ValidateScript({ [int]$_ -ge 1 -and [int]$_ -le 365000 })][string]$DaysBack,
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

function Export-SwpReportPair {
    param([object[]]$ApplicationRows, [object[]]$CredentialRows, [string[]]$ApplicationColumns, [string[]]$CredentialColumns)
    $completed = [System.Collections.Generic.List[string]]::new()
    try {
        # Each writer uses CreateNew. Track only paths created successfully by this
        # run, so a collision in either filename can never delete an older report.
        $applicationPath = Export-SwpCsv -Rows $ApplicationRows -Prefix 'Entra_AppRegistrations' -Columns $ApplicationColumns
        [void]$completed.Add($applicationPath)
        $credentialPath = Export-SwpCsv -Rows $CredentialRows -Prefix 'Entra_AppRegistrationCredentials' -Columns $CredentialColumns
        [void]$completed.Add($credentialPath)
        return [pscustomobject]@{ ApplicationPath=$applicationPath; CredentialPath=$credentialPath }
    } catch {
        $failure = $_
        $uncleaned = [System.Collections.Generic.List[string]]::new()
        foreach ($path in $completed) {
            try { if ([IO.File]::Exists($path)) { [IO.File]::Delete($path) } }
            catch { [void]$uncleaned.Add($path) }
        }
        if ($uncleaned.Count) { throw "CSV pair export failed; cleanup also failed at '$($uncleaned -join "', '")'. Remove these incomplete report files explicitly. Original failure: $($failure.Exception.Message)" }
        throw "CSV pair export failed; this run did not complete. $($failure.Exception.Message)"
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

if (-not $PSBoundParameters.ContainsKey('DaysBack')) {
    $answer = Read-Host 'How many days back should I look for newly created application registrations?'
    $number = 0
    if ($answer -notmatch '^\s*[0-9]+\s*$' -or -not [int]::TryParse($answer, [ref]$number) -or $number -lt 1 -or $number -gt 365000) {
        throw 'DaysBack must be a positive whole number from 1 through 365000.'
    }
    $DaysBack = $number
}
$script:SwpRequiredScopes = @('Application.Read.All', 'User.ReadBasic.All')
$context = Assert-SwpContext $script:SwpRequiredScopes
$script:SwpInitialAccount = [string](Get-SwpValue $context 'Account')
$script:SwpInitialClientId = [string](Get-SwpValue $context 'ClientId')
if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) { throw 'Microsoft.Graph.Authentication is required.' }
$asOf = [DateTimeOffset]::UtcNow

$start = $asOf.AddDays(-([int]$DaysBack))
Write-Host "UTC creation window: $($start.ToString('o')) through $($asOf.ToString('o'))"
# Retrieve only relevant application properties; follow all application pages.
$apps = Get-SwpPages 'https://graph.microsoft.com/v1.0/applications?$select=id,appId,displayName,createdDateTime,signInAudience,publisherDomain,description,keyCredentials,passwordCredentials'
$missingDates = 0
$warnings = 0
$credentialWarnings = 0
$credentialRows = [System.Collections.Generic.List[object]]::new()
$rows = @(foreach ($app in $apps) {
    $created = ConvertTo-SwpUtc (Get-SwpValue $app 'createdDateTime')
    if ($null -eq $created) { $missingDates++; continue }
    if ($created -lt $start -or $created -gt $asOf) { continue }
    $id = [guid](Get-SwpValue $app 'id')
    $type = Get-SwpValue $app '@odata.type'
    if (-not $type) { $type = '#microsoft.graph.application' }
    $owners = [System.Collections.Generic.List[object]]::new()
    $status = 'Complete'
    $count = $null
    try {
        # Mixed directoryObject types have different properties; the default
        # owners response preserves type and basic details without an invalid mixed select.
        $null = Get-SwpPages "https://graph.microsoft.com/v1.0/applications/$id/owners" -Items $owners
        $count = $owners.Count
    } catch {
        # A context switch must abort the whole report, not become a row warning.
        $verified = Assert-SwpContext $script:SwpRequiredScopes
        if ([string](Get-SwpValue $verified 'Account') -ine $script:SwpInitialAccount -or [string](Get-SwpValue $verified 'ClientId') -ine $script:SwpInitialClientId) { throw 'Graph identity changed.' }
        if ($owners.Count -gt 0) { $status = 'Incomplete' }
        else { $status = 'FailedOrDenied' }
        Write-Warning "Owner lookup $status for application $id; known owners retained, count unavailable."
    }
    $details = @(foreach ($owner in $owners) {
        $ownerType = Get-SwpValue $owner '@odata.type'
        $name = Get-SwpValue $owner 'displayName'
        $upn = Get-SwpValue $owner 'userPrincipalName'
        $mail = Get-SwpValue $owner 'mail'
        if ($status -eq 'Complete' -and (-not $ownerType -or -not $name -or ($ownerType -eq '#microsoft.graph.user' -and -not $upn))) { $status = 'LimitedDetails' }
        [pscustomobject][ordered]@{
            ObjectId = Get-SwpValue $owner 'id'
            ObjectType = $ownerType
            DisplayName = $name
            UserPrincipalName = $upn
            Mail = $mail
        }
    })
    if ($status -ne 'Complete') { $warnings++ }
    if ($status -eq 'LimitedDetails') { Write-Warning "Some owner details are unavailable for application $id; returned owner count retained." }
    $credentialMetadata = Get-SwpCredentialMetadata $app
    foreach ($credentialRow in $credentialMetadata.Rows) { [void]$credentialRows.Add($credentialRow) }
    if ($credentialMetadata.CredentialMetadataStatus -ne 'Complete') { $credentialWarnings++ }
    [pscustomobject][ordered]@{
        AppName = Get-SwpValue $app 'displayName'
        ApplicationClientId = Get-SwpValue $app 'appId'
        ApplicationObjectId = $id
        ApplicationObjectType = $type
        CreatedDateTimeUtc = Format-SwpUtc $created
        AgeDays = Get-SwpAge $created
        SignInAudience = Get-SwpValue $app 'signInAudience'
        PublisherDomain = Get-SwpValue $app 'publisherDomain'
        Description = Get-SwpValue $app 'description'
        OwnerCount = $count
        OwnerSummary = ConvertTo-Json -InputObject $details -Compress -Depth 4
        OwnerLookupStatus = $status
        ReportAsOfUtc = Format-SwpUtc $asOf
        CertificateCount = $credentialMetadata.CertificateCount
        ClientSecretCount = $credentialMetadata.ClientSecretCount
        OtherKeyCredentialCount = $credentialMetadata.OtherKeyCredentialCount
        ExpiredCredentialCount = $credentialMetadata.ExpiredCredentialCount
        CredentialMetadataStatus = $credentialMetadata.CredentialMetadataStatus
    }
})
$columns = @('AppName','ApplicationClientId','ApplicationObjectId','ApplicationObjectType','CreatedDateTimeUtc','AgeDays','SignInAudience','PublisherDomain','Description','OwnerCount','OwnerSummary','OwnerLookupStatus','ReportAsOfUtc','CertificateCount','ClientSecretCount','OtherKeyCredentialCount','ExpiredCredentialCount','CredentialMetadataStatus')
$credentialColumns = @('AppName','ApplicationClientId','ApplicationObjectId','CredentialKind','CredentialDescription','CredentialId','StartDateTimeUtc','EndDateTimeUtc','DaysUntilExpiry','ValidityStatus','KeyType','KeyUsage','CustomKeyIdentifierHex','IdentifierStatus','ReportAsOfUtc')
$paths = Export-SwpReportPair -ApplicationRows $rows -CredentialRows $credentialRows.ToArray() -ApplicationColumns $columns -CredentialColumns $credentialColumns
Write-Host "Applications scanned: $($apps.Count); matched: $($rows.Count); unusable dates excluded: $missingDates; owner lookup warnings: $warnings"
Write-Host "Credential rows: $($credentialRows.Count); applications with unavailable/incomplete credential metadata: $credentialWarnings"
Write-Host "Saved: $($paths.ApplicationPath)"
Write-Host "Saved: $($paths.CredentialPath)"
