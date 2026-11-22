#requires -Version 7.0
<#
.SYNOPSIS
Read one explicit Entra user or a bounded, explicitly supplied OData filter.
.DESCRIPTION
Requires Microsoft.Graph.Authentication and a delegated Global/Process context
for TenantId with User.Read.All. Connect separately; this script never signs in,
installs modules or changes directory data. No unfiltered tenant-wide default.
UPNs are escaped as OData string literals; IDs may be GUIDs. Filters are data sent
to Graph, never executed as PowerShell. Unsupported filters fail rather than widen.
Paging, record counts and 429/503 retries are bounded; failures yield no partial
report or CSV. Optional CSV is a new timestamped file in your chosen filesystem directory, never overwritten;
formula-like strings are prefixed with an apostrophe and nulls become [null].
CSV escaping is transport hygiene, not an assurance about every spreadsheet import.
.EXAMPLE
./Get-SwpEntraUser.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -UserId 'alex@example.invalid'
.EXAMPLE
./Get-SwpEntraUser.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -Filter "startsWith(displayName,'Alex')" -MaxRecords 100
.NOTES
Choose a protected export directory; filesystem paths may be network-mounted or
cloud-synced. CSV contains directory data and must be handled accordingly.
MaxRetries bounds this script's loop; the Graph SDK may apply its own retry policy.
Read-only Graph v1.0. Selected profile fields only; no sign-in history, licenses,
passwords, roles or PIM. Null properties remain unknown, not false. Consistency
is a point-in-time read, not an immutable snapshot. Live-tenant validation pending.
#>
[CmdletBinding(DefaultParameterSetName='User')]
param(
    [Parameter(Mandatory,ParameterSetName='User')][ValidateLength(1,320)][string]$UserId,
    [Parameter(Mandatory,ParameterSetName='Filter')][ValidateLength(1,500)][ValidatePattern('\S')][string]$Filter,

    [Parameter(Mandatory)][guid]$TenantId,
    [ValidateRange(1,100)][int]$MaxPages = 20,
    [ValidateRange(1,10000)][int]$MaxRecords = 1000,
    [ValidateRange(0,5)][int]$MaxRetries = 3,
    [string]$CsvDirectory
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
    if ($null -eq $context) { throw 'No Graph context. First run Connect-SwpEntra.ps1 for the explicit tenant and purpose.' }
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
    param([string]$Address)
    $expectedPath = ([uri]$Address).AbsolutePath
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $items = [System.Collections.Generic.List[object]]::new()
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
    param([object[]]$Rows, [string]$Prefix)
    if ([string]::IsNullOrWhiteSpace($CsvDirectory)) { return $null }
    $directory = Get-Item -LiteralPath $CsvDirectory -ErrorAction Stop
    if (-not $directory.PSIsContainer -or $directory.PSProvider.Name -ne 'FileSystem') { throw 'CsvDirectory must be an existing filesystem directory.' }
    $stamp = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ')
    $leaf = $Prefix + '-' + $stamp + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.csv'
    $path = Join-Path $directory.FullName $leaf
    $safeRows = foreach ($row in $Rows) {
        $safe = [ordered]@{}
        foreach ($property in $row.PSObject.Properties) { $safe[$property.Name] = ConvertTo-SwpCsvCell $property.Value }
        [pscustomobject]$safe
    }
    $lines = @($safeRows | ConvertTo-Csv -NoTypeInformation)
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

if ($TenantId -eq [guid]::Empty) { throw 'TenantId cannot be empty.' }
$script:SwpRequiredScopes = @('User.Read.All')
$context = Assert-SwpContext $script:SwpRequiredScopes
$script:SwpInitialAccount = [string](Get-SwpValue $context 'Account')
$script:SwpInitialClientId = [string](Get-SwpValue $context 'ClientId')
if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) { throw 'Invoke-MgGraphRequest is required from Microsoft.Graph.Authentication.' }
$select = 'id,displayName,userPrincipalName,mail,accountEnabled,userType'
if ($PSCmdlet.ParameterSetName -eq 'User') {
    if ($UserId -match '[\x00-\x1f]') { throw 'UserId cannot contain control characters.' }
    $id = [guid]::Empty
    if ([guid]::TryParse($UserId, [ref]$id)) { $address = 'https://graph.microsoft.com/v1.0/users/' + $id.ToString() + '?$select=' + $select }
    else {
        if ($UserId -notmatch '^[^\s@]+@[^\s@]+$') { throw 'UserId must be an object GUID or an explicit UPN.' }
        $literal = [uri]::EscapeDataString($UserId.Replace("'", "''"))
        $address = "https://graph.microsoft.com/v1.0/users('$literal')?`$select=$select"
    }
    $raw = @(Invoke-SwpGet $address)
    if ([string]::IsNullOrWhiteSpace([string](Get-SwpValue $raw[0] 'id'))) { throw 'User response has no ID; no report was produced.' }
    if ($id -ne [guid]::Empty -and [string](Get-SwpValue $raw[0] 'id') -ine $id.ToString()) { throw 'User response ID does not match the requested object GUID.' }
} else {
    if ($Filter -match '[\x00-\x1f]') { throw 'Filter cannot contain control characters.' }
    $address = 'https://graph.microsoft.com/v1.0/users?$select=' + $select + '&$top=100&$filter=' + [uri]::EscapeDataString($Filter)
    $raw = Get-SwpPages $address
}
$retrieved = [DateTime]::UtcNow.ToString('o')
$coverage = 'Selected visible user profile fields from the requested v1.0 query. Null means unknown/unavailable; not a sign-in, license or privilege audit.'
$users = @($raw | ForEach-Object {
    [pscustomobject]@{ Id = Get-SwpValue $_ 'id'; DisplayName = Get-SwpValue $_ 'displayName'; UserPrincipalName = Get-SwpValue $_ 'userPrincipalName'; Mail = Get-SwpValue $_ 'mail'; AccountEnabled = Get-SwpValue $_ 'accountEnabled'; UserType = Get-SwpValue $_ 'userType' }
})
$rows = [System.Collections.Generic.List[object]]::new()
[void]$rows.Add([pscustomobject]@{ RecordType='Summary'; TenantId=$TenantId.ToString(); RetrievedAtUtc=$retrieved; Coverage=$coverage; Count=$users.Count; Id=$null; DisplayName=$null; UserPrincipalName=$null; Mail=$null; AccountEnabled=$null; UserType=$null })
foreach ($user in $users) { [void]$rows.Add([pscustomobject]@{ RecordType='User'; TenantId=$TenantId.ToString(); RetrievedAtUtc=$retrieved; Coverage=$coverage; Count=$null; Id=$user.Id; DisplayName=$user.DisplayName; UserPrincipalName=$user.UserPrincipalName; Mail=$user.Mail; AccountEnabled=$user.AccountEnabled; UserType=$user.UserType }) }
$csvPath = Export-SwpCsv $rows.ToArray() 'ShellWePlay-EntraUsers'
[pscustomobject]@{ TenantId=$TenantId.ToString(); Account=Get-SwpValue $context 'Account'; RetrievedAtUtc=$retrieved; Coverage=$coverage; RecordCount=$users.Count; Users=$users; CsvPath=$csvPath }
