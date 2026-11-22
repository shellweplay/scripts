#requires -Version 7.0
<#
.SYNOPSIS
Read one group's visible owners and direct OR flat transitive members.
.DESCRIPTION
Requires Microsoft.Graph.Authentication, explicit TenantId/GroupId and delegated
Global/Process context with GroupMember.Read.All. No connection, installation,
write requests, roles/PIM, or permission escalation is performed.
IMPORTANT: v1.0 direct members and owners have documented service-principal
omissions. This report is coverage-limited even when every returned page is read.
It is NOT a full inventory, access certification, or proof of no members/owners.
The documented direct-member beta/$expand workaround is not silently applied:
expanded collections can truncate and beta has different support expectations.
Transitive is a flat list of nested members, not a path/provenance report.
Hidden membership additionally requires Member.Read.Hidden AND sufficient user
rights; this script never adds that permission. Owners may be unavailable for
Exchange-created, distribution or on-premises-synced groups. Known unsupported
owner sources are labeled unavailable rather than reported as having no owners.
Optional CSV: timestamped new file in your chosen filesystem directory, formula-like cells neutralized, [null]
marks unavailable values. All requested calls must finish before output/export.
.EXAMPLE
./Get-SwpEntraGroupReport.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -GroupId '22222222-2222-4222-8222-222222222222' -Membership Direct
.NOTES
Choose a protected export directory; filesystem paths may be network-mounted or
cloud-synced. CSV contains directory data and must be handled accordingly.
MaxRetries bounds this script's loop; the Graph SDK may apply its own retry policy.
Sources: learn.microsoft.com/graph/api/group-list-members and group-list-owners
(v1.0). No beta or complete-inventory claim. Live-tenant validation pending.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][guid]$GroupId,
    [ValidateSet('Direct','Transitive')][string]$Membership = 'Direct',

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

if ($TenantId -eq [guid]::Empty -or $GroupId -eq [guid]::Empty) { throw 'TenantId and GroupId cannot be empty.' }
$script:SwpRequiredScopes = @('GroupMember.Read.All')
$context = Assert-SwpContext $script:SwpRequiredScopes
$script:SwpInitialAccount = [string](Get-SwpValue $context 'Account')
$script:SwpInitialClientId = [string](Get-SwpValue $context 'ClientId')
if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) { throw 'Invoke-MgGraphRequest is required from Microsoft.Graph.Authentication.' }
$base = 'https://graph.microsoft.com/v1.0/groups/' + $GroupId.ToString()
$group = Invoke-SwpGet ($base + '?$select=id,displayName,groupTypes,mailEnabled,securityEnabled,onPremisesSyncEnabled,visibility')
if ([string](Get-SwpValue $group 'id') -ine $GroupId.ToString()) { throw 'Group response ID does not match the requested group.' }
if ([string](Get-SwpValue $group 'visibility') -ieq 'HiddenMembership' -and 'Member.Read.Hidden' -notin (Get-SwpValue $context 'Scopes')) { throw 'Hidden membership needs separately approved Member.Read.Hidden and appropriate user rights. No membership report was produced.' }
$groupTypes = Get-SwpValue $group 'groupTypes'
$ownersUnavailable = (Get-SwpValue $group 'onPremisesSyncEnabled') -eq $true -or ((Get-SwpValue $group 'mailEnabled') -eq $true -and 'Unified' -notin $groupTypes)
$ownerStatus = 'Visible v1.0 owners only; service-principal owners can be omitted. Exchange-created sources may not support this relationship.'
$ownerRaw = @()
if ($ownersUnavailable) { $ownerStatus = 'Unavailable for this on-premises-synced or non-Unified mail-enabled source; an empty Owners array is NOT evidence of no owners.' }
else { $ownerRaw = Get-SwpPages ($base + '/owners?$top=100') }
$relation = if ($Membership -eq 'Direct') { 'members' } else { 'transitiveMembers' }
$memberRaw = Get-SwpPages ($base + '/' + $relation + '?$top=100')
function ConvertTo-SwpDirectoryObject {
    param($Object)
    $name = Get-SwpValue $Object 'displayName'
    $type = Get-SwpValue $Object '@odata.type'
    [pscustomobject]@{ Id=Get-SwpValue $Object 'id'; ObjectType=if ($type) { $type } else { '[unknown type]' }; DisplayName=$name; UserPrincipalName=Get-SwpValue $Object 'userPrincipalName'; Mail=Get-SwpValue $Object 'mail'; Visibility=if ($null -eq $name) { 'Name unavailable; possibly permission-limited' } else { 'Returned fields only' } }
}
$owners = @($ownerRaw | ForEach-Object { ConvertTo-SwpDirectoryObject $_ })
$members = @($memberRaw | ForEach-Object { ConvertTo-SwpDirectoryObject $_ })
$coverage = 'Coverage-limited visible v1.0 results, not a complete inventory. Owners can omit service principals; Direct members also have a documented service-principal omission. Transitive results are flat, not membership paths. Hidden/other object fields depend on permissions.'
$retrieved = [DateTime]::UtcNow.ToString('o')
$rows = [System.Collections.Generic.List[object]]::new()
[void]$rows.Add([pscustomobject]@{ RecordType='Summary'; TenantId=$TenantId.ToString(); GroupId=$GroupId.ToString(); MembershipMode=$Membership; RetrievedAtUtc=$retrieved; Coverage=$coverage; OwnerCoverage=$ownerStatus; Count=$members.Count; Id=$null; ObjectType=$null; DisplayName=Get-SwpValue $group 'displayName'; UserPrincipalName=$null; Mail=$null; Visibility=$null })
foreach ($kind in @('Owner','Member')) {
    $objects = if ($kind -eq 'Owner') { $owners } else { $members }
    foreach ($item in $objects) { [void]$rows.Add([pscustomobject]@{ RecordType=$kind; TenantId=$TenantId.ToString(); GroupId=$GroupId.ToString(); MembershipMode=$Membership; RetrievedAtUtc=$retrieved; Coverage=$coverage; OwnerCoverage=$ownerStatus; Count=$null; Id=$item.Id; ObjectType=$item.ObjectType; DisplayName=$item.DisplayName; UserPrincipalName=$item.UserPrincipalName; Mail=$item.Mail; Visibility=$item.Visibility }) }
}
$csvPath = Export-SwpCsv $rows.ToArray() 'ShellWePlay-EntraGroup'
[pscustomobject]@{ TenantId=$TenantId.ToString(); Account=Get-SwpValue $context 'Account'; GroupId=$GroupId.ToString(); GroupDisplayName=Get-SwpValue $group 'displayName'; MembershipMode=$Membership; RetrievedAtUtc=$retrieved; Coverage=$coverage; OwnerCoverage=$ownerStatus; OwnerCount=$owners.Count; MemberCount=$members.Count; Owners=$owners; Members=$members; CsvPath=$csvPath }
