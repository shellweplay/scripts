#requires -Version 7.0
<#
.SYNOPSIS
Connect explicitly to one Microsoft Entra tenant with delegated read scopes.
.DESCRIPTION
Requires Microsoft.Graph.Authentication, installed separately after review.
No module installation, credentials in files, profile edits, or directory writes.
An interactive delegated sign-in/consent is an intentional action when YOU run this file.
Existing wrong-tenant context is refused, not silently reused or disconnected.
Process scope limits SDK context persistence; it does not revoke browser sign-in sessions.
Scope/context checks do not prove every server-side permission or directory role.
.EXAMPLE
./Connect-SwpEntra.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -Purpose Users
.NOTES
Global cloud only. Examples use fictional IDs. No live-tenant validation is claimed.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][guid]$TenantId,
    [ValidateSet('Users','Groups','Both')][string]$Purpose = 'Users',
    [string]$ExpectedAccount,
    [switch]$UseDeviceCode
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

if ($TenantId -eq [guid]::Empty) { throw 'TenantId cannot be the empty GUID.' }
foreach ($name in @('Connect-MgGraph','Get-MgContext')) { if (-not (Get-Command $name -ErrorAction SilentlyContinue)) { throw 'Microsoft.Graph.Authentication is required; install it separately after review.' } }
$existing = Get-MgContext
if ($null -ne $existing -and [string](Get-SwpValue $existing 'TenantId') -ine $TenantId.ToString()) { throw 'Existing context is for another tenant. Run Disconnect-MgGraph explicitly, then retry with the intended tenant.' }
$required = switch ($Purpose) { 'Users' { @('User.Read.All') }; 'Groups' { @('GroupMember.Read.All') }; 'Both' { @('User.Read.All','GroupMember.Read.All') } }
$connect = @{ TenantId = $TenantId.ToString(); Scopes = $required; ContextScope = 'Process'; Environment = 'Global'; NoWelcome = $true; ErrorAction = 'Stop' }
if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
$null = Connect-MgGraph @connect
$context = Assert-SwpContext $required
$account = [string](Get-SwpValue $context 'Account')
if ($ExpectedAccount -and $account -ine $ExpectedAccount) { throw 'The signed-in account does not match ExpectedAccount. No directory query was made; disconnect and choose the intended account.' }
[pscustomobject]@{
    TenantId = $TenantId.ToString(); Account = $account; AuthType = 'Delegated'; ContextScope = 'Process'; Environment = 'Global'
    RequestedReadScopes = @($required); ContextScopes = Get-SwpValue $context 'Scopes'
    Verification = 'Context matches requested tenant, delegated account, cloud and scope requirements. Verify the displayed account; this is not an effective-access audit.'
}
