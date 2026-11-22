#Requires -Version 5.1
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
Exports service-account users and their owner information to a dated CSV.

.DESCRIPTION
Reports every AD user in the specified Service Accounts OU, including child
OUs by default. OU membership defines the report scope; the script does not
check whether an account actually runs a service. Includes enabled and
disabled users. Managed service accounts (sMSAs/gMSAs) are outside this scope.

Reads the literal LDAP attribute 'owner' and resolves its distinguished name
to the referenced object's displayName. This is not the Security tab owner,
managedBy, or proof of who created the account. Use of 'owner' on user objects
depends on your schema; a missing or unreadable value is reported as NotReturned.

Uses one domain controller and caches owner lookups. Makes no AD changes.
Creates a UTF-8 CSV only after the account query and report are complete.

.PARAMETER SearchBase
Distinguished name of the Service Accounts OU.

.PARAMETER Server
AD domain DNS name or domain controller hostname. Defaults to mylab.dev.
The responding controller is used for all subsequent directory queries.

.PARAMETER OutputDirectory
Folder for the CSV. Defaults to a Reports folder in the current directory.
The folder is created if necessary.

.PARAMETER SearchScope
Subtree includes child OUs. OneLevel includes only users directly in the OU.

.EXAMPLE
.\Export-SwpADServiceAccounts.ps1

.EXAMPLE
.\Export-SwpADServiceAccounts.ps1 -OutputDirectory 'C:\Reports' -Verbose

.EXAMPLE
.\Export-SwpADServiceAccounts.ps1 -Server 'dc01.mylab.dev' -SearchScope OneLevel

.NOTES
Author: ShellWePlay
Version: 1.0.0
Requires: Windows PowerShell 5.1, the ActiveDirectory module (RSAT), AD network
access, directory read permissions, and permission to write the output folder.
Uses the current Windows credentials. No administrator role is inherently
required if those permissions are already granted.

CSV columns: SamAccountName, CreatedDate, Enabled, OU, Owner,
OwnerDisplayName, OwnerResolutionStatus.
CreatedDate is UTC in ISO 8601 format; the filename uses the local run date.
OwnerResolutionStatus: Resolved, NotReturned, DisplayNameNotReturned, Unresolved.
An owner in another domain may not resolve through the selected controller.
If no users are found, a warning is shown and no CSV is created.
Same-day reruns add a time suffix if the date-only filename already exists.
Validation status: Live Active Directory validation is pending.

.LINK
https://learn.microsoft.com/en-us/windows/win32/adschema/a-owner
#>

[CmdletBinding()]
param (
    [ValidateNotNullOrEmpty()]
    [string]$SearchBase = 'OU=Service Accounts,DC=mylab,DC=dev',

    [ValidateNotNullOrEmpty()]
    [string]$Server = 'mylab.dev',

    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory = '.\Reports',

    [ValidateSet('Subtree', 'OneLevel')]
    [string]$SearchScope = 'Subtree'
)

# Stop on query/export errors instead of presenting an incomplete report.
$ErrorActionPreference = 'Stop'

# Resolve one controller so all reads use the same directory replica.
$rootDse = Get-ADRootDSE -Server $Server -ErrorAction Stop
$directoryServer = [string]$rootDse.dnsHostName
if ([string]::IsNullOrWhiteSpace($directoryServer)) {
    throw "Could not determine a domain controller from '$Server'."
}

Write-Verbose "Domain controller: $directoryServer"
Write-Verbose "Search base: $SearchBase; search scope: $SearchScope"

# SamAccountName, Enabled, and DistinguishedName are default user properties.
# Request only the extra fields needed; do not retrieve every AD attribute.
$queryParameters = @{
    Filter         = '*'
    SearchBase     = $SearchBase
    SearchScope    = $SearchScope
    Server         = $directoryServer
    Properties     = @('whenCreated', 'owner')
    ResultPageSize = 500
    ResultSetSize  = $null
    ErrorAction    = 'Stop'
}
$accounts = @(Get-ADUser @queryParameters)

if ($accounts.Count -eq 0) {
    Write-Warning "No user accounts were found in '$SearchBase'. No CSV was created."
    return
}

# A case-insensitive cache avoids querying the same owner for every account.
$ownerCache = @{}

$report = @(
    foreach ($account in $accounts) {
        $ownerDn = [string]$account.owner
        $ownerDisplayName = ''
        $ownerStatus = 'NotReturned'

        if (-not [string]::IsNullOrWhiteSpace($ownerDn)) {
            if (-not $ownerCache.ContainsKey($ownerDn)) {
                try {
                    # Get-ADObject also supports owners that are groups or other
                    # object types; it does not assume the owner is a user.
                    $ownerObject = Get-ADObject -Identity $ownerDn -Server $directoryServer `
                        -Properties displayName -ErrorAction Stop

                    $displayName = [string]$ownerObject.displayName
                    $status = 'Resolved'
                    if ([string]::IsNullOrWhiteSpace($displayName)) {
                        $status = 'DisplayNameNotReturned'
                    }

                    $ownerCache[$ownerDn] = [pscustomobject]@{
                        DisplayName = $displayName
                        Status      = $status
                    }
                }
                catch {
                    # Preserve the account and raw owner even if this lookup fails.
                    # Cache the failure too, so it is reported once per owner.
                    $ownerCache[$ownerDn] = [pscustomobject]@{
                        DisplayName = ''
                        Status      = 'Unresolved'
                    }
                    Write-Warning "Could not resolve owner '$ownerDn': $($_.Exception.Message)"
                }
            }

            $ownerDisplayName = $ownerCache[$ownerDn].DisplayName
            $ownerStatus = $ownerCache[$ownerDn].Status
        }

        # Remove the first RDN while respecting escaped commas/backslashes.
        # This preserves the actual parent OU, including nested OU paths.
        $parentOu = $account.DistinguishedName -replace '^(?:\\.|[^\\,])+,', ''

        $createdDate = ''
        if ($null -ne $account.whenCreated) {
            $createdDate = ([datetime]$account.whenCreated).ToUniversalTime().ToString(
                "yyyy-MM-ddTHH:mm:ss'Z'", [System.Globalization.CultureInfo]::InvariantCulture
            )
        }

        [pscustomobject][ordered]@{
            SamAccountName        = $account.SamAccountName
            CreatedDate           = $createdDate
            Enabled               = $account.Enabled
            OU                    = $parentOu
            Owner                 = $ownerDn
            OwnerDisplayName      = $ownerDisplayName
            OwnerResolutionStatus = $ownerStatus
        }
    }
)

# Resolve the destination as a filesystem path and create it if necessary.
$outputFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
$null = [System.IO.Directory]::CreateDirectory($outputFolder)

$dateStamp = Get-Date -Format 'yyyy-MM-dd'
$csvPath = Join-Path -Path $outputFolder -ChildPath "ServiceAccounts_$dateStamp.csv"

# Keep previous reports. NoClobber also protects against a filename collision.
if (Test-Path -LiteralPath $csvPath) {
    $timeStamp = Get-Date -Format 'HHmmssfff'
    $csvPath = Join-Path -Path $outputFolder -ChildPath "ServiceAccounts_${dateStamp}_$timeStamp.csv"
}

$report | Sort-Object -Property SamAccountName |
    Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8 -NoClobber -ErrorAction Stop

$missingOwnerCount = @($report | Where-Object { $_.OwnerResolutionStatus -eq 'NotReturned' }).Count
if ($missingOwnerCount -gt 0) {
    Write-Warning "$missingOwnerCount account(s) returned no owner value. It may be unset, inapplicable, or unreadable."
}

Write-Host "Exported $($report.Count) service-account users."
Write-Host "CSV: $csvPath"

# Return the file object so callers can inspect or pass on the finished CSV.
Get-Item -LiteralPath $csvPath
