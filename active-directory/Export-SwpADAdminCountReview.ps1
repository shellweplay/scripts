<#
.SYNOPSIS
Reviews existing adminCount=1 flags for a supplied roster.
.DESCRIPTION
Compares flagged users to an explicit list of approved group DNs or SIDs.
Recursive membership can include nested groups. Missing matches are not
proof of no access: the selected list, primary group, cross-domain groups,
directory visibility and effective permissions are not fully evaluated.
This is a review cue, not a determination that a flag can be cleared.
Writes a new sensitive local CSV. Version 1.0.0.
.EXAMPLE
./Export-SwpADAdminCountReview.ps1 -Server dc01.example.test -InputPath ./accounts.csv -ApprovedGroup 'CN=Example Review Group,OU=Groups,DC=example,DC=test' -OutputPath ./review.csv
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Server,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $InputPath,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string[]] $ApprovedGroup,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $OutputPath
)
$ErrorActionPreference = 'Stop'
Import-Module ActiveDirectory

$accounts = @(Import-Csv -LiteralPath $InputPath)
if (!$accounts.Count -or $accounts[0].PSObject.Properties.Name -notcontains 'SamAccountName') {
    throw 'Provide a CSV with a SamAccountName column and at least one row.'
}
$roster = @($accounts | ForEach-Object { ([string]$_.SamAccountName).Trim() } | Where-Object { $_ } | Select-Object -Unique)
if (!$roster.Count) { throw 'The roster has no nonblank account names.' }
if (Test-Path -LiteralPath $OutputPath) { throw 'Choose a new output file.' }

# Keep only user SIDs for comparison; never export group members or the approved group list.
$memberSids = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$groupErrors = 0
foreach ($group in ($ApprovedGroup | Select-Object -Unique)) {
    try {
        if ($group -notmatch '^(?i:CN=|S-1-)') { throw 'Use a group DN or SID.' }
        Get-ADGroupMember -Identity $group -Recursive -Server $Server | ForEach-Object {
            if ($_.ObjectClass -eq 'user' -and $_.SID) { [void]$memberSids.Add([string]$_.SID) }
        }
    } catch {
        $groupErrors += 1
    }
}
if ($groupErrors) { Write-Warning 'One or more approved group lookups failed; non-matches may be inconclusive.' }
$results = foreach ($name in $roster) {
    $user = $null
    try {
        $escaped = $name.Replace('\', '\5c').Replace('*', '\2a').Replace('(', '\28').Replace(')', '\29').Replace([string][char]0, '\00')
        $users = @(Get-ADUser -LDAPFilter ('(sAMAccountName=' + $escaped + ')') -Server $Server -Properties adminCount)
        if ($users.Count -gt 1) { throw 'The queried server returned an ambiguous account match.' }
        $user = $users | Select-Object -First 1
        $status = if (!$user) { 'NotFound' }
        elseif ($user.adminCount -ne 1) { 'NotFlagged' }
        elseif (!$user.SID) { 'MembershipUnknown' }
        elseif ($memberSids.Contains([string]$user.SID)) { 'ApprovedGroupMatch' }
        elseif ($groupErrors) { 'MembershipUnknown' }
        else { 'FlaggedWithoutApprovedGroupMatch' }
    } catch { $status = 'LookupFailed' }
    $safeName = if ($name -match '^[\s\x00-\x1f]*[=+@-]|^[\t\r\n]') { "'" + $name } else { $name }
    [pscustomobject]@{ SamAccountName = $safeName; AdminCount = $user.adminCount; ReviewStatus = $status }
}
$results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8 -NoClobber
Write-Host ('Exported {0} adminCount review row(s).' -f @($results).Count)
