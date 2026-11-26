# Entra ID scripts

These PowerShell 7 tools use Microsoft Graph to inspect selected directory objects.
They are review previews: **live tenant validation is pending**. Existing website
CI exercises authentication/context checks, paging, failures, and exports with
mocked Graph responses; it never connects to a tenant.

| Script | Purpose | Website guide |
| --- | --- | --- |
| [Connect-SwpEntra.ps1](Connect-SwpEntra.ps1) | Explicit delegated sign-in to the expected tenant. | [Connection guide](https://shellweplay.vercel.app/scripts/entra-connect) |
| [Get-SwpEntraUser.ps1](Get-SwpEntraUser.ps1) | Selected user properties or an explicitly bounded filter. | [User guide](https://shellweplay.vercel.app/scripts/entra-user) |
| [Get-SwpEntraGroupReport.ps1](Get-SwpEntraGroupReport.ps1) | Visible owners and direct or transitive members of one group. | [Group guide](https://shellweplay.vercel.app/scripts/entra-group-report) |

## Prerequisites

Use an approved test tenant/account, PowerShell 7, and the
`Microsoft.Graph.Authentication` module installed through your approved process.
These scripts do not install modules or grant consent. They support the Global
Microsoft Graph cloud and require a process-scoped delegated context for the
explicit tenant ID.

The connection helper requests `User.Read.All` for `-Purpose Users` or
`GroupMember.Read.All` for `-Purpose Groups`; `Both` is an explicit option.
Use the permissions approved for your task. The reports validate the tenant,
context, and required scopes before querying. The connection helper also checks
the account when you supply `-ExpectedAccount`; otherwise, verify its displayed
account before running a report.

## Examples

Run from the repository root. The GUIDs and user name below are fictional lab
placeholders; replace them locally with the identifiers for your approved test.

```powershell
# Connect for a selected user read, then inspect that user.
./entra-id/Connect-SwpEntra.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -Purpose Users
./entra-id/Get-SwpEntraUser.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -UserId 'alex@mylab.dev'

# Separately connect for group reads, then inspect direct members and visible owners.
./entra-id/Connect-SwpEntra.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -Purpose Groups
./entra-id/Get-SwpEntraGroupReport.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -GroupId '22222222-2222-4222-8222-222222222222' -Membership Direct

# End the dedicated Graph session when finished.
Disconnect-MgGraph
```

`Get-SwpEntraUser.ps1` also accepts an explicit `-Filter` instead of `-UserId`.
Both report scripts bound page/record counts and retries. Optional
`-CsvDirectory` must point to an **existing filesystem directory**. Exports use new
timestamped filenames and preserve existing files. They can contain directory
data; select appropriate protected storage. The website never receives reports.

## Interpreting the results

The group report returns visible relationships, not a complete access inventory.
Graph v1.0 member/owner endpoints have service-principal and group-source
limitations. Hidden membership needs additional permissions and account rights;
the script does not request them automatically. Missing names can indicate limited
visibility, not deleted accounts. Transitive membership is a flat list rather than
the paths through nested groups. These reports do not inspect directory roles,
PIM eligibility, Azure resource RBAC, or all effective access.

See the embedded help and website guides for the complete parameters and limits,
plus Microsoft's [Graph authentication documentation](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands),
[member endpoint documentation](https://learn.microsoft.com/en-us/graph/api/group-list-members?view=graph-rest-1.0),
and [owner endpoint documentation](https://learn.microsoft.com/en-us/graph/api/group-list-owners?view=graph-rest-1.0).
## November 2026 launch reports

Each Entra report requires explicitly approved object IDs and a verified delegated Graph context in the specified tenant. Reports are read-only and export selected status fields to local CSV.

| Script | Purpose | Website guide |
| --- | --- | --- |
| [`Export-SwpEntraRosterAccountState.ps1`](Export-SwpEntraRosterAccountState.ps1) | Check the reported enabled state for a selected onboarding, transfer, or offboarding roster. | [Guide](https://shellweplay.vercel.app/scripts/entra-roster-account-state) |
| [`Export-SwpEntraOrganizationProfileGaps.ps1`](Export-SwpEntraOrganizationProfileGaps.ps1) | Find missing department, job title, and office fields on a known roster without copying field values. | [Guide](https://shellweplay.vercel.app/scripts/entra-organization-profile-gaps) |
| [`Export-SwpEntraPrimaryMailReview.ps1`](Export-SwpEntraPrimaryMailReview.ps1) | Review whether the Graph mail field is populated and how it compares with the sign-in name, without exporting either value. | [Guide](https://shellweplay.vercel.app/scripts/entra-primary-mail-review) |
| [`Export-SwpEntraUsageLocationReadiness.ps1`](Export-SwpEntraUsageLocationReadiness.ps1) | Check whether selected users have a usage location in the expected two-letter form before a licensing review. | [Guide](https://shellweplay.vercel.app/scripts/entra-usage-location-readiness) |
| [`Export-SwpEntraLicenseAssignmentPresence.ps1`](Export-SwpEntraLicenseAssignmentPresence.ps1) | See whether Graph reports any assigned license SKU for each selected user before investigating entitlement or onboarding gaps. | [Guide](https://shellweplay.vercel.app/scripts/entra-license-assignment-presence) |
| [`Export-SwpEntraHybridSyncFreshness.ps1`](Export-SwpEntraHybridSyncFreshness.ps1) | Review the last recorded sync timestamp for selected hybrid users against a configurable elapsed-day threshold. | [Guide](https://shellweplay.vercel.app/scripts/entra-hybrid-sync-freshness) |
| [`Export-SwpEntraPreferredLanguageReadiness.ps1`](Export-SwpEntraPreferredLanguageReadiness.ps1) | Check if the preferredLanguage field is populated before planning communications to a selected roster. | [Guide](https://shellweplay.vercel.app/scripts/entra-preferred-language-readiness) |
| [`Export-SwpEntraGroupDocumentationGaps.ps1`](Export-SwpEntraGroupDocumentationGaps.ps1) | Check selected groups for missing descriptions or names without exposing the contents of either. | [Guide](https://shellweplay.vercel.app/scripts/entra-group-documentation-gaps) |
| [`Export-SwpEntraGroupMailReadiness.ps1`](Export-SwpEntraGroupMailReadiness.ps1) | Cross-check mailEnabled and mail field presence for selected groups before handing off an email-address review. | [Guide](https://shellweplay.vercel.app/scripts/entra-group-mail-readiness) |
| [`Export-SwpEntraGroupExpirationReview.ps1`](Export-SwpEntraGroupExpirationReview.ps1) | Check the policy-reported expiry of selected Microsoft 365 groups against a configurable review window. | [Guide](https://shellweplay.vercel.app/scripts/entra-group-expiration-review) |

Live directory validation is pending. Do not infer a complete inventory or use these reports alone to change identity state.
