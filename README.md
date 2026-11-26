# ShellWePlay Scripts

Documented administration scripts from [ShellWePlay](https://shellweplay.vercel.app/).
This repository holds the script source. The website provides guides, examples,
and matching copy/download controls.

## Script collection

| Script | Purpose | Guide | Validation status |
| --- | --- | --- | --- |
| [Export-SwpADServiceAccounts.ps1](active-directory/Export-SwpADServiceAccounts.ps1) | Export user-based service accounts from an OU, including account dates, enabled state, location, and literal `owner` details. | [Service-account report](https://shellweplay.vercel.app/scripts/export-ad-service-accounts) | Lab validation pending |
| [Get-ADGroupTypeCount.ps1](active-directory/Get-ADGroupTypeCount.ps1) | Count security and distribution groups with one paged query; optionally export the totals. | [AD group counts](https://shellweplay.vercel.app/scripts/count-ad-groups) | Live AD and compatibility validation pending |
| [Connect-SwpEntra.ps1](entra-id/Connect-SwpEntra.ps1) | Connect explicitly to the intended Entra tenant with selected delegated read scopes. | [Connect to Entra](https://shellweplay.vercel.app/scripts/entra-connect) | Preview; live tenant validation pending |
| [Get-SwpEntraUser.ps1](entra-id/Get-SwpEntraUser.ps1) | Inspect one selected Entra user or an explicit bounded filter. | [Entra user details](https://shellweplay.vercel.app/scripts/entra-user) | Preview; live tenant validation pending |
| [Get-SwpEntraGroupReport.ps1](entra-id/Get-SwpEntraGroupReport.ps1) | Report one group's visible owners and direct or flat transitive members. | [Entra group report](https://shellweplay.vercel.app/scripts/entra-group-report) | Preview; live tenant validation pending |
| [ShellWePlay.Prompt.ps1](powershell/ShellWePlay.Prompt.ps1) | Add a local date/time prompt to the current PowerShell session. | [Shell Lab](https://shellweplay.vercel.app/shell-lab) | Interactive Windows/terminal validation pending |

Category documentation: [Active Directory](active-directory/README.md),
[Entra ID](entra-id/README.md), and [PowerShell customization](powershell/README.md).

## Read before running

Each script contains help and comments. Review its source, required modules,
permissions, parameters, and output behavior before running it in your own lab.
Examples use lab names and fictional identifiers. Replace those values locally.

- AD and Entra reports make no directory changes, but can perform substantial reads and write local CSV files.
- The Entra connection helper performs an explicit delegated sign-in when you run it; the report scripts require that connection separately.
- The prompt replaces the prompt function in the current session. It does not install a theme or edit a profile.
- Keep real identities, credentials, environment configuration, and generated reports out of the repository.

For example, from the repository root:

```powershell
Get-Help .\active-directory\Export-SwpADServiceAccounts.ps1 -Full
```

## What validation means here

The initial six scripts are unchanged copies of the existing website downloads.
Existing website CI has exercised the Entra tools with mocked Graph responses and
the prompt in an isolated PowerShell process. Those checks do not establish live
tenant, Active Directory, Windows PowerShell, or interactive terminal validation.
The service-account report has received static review; its PowerShell runtime and
live AD testing remain pending. The status beside each script stays explicit until
real environment results are recorded.

New categories will be added with their first published script.

## November 2026 launch reports

Ten on-prem AD and ten Entra identity reports focused on operational data quality and lifecycle checks. These are local tools, not a hosted directory connection. Protect output files.

| Script | Purpose | Website guide | Validation |
| --- | --- | --- | --- |
| [`Export-SwpADProfileCompleteness.ps1`](active-directory/Export-SwpADProfileCompleteness.ps1) | Count populated department, office, title, and company fields for visible user objects in an OU. | [Guide](https://shellweplay.vercel.app/scripts/ad-profile-completeness) | Lab validation pending |
| [`Export-SwpADGroupDocumentationCoverage.ps1`](active-directory/Export-SwpADGroupDocumentationCoverage.ps1) | Count how many visible groups have a description and a managed-by reference, without exposing either field. | [Guide](https://shellweplay.vercel.app/scripts/ad-group-documentation-coverage) | Lab validation pending |
| [`Export-SwpADAccountStateTotals.ps1`](active-directory/Export-SwpADAccountStateTotals.ps1) | Summarize enabled, disabled, and unknown states for visible user objects in an OU. | [Guide](https://shellweplay.vercel.app/scripts/ad-account-state-totals) | Lab validation pending |
| [`Export-SwpADAccountCreationTrend.ps1`](active-directory/Export-SwpADAccountCreationTrend.ps1) | Count currently visible accounts by UTC creation month within a configurable recent window. | [Guide](https://shellweplay.vercel.app/scripts/ad-account-creation-trend) | Lab validation pending |
| [`Export-SwpADAccountExpiryBuckets.ps1`](active-directory/Export-SwpADAccountExpiryBuckets.ps1) | Bucket visible user objects by configured AD account-expiration time, without exposing identities. | [Guide](https://shellweplay.vercel.app/scripts/ad-account-expiry-buckets) | Lab validation pending |
| [`Export-SwpADDepartmentDistribution.ps1`](active-directory/Export-SwpADDepartmentDistribution.ps1) | Count visible users mapped to administrator-supplied department labels; all other values stay aggregated. | [Guide](https://shellweplay.vercel.app/scripts/ad-department-distribution) | Lab validation pending |
| [`Export-SwpADOfficeDistribution.ps1`](active-directory/Export-SwpADOfficeDistribution.ps1) | Count visible users mapped to administrator-supplied office labels; all other values stay aggregated. | [Guide](https://shellweplay.vercel.app/scripts/ad-office-distribution) | Lab validation pending |
| [`Export-SwpADDuplicateMailCounts.ps1`](active-directory/Export-SwpADDuplicateMailCounts.ps1) | Count duplicates in the visible mail field without revealing any values or account identities. | [Guide](https://shellweplay.vercel.app/scripts/ad-duplicate-mail-counts) | Lab validation pending |
| [`Export-SwpADDuplicateEmployeeIdCounts.ps1`](active-directory/Export-SwpADDuplicateEmployeeIdCounts.ps1) | Count duplicates in the visible employee ID field without revealing any values or account identities. | [Guide](https://shellweplay.vercel.app/scripts/ad-duplicate-employee-id-counts) | Lab validation pending |
| [`Export-SwpADEmployeeTypeCoverage.ps1`](active-directory/Export-SwpADEmployeeTypeCoverage.ps1) | Measure how many user objects have a known, other, or unspecified employee classification against an approved list. | [Guide](https://shellweplay.vercel.app/scripts/ad-employee-type-coverage) | Lab validation pending |
| [`Export-SwpEntraRosterAccountState.ps1`](entra-id/Export-SwpEntraRosterAccountState.ps1) | Check the reported enabled state for a selected onboarding, transfer, or offboarding roster. | [Guide](https://shellweplay.vercel.app/scripts/entra-roster-account-state) | Lab validation pending |
| [`Export-SwpEntraOrganizationProfileGaps.ps1`](entra-id/Export-SwpEntraOrganizationProfileGaps.ps1) | Find missing department, job title, and office fields on a known roster without copying field values. | [Guide](https://shellweplay.vercel.app/scripts/entra-organization-profile-gaps) | Lab validation pending |
| [`Export-SwpEntraPrimaryMailReview.ps1`](entra-id/Export-SwpEntraPrimaryMailReview.ps1) | Review whether the Graph mail field is populated and how it compares with the sign-in name, without exporting either value. | [Guide](https://shellweplay.vercel.app/scripts/entra-primary-mail-review) | Lab validation pending |
| [`Export-SwpEntraUsageLocationReadiness.ps1`](entra-id/Export-SwpEntraUsageLocationReadiness.ps1) | Check whether selected users have a usage location in the expected two-letter form before a licensing review. | [Guide](https://shellweplay.vercel.app/scripts/entra-usage-location-readiness) | Lab validation pending |
| [`Export-SwpEntraLicenseAssignmentPresence.ps1`](entra-id/Export-SwpEntraLicenseAssignmentPresence.ps1) | See whether Graph reports any assigned license SKU for each selected user before investigating entitlement or onboarding gaps. | [Guide](https://shellweplay.vercel.app/scripts/entra-license-assignment-presence) | Lab validation pending |
| [`Export-SwpEntraHybridSyncFreshness.ps1`](entra-id/Export-SwpEntraHybridSyncFreshness.ps1) | Review the last recorded sync timestamp for selected hybrid users against a configurable elapsed-day threshold. | [Guide](https://shellweplay.vercel.app/scripts/entra-hybrid-sync-freshness) | Lab validation pending |
| [`Export-SwpEntraPreferredLanguageReadiness.ps1`](entra-id/Export-SwpEntraPreferredLanguageReadiness.ps1) | Check if the preferredLanguage field is populated before planning communications to a selected roster. | [Guide](https://shellweplay.vercel.app/scripts/entra-preferred-language-readiness) | Lab validation pending |
| [`Export-SwpEntraGroupDocumentationGaps.ps1`](entra-id/Export-SwpEntraGroupDocumentationGaps.ps1) | Check selected groups for missing descriptions or names without exposing the contents of either. | [Guide](https://shellweplay.vercel.app/scripts/entra-group-documentation-gaps) | Lab validation pending |
| [`Export-SwpEntraGroupMailReadiness.ps1`](entra-id/Export-SwpEntraGroupMailReadiness.ps1) | Cross-check mailEnabled and mail field presence for selected groups before handing off an email-address review. | [Guide](https://shellweplay.vercel.app/scripts/entra-group-mail-readiness) | Lab validation pending |
| [`Export-SwpEntraGroupExpirationReview.ps1`](entra-id/Export-SwpEntraGroupExpirationReview.ps1) | Check the policy-reported expiry of selected Microsoft 365 groups against a configurable review window. | [Guide](https://shellweplay.vercel.app/scripts/entra-group-expiration-review) | Lab validation pending |

Offline verification: run `pwsh -NoProfile -File ./tests/Invoke-SwpADReportFixtureTests.ps1` and `pwsh -NoProfile -File ./tests/Invoke-SwpEntraReportFixtureTests.ps1` from the repository root. The fixtures do not connect to real AD or Graph. Windows RSAT, Windows PowerShell 5.1, and live directory validation remain pending.

Keep filenames stable. Extend an existing script when the operational question stays the same; use a new filename only for a materially different question or scope. Review complete source and check results before running in a live environment.
