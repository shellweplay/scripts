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
| [Export-SwpEntraRecentAppRegistrations.ps1](entra-id/Export-SwpEntraRecentAppRegistrations.ps1) | Current applications registered in a UTC lookback window, with current owners and certificate/client-secret metadata in a companion CSV. | [Recent app registrations](https://shellweplay.vercel.app/scripts/entra-recent-app-registrations) | v1.1.0; lab validation pending |
| [Export-SwpEntraPendingGuestInvitations.ps1](entra-id/Export-SwpEntraPendingGuestInvitations.ps1) | Current guests pending acceptance, with account/state dates and optional account age. | [Pending guest invitations](https://shellweplay.vercel.app/scripts/entra-pending-guest-invitations) | Lab validation pending |
| [ShellWePlay.Prompt.ps1](powershell/ShellWePlay.Prompt.ps1) | Add a local date/time prompt to the current PowerShell session. | [Shell Lab](https://shellweplay.vercel.app/shell-lab) | Interactive Windows/terminal validation pending |

Category documentation: [Active Directory](active-directory/README.md),
[Entra ID](entra-id/README.md), and [PowerShell customization](powershell/README.md).

## Improving an existing script

Keep one maintained script for each operational question. Add related fields,
filters, or output details to that script when its purpose and query scope remain
the same. Create a separate script when the question or scope changes materially.
For example, credential details belong in the recent app-registration report;
finding credentials that expire soon across all applications would be a distinct
report, because app creation age must not exclude older applications.

Keep filenames stable. Record the version and revision notes in the complete
script and use Git history to retain earlier revisions; do not create `_v2`,
`_new`, or `_final` copies. Use patch versions for fixes, minor versions for
compatible additions, and major versions for breaking parameter, output, or
behavior changes with migration notes. Document any added CSV columns or files.

Continue the existing PR while it is open. After it is merged, use a new focused
PR against the same script. Each revision includes updated help, examples,
meaningful offline checks, and the matching website guide and exact source pin.
Keep lab-validation status explicit; passing mocked checks does not validate a
script against a live directory.

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
