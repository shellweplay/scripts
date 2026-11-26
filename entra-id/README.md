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
| [Export-SwpEntraRecentAppRegistrations.ps1](Export-SwpEntraRecentAppRegistrations.ps1) | Current applications registered in a UTC lookback window, current owners, and certificate/client-secret metadata in two CSVs. | [Recent app registrations](https://shellweplay.vercel.app/scripts/entra-recent-app-registrations) |
| [Export-SwpEntraPendingGuestInvitations.ps1](Export-SwpEntraPendingGuestInvitations.ps1) | Current guests pending acceptance, with account/state dates and optional account age. | [Pending guest invitations](https://shellweplay.vercel.app/scripts/entra-pending-guest-invitations) |

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

## November 25, 2026 reports — lab validation pending

Use an existing protected Reports folder in a dedicated PowerShell 7 process. Example GUID and operator domain mylab.dev are fictional. Replace them locally; review approved consent before connecting. The two reports never connect themselves.

```powershell
Connect-MgGraph -TenantId '11111111-1111-4111-8111-111111111111' -Scopes 'Application.Read.All','User.ReadBasic.All' -ContextScope Process
./entra-id/Export-SwpEntraRecentAppRegistrations.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -DaysBack 30 -CsvDirectory './Reports'
Disconnect-MgGraph
Connect-MgGraph -TenantId '11111111-1111-4111-8111-111111111111' -Scopes 'User.Read.All' -ContextScope Process
./entra-id/Export-SwpEntraPendingGuestInvitations.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -MinimumAccountAgeDays 30 -CsvDirectory './Reports'
Disconnect-MgGraph
```

The app report prompts when DaysBack is omitted. It scans current applications, not deleted objects, and includes derived types. Owners are current metadata, not creators. A failed owner lookup does not count as zero; known IDs/types survive a later-page failure. The guest report selects Guest objects and checks PendingAcceptance, not invitation send time or proof of never redeeming. MinimumAccountAgeDays defaults to zero, including unknown creation ages. Both reports follow all returned pages within explicit bounds; a failed main collection aborts. CSV uses UTC date, no overwrite, [null] for unavailable cells, and formula protection. Neither report exports secret values or raw certificate/key material; neither queries audit logs or sign-ins. Graph visibility and eventual consistency still apply; no additional license-dependent API is requested. See embedded help for Microsoft references and parameters. Website guides are proposed separately and are not live yet.

### Recent app registrations v1.1.0 — credential metadata

This is an enhancement to the same report and filename. `DaysBack` still means
**applications created in the last X days**; it does not filter when credentials
were added or which credentials expire soon. Only matching applications' current
credential metadata appears. It is not a tenant-wide expiry inventory.

The same run writes two files in the existing `CsvDirectory`:

| File | Rows and additions |
| --- | --- |
| `Entra_AppRegistrations_yyyy-MM-dd.csv` | One row per matching app, including apps with no credentials. All original columns and owner metadata remain; appended columns are `CertificateCount`, `ClientSecretCount`, `OtherKeyCredentialCount`, `ExpiredCredentialCount`, and `CredentialMetadataStatus`. |
| `Entra_AppRegistrationCredentials_yyyy-MM-dd.csv` | One row per returned credential, joined by `ApplicationObjectId` or `ApplicationClientId`. Apps with no credentials produce no invented detail rows. |

The credential CSV columns, in order, are `AppName`, `ApplicationClientId`,
`ApplicationObjectId`, `CredentialKind`, `CredentialDescription`, `CredentialId`,
`StartDateTimeUtc`, `EndDateTimeUtc`, `DaysUntilExpiry`, `ValidityStatus`, `KeyType`,
`KeyUsage`, `CustomKeyIdentifierHex`, `IdentifierStatus`, and `ReportAsOfUtc`.

`CredentialDescription` comes from the credential's optional `displayName`.
`CredentialId` is its `keyId`. The start/end columns are **validity endpoints**,
not credential creation or upload dates. `DaysUntilExpiry` is signed fractional
days rounded to two decimals against the same frozen UTC timestamp in both files;
unusable endpoints or an invalid range leave it `[null]`. Existing CSV formula
protection also prefixes negative day text with an apostrophe.

`AsymmetricX509Cert` and `X509Cert` keys are `Certificate`; other key types are
`OtherKey`. Entries in `passwordCredentials` are `ClientSecret`. A missing key
type remains an `OtherKey` detail row with unknown certificate/other-key counts,
so it is never presented as a confirmed non-certificate inventory.

| Status field | Meaning |
| --- | --- |
| `ValidityStatus` | `Expired` when end is at/before report time; `NotYetValid` when start is later; otherwise `WithinValidityPeriod`. `UnknownDates` means an endpoint is missing/unusable. `InvalidDateRange` means start is at/after end. These describe timestamps, not whether an application works or which credential it uses. |
| `CredentialMetadataStatus` | `Complete` means both collections were returned with usable metadata. `LimitedDetails` means the collections are known but a date, ID, type, usage or malformed field is unusable. `Incomplete` means a collection/item was missing or malformed. `Unavailable` means both collections were absent/null. Optional absent descriptions/custom identifiers alone do not downgrade the status. |
| `IdentifierStatus` | Two independent tokens: `KeyId=Valid`, `KeyId=Missing`, or `KeyId=Invalid`; and `CustomKeyIdentifier=Decoded`, `CustomKeyIdentifier=Absent`, `CustomKeyIdentifier=Invalid`, or `CustomKeyIdentifier=NotApplicable`, separated by a semicolon. Client-secret custom identifiers are ignored and marked `NotApplicable`. |

Successfully returned, usable collections can produce zero counts; only explicit
empty arrays establish that a collection has no credentials. An unavailable or
malformed collection leaves its affected counts `[null]`; usable rows from a
partially malformed collection are retained. An unknown validity interval or
incomplete collection leaves `ExpiredCredentialCount` `[null]` rather than
claiming zero expired credentials. Known credential counts remain usable even
when an optional description is absent or another detail is unavailable.

`CustomKeyIdentifierHex` converts a key credential's returned base64/binary custom
identifier to hex. It is `[null]` for client secrets because Microsoft marks that
password property "Do not use." It is **not a verified certificate thumbprint**: the identifier can be
supplied independently. `secretText`, secret hints, raw certificates and private
key material are never exported. Existing secret values cannot be read back with
this report. No additional permission beyond the original delegated
`Application.Read.All` and `User.ReadBasic.All` is requested. The selected
`keyCredentials` property has a tenant-level Graph throttle (currently 150
requests/minute); bounded GET retries honor `Retry-After` and stop if the budget
is exceeded.

Zero matching apps produce two header-only files. Both filenames are protected
by atomic create-new writes. If the companion export fails, the script attempts
to remove files created by that run, reports any cleanup failure explicitly, and
does not print a successful pair. It preserves existing files; use another
existing output directory when repeating a report on the same UTC date.

The report remains **lab validation pending**. Offline mocks exercise formatting,
null/empty distinctions, date boundaries and failures, but do not validate live
tenant permissions or service behavior. Microsoft's
[application list](https://learn.microsoft.com/en-us/graph/api/application-list?view=graph-rest-1.0),
[keyCredential](https://learn.microsoft.com/en-us/graph/api/resources/keycredential?view=graph-rest-1.0),
and [passwordCredential](https://learn.microsoft.com/en-us/graph/api/resources/passwordcredential?view=graph-rest-1.0)
documentation describe the source fields and limits.

## App credential expiration across current registrations (v1.0.0)

[Export-SwpEntraAppCredentialExpiry.ps1](Export-SwpEntraAppCredentialExpiry.ps1) is distinct from the recent registration report: an older app can have a credential expiring tomorrow. It follows all visible application list pages, uses only delegated Application.Read.All, includes certificates/client secrets/other keys, and does not read owners, sign-ins, audit logs or enterprise application service principals. Lab validation pending.

```powershell
Connect-MgGraph -TenantId '11111111-1111-4111-8111-111111111111' -Scopes 'Application.Read.All' -ContextScope Process
./entra-id/Export-SwpEntraAppCredentialExpiry.ps1 -TenantId '11111111-1111-4111-8111-111111111111' -DaysAhead 30 -IncludeExpired -CsvDirectory './Reports'
```

The fictional mylab.dev tenant GUID must be replaced locally. Default: future 30-day validity-end window; IncludeExpired also includes already-expired credentials. Unknown/invalid dates are retained regardless of window. CoverageWarning rows distinguish unavailable, incomplete and limited metadata; they are not credentials. One dated Entra_AppCredentialExpiry_yyyy-MM-dd.csv, UTF-8, never overwrites, header-only if no rows. See complete help for columns and bounds. The Graph collection endpoint returns metadata; raw certificate data is only available through a single-object selection, which is not used. Secret hints and secret values are never exported. Credential validity is not application usage, health, or upload time; a custom identifier is not a verified thumbprint. Current application objects only, including derived types; not an atomic snapshot or a full workload identity inventory. Validate known expiry dates, permissions and coverage warnings in a lab.

References: [List applications](https://learn.microsoft.com/en-us/graph/api/application-list?view=graph-rest-1.0), [keyCredential](https://learn.microsoft.com/en-us/graph/api/resources/keycredential?view=graph-rest-1.0), [passwordCredential](https://learn.microsoft.com/en-us/graph/api/resources/passwordcredential?view=graph-rest-1.0).
