# Active Directory scripts

| Script | Purpose | Website guide |
| --- | --- | --- |
| [Export-SwpADServiceAccounts.ps1](Export-SwpADServiceAccounts.ps1) | Export user-based service accounts and literal LDAP owner details. | [Service-account report](https://shellweplay.vercel.app/scripts/export-ad-service-accounts) |
| [Get-ADGroupTypeCount.ps1](Get-ADGroupTypeCount.ps1) | Count security and distribution groups in a selected scope. | [AD group counts](https://shellweplay.vercel.app/scripts/count-ad-groups) |

## Export AD service-account users

[Export-SwpADServiceAccounts.ps1](Export-SwpADServiceAccounts.ps1) exports enabled and disabled AD user accounts from a designated Service Accounts OU. By default, it includes child OUs.

**Validation: live Active Directory testing is pending.** The source has received static review, but PowerShell runtime execution has not yet been verified.

### Scope and requirements

OU placement defines which accounts are reported. The script does not detect whether an account actually runs a service. It covers user-based accounts; standalone and group managed service accounts (sMSAs/gMSAs) are outside its scope.

The script is read-only in Active Directory and creates a local CSV report. It uses the current Windows credentials and requires:

- Windows PowerShell 5.1.
- The `ActiveDirectory` module, available through RSAT.
- Network access to AD and permission to read the relevant objects and attributes.
- Permission to create or write the output folder.

### Run the script

Run these examples from the repository root. All domain values below use the `mylab.dev` lab.

```powershell
# Use OU=Service Accounts,DC=mylab,DC=dev and include child OUs.
# Write the report to a Reports folder in the current directory.
.\active-directory\Export-SwpADServiceAccounts.ps1

# Specify the lab scope and output folder; show the selected controller.
.\active-directory\Export-SwpADServiceAccounts.ps1 `
    -Server 'mylab.dev' `
    -SearchBase 'OU=Service Accounts,DC=mylab,DC=dev' `
    -OutputDirectory 'C:\Reports' `
    -Verbose

# Query one controller and include only users directly in the selected OU.
.\active-directory\Export-SwpADServiceAccounts.ps1 `
    -Server 'dc01.mylab.dev' `
    -SearchScope OneLevel

# Read the complete embedded help.
Get-Help .\active-directory\Export-SwpADServiceAccounts.ps1 -Full
```

`SearchScope` accepts `Subtree` (the default) or `OneLevel`. One controller is selected for the run and used for both the account query and owner lookups. Repeated owner DNs are resolved once and cached.

### CSV fields

| Column | Meaning |
| --- | --- |
| `SamAccountName` | The account's AD logon name. |
| `CreatedDate` | The object's `whenCreated` timestamp, converted to UTC ISO 8601 format. |
| `Enabled` | Whether the user account is enabled. |
| `OU` | The immediate parent distinguished name, preserving the actual child OU path. |
| `Owner` | The literal LDAP `owner` value, as a distinguished name. |
| `OwnerDisplayName` | The referenced object's `displayName`, when returned. |
| `OwnerResolutionStatus` | The result of reading and resolving the owner, as described below. |

### What owner means

This script reads the exact LDAP attribute **`owner`**. Its use on user objects depends on the directory schema. It is separate from the owner shown on the object's **Security** tab, separate from `managedBy`, and does not establish who created the account.

A blank result can mean the attribute is unset, inapplicable, or unreadable. It is not proof that the account has no responsible owner. See Microsoft's [Owner attribute documentation](https://learn.microsoft.com/en-us/windows/win32/adschema/a-owner).

| Status | Meaning |
| --- | --- |
| `NotReturned` | No nonblank `owner` value was returned. |
| `Resolved` | The owner DN was resolved and a nonblank `displayName` was returned. |
| `DisplayNameNotReturned` | The owner object was resolved, but no nonblank `displayName` was returned. |
| `Unresolved` | The owner lookup failed on the selected controller. The raw owner DN is retained. |

An owner in another domain may not resolve through the selected controller. `Unresolved` does not establish that the referenced object has been deleted.

### Output behavior

The script creates a UTF-8 CSV named `ServiceAccounts_yyyy-MM-dd.csv`, using the local run date. If that filename already exists, it adds an `HHmmssfff` time suffix. Existing files are not overwritten.

Rows are sorted by `SamAccountName`. The account query and report preparation finish before export begins. If the initial query fails or finds no users, no CSV is created. Individual owner lookup failures retain the account row and record an `Unresolved` status.

After a successful export, the script displays the row count and file path, then returns the CSV file object to the PowerShell pipeline.

## Count AD groups

[Get-ADGroupTypeCount.ps1](Get-ADGroupTypeCount.ps1) counts security and distribution groups with one paged, subtree query. It counts each group as it arrives, without expanding membership or retaining the full set of group objects.

**Validation: live Active Directory and PowerShell compatibility testing are pending.** This preview is intended for Windows PowerShell 5.1, or PowerShell 7 on Windows with a compatible `ActiveDirectory` module. It requires read access to the selected scope and network access to a domain controller.

### Parameters and examples

| Parameter | Behavior |
| --- | --- |
| `Server` | Optional controller; normal AD discovery is used when omitted. Subsequent reads use the same responding controller. |
| `SearchBase` | Optional OU/container DN, including children. Defaults to the selected controller's default naming context. |
| `CsvPath` | Optional new CSV filename. The parent folder must already exist; existing files are never overwritten. |

Run from the repository root, replacing these `mylab.dev` examples with your lab values:

```powershell
# Return totals for one lab OU and its children.
.\active-directory\Get-ADGroupTypeCount.ps1 `
    -Server 'dc01.mylab.dev' `
    -SearchBase 'OU=Lab,DC=mylab,DC=dev'

# Also save the totals to a new CSV in the current directory.
.\active-directory\Get-ADGroupTypeCount.ps1 `
    -Server 'dc01.mylab.dev' `
    -SearchBase 'OU=Lab,DC=mylab,DC=dev' `
    -CsvPath '.\GroupCounts.csv'
```

The result contains `Server`, `SearchBase`, `Security`, `Distribution`, `Total`, `StartedAt`, `CompletedAt`, and `ElapsedSeconds`. A query failure or unexpected group category aborts the result, so partial counts are not reported as complete.

Counts cover readable groups within the chosen scope during the query. They are not a forest-wide or atomic snapshot. The script makes no AD changes, but large reads consume domain-controller resources; begin with a small lab OU.
