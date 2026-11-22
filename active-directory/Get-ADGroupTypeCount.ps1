#requires -Version 5.1
#requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Counts security and distribution groups using one paged AD query.
.DESCRIPTION
    Read-only against Active Directory. Groups are counted as they arrive;
    group objects and memberships are not accumulated in a collection.
    An optional CSV is a LOCAL write, never a directory modification.
    Large queries still consume domain-controller resources.
.PARAMETER Server
    Optional AD domain controller. If omitted, normal AD discovery is used.
.PARAMETER SearchBase
    Optional OU/container distinguished name. Defaults to the chosen
    controller's default naming context (not every domain in the forest).
.PARAMETER CsvPath
    Optional new CSV file. Its parent directory must exist. Never overwrites.
.EXAMPLE
    .\Get-ADGroupTypeCount.ps1 -SearchBase 'OU=Lab,DC=example,DC=com'
.EXAMPLE
    .\Get-ADGroupTypeCount.ps1 -Server 'dc01.example.com' -CsvPath '.\Counts.csv'
.NOTES
    Version: 1.0.0 / ShellWePlay preview
    Intended: Windows PowerShell 5.1; PowerShell 7 on Windows with AD module.
    Validation: Live Active Directory and compatibility testing PENDING.
    Permissions: Read access to the requested scope. No elevation required
    by the script itself; module installation is a separate prerequisite.
    Counts include only readable objects, not an atomic directory snapshot.
    Errors abort the result so partial counts are not reported as complete.
#>
[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()][string]$Server,
    [ValidateNotNullOrEmpty()][string]$SearchBase,
    [ValidateNotNullOrEmpty()][string]$CsvPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$started = Get-Date
$timer = [System.Diagnostics.Stopwatch]::StartNew()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $rootArgs = @{ ErrorAction = 'Stop' }
    if ($PSBoundParameters.ContainsKey('Server')) { $rootArgs.Server = $Server }
    $root = Get-ADRootDSE @rootArgs
    # Pin all subsequent reads to the same controller.
    $controller = [string]$root.dnsHostName
    $domainBase = [string]$root.defaultNamingContext
    if ([string]::IsNullOrWhiteSpace($controller) -or
        [string]::IsNullOrWhiteSpace($domainBase)) {
        throw 'The selected controller did not return a default naming context.'
    }
    if (-not $PSBoundParameters.ContainsKey('SearchBase')) { $SearchBase = $domainBase }
    # Validate existence before starting the potentially expensive query.
    $null = Get-ADObject -Identity $SearchBase -Server $controller -ErrorAction Stop

    if ($PSBoundParameters.ContainsKey('CsvPath')) {
        $CsvPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CsvPath)
        if (Test-Path -LiteralPath $CsvPath) { throw "Output already exists: $CsvPath" }
        $parent = Split-Path -Path $CsvPath -Parent
        if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
            throw "Output directory does not exist: $parent"
        }
    }

    # Mutable counters avoid retaining hundreds of thousands of AD objects.
    $counts = @{ Security = [long]0; Distribution = [long]0 }
    $query = @{
        Filter = '*'
        Server = $controller
        SearchBase = $SearchBase
        SearchScope = 'Subtree'
        ResultPageSize = 500
        ResultSetSize = $null
        ErrorAction = 'Stop'
    }
    Write-Verbose "Reading groups below $SearchBase from $controller"
    Get-ADGroup @query | ForEach-Object {
        switch ([string]$_.GroupCategory) {
            'Security' { $counts.Security++ }
            'Distribution' { $counts.Distribution++ }
            default { throw "Unrecognized category on group $($_.DistinguishedName). No totals emitted." }
        }
    }
    $timer.Stop()
    $result = [pscustomobject][ordered]@{
        Server = $controller
        SearchBase = $SearchBase
        Security = $counts.Security
        Distribution = $counts.Distribution
        Total = $counts.Security + $counts.Distribution
        StartedAt = $started.ToString('o')
        CompletedAt = (Get-Date).ToString('o')
        ElapsedSeconds = [math]::Round($timer.Elapsed.TotalSeconds, 2)
    }
    if ($PSBoundParameters.ContainsKey('CsvPath')) {
        $result | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8 -NoClobber -ErrorAction Stop
    }
    # Emit the final result only after the entire operation succeeds.
    $result
}
catch {
    throw "Group count did not complete; no successful total was emitted. $($_.Exception.Message)"
}
finally {
    $timer.Stop()
}
