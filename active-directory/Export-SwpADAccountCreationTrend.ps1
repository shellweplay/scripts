#requires -Version 5.1
<#
.SYNOPSIS
    AD account creation trend.
.DESCRIPTION
    Count currently visible accounts by UTC creation month within a configurable recent window.
    Makes a paged, read-only query against one explicit organizational unit.
    Writes aggregate counts only; review and protect the CSV before sharing.
    The window includes the current UTC calendar month.
    This is not a historical audit log: deleted or moved-out users are not counted. Future months or unreadable dates use Unknown.
.PARAMETER SearchBase
    Distinguished name of the authorized OU to query. No default domain scope.
.PARAMETER SearchScope
    OneLevel (default) or Subtree (includes nested OUs).
.PARAMETER Server
    Optional approved AD DS domain controller. Otherwise use normal module selection.
.PARAMETER OutputDirectory
    Existing local filesystem folder. Defaults to the current directory.
.PARAMETER Months
    Number of UTC calendar months, including this month (1 to 60).
.EXAMPLE
    .\Export-SwpADAccountCreationTrend.ps1 -SearchBase "OU=Lab,DC=example,DC=test" -Months 6
.NOTES
    ShellWePlay version 1.0.0. Offline fixture checks do not replace live lab validation.
.LINK
    https://learn.microsoft.com/powershell/module/activedirectory/get-aduser
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('(?i)^(OU=).+')]
    [string]$SearchBase,

    [ValidateSet('OneLevel', 'Subtree')]
    [string]$SearchScope = 'OneLevel',

    [ValidateNotNullOrEmpty()]
    [string]$Server,

    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory = '.'
,

    [ValidateRange(1, 60)]
    [int]$Months = 12
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Resolve a literal directory before querying; no folders are created automatically.
$destination = Get-Item -LiteralPath $OutputDirectory
if ($destination -isnot [System.IO.DirectoryInfo]) {
    throw 'OutputDirectory must be an existing filesystem directory.'
}

Import-Module ActiveDirectory -ErrorAction Stop
$query = @{
    Filter = '*'
    SearchBase = $SearchBase
    SearchScope = $SearchScope
    ResultPageSize = 500
    ErrorAction = 'Stop'
    Properties = @('whenCreated')
}
if ($PSBoundParameters.ContainsKey('Server')) {
    $query.Server = $Server
}

# Do not emit a partial CSV if the directory query fails partway through.
$monthNow = [DateTime]::UtcNow.ToString('yyyy-MM', [System.Globalization.CultureInfo]::InvariantCulture)
$counts = [ordered]@{}
$firstMonth = [DateTime]::ParseExact(("$monthNow-01"), 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture).AddMonths(1 - $Months)
for ($i = 0; $i -lt $Months; $i++) {
    $key = $firstMonth.AddMonths($i).ToString('yyyy-MM', [System.Globalization.CultureInfo]::InvariantCulture)
    $counts[$key] = 0L
}
$counts['Earlier'] = 0L
$counts['Unknown'] = 0L
Get-ADUser @query | ForEach-Object {
    $property = $_.PSObject.Properties['whenCreated']
    if ($null -eq $property -or $property.Value -isnot [datetime]) { $counts['Unknown']++ }
    else {
        $key = $property.Value.ToUniversalTime().ToString('yyyy-MM', [System.Globalization.CultureInfo]::InvariantCulture)
        if ($counts.Contains($key)) { $counts[$key]++ }
        elseif ([string]::CompareOrdinal($key, $monthNow) -gt 0) { $counts['Unknown']++ }
        else { $counts['Earlier']++ }
    }
}
$rows = foreach ($key in $counts.Keys) {
    [pscustomobject]@{ PeriodUtc = $key; Count = $counts[$key] }
}

# Text from the directory is never written to these reports. The first row is a
# fixed header even if this OU contains no visible objects. CreateNew is atomic
# and never overwrites an existing report. Acquire it before managing cleanup.
$columns = @('PeriodUtc', 'Count')
$header = ($columns | ForEach-Object { '"{0}"' -f $_ }) -join ','
$stamp = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ', [System.Globalization.CultureInfo]::InvariantCulture)
$path = Join-Path -Path $destination.FullName -ChildPath ('Export-SwpADAccountCreationTrend-{0}.csv' -f $stamp)
$file = [System.IO.File]::Open($path, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
try {
    try {
        $writer = [System.IO.StreamWriter]::new($file, [System.Text.UTF8Encoding]::new($true))
        try {
            $writer.WriteLine($header)
            $rows | Select-Object -Property $columns | ConvertTo-Csv -NoTypeInformation |
                Select-Object -Skip 1 | ForEach-Object { $writer.WriteLine($_) }
        }
        finally {
            $writer.Dispose()
        }
    }
    finally {
        $file.Dispose()
    }
}
catch {
    try {
        [System.IO.File]::Delete($path)
    }
    catch {
        Write-Warning 'Could not remove the incomplete local report.'
    }
    throw
}
Get-Item -LiteralPath $path
