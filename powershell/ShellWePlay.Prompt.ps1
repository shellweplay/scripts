#requires -Version 5.1
<#
.SYNOPSIS
    A local date/time prompt for the current PowerShell session.
.DESCRIPTION
    Review this file, then dot-source it in a new test window. No files,
    policies, modules, network connections, or directory objects are changed.
    Closing that window restores normal behavior unless you explicitly add
    the file to your profile. This replaces the session's prompt function.

    The timestamp uses local time and refreshes when PowerShell draws a new
    prompt. It is not a constantly ticking clock or a command-duration timer.

    Colors use the host's console palette, not embedded ANSI escape sequences.
    With the companion ShellWePlay Dracula Windows Terminal scheme selected:
    DarkBlue = purple (#BD93F9), DarkCyan = cyan (#8BE9FD), Gray = #F8F8F2.
    Other terminal schemes determine their own colors for these same slots.
    This script does not install or select a terminal scheme.
.NOTES
    ShellWePlay / version 1.1.0. Windows PowerShell 5.1 / PowerShell 7 syntax.
    Interactive rendering depends on the host; test in a separate window.
#>
function global:prompt {
    # Read the local clock inside the function, once for each new prompt.
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host ''
    Write-Host ('[' + $timestamp + '] ') -NoNewline -ForegroundColor DarkBlue
    Write-Host 'OPERATOR ' -NoNewline -ForegroundColor DarkCyan
    Write-Host (Get-Location).Path -ForegroundColor Gray
    # Return a prompt string so PowerShell knows where input begins.
    return 'SHELL> '
}
Write-Host 'CONNECTION ESTABLISHED. GREETINGS, OPERATOR.' -ForegroundColor DarkCyan
