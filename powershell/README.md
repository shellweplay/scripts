# PowerShell customization

[ShellWePlay.Prompt.ps1](ShellWePlay.Prompt.ps1), version 1.1.0, displays a local
date/time, an operator label, the current path, and a `SHELL>` input prompt.
The timestamp refreshes each time PowerShell draws the next prompt.

**Validation:** Existing website CI checks the prompt in an isolated PowerShell
process, including the timestamp refresh, literal path rendering, palette slots,
and returned prompt string. Interactive Windows PowerShell / Windows Terminal
validation remains pending. Host rendering can differ.

## Try it in a separate window

Review the file, open a fresh Windows PowerShell 5.1 or PowerShell 7 test window,
and run this from the repository root:

```powershell
# Dot-source the reviewed file to replace this session's prompt function.
. ./powershell/ShellWePlay.Prompt.ps1
```

The script does not write files, edit a profile, install modules, change execution
policy, or select a terminal theme. The prompt change lasts for that session.
If policy blocks the file, use your approved
process rather than weakening the policy for a prompt.

Colors come from your terminal's current console palette. The optional Windows
Terminal schemes, profile backup/persistence instructions, and undo steps are
available in the [Shell Lab guide](https://shellweplay.vercel.app/shell-lab).
