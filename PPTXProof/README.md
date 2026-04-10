# PPTX Proofreading Language Tool

A lightweight Windows GUI tool that sets the proofing (spell-check) language on **every text element** in an open PowerPoint presentation — slides, notes, tables, grouped shapes, masters and layouts.

Useful when working on a system where the OS language differs from the presentation language (e.g. German Windows but English slides), causing red squiggly spell-check marks everywhere.

## Features

- Auto-detects the currently open PowerPoint presentation
- Sets proofing language on all text elements in one click:
  - Slide shapes (text boxes, titles, callouts, etc.)
  - Grouped shapes (recursive)
  - Table cells
  - Notes pages
  - Slide masters and custom layouts
- Supported languages: **English (US)** and **German (DE)**
- Saves the presentation after patching
- No installation required — single EXE, zero dependencies

## Requirements

- **Windows** (uses PowerShell, WinForms, and COM automation — not compatible with macOS/Linux)
- **Microsoft PowerPoint** must be installed and running with a presentation open
- No admin rights needed

## Usage

1. Open your `.pptx` file in PowerPoint
2. Double-click `PPTXProof.exe`
3. Select the target language from the dropdown
4. Click **"Set Proofreading Language"**
5. Done — all text elements are patched and the file is saved

## Files

| File | Description |
|------|-------------|
| `PPTXProof.exe` | Compiled EXE — just double-click to run |
| `PPTX_Language_Tool.ps1` | PowerShell source code |
| `app.ico` | Application icon |

## Adding a Language

Edit `PPTX_Language_Tool.ps1` and add an entry to the `$Languages` hashtable:

```powershell
$Languages = [ordered]@{
    "English (US)" = 1033
    "German (DE)"  = 1031
    "French (FR)"  = 1036   # <-- add new languages here
}
```

Language IDs are `msoLanguageID` constants. Common values:

| Language | ID |
|----------|----|
| English (US) | 1033 |
| English (UK) | 2057 |
| German (DE) | 1031 |
| French (FR) | 1036 |
| Spanish (ES) | 3082 |
| Italian (IT) | 1040 |
| Portuguese (BR) | 1046 |
| Dutch (NL) | 1043 |
| Polish (PL) | 1045 |
| Chinese (Simplified) | 2052 |
| Japanese (JP) | 1041 |

Full list: [Microsoft msoLanguageID docs](https://learn.microsoft.com/en-us/office/vba/api/office.msolanguageid)

## Recompiling the EXE

Install the `ps2exe` PowerShell module (one-time):

```powershell
Install-Module -Name ps2exe -Scope CurrentUser -Force
```

Then compile:

```powershell
Invoke-ps2exe -InputFile PPTX_Language_Tool.ps1 -OutputFile PPTXProof.exe -noConsole -iconFile app.ico -title "PPTX Proofreading Language Tool"
```

## How It Works

The tool connects to the running PowerPoint instance via COM automation (`GetActiveObject("PowerPoint.Application")`), iterates every shape on every slide/master/layout, and sets `TextFrame.TextRange.LanguageID` to the selected language code. This is the same approach as doing it manually in PowerPoint via *Review > Language > Set Proofing Language*, but applied to all elements at once.

## License

MIT
