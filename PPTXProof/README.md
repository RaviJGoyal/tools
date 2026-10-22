# PPTXProof

Sets **one proofing (spell-check) language on everything** in PowerPoint presentations, so the wavy red underlines caused by text tagged with the wrong language disappear.

It is a single PowerShell script with a small window. There is no exe and nothing to install, and you don't need admin rights.

## Quick start

1. Double-click **`PPTXProof.cmd`**.
2. Drop presentations (or whole folders) into the window, or click **Add files...**.
   If PowerPoint is open, the presentation you are working on is already in the list.
3. Pick the language (German or English, etc.) and click **Set language**.

Each file shows the result, for example *"812 of 1,234 text elements set to German (Germany) (before: en-US 790, none 22)"*.

If a presentation is open in PowerPoint, PPTXProof asks whether to save and close it, applies the language, and opens it again.

### Shortcuts and right-click menu

Click **Create shortcuts** at the bottom of the window to add PPTXProof to the Start menu, the Desktop and Explorer's **Send to** menu. Then you can right-click one or more presentations and choose **Show more options > Send to > PPTXProof**.

You can also drag presentations onto `PPTXProof.cmd`. If you move the folder, create the shortcuts again.

## What gets changed

All text in the file is set to the chosen language:

- Slides: text boxes, placeholders, grouped shapes (any depth), tables, charts and SmartArt
- Speaker notes, notes master and handout master
- Slide masters and layouts
- Empty placeholders and paragraph ends, so text you type there later also gets the right language
- Presentation and theme defaults for new text
- Text without any language tag. It would otherwise fall back to PowerPoint's editing language.

Old spell-check results are cleared for the changed text, so PowerPoint checks it again.

Your file is handled carefully:

- PPTXProof edits a temporary copy, re-reads it to make sure it is intact, and only then replaces the original.
- Check **Keep a backup copy** to also keep the untouched original next to it, as `Name (backup 2026-10-22 153000).pptx`.
- Running it twice changes nothing.

Supported files: `.pptx`, `.pptm`, `.ppsx`, `.ppsm`, `.potx`, `.potm`. Old `.ppt` files must be saved as `.pptx` first. Password-protected or encrypted files (for example with a sensitivity label) cannot be edited.

## If typed text still gets the wrong language

PowerPoint tags newly typed text with the language of your **keyboard layout**. If you write English slides with a German keyboard layout active, new text may be marked as German again.

- Switch the layout with `Win + Space`, or remove layouts you don't use (Settings > Time & language > Language & region).
- Run PPTXProof again before you finish, so everything ends up in one language.

## Command line

```powershell
.\PPTXProof.ps1                                              # open the window
.\PPTXProof.ps1 .\Talk.pptx -Language en-GB                  # window with file and language preselected
.\PPTXProof.ps1 .\Talk.pptx .\Slides -Language de-DE -NoGui  # no window; folders are searched recursively
.\PPTXProof.ps1 .\Talk.pptx -Language en-US -Backup -NoGui   # keep a backup copy
```

`-Language` takes any language tag (`fr-FR`, `es-ES`, ...). With `-NoGui`, the script exits with code 1 if any file failed.

## Troubleshooting

- **"Windows protected your PC" / security warning:** the files were downloaded from the internet. Right-click `PPTXProof.cmd` and `PPTXProof.ps1`, open **Properties**, tick **Unblock**, and click OK.
- **Nothing happens / PowerShell is blocked:** some company policies block PowerShell scripts entirely (the launcher already bypasses the normal execution policy). In that case ask IT to allow it.
- **"The file is open in another program":** close the file in that program and try again. For PowerPoint, PPTXProof offers to do this for you.

## Adding a language to the window

Add an entry to `$Languages` at the top of `PPTXProof.ps1`:

```powershell
$Languages = [ordered]@{
    'de-DE' = 'German (Germany)'
    ...
    'fr-FR' = 'French (France)'   # <-- new entry
}
```

## How it works

A `.pptx` file is a ZIP archive of XML files. Every piece of text has a `lang` attribute, for example `<a:rPr lang="en-US"/>`. PPTXProof opens the archive and sets that attribute on all text, paragraph ends and text defaults in every slide, notes page, master, layout, chart and SmartArt part. It also adds the attribute where it is missing.

Because it edits the file directly, it doesn't depend on PowerPoint's automation interface. PowerPoint is only used, when it is running, to close and reopen a presentation you have open.

## Files

| File | Description |
|------|-------------|
| `PPTXProof.cmd` | Launcher: double-click it, or drop presentations onto it |
| `PPTXProof.ps1` | The tool (window and command line) |
| `app.ico` | Window and shortcut icon |

## License

MIT
