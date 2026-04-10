Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Languages = [ordered]@{
    "English (US)" = 1033
    "German (DE)"  = 1031
}

function Get-PowerPointApp {
    try { [System.Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
    catch { $null }
}

function Set-ShapeLanguage($shape, [int]$langId) {
    $count = 0
    try {
        if ($shape.HasTextFrame) {
            $shape.TextFrame.TextRange.LanguageID = $langId
            $count++
        }
    } catch {}
    try {
        if ($shape.Type -eq 6) {
            for ($i = 1; $i -le $shape.GroupItems.Count; $i++) {
                $count += Set-ShapeLanguage $shape.GroupItems.Item($i) $langId
            }
        }
    } catch {}
    try {
        if ($shape.HasTable) {
            $tbl = $shape.Table
            for ($r = 1; $r -le $tbl.Rows.Count; $r++) {
                for ($c = 1; $c -le $tbl.Rows($r).Cells.Count; $c++) {
                    $tbl.Rows($r).Cells($c).Shape.TextFrame.TextRange.LanguageID = $langId
                    $count++
                }
            }
        }
    } catch {}
    return $count
}

function Invoke-PatchPresentation([int]$langId) {
    $app = Get-PowerPointApp
    if (-not $app) { throw "PowerPoint is not running." }
    try { $pres = $app.ActivePresentation } catch { throw "No presentation is open." }
    $filepath = $pres.FullName
    if (-not $filepath -or -not (Test-Path $filepath)) {
        throw "The presentation has not been saved yet.`nPlease save it first."
    }
    $total = 0
    for ($j = 1; $j -le $pres.Slides.Count; $j++) {
        $slide = $pres.Slides($j)
        for ($k = 1; $k -le $slide.Shapes.Count; $k++) {
            $total += Set-ShapeLanguage $slide.Shapes($k) $langId
        }
        try {
            $notes = $slide.NotesPage
            for ($k = 1; $k -le $notes.Shapes.Count; $k++) {
                if ($notes.Shapes($k).HasTextFrame) {
                    $notes.Shapes($k).TextFrame.TextRange.LanguageID = $langId
                    $total++
                }
            }
        } catch {}
    }
    try {
        for ($m = 1; $m -le $pres.SlideMasters.Count; $m++) {
            $master = $pres.SlideMasters($m)
            for ($k = 1; $k -le $master.Shapes.Count; $k++) {
                $total += Set-ShapeLanguage $master.Shapes($k) $langId
            }
            for ($lay = 1; $lay -le $master.CustomLayouts.Count; $lay++) {
                $layout = $master.CustomLayouts($lay)
                for ($k = 1; $k -le $layout.Shapes.Count; $k++) {
                    $total += Set-ShapeLanguage $layout.Shapes($k) $langId
                }
            }
        }
    } catch {}
    $pres.Save()
    return @{ Path = $filepath; Count = $total }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "PPTX Proofreading Language Tool"
$form.Size = New-Object System.Drawing.Size(480, 280)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Set Proofreading Language for Entire Presentation"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(15, 12)
$lblTitle.AutoSize = $true
$form.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = "Changes the spell-check language on every text element`n(slides, notes, tables, groups, masters) in the open presentation."
$lblSub.ForeColor = [System.Drawing.Color]::Gray
$lblSub.Location = New-Object System.Drawing.Point(15, 42)
$lblSub.AutoSize = $true
$form.Controls.Add($lblSub)

$lblFileCaption = New-Object System.Windows.Forms.Label
$lblFileCaption.Text = "Open file:"
$lblFileCaption.Location = New-Object System.Drawing.Point(15, 90)
$lblFileCaption.AutoSize = $true
$form.Controls.Add($lblFileCaption)

$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text = "(detecting...)"
$lblFile.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$lblFile.Location = New-Object System.Drawing.Point(85, 90)
$lblFile.MaximumSize = New-Object System.Drawing.Size(370, 40)
$lblFile.AutoSize = $true
$form.Controls.Add($lblFile)

$lblLang = New-Object System.Windows.Forms.Label
$lblLang.Text = "Set language to:"
$lblLang.Location = New-Object System.Drawing.Point(15, 135)
$lblLang.AutoSize = $true
$form.Controls.Add($lblLang)

$cboLang = New-Object System.Windows.Forms.ComboBox
$cboLang.DropDownStyle = "DropDownList"
$cboLang.Location = New-Object System.Drawing.Point(125, 132)
$cboLang.Size = New-Object System.Drawing.Size(150, 24)
$Languages.Keys | ForEach-Object { $cboLang.Items.Add($_) } | Out-Null
$cboLang.SelectedIndex = 0
$form.Controls.Add($cboLang)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Set Proofreading Language"
$btnRun.Location = New-Object System.Drawing.Point(15, 180)
$btnRun.Size = New-Object System.Drawing.Size(200, 35)
$btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRun)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = ""
$lblStatus.ForeColor = [System.Drawing.Color]::Gray
$lblStatus.Location = New-Object System.Drawing.Point(225, 188)
$lblStatus.AutoSize = $true
$form.Controls.Add($lblStatus)

$detectAction = {
    $app = Get-PowerPointApp
    if (-not $app) { $lblFile.Text = "PowerPoint is not running"; return }
    try {
        $name = $app.ActivePresentation.FullName
        if ($name) { $lblFile.Text = $name } else { $lblFile.Text = "(unsaved presentation)" }
    } catch { $lblFile.Text = "No presentation open" }
}

$btnRun.Add_Click({
    $langLabel = $cboLang.SelectedItem.ToString()
    $langId = $Languages[$langLabel]
    $btnRun.Enabled = $false
    $lblStatus.Text = "Processing..."
    $form.Refresh()
    try {
        $result = Invoke-PatchPresentation $langId
        $filename = [System.IO.Path]::GetFileName($result.Path)
        $lblStatus.Text = "Done - $($result.Count) elements patched"
        [System.Windows.Forms.MessageBox]::Show(
            "Proofreading language for the entire presentation`n  `"$filename`"`nhas been set to $langLabel.`n`n$($result.Count) text elements patched.",
            "Proofreading Language Set",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    } catch {
        $lblStatus.Text = "Error"
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message, "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    $btnRun.Enabled = $true
})

$form.Add_Shown({ & $detectAction })
[void]$form.ShowDialog()
