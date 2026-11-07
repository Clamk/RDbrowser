# RD Browser (fenêtre classique) - tapez un numéro de RD, ouvrez son dossier ou ses fichiers.
# Les réglages sont dans RD_Browser.config.json, à côté de ce script.
# Lancement avec RD_Browser.cmd (aucune installation nécessaire).

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------- settings
$scriptDir  = $PSScriptRoot
$configPath = Join-Path $scriptDir 'RD_Browser.config.json'

$config = $null
if (Test-Path -LiteralPath $configPath) {
    try { $config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json }
    catch { [System.Windows.Forms.MessageBox]::Show("Impossible de lire le fichier de réglages :`n$configPath`n`n$($_.Exception.Message)", 'RD Browser') | Out-Null }
}
if (-not $config) {
    $config = [pscustomobject]@{
        RootPath = ''
        Buttons  = @(
            [pscustomobject]@{ Label = 'Dossier'; Type = 'folder' }
            [pscustomobject]@{ Label = 'Manip';  Type = 'file'; Pattern = 'manip'; Extensions = @('.xlsx', '.xlsm', '.xls') }
        )
    }
}

# Empty RootPath = the folder this script is in. Relative paths are relative to the script.
$root = [string]$config.RootPath
if ([string]::IsNullOrWhiteSpace($root))              { $root = $scriptDir }
elseif (-not [System.IO.Path]::IsPathRooted($root))   { $root = Join-Path $scriptDir $root }
try { $root = [System.IO.Path]::GetFullPath($root) } catch { }

# ---------------------------------------------------------------- logic
function Get-RdName([string]$text) {
    # "3", "03", "rd 3", "RD003" -> "RD003"; "100" -> "RD100"
    $digits = $text -replace '\D', ''
    if (-not $digits -or $digits.Length -gt 6) { return $null }
    'RD{0:D3}' -f [int]$digits
}

function Find-RdFolder([string]$rd) {
    $exact = Join-Path $root $rd
    if (Test-Path -LiteralPath $exact -PathType Container) { return Get-Item -LiteralPath $exact }
    # Also accept "RD003 - Something", but never "RD0030"
    Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match "^$rd(\D|$)" } |
        Select-Object -First 1
}

function Find-Files($folder, $btn) {
    $exts = @('.xlsx', '.xlsm', '.xls')
    if ($btn.Extensions) { $exts = @($btn.Extensions | ForEach-Object { $_.ToLower() }) }
    $pattern = [regex]::Escape([string]$btn.Pattern)
    @(Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -notlike '~$*' -and
            $exts -contains $_.Extension.ToLower() -and
            $_.BaseName -match $pattern
        } |
        Sort-Object LastWriteTime -Descending)
}

function Get-RdList {
    Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^RD\d{3}' } |
        Sort-Object Name |
        ForEach-Object { $_.Name }
}

# ---------------------------------------------------------------- UI helpers
function Set-Status([string]$msg, [string]$kind = 'info') {
    $status.Text = $msg
    $status.ForeColor = switch ($kind) {
        'ok'    { [System.Drawing.Color]::DarkGreen }
        'warn'  { [System.Drawing.Color]::DarkOrange }
        'error' { [System.Drawing.Color]::Firebrick }
        default { [System.Drawing.Color]::DimGray }
    }
}

function Select-FromList([string]$title, $items) {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $title
    $f.Font = $form.Font
    $f.StartPosition   = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'
    $f.MinimizeBox = $false; $f.MaximizeBox = $false
    $f.ClientSize  = New-Object System.Drawing.Size(380, 230)

    $lb = New-Object System.Windows.Forms.ListBox
    $lb.SetBounds(12, 12, 356, 170)
    foreach ($i in $items) { [void]$lb.Items.Add($i.Name) }
    $lb.SelectedIndex = 0
    $lb.Add_DoubleClick({ $f.DialogResult = 'OK' })

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Ouvrir'; $ok.DialogResult = 'OK'; $ok.SetBounds(196, 190, 82, 30)
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Annuler'; $cancel.DialogResult = 'Cancel'; $cancel.SetBounds(286, 190, 82, 30)

    $f.Controls.AddRange(@($lb, $ok, $cancel))
    $f.AcceptButton = $ok; $f.CancelButton = $cancel
    if ($f.ShowDialog($form) -eq 'OK') { return $items[$lb.SelectedIndex] }
    return $null
}

function Invoke-ButtonAction($btn) {
    $rd = Get-RdName $combo.Text
    if (-not $rd) { Set-Status 'Tapez d''abord un numéro de RD.' 'warn'; $combo.Focus(); return }

    $folder = Find-RdFolder $rd
    if (-not $folder) { Set-Status "$rd introuvable dans $root" 'error'; return }

    switch ($btn.Type) {
        'folder' {
            Invoke-Item -LiteralPath $folder.FullName
            Set-Status "Dossier $($folder.Name) ouvert" 'ok'
        }
        'file' {
            $files = Find-Files $folder $btn
            if ($files.Count -eq 0) {
                $answer = [System.Windows.Forms.MessageBox]::Show(
                    "Aucun fichier $($btn.Label) trouvé dans $($folder.Name).`n`nOuvrir le dossier à la place ?",
                    'RD Browser', 'YesNo', 'Question')
                if ($answer -eq 'Yes') { Invoke-Item -LiteralPath $folder.FullName }
                Set-Status "Aucun fichier $($btn.Label) dans $($folder.Name)" 'warn'
                return
            }
            $file = if ($files.Count -eq 1) { $files[0] }
                    else { Select-FromList "$($files.Count) fichiers $($btn.Label) dans $($folder.Name)" $files }
            if ($file) {
                Invoke-Item -LiteralPath $file.FullName
                Set-Status "Ouvert : $($folder.Name)\$($file.Name)" 'ok'
            }
        }
        default { Set-Status "« $($btn.Label) » n'est pas encore configuré." 'info' }
    }
}

# ---------------------------------------------------------------- window
$form = New-Object System.Windows.Forms.Form
$form.Text = 'RD Browser'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$form.AutoScaleMode   = 'Font'
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox     = $false
$form.ClientSize      = New-Object System.Drawing.Size(460, 225)

$tips = New-Object System.Windows.Forms.ToolTip

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = 'Numéro RD :'; $lbl.AutoSize = $true
$lbl.Location = New-Object System.Drawing.Point(20, 23)

$combo = New-Object System.Windows.Forms.ComboBox
$combo.DropDownStyle      = 'DropDown'
$combo.AutoCompleteMode   = 'SuggestAppend'
$combo.AutoCompleteSource = 'ListItems'
$combo.SetBounds(115, 19, 170, 28)
foreach ($n in Get-RdList) { [void]$combo.Items.Add($n) }
$combo.Add_DropDown({
    $current = $combo.Text
    $combo.Items.Clear()
    foreach ($n in Get-RdList) { [void]$combo.Items.Add($n) }
    $combo.Text = $current
})

$preview = New-Object System.Windows.Forms.Label
$preview.AutoSize = $true
$preview.ForeColor = [System.Drawing.Color]::DimGray
$preview.Location = New-Object System.Drawing.Point(295, 23)
$combo.Add_TextChanged({
    $rd = Get-RdName $combo.Text
    $preview.Text = if ($rd) { "-> $rd" } else { '' }
})

$form.Controls.AddRange(@($lbl, $combo, $preview))

# Buttons: 3 per row, built from the settings file
$buttons = @($config.Buttons)
for ($i = 0; $i -lt $buttons.Count; $i++) {
    $cfg = $buttons[$i]
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $cfg.Label
    $b.Tag  = $cfg
    $b.SetBounds(20 + ($i % 3) * 145, 65 + [math]::Floor($i / 3) * 50, 135, 40)
    if ($cfg.Type -eq 'none' -or -not $cfg.Type) {
        $b.Enabled = $false
        $tips.SetToolTip($b, 'Pas encore configuré')
    }
    $b.Add_Click({ Invoke-ButtonAction $this.Tag })
    $form.Controls.Add($b)
}

# Grow the window if there are more than 6 buttons
$rows = [math]::Max(2, [math]::Ceiling($buttons.Count / 3))
$statusTop = 65 + $rows * 50 + 10
$form.ClientSize = New-Object System.Drawing.Size(460, ($statusTop + 40))

$status = New-Object System.Windows.Forms.Label
$status.SetBounds(20, $statusTop, 420, 30)
$status.AutoEllipsis = $true
$form.Controls.Add($status)

# Enter = first button (Folder)
$combo.Add_KeyDown({
    if ($_.KeyCode -eq 'Enter') {
        $_.SuppressKeyPress = $true
        if ($buttons.Count -gt 0) { Invoke-ButtonAction $buttons[0] }
    }
})

if (Test-Path -LiteralPath $root -PathType Container) {
    Set-Status "Recherche dans $root"
} else {
    Set-Status "Impossible d'accéder à $root - vérifiez RD_Browser.config.json" 'error'
}

$form.Add_Shown({ $combo.Focus() })
[void]$form.ShowDialog()
