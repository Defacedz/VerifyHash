#requires -Version 5.1
<#
.SYNOPSIS
    Verifies a file hash from the Windows right-click menu.

.DESCRIPTION
    A single file with two modes, picked automatically:

      * started without an argument - the installer window, which adds or
        removes the "Verify hash" entry in the right-click menu of every file.

      * started with a file path - the verification window, which hashes that
        file and compares the result to the hash you paste.

    What it does:

      - MD5, SHA-1, SHA-256 and SHA-512.
      - The algorithm follows the hash you paste: 32 characters switch to MD5,
        40 to SHA-1, 64 to SHA-256, 128 to SHA-512.
      - Hashing is progressive. The window is up immediately and a progress bar
        moves, even on a file of several gigabytes.
      - Pasted text is cleaned up: spaces, line breaks and prefixes such as
        "SHA256:" are dropped, and the case is normalised.
      - Characters that differ are highlighted in both fields, with their
        position.
      - Another file dropped on the window replaces the current one.

    The script never writes to the clipboard; the Paste button only reads it.
    The menu entry lives under HKEY_CURRENT_USER, so no administrator rights
    are needed and nothing is installed for other users.

.PARAMETER Path
    File to verify. Omit it to open the installer window instead.

.PARAMETER Language
    'en' or 'fr'. Defaults to the Windows display language.

.EXAMPLE
    .\VerifyHash.ps1

    Opens the installer window.

.EXAMPLE
    .\VerifyHash.ps1 'C:\Downloads\linux.iso'

    Opens the verification window for that file.

.NOTES
    Save this file as UTF-8 with a BOM. Windows PowerShell 5.1 does not detect
    UTF-8 without one, and the accented French strings would be mangled.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Path,

    [ValidateSet('en', 'fr')]
    [string]$Language
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -Namespace VH -Name Native -MemberDefinition @'
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();

    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
'@

# Declared before any window exists: Windows then hands us real pixels instead
# of a blurry bitmap stretched by the compositor on a scaled display.
[void][VH.Native]::SetProcessDPIAware()
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Constants ---------------------------------------------------------------

$RegPath = 'Software\Classes\*\shell\VerifyHash'

# Keys written by earlier versions of this script, removed on uninstall.
$RegPathLegacy = @(
    'Software\Classes\*\shell\HashCheck',
    'Software\Classes\*\shell\SHA256Check',
    'Software\Classes\*\shell\SHA256'
)

# Hash length in characters -> matching algorithm.
$LenToAlgo = @{ 32 = 'MD5'; 40 = 'SHA1'; 64 = 'SHA256'; 128 = 'SHA512' }

$ChunkSize = 4MB

# --- Strings -----------------------------------------------------------------

$Strings = @{
    en = @{
        MenuLabel     = 'Verify hash'
        AppTitle      = 'VerifyHash'
        Algorithm     = 'ALGORITHM'
        Computed      = 'COMPUTED HASH'
        Expected      = 'EXPECTED HASH — paste it here'
        Paste         = 'Paste'
        Hint          = 'Esc closes the window   ·   drop another file on it to check that one'
        Computing     = 'Hashing…'
        AwaitExpected = 'Waiting for the expected hash'
        Typing        = 'Matching so far — {0} of {1} characters'
        Match         = 'MATCH — the file is intact'
        MatchDetail   = '{0} characters, exact match ({1}).'
        DiffTitle     = 'DIFFERENT — {0} character(s) do not match'
        AllDiffTitle  = 'COMPLETELY DIFFERENT'
        AllDiffNote   = 'not one character matches: wrong file, or a hash taken from another version'
        Positions     = 'position(s): {0}'
        LengthNote    = 'length {0} instead of {1}'
        ReadFailed    = 'Could not read the file'
        OpenFailed    = "Unreadable file:`n{0}"
        InstallTitle  = 'VerifyHash — installation'
        InstallIntro  = 'Adds "Verify hash" to the right-click menu of every file.'
        ScriptPath    = 'LOCATION OF THIS SCRIPT'
        NotInstalled  = 'Not installed'
        UpToDate      = 'Installed and up to date'
        Stale         = "Installed, but pointing at another location.`nClick Install to update it."
        BtnInstall    = 'Install'
        BtnUninstall  = 'Uninstall'
        BtnClose      = 'Close'
        Failed        = 'Failed: {0}'
    }
    fr = @{
        MenuLabel     = "Vérifier l'empreinte"
        AppTitle      = 'VerifyHash'
        Algorithm     = 'ALGORITHME'
        Computed      = 'EMPREINTE CALCULÉE'
        Expected      = 'EMPREINTE ATTENDUE — collez-la ici'
        Paste         = 'Coller'
        Hint          = 'Échap ferme la fenêtre   ·   déposez-y un autre fichier pour le vérifier'
        Computing     = 'Calcul en cours…'
        AwaitExpected = "En attente de l'empreinte attendue"
        Typing        = 'Concordant jusqu''ici — {0} caractères sur {1}'
        Match         = 'IDENTIQUE — le fichier est intact'
        MatchDetail   = '{0} caractères, correspondance totale ({1}).'
        DiffTitle     = 'DIFFÉRENT — {0} caractère(s) en écart'
        AllDiffTitle  = 'TOTALEMENT DIFFÉRENT'
        AllDiffNote   = "aucun caractère ne correspond : mauvais fichier, ou empreinte issue d'une autre version"
        Positions     = 'position(s) : {0}'
        LengthNote    = 'longueur {0} au lieu de {1}'
        ReadFailed    = 'Lecture impossible'
        OpenFailed    = "Fichier illisible :`n{0}"
        InstallTitle  = 'VerifyHash — installation'
        InstallIntro  = "Ajoute « Vérifier l'empreinte » au menu clic droit de tous les fichiers."
        ScriptPath    = 'EMPLACEMENT DE CE SCRIPT'
        NotInstalled  = 'Non installé'
        UpToDate      = 'Installé et à jour'
        Stale         = "Installé, mais pointe vers un autre emplacement.`nCliquez sur Installer pour corriger."
        BtnInstall    = 'Installer'
        BtnUninstall  = 'Désinstaller'
        BtnClose      = 'Fermer'
        Failed        = 'Échec : {0}'
    }
}

if (-not $Language) {
    $Language = if ([System.Globalization.CultureInfo]::CurrentUICulture.TwoLetterISOLanguageName -eq 'fr') { 'fr' } else { 'en' }
}
$T = $Strings[$Language]

# --- Theme -------------------------------------------------------------------

function RGB([int]$r, [int]$g, [int]$b) { [System.Drawing.Color]::FromArgb($r, $g, $b) }

$colBg      = RGB  21  23  26     # window, idle
$colBgOk    = RGB  19  38  27     # window, hashes match
$colBgBad   = RGB  40  21  23     # window, hashes differ
$colField   = RGB  28  31  36
$colLine    = RGB  42  47  54
$colFg      = RGB 230 232 235
$colDim     = RGB 139 146 155
$colBandOk  = RGB  30 122  70
$colBandBad = RGB 163  42  46
$colDiff    = RGB 255 107 107
$colMatch   = RGB  85 217 141
$colBtn     = RGB  42  47  54
$colBtnHot  = RGB  56  63  72
$colAccent  = RGB  56 110 186

# All layout numbers below are written for 96 DPI and scaled through Px().
$Scale = 1.0
$gfx = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
try { $Scale = $gfx.DpiX / 96.0 } finally { $gfx.Dispose() }

function Px([double]$v) { [int][Math]::Round($v * $Scale) }

# Point sizes are resolution-independent, so fonts must not be scaled again.
function New-Font([string[]]$families, [double]$size, [System.Drawing.FontStyle]$style) {
    foreach ($name in $families) {
        $f = New-Object System.Drawing.Font($name, $size, $style)
        if ($f.Name -eq $name) { return $f }
        $f.Dispose()
    }
    return (New-Object System.Drawing.Font('Segoe UI', $size, $style))
}

$reg  = [System.Drawing.FontStyle]::Regular
$bold = [System.Drawing.FontStyle]::Bold

$fontMono   = New-Font @('Cascadia Mono', 'Consolas', 'Courier New') 11   $reg
$fontUi     = New-Font @('Segoe UI Variable Text', 'Segoe UI')        9.5 $reg
$fontHead   = New-Font @('Segoe UI Variable Text', 'Segoe UI')        8.5 $bold
$fontFile   = New-Font @('Segoe UI Variable Display', 'Segoe UI')     11  $bold
$fontStatus = New-Font @('Segoe UI Variable Display', 'Segoe UI')     13  $bold

# --- UI helpers --------------------------------------------------------------

function New-Form($title, $w, $h) {
    $f                 = New-Object System.Windows.Forms.Form
    $f.Text            = $title
    # The layout is scaled by hand through Px(). WinForms must not scale it a
    # second time, which its default font-based auto-scaling would do.
    $f.AutoScaleMode   = 'None'
    $f.ClientSize      = New-Object System.Drawing.Size((Px $w), (Px $h))
    $f.StartPosition   = 'CenterScreen'
    $f.BackColor       = $colBg
    $f.ForeColor       = $colFg
    $f.Font            = $fontUi
    $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox     = $false
    $f.MinimizeBox     = $false
    $f.TopMost         = $true
    $f.KeyPreview      = $true
    $f.Add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $this.Close() } })
    $f.Add_HandleCreated({
        # Dark title bar. Attribute 20 on Windows 10 2004+, 19 on 1809-1909,
        # and simply ignored on anything older.
        $on = 1
        foreach ($attr in 20, 19) {
            [void][VH.Native]::DwmSetWindowAttribute($this.Handle, $attr, [ref]$on, 4)
        }
    })
    $f.Add_Shown({ $this.Activate(); $this.BringToFront() })
    return $f
}

function Add-Label($parent, $text, $x, $y, $w, $h, $font, $color) {
    $l           = New-Object System.Windows.Forms.Label
    $l.Text      = $text
    $l.Location  = New-Object System.Drawing.Point((Px $x), (Px $y))
    $l.Size      = New-Object System.Drawing.Size((Px $w), (Px $h))
    $l.Font      = $font
    $l.ForeColor = $color
    $l.BackColor = [System.Drawing.Color]::Transparent
    $parent.Controls.Add($l)
    return $l
}

function Add-Button($parent, $text, $x, $y, $w, $h) {
    $b           = New-Object System.Windows.Forms.Button
    $b.Text      = $text
    $b.Location  = New-Object System.Drawing.Point((Px $x), (Px $y))
    $b.Size      = New-Object System.Drawing.Size((Px $w), (Px $h))
    $b.FlatStyle = 'Flat'
    $b.BackColor = $colBtn
    $b.ForeColor = $colFg
    $b.FlatAppearance.BorderSize         = 0
    $b.FlatAppearance.MouseOverBackColor = $colBtnHot
    $b.FlatAppearance.MouseDownBackColor = $colLine
    $b.UseVisualStyleBackColor = $false
    $parent.Controls.Add($b)
    return $b
}

# A borderless RichTextBox inset in a panel: the panel supplies the padding and
# the background, which a RichTextBox cannot do on its own.
function Add-HashBox($parent, $x, $y, $w, $h, $readOnly) {
    $pad = 10

    $panel           = New-Object System.Windows.Forms.Panel
    $panel.Location  = New-Object System.Drawing.Point((Px $x), (Px $y))
    $panel.Size      = New-Object System.Drawing.Size((Px $w), (Px $h))
    $panel.BackColor = $colField
    $parent.Controls.Add($panel)

    $b             = New-Object System.Windows.Forms.RichTextBox
    $b.Location    = New-Object System.Drawing.Point((Px $pad), (Px (($h - 20) / 2)))
    $b.Size        = New-Object System.Drawing.Size((Px ($w - 2 * $pad)), (Px 20))
    $b.Font        = $fontMono
    $b.BackColor   = $colField
    $b.ForeColor   = $colFg
    $b.BorderStyle = 'None'
    $b.Multiline   = $true
    $b.WordWrap    = $false
    $b.ReadOnly    = $readOnly
    $b.ScrollBars  = 'None'
    $panel.Controls.Add($b)

    # Clicking anywhere in the padding still lands in the field.
    if (-not $readOnly) { $panel.Add_Click({ $b.Focus() }.GetNewClosure()) }

    return $b
}

function Format-Size($bytes) {
    if     ($bytes -ge 1GB) { '{0:N2} GB' -f ($bytes / 1GB) }
    elseif ($bytes -ge 1MB) { '{0:N1} MB' -f ($bytes / 1MB) }
    elseif ($bytes -ge 1KB) { '{0:N0} KB' -f ($bytes / 1KB) }
    else                    { "$bytes B" }
}

# --- Registry ----------------------------------------------------------------
# The .NET API is used rather than New-Item / Get-Item: the * in the key path is
# a wildcard for the PowerShell Registry provider, which would break the call.

function Get-InstalledCommand {
    $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("$RegPath\command")
    if ($null -eq $k) { return $null }
    $v = $k.GetValue('')
    $k.Close()
    return $v
}

function Install-ContextMenu($scriptPath) {
    $cmd = 'powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" "%1"' -f $scriptPath
    $k = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($RegPath)
    $k.SetValue('', $T.MenuLabel)
    $k.SetValue('Icon', 'powershell.exe')
    $c = $k.CreateSubKey('command')
    $c.SetValue('', $cmd)
    $c.Close()
    $k.Close()
}

function Uninstall-ContextMenu {
    foreach ($p in @($RegPath) + $RegPathLegacy) {
        try { [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($p, $false) } catch { }
    }
}

# --- Verification window -----------------------------------------------------

function Show-HashWindow($filePath) {

    $script:FilePath = $filePath
    $script:Algo     = 'SHA256'
    $script:Computed = ''
    $script:Busy     = $false
    $script:Abort    = $false
    $script:Pending  = $false
    $script:Closing  = $false

    $W = 760

    $form = New-Form $T.AppTitle $W 390
    $form.AllowDrop = $true

    $lblFile = Add-Label $form '' 24 18 400 24 $fontFile $colFg
    $lblFile.AutoSize = $true
    $lblSize = Add-Label $form '' 24 20 200 22 $fontUi $colDim
    $lblSize.AutoSize = $true

    # A segmented control rather than a ComboBox: a DropDownList keeps the
    # system's white background whatever BackColor says, which no dark theme
    # survives.
    Add-Label $form $T.Algorithm 24 58 300 16 $fontHead $colDim | Out-Null
    $algos   = @('MD5', 'SHA1', 'SHA256', 'SHA512')
    $btnAlgo = @{}
    for ($i = 0; $i -lt $algos.Count; $i++) {
        $btnAlgo[$algos[$i]] = Add-Button $form $algos[$i] (24 + $i * 78) 78 72 30
    }

    $syncAlgo = {
        foreach ($a in $algos) {
            $b = $btnAlgo[$a]
            if ($a -eq $script:Algo) {
                $b.BackColor = $colAccent
                $b.ForeColor = [System.Drawing.Color]::White
                $b.FlatAppearance.MouseOverBackColor = $colAccent
            }
            else {
                $b.BackColor = $colBtn
                $b.ForeColor = $colDim
                $b.FlatAppearance.MouseOverBackColor = $colBtnHot
            }
        }
    }
    & $syncAlgo

    $bar          = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point((Px 350), (Px 89))
    $bar.Size     = New-Object System.Drawing.Size((Px ($W - 374)), (Px 8))
    $bar.Style    = 'Continuous'
    $bar.Visible  = $false
    $form.Controls.Add($bar)

    Add-Label $form $T.Computed 24 130 400 16 $fontHead $colDim | Out-Null
    $boxComputed = Add-HashBox $form 24 150 ($W - 48) 40 $true

    Add-Label $form $T.Expected 24 206 400 16 $fontHead $colDim | Out-Null
    $boxExpected = Add-HashBox $form 24 226 ($W - 152) 40 $false
    $btnPaste    = Add-Button  $form $T.Paste ($W - 120) 226 96 40

    $band           = New-Object System.Windows.Forms.Panel
    $band.Location  = New-Object System.Drawing.Point(0, (Px 288))
    $band.Size      = New-Object System.Drawing.Size((Px $W), (Px 74))
    $band.BackColor = $colField
    $form.Controls.Add($band)

    $lblStatus = Add-Label $band '' 24  12 ($W - 48) 26 $fontStatus $colDim
    $lblDetail = Add-Label $band '' 24  42 ($W - 48) 22 $fontUi     $colDim

    Add-Label $form $T.Hint 24 370 ($W - 48) 16 $fontUi $colDim | Out-Null

    # Colours the characters listed in $idx, leaves the rest alone.
    $paint = {
        param($box, $idx, $allGood)
        $sel = $box.SelectionStart
        $box.SelectAll()
        $box.SelectionColor = if ($allGood) { $colMatch } else { $colFg }
        foreach ($i in $idx) {
            if ($i -lt $box.TextLength) {
                $box.Select($i, 1)
                $box.SelectionColor = $colDiff
            }
        }
        $box.Select($sel, 0)
    }

    $setBand = {
        param($bg, $band2, $status, $statusColor, $detail)
        $form.BackColor      = $bg
        $band.BackColor      = $band2
        $lblStatus.ForeColor = $statusColor
        $lblStatus.Text      = $status
        $lblDetail.Text      = $detail
    }

    $compare = {
        $expected = ($boxExpected.Text -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
        $computed = $script:Computed
        $diff     = New-Object System.Collections.ArrayList

        if ($computed -eq '' -or $expected.Length -eq 0) {
            $msg = if ($computed -eq '') { $T.Computing } else { $T.AwaitExpected }
            & $setBand $colBg $colField $msg $colDim ''
            & $paint $boxComputed $diff $false
            & $paint $boxExpected $diff $false
            return
        }

        $n = [Math]::Min($expected.Length, $computed.Length)
        for ($i = 0; $i -lt $n; $i++) {
            if ($expected[$i] -ne $computed[$i]) { [void]$diff.Add($i) }
        }
        $lenDelta = $expected.Length - $computed.Length

        # Still being typed and correct so far: stay neutral rather than
        # flashing red on every keystroke.
        if ($diff.Count -eq 0 -and $lenDelta -lt 0) {
            & $setBand $colBg $colField ($T.Typing -f $expected.Length, $computed.Length) $colDim ''
            & $paint $boxComputed $diff $false
            & $paint $boxExpected $diff $false
            return
        }

        if ($diff.Count -eq 0 -and $lenDelta -eq 0) {
            & $setBand $colBgOk $colBandOk $T.Match ([System.Drawing.Color]::White) `
                ($T.MatchDetail -f $computed.Length, $script:Algo)
            $lblDetail.ForeColor = RGB 205 240 218
            & $paint $boxComputed $diff $true
            & $paint $boxExpected $diff $true
            return
        }

        $notes = @()
        if ($lenDelta -ne 0) { $notes += $T.LengthNote -f $expected.Length, $computed.Length }

        if ($n -gt 0 -and $diff.Count -eq $n) {
            $title = $T.AllDiffTitle
            $notes += $T.AllDiffNote
        }
        else {
            $title = $T.DiffTitle -f $diff.Count
            $first = ($diff | Select-Object -First 6) | ForEach-Object { $_ + 1 }
            $suite = if ($diff.Count -gt 6) { ', …' } else { '' }
            $notes += $T.Positions -f "$($first -join ', ')$suite"
        }

        & $setBand $colBgBad $colBandBad $title ([System.Drawing.Color]::White) ($notes -join '   ·   ')
        $lblDetail.ForeColor = RGB 255 214 214
        & $paint $boxComputed $diff $false
        & $paint $boxExpected $diff $false
    }

    # One hashing pass. The file is read in chunks and the message loop is
    # pumped between them, so the window stays alive on a multi-gigabyte file
    # without a second thread.
    $computeCore = {
        $script:Abort     = $false
        $script:Computed  = ''
        $boxComputed.Text = ''
        foreach ($a in $algos) { $btnAlgo[$a].Enabled = $false }
        $bar.Value        = 0
        $bar.Visible      = $true
        & $compare

        $algo = [System.Security.Cryptography.HashAlgorithm]::Create($script:Algo)
        $fs   = $null
        try {
            $fs    = [System.IO.File]::Open($script:FilePath, 'Open', 'Read', 'ReadWrite')
            $total = [Math]::Max($fs.Length, 1)
            $buf   = New-Object byte[] ($ChunkSize)
            $done  = 0L
            $tick  = 0

            while (($read = $fs.Read($buf, 0, $buf.Length)) -gt 0) {
                [void]$algo.TransformBlock($buf, 0, $read, $null, 0)
                $done += $read
                if ((++$tick % 2) -eq 0) {
                    $bar.Value = [int](100 * $done / $total)
                    [System.Windows.Forms.Application]::DoEvents()
                    if ($script:Abort -or $form.IsDisposed) { break }
                }
            }

            if (-not $script:Abort -and -not $form.IsDisposed) {
                [void]$algo.TransformFinalBlock($buf, 0, 0)
                $script:Computed  = ([BitConverter]::ToString($algo.Hash) -replace '-', '')
                $boxComputed.Text = $script:Computed
            }
        }
        catch {
            & $setBand $colBgBad $colBandBad $T.ReadFailed ([System.Drawing.Color]::White) $_.Exception.Message
            $lblDetail.ForeColor = RGB 255 214 214
        }
        finally {
            if ($fs)   { $fs.Dispose() }
            if ($algo) { $algo.Dispose() }
            if (-not $form.IsDisposed -and -not $script:Pending) {
                $bar.Visible = $false
                foreach ($a in $algos) { $btnAlgo[$a].Enabled = $true }
                if ($script:Computed -ne '') { & $compare }
            }
        }
    }

    # Asking for a hash while one is running cancels it and starts the new one,
    # rather than being dropped on the floor.
    $compute = {
        if ($script:Busy) {
            $script:Abort   = $true
            $script:Pending = $true
            return
        }
        $script:Busy = $true
        try {
            do {
                $script:Pending = $false
                & $computeCore
            } while ($script:Pending -and -not $script:Closing -and -not $form.IsDisposed)
        }
        finally { $script:Busy = $false }
    }

    $load = {
        param($p)
        try { $item = Get-Item -LiteralPath $p -ErrorAction Stop }
        catch {
            [System.Windows.Forms.MessageBox]::Show(($T.OpenFailed -f $p), $T.AppTitle, 'OK', 'Error') | Out-Null
            return
        }
        $script:FilePath = $item.FullName
        $lblFile.Text    = $item.Name
        $lblSize.Text    = Format-Size $item.Length
        $lblSize.Left    = $lblFile.Right + (Px 14)
        $form.Text       = "$($T.AppTitle) — $($item.Name)"
        & $compute
    }

    $pick = {
        param($algoName)
        if ($algoName -eq $script:Algo) { return }
        $script:Algo = $algoName
        & $syncAlgo
        & $compute
    }

    foreach ($a in $algos) {
        $btnAlgo[$a].Add_Click({ & $pick $this.Text }.GetNewClosure())
    }

    $boxExpected.Add_TextChanged({
        $raw   = $boxExpected.Text
        $clean = ($raw -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
        if ($raw -cne $clean) {
            $boxExpected.Text = $clean
            $boxExpected.Select($clean.Length, 0)
            return   # this assignment fires the handler again
        }

        # The pasted length decides the algorithm.
        $want = $LenToAlgo[$clean.Length]
        if ($want -and $want -ne $script:Algo) {
            & $pick $want   # triggers a fresh hash
            return
        }

        & $compare
    })

    $btnPaste.Add_Click({
        # Read-only: nothing is ever put back on the clipboard.
        if ([System.Windows.Forms.Clipboard]::ContainsText()) {
            $boxExpected.Text = [System.Windows.Forms.Clipboard]::GetText()
        }
        $boxExpected.Focus()
        $boxExpected.Select($boxExpected.TextLength, 0)
    })

    $form.Add_DragEnter({
        if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        }
    })
    $form.Add_DragDrop({
        $dropped = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        if ($dropped -and $dropped.Count -gt 0) { & $load $dropped[0] }
    })

    $form.Add_FormClosing({
        $script:Closing = $true
        $script:Abort   = $true
        $script:Pending = $false
    })
    $form.Add_Shown({ $boxExpected.Focus(); & $load $script:FilePath })

    [void]$form.ShowDialog()
    $form.Dispose()
}

# --- Installer window --------------------------------------------------------

function Show-Installer {

    $self = $PSCommandPath
    if ([string]::IsNullOrWhiteSpace($self)) { $self = $MyInvocation.MyCommand.Path }

    $W = 620

    $form = New-Form $T.InstallTitle $W 268

    Add-Label $form $T.InstallIntro 24 22 ($W - 48) 22 $fontFile $colFg  | Out-Null
    Add-Label $form $T.ScriptPath   24 62 ($W - 48) 16 $fontHead $colDim | Out-Null

    $pathPanel           = New-Object System.Windows.Forms.Panel
    $pathPanel.Location  = New-Object System.Drawing.Point((Px 24), (Px 82))
    $pathPanel.Size      = New-Object System.Drawing.Size((Px ($W - 48)), (Px 40))
    $pathPanel.BackColor = $colField
    $form.Controls.Add($pathPanel)
    Add-Label $pathPanel $self 10 11 ($W - 68) 20 $fontUi $colFg | Out-Null

    $lblState = Add-Label $form '' 24 138 ($W - 48) 44 $fontStatus $colDim

    $btnW = [int](($W - 48 - 24) / 3)
    $btnIn  = Add-Button $form $T.BtnInstall   24                    198 $btnW 40
    $btnIn.BackColor = $colAccent
    $btnIn.ForeColor = [System.Drawing.Color]::White
    $btnIn.FlatAppearance.MouseOverBackColor = RGB 72 128 205
    $btnOut = Add-Button $form $T.BtnUninstall (24 + $btnW + 12)     198 $btnW 40
    $btnEnd = Add-Button $form $T.BtnClose     (24 + 2 * $btnW + 24) 198 $btnW 40

    $refresh = {
        $cur = Get-InstalledCommand
        if ($null -eq $cur) {
            $lblState.Text      = $T.NotInstalled
            $lblState.ForeColor = $colDim
        }
        elseif ($cur -like "*$self*") {
            $lblState.Text      = $T.UpToDate
            $lblState.ForeColor = $colMatch
        }
        else {
            $lblState.Text      = $T.Stale
            $lblState.ForeColor = $colDiff
        }
    }

    $btnIn.Add_Click({
        try { Install-ContextMenu $self; & $refresh }
        catch {
            [System.Windows.Forms.MessageBox]::Show(($T.Failed -f $_.Exception.Message), $T.AppTitle, 'OK', 'Error') | Out-Null
        }
    })
    $btnOut.Add_Click({ Uninstall-ContextMenu; & $refresh })
    $btnEnd.Add_Click({ $form.Close() })

    & $refresh
    [void]$form.ShowDialog()
    $form.Dispose()
}

# --- Entry point -------------------------------------------------------------

if ([string]::IsNullOrWhiteSpace($Path)) { Show-Installer }
else                                     { Show-HashWindow $Path }
