# Automated checks for an installed Typewriter (run build\build.ps1 -Install
# first). Opens Word, drives typewriter mode through its macros, and prints
# PASS/FAIL lines followed by the built-in Diagnostics report.
#
#   pwsh tests\test.ps1
$ErrorActionPreference = 'Continue'
if (Get-Process WINWORD -ErrorAction SilentlyContinue) { throw 'Close Word before testing.' }

$w = New-Object -ComObject Word.Application
function T($a, $n = 0) { $w.Run('Typewriter.TypewriterTest', $a, $n) }
function M($name) { $w.Run("Typewriter.$name") | Out-Null }
function Check($label, $cond) { if ($cond) { "PASS  $label" } else { "FAIL  $label" } }
# Settings without the parts that legitimately change (status, pending timer tick)
function Settings($state) { $state -replace 'status=\S+ ', '' -replace ' tickPending=\S+', '' }

try {
    $w.Visible = $true
    'add-ins: ' + (($w.AddIns | ForEach-Object { $_.Name + '(' + $_.Installed + ')' }) -join ', ')
    $d = $w.Documents.Add()
    $w.Selection.TypeText('Hello world. This is the start.')
    $orig = $d.Content.Text
    $before = T 'state'

    # Typewriter mode on
    [void](T 'start')
    Check 'session starts' ((T 'state') -like 'on=True*')
    Check 'window title shows the timer' ($d.Windows(1).Caption -like '*TYPEWRITER*')
    $d.Range(0, 5).Select()
    Check 'selecting old text snaps the cursor to the end' (
        $w.Selection.Start -eq ($d.Content.End - 1) -and $w.Selection.End -eq $w.Selection.Start)
    foreach ($m in 'TypewriterBackspace', 'TypewriterCtrlBackspace', 'TypewriterAltBackspace',
                   'TypewriterDelete', 'TypewriterCtrlDelete', 'EditClear', 'DeleteWord',
                   'DeleteBackWord', 'EditCut', 'EditUndo') { M $m }
    Check 'delete, cut and undo leave text unchanged' ($d.Content.Text -eq $orig)
    $w.Selection.TypeText(' More.')
    Check 'typing still appends' ($d.Content.Text.TrimEnd() -eq ($orig.TrimEnd() + ' More.'))
    $s1 = T 'state'; Start-Sleep 3; $s2 = T 'state'
    Check 'countdown ticks' ($s1 -ne $s2)

    # Typewriter mode off
    [void](T 'stop')
    Check 'window title restored' ($d.Windows(1).Caption -notlike '*TYPEWRITER*')
    Check 'settings restored' ((Settings (T 'state')) -eq (Settings $before))
    $d.Range(5, 5).Select(); M 'TypewriterBackspace'
    Check 'backspace works when off' ($d.Content.Text.StartsWith('Hell world'))
    $d.Range(0, 0).Select(); M 'EditClear'
    Check 'delete works when off' ($d.Content.Text.StartsWith('ell world'))
    M 'EditUndo'
    Check 'undo works when off' ($d.Content.Text.StartsWith('Hell world'))
    $d.Range(0, 4).Select(); M 'EditCut'
    Check 'cut works when off' ($d.Content.Text.StartsWith('world'))   # smart cut drops the space

    # Updates
    $v = T 'versions'
    Check 'version comparison' ($v -notmatch 'FAIL')
    'update check: ' + (T 'update')

    ''
    'DIAGNOSTICS:'
    ($w.Run('Typewriter.TypewriterDiagnostics')) -split "`r" | Where-Object { $_ }
    $d.Close(0)
}
catch { 'ERROR: ' + $_.Exception.Message }
finally { $w.Quit(0) }
