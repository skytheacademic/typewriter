# Builds dist\Typewriter.dotm from the files in src\.
#
# Needs Word for Windows, and "Trust access to the VBA project object model"
# switched on (Word > File > Options > Trust Center > Trust Center Settings >
# Macro Settings). Switch it off again when you're done.
#
#   pwsh build\build.ps1            build dist\Typewriter.dotm
#   pwsh build\build.ps1 -Install   also copy it into Word's Startup folder
param(
    [string]$Out = (Join-Path $PSScriptRoot '..\dist\Typewriter.dotm'),
    [switch]$Install
)
$ErrorActionPreference = 'Stop'
$src = Resolve-Path (Join-Path $PSScriptRoot '..\src')
$Out = [IO.Path]::GetFullPath($Out)
$startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'

if (Get-Process WINWORD -ErrorAction SilentlyContinue) { throw 'Close Word before building.' }
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
if (Test-Path $Out) { Remove-Item $Out }

$w = New-Object -ComObject Word.Application
try {
    # An installed copy would answer the macro calls below instead of the new one
    foreach ($a in $w.AddIns) { if ($a.Name -eq 'Typewriter.dotm') { $a.Installed = $false } }

    $d = $w.Documents.Add()
    $d.VBProject.Name = 'TypewriterAddin'
    [void]$d.VBProject.VBComponents.Import((Join-Path $src 'Typewriter.bas'))
    [void]$d.VBProject.VBComponents.Import((Join-Path $src 'TypewriterEvents.cls'))
    $d.SaveAs2([string]$Out, 15)   # 15 = macro-enabled template

    # Compiles the code and stores the Windows keyboard shortcuts in the file
    [void]$w.Run('Typewriter.TypewriterTest', 'install')
    $d.Save()
    $d.Close(0)
}
finally {
    $w.Quit(0)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
}

# Word takes a moment to quit and let go of the file
$deadline = (Get-Date).AddSeconds(60)
while ((Get-Process WINWORD -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 500
}
& (Join-Path $PSScriptRoot 'add-ribbon.ps1') -Path $Out
Write-Host "Built $Out"

if ($Install) {
    New-Item -ItemType Directory -Force $startup | Out-Null
    Copy-Item $Out (Join-Path $startup 'Typewriter.dotm') -Force
    Write-Host "Installed to $startup"
}
