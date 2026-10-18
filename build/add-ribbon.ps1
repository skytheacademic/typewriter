# Adds the Typewriter ribbon tab (customUI14.xml) to Typewriter.dotm.
# Called by build.ps1. Safe to run again after editing the XML.
param(
    [Parameter(Mandatory)][string]$Path,
    [string]$Ui = (Join-Path $PSScriptRoot '..\src\customUI14.xml')
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$uiSource = $Ui
$uiPart = 'customUI/customUI14.xml'
$relType = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
$relNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
$utf8 = New-Object System.Text.UTF8Encoding $false

if (-not (Test-Path $Path)) { throw "Can't find $Path." }
if (Get-Process WINWORD -ErrorAction SilentlyContinue) { throw 'Word is still running. Close it completely and run this again.' }

function Read-Entry($zip, $name) {
    $reader = New-Object System.IO.StreamReader($zip.GetEntry($name).Open())
    try { $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Write-Entry($zip, $name, $text) {
    $old = $zip.GetEntry($name)
    if ($old) { $old.Delete() }
    $writer = New-Object System.IO.StreamWriter($zip.CreateEntry($name).Open(), $utf8)
    try { $writer.Write($text) } finally { $writer.Dispose() }
}

$zip = [System.IO.Compression.ZipFile]::Open($Path, 'Update')
try {
    Write-Entry $zip $uiPart ([System.IO.File]::ReadAllText($uiSource))

    # Point the package at the ribbon XML, replacing any earlier link
    [xml]$rels = Read-Entry $zip '_rels/.rels'
    foreach ($r in @($rels.DocumentElement.ChildNodes | Where-Object { $_.Type -eq $relType })) {
        [void]$rels.DocumentElement.RemoveChild($r)
    }
    $rel = $rels.CreateElement('Relationship', $relNs)
    $rel.SetAttribute('Id', 'rTypewriterUI')
    $rel.SetAttribute('Type', $relType)
    $rel.SetAttribute('Target', $uiPart)
    [void]$rels.DocumentElement.AppendChild($rel)
    Write-Entry $zip '_rels/.rels' $rels.OuterXml

    # Make sure .xml parts have a content type
    [xml]$types = Read-Entry $zip '[Content_Types].xml'
    $hasXml = $types.DocumentElement.ChildNodes | Where-Object { $_.LocalName -eq 'Default' -and $_.Extension -eq 'xml' }
    if (-not $hasXml) {
        $d = $types.CreateElement('Default', $types.DocumentElement.NamespaceURI)
        $d.SetAttribute('Extension', 'xml')
        $d.SetAttribute('ContentType', 'application/xml')
        [void]$types.DocumentElement.PrependChild($d)
        Write-Entry $zip '[Content_Types].xml' $types.OuterXml
    }
}
finally {
    $zip.Dispose()
}

Write-Host "Ribbon tab added to $Path"
