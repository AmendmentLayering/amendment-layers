# For each sample .docx: write a flat package (document, numbering, styles) and
# Word's own list strings (read-only COM) into the test\tmp folder beside this script.
param([string[]]$Docs)
$out = Join-Path $PSScriptRoot "tmp"
New-Item -ItemType Directory -Force $out | Out-Null
$work = Join-Path $env:TEMP ("amd-real-" + [guid]::NewGuid())
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$i = 0
$index = @()
foreach ($d in $Docs) {
  $i++
  $x = Join-Path $work "d$i"
  New-Item -ItemType Directory -Force $x | Out-Null
  Copy-Item -LiteralPath $d "$x\a.zip"
  Expand-Archive "$x\a.zip" "$x\u"
  $parts = ""
  foreach ($p in "document","numbering","styles") {
    $f = "$x\u\word\$p.xml"
    if (Test-Path $f) {
      $t = [IO.File]::ReadAllText($f) -replace '^\s*<\?xml[^>]*\?>', ''
      $parts += "<pkg:part pkg:name=""/word/$p.xml"" pkg:contentType=""x""><pkg:xmlData>$t</pkg:xmlData></pkg:part>"
    }
  }
  $pkg = "<pkg:package xmlns:pkg=""http://schemas.microsoft.com/office/2006/xmlPackage"">$parts</pkg:package>"
  [IO.File]::WriteAllText("$out\d$i.xml", $pkg, (New-Object Text.UTF8Encoding $false))
  $doc = $w.Documents.Open($d, $false, $true, $false)
  $ls = @()
  # Skip table end-of-row marks: COM counts them as paragraphs, Office.js doesn't.
  foreach ($para in $doc.Paragraphs) { if (-not ($para.Range.Information(12) -and $para.Range.Cells.Count -eq 0)) { $ls += [string]$para.Range.ListFormat.ListString } }
  $doc.Close($false)
  ($ls | ConvertTo-Json -Compress) | Out-File "$out\d$i.json" -Encoding utf8
  $index += [IO.Path]::GetFileName($d)
}
$w.Quit()
($index | ConvertTo-Json -Compress) | Out-File "$out\index.json" -Encoding utf8
Remove-Item -Recurse -Force $work
"done $i"
