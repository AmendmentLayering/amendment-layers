# Has Word fill in a document's table of contents (page numbers and links),
# then saves it. Called by build-docs.js after it writes the User Manual.
#   powershell -ExecutionPolicy Bypass -File docs\update-toc.ps1 "<full path to .docx>"
param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'

$word = New-Object -ComObject Word.Application
# If Word handed back a copy that already has documents open, it is the
# user's own Word: close only our document and leave Word running.
$ours = $word.Documents.Count -eq 0
if ($ours) { $word.Visible = $false }
$word.DisplayAlerts = 0
try {
  # Word's optional arguments must be passed as [ref] from PowerShell, so
  # rely on the defaults and close after saving.
  $doc = $word.Documents.Open($Path)
  try {
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
    $doc.Save()
  } finally {
    $doc.Close()
  }
} finally {
  if ($ours) { $word.Quit() }
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
