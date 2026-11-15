# Checks that the add-in files the live site serves are exactly the files in
# this folder's web\ folder (the released copy). Word runs whatever the site
# serves, so a difference means the site holds something other than this copy:
# a newer or older deploy, or a change made by someone else.
#
# Usage (from this folder):
#   powershell -ExecutionPolicy Bypass -File check-live.ps1
#   powershell -ExecutionPolicy Bypass -File check-live.ps1 -Site https://example.netlify.app
#
# It only downloads and compares; it changes nothing, here or on the site.
# Files on the site that aren't in web\ can't be listed from outside, so they
# aren't checked; web\index.html and web\presenter.html load only the files
# listed in web\ (plus Microsoft's office.js), which a reviewer can confirm.
param(
  [string]$Site = "https://amendment-layers.netlify.app",
  [string]$Folder = (Join-Path $PSScriptRoot "web")
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$sha = [Security.Cryptography.SHA256]::Create()
function Hash([byte[]]$b) { ([BitConverter]::ToString($sha.ComputeHash($b))).Replace("-", "") }

$client = New-Object Net.WebClient
$client.Headers["Cache-Control"] = "no-cache"
$stamp = [DateTime]::UtcNow.Ticks
$files = Get-ChildItem -LiteralPath $Folder -Recurse -File | Sort-Object FullName
$bad = 0

Write-Host "Comparing $Site with $Folder"
foreach ($f in $files) {
  $rel = $f.FullName.Substring($Folder.TrimEnd('\').Length + 1).Replace('\', '/')
  $local = Hash ([IO.File]::ReadAllBytes($f.FullName))
  try {
    # The extra query string asks for a fresh copy rather than a cached one.
    $live = Hash ($client.DownloadData("$Site/$rel" + "?check=$stamp"))
  } catch {
    Write-Host ("  MISSING  {0}  ({1})" -f $rel, $_.Exception.Message) -ForegroundColor Red
    $bad++
    continue
  }
  if ($live -eq $local) {
    Write-Host ("  same     {0}" -f $rel)
  } else {
    Write-Host ("  DIFFERS  {0}" -f $rel) -ForegroundColor Red
    $bad++
  }
}

if ($bad -eq 0) {
  Write-Host "All $($files.Count) files on the site match the released copy." -ForegroundColor Green
  exit 0
}
Write-Host "$bad of $($files.Count) files don't match. If you didn't publish a different version, check the site's Deploys page." -ForegroundColor Red
exit 1
