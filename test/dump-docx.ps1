# Prints the runs (text + amendment formatting) and add-in content controls of
# a .docx, plus the add-in's stored settings. Diagnostic aid.
# Usage:  powershell -File test\dump-docx.ps1 <file.docx>
param([Parameter(Mandatory)][string]$Path)
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead((Resolve-Path $Path))
try {
  function Read-Entry($name) {
    $e = $zip.GetEntry($name); if (-not $e) { return $null }
    $r = New-Object IO.StreamReader($e.Open()); try { $r.ReadToEnd() } finally { $r.Close() }
  }
  [xml]$doc = Read-Entry "word/document.xml"
  $ns = New-Object Xml.XmlNamespaceManager $doc.NameTable
  $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

  function Describe-Run($r) {
    $rPr = $r.SelectSingleNode("w:rPr", $ns)
    $bits = @()
    foreach ($n in "i", "strike", "dstrike", "noProof") { if ($rPr -and $rPr.SelectSingleNode("w:$n", $ns)) { $bits += $n } }
    if ($rPr) {
      $c = $rPr.SelectSingleNode("w:color", $ns); if ($c) { $bits += "color=" + $c.GetAttribute("val", $ns.LookupNamespace("w")) }
      $u = $rPr.SelectSingleNode("w:u", $ns); if ($u) { $bits += "u=" + $u.GetAttribute("val", $ns.LookupNamespace("w")) }
      $s = $rPr.SelectSingleNode("w:shd", $ns); if ($s) { $bits += "shd=" + $s.GetAttribute("fill", $ns.LookupNamespace("w")) }
      $st = $rPr.SelectSingleNode("w:rStyle", $ns); if ($st) { $bits += "style=" + $st.GetAttribute("val", $ns.LookupNamespace("w")) }
    }
    $text = ($r.SelectNodes("w:t", $ns) | ForEach-Object { $_.InnerText }) -join ""
    "[{0}] {1}" -f $text, ($bits -join ",")
  }

  function Walk($node, $depth) {
    foreach ($c in $node.ChildNodes) {
      if ($c.LocalName -eq "r") { ("  " * $depth) + (Describe-Run $c) }
      elseif ($c.LocalName -eq "sdt") {
        $tag = $c.SelectSingleNode("w:sdtPr/w:tag", $ns)
        $tv = if ($tag) { $tag.GetAttribute("val", $ns.LookupNamespace("w")) } else { "(no tag)" }
        $sdtRpr = $c.SelectSingleNode("w:sdtPr/w:rPr", $ns)
        ("  " * $depth) + "<sdt $tv>" + $(if ($sdtRpr) { " sdtPr/rPr=" + $sdtRpr.InnerXml } else { "" })
        Walk $c.SelectSingleNode("w:sdtContent", $ns) ($depth + 1)
        ("  " * $depth) + "</sdt>"
      }
      elseif ($c.LocalName -in "p", "hyperlink", "ins", "smartTag") {
        if ($c.LocalName -eq "p") {
          $txt = $c.InnerText
          $num = if ($c.SelectSingleNode("w:pPr/w:numPr", $ns)) { " numbered" } else { "" }
          $markRun = $c.SelectSingleNode("w:pPr", $ns)
          $mark = if ($markRun -and $markRun.SelectSingleNode("w:rPr", $ns)) { " mark " + (Describe-Run $markRun) } else { "" }
          if (-not $txt.Trim() -and -not $num) { continue }
          ("  " * $depth) + "P:$num$mark"
        }
        Walk $c ($depth + 1)
      }
      elseif ($c.LocalName -in "body", "tbl", "tr", "tc") { Walk $c $depth }
    }
  }
  Walk $doc.DocumentElement.SelectSingleNode("w:body", $ns) 0

  "---- add-in settings"
  foreach ($e in $zip.Entries | Where-Object { $_.FullName -like "word/webextensions/webextension*.xml" }) {
    [xml]$we = Read-Entry $e.FullName
    foreach ($p in $we.GetElementsByTagName("we:property")) {
      $name = $p.GetAttribute("name"); $val = $p.GetAttribute("value")
      if ($name -eq "amendmentLayers") {
        $state = $val | ConvertFrom-Json
        "activeLevel=$($state.activeLevel) openThreadId=$($state.openThreadId) seq=$($state.seq)"
        foreach ($t in $state.threads.PSObject.Properties.Value) {
          "thread $($t.id) level=$($t.level) parent=$($t.parentId) status=$($t.status)"
          foreach ($a in $t.anchors) {
            "   anchor $($a.tag)"
            if ($a.prior) {
              [xml]$px = $a.prior
              $pns = New-Object Xml.XmlNamespaceManager $px.NameTable
              $pns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
              $script:ns = $pns
              foreach ($r in $px.SelectNodes("//w:body//w:r", $pns)) { "      prior " + (Describe-Run $r) }
              $script:ns = $ns
            } elseif ($null -ne $a.prior) { "      prior (empty)" }
          }
        }
      } elseif ($name -eq "amendmentLayersLastMiss") {
        $miss = $val | ConvertFrom-Json
        "last 'no amendment at cursor': $($miss.at) selectionEmpty=$($miss.selectionEmpty)"
        foreach ($c in $miss.checked) { "   checked $c" }
        foreach ($t in $miss.notInDocument) { "   NOT IN DOCUMENT: $t" }
      } else { "$name = $val" }
    }
  }
} finally { $zip.Dispose() }
