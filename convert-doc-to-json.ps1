Add-Type -AssemblyName System.IO.Compression.FileSystem

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Api = Join-Path $Root "api"
New-Item -ItemType Directory -Force -Path $Api | Out-Null

function Remove-Diacritics([string]$Text) {
  if ($null -eq $Text) { return "" }
  $normalized = $Text.Normalize([Text.NormalizationForm]::FormD)
  return ([regex]::Replace($normalized, '\p{Mn}', '')).ToLowerInvariant().Trim()
}

function Clean([object]$Value) {
  if ($null -eq $Value) { return "" }
  return ([string]$Value -replace '\s+', ' ').Trim()
}

function Excel-Date([object]$Value) {
  if ($null -eq $Value -or "$Value" -eq "") { return $null }
  $number = 0.0
  if ([double]::TryParse("$Value", [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
    try { return ([DateTime]::FromOADate($number)).ToString("yyyy-MM-dd") } catch { return $null }
  }
  $date = [DateTime]::MinValue
  if ([DateTime]::TryParse("$Value", [Globalization.CultureInfo]::GetCultureInfo("pt-BR"), [Globalization.DateTimeStyles]::None, [ref]$date)) {
    return $date.ToString("yyyy-MM-dd")
  }
  return $null
}

function Sheet-CellIndex([string]$Ref) {
  if ($Ref -notmatch '^([A-Z]+)') { return 0 }
  $letters = $Matches[1]
  $index = 0
  foreach ($char in $letters.ToCharArray()) {
    $index = ($index * 26) + ([int][char]$char - [int][char]'A' + 1)
  }
  return $index - 1
}

function Read-Xlsx([string]$Path) {
  $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
  try {
    $shared = New-Object System.Collections.ArrayList
    $sharedEntry = $zip.GetEntry("xl/sharedStrings.xml")
    if ($sharedEntry) {
      $reader = New-Object IO.StreamReader($sharedEntry.Open())
      [xml]$sharedXml = $reader.ReadToEnd()
      $reader.Close()
      foreach ($si in $sharedXml.GetElementsByTagName("si")) {
        [void]$shared.Add($si.InnerText)
      }
    }

    $reader = New-Object IO.StreamReader($zip.GetEntry("xl/workbook.xml").Open())
    $workbookXml = $reader.ReadToEnd()
    $reader.Close()
    $reader = New-Object IO.StreamReader($zip.GetEntry("xl/_rels/workbook.xml.rels").Open())
    $relsXml = $reader.ReadToEnd()
    $reader.Close()

    $relMap = @{}
    [regex]::Matches($relsXml, '<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"') | ForEach-Object {
      $relMap[$_.Groups[1].Value] = $_.Groups[2].Value
    }

    $sheets = @{}
    [regex]::Matches($workbookXml, '<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"') | ForEach-Object {
      $name = [System.Net.WebUtility]::HtmlDecode($_.Groups[1].Value)
      $target = $relMap[$_.Groups[2].Value]
      if (-not $target) { return }
      $entry = $zip.GetEntry(("xl/" + $target.TrimStart("/")))
      if (-not $entry) { return }
      $reader = New-Object IO.StreamReader($entry.Open())
      [xml]$sheetXml = $reader.ReadToEnd()
      $reader.Close()

      $matrix = New-Object System.Collections.ArrayList
      foreach ($rowNode in $sheetXml.GetElementsByTagName("row")) {
        $rowIndex = [int]$rowNode.GetAttribute("r") - 1
        while ($matrix.Count -le $rowIndex) { [void]$matrix.Add((New-Object System.Collections.ArrayList)) }
        $row = $matrix[$rowIndex]
        foreach ($cell in $rowNode.GetElementsByTagName("c")) {
          $cellIndex = Sheet-CellIndex $cell.GetAttribute("r")
          while ($row.Count -le $cellIndex) { [void]$row.Add($null) }
          $value = ""
          $vNodes = $cell.GetElementsByTagName("v")
          if ($vNodes.Count -gt 0) { $value = $vNodes[0].InnerText }
          elseif ($cell.GetElementsByTagName("is").Count -gt 0) { $value = $cell.GetElementsByTagName("is")[0].InnerText }
          if ($cell.GetAttribute("t") -eq "s" -and $value -match '^\d+$') {
            $idx = [int]$value
            if ($idx -lt $shared.Count) { $value = $shared[$idx] }
          }
          $row[$cellIndex] = $value
        }
      }
      $sheets[$name] = $matrix
    }
    return $sheets
  }
  finally {
    $zip.Dispose()
  }
}

function Discipline-Base([string]$Value) {
  $n = Remove-Diacritics ($Value -replace '\([^)]*\)', '')
  if ($n -match 'eletr') { return "Elétrica" }
  if ($n -match 'hidraul|ar-condicionado|exaust') { return "Hidráulica" }
  if ($n -match 'revest') { return "Revestimento" }
  if ($n -match 'acab|drywall|pintura|forro') { return "Acabamentos" }
  if ($n -match 'estrut|piso|laminado') { return "Estruturas" }
  $clean = Clean $Value
  if ($clean) { return $clean }
  return "Não informado"
}

function Infer-Discipline([string]$Text) {
  $n = Remove-Diacritics $Text
  if ($n -match 'eletr') { return "Elétrica" }
  if ($n -match 'hidraul|exaust|ar-condicionado|loucas') { return "Hidráulica" }
  if ($n -match 'revest|impermeabil') { return "Revestimento" }
  if ($n -match 'acab|drywall|pintura|forro') { return "Acabamentos" }
  if ($n -match 'estrut|laminado') { return "Estruturas" }
  return "Não informado"
}

function Infer-Pavimento([string]$Text) {
  $n = Remove-Diacritics $Text
  $found = New-Object System.Collections.ArrayList
  if ($n -match 'piso 03|3 andar|3°|3º') { [void]$found.Add("3º") }
  if ($n -match 'piso 02|2 andar|2°|2º') { [void]$found.Add("2º") }
  if ($n -match 'piso 01|1 andar|1°|1º') { [void]$found.Add("1º") }
  if ($n -match 'terreo|-t| t ') { [void]$found.Add("Térreo") }
  return (($found | Select-Object -Unique) -join " / ")
}

function Find-HeaderRow($Matrix) {
  $best = -1
  $bestScore = 0
  for ($i = 0; $i -lt $Matrix.Count; $i++) {
    $text = Remove-Diacritics (($Matrix[$i] | ForEach-Object { Clean $_ }) -join " | ")
    $score = 0
    foreach ($k in @("codigo", "item", "descricao", "qtd", "quantidade", "dimensao", "estoque")) {
      if ($text.Contains($k)) { $score++ }
    }
    if ($score -gt $bestScore) { $best = $i; $bestScore = $score }
  }
  if ($bestScore -ge 2) { return $best }
  return -1
}

function Find-Groups($Headers) {
  $groups = New-Object System.Collections.ArrayList
  for ($i = 0; $i -lt $Headers.Count; $i++) {
    $h = Remove-Diacritics $Headers[$i]
    if (-not $h.Contains("codigo")) { continue }
    $find = {
      param($start, $limit, $patterns)
      for ($j = $start; $j -le [Math]::Min($Headers.Count - 1, $start + $limit); $j++) {
        $x = Remove-Diacritics $Headers[$j]
        foreach ($p in $patterns) { if ($x -match $p) { return $j } }
      }
      return -1
    }
    $desc = & $find $i 5 @("descricao", "af e aq", "exaustao", "infra")
    $qty = & $find $i 6 @("^qtd$", "quantidade")
    if ($desc -ge 0 -and $qty -ge 0) {
      [void]$groups.Add([pscustomobject]@{
        code = $i
        description = $desc
        quantity = $qty
        dimension = (& $find $i 7 @("dimensao", "^und$", "coluna"))
        delivery = (& $find $i 9 @("entrega"))
        order = (& $find $i 9 @("encomendar", "reordenar"))
        stock = (& $find $i 9 @("estoque"))
      })
    }
  }
  return $groups
}

function Find-NearbyHeader($Matrix, [int]$HeaderIndex, [int]$Column, [string]$Keyword) {
  for ($r = 0; $r -lt $HeaderIndex; $r++) {
    $row = $Matrix[$r]
    for ($c = [Math]::Max(0, $Column - 3); $c -le [Math]::Min($row.Count - 1, $Column + 3); $c++) {
      $value = Clean $row[$c]
      if ((Remove-Diacritics $value).Contains($Keyword)) {
        return ($value -replace '^Nome do Kit:\s*', '' -replace '^LOCAL:\s*', '')
      }
    }
  }
  return ""
}

function Needs-Order([string]$Value) {
  $n = Remove-Diacritics $Value
  if (-not $n -or $n -in @("nao", "não", "no", "false", "0", "-")) { return $false }
  $num = 0.0
  if ([double]::TryParse(($Value -replace ',', '.' -replace '[^\d.-]', ''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
    if ($num -gt 0) { return $true }
  }
  return ($n -match 'sim|comprar|pedido')
}

function No-Stock([string]$Value) {
  $n = Remove-Diacritics $Value
  return ($n -match 'nao|sem' -or $n -eq "0")
}

$cronogramaName = "Cronograma Residenciais.xlsx"
$cronogramaSheets = Read-Xlsx (Join-Path $Root $cronogramaName)
$cronogramaSheet = ($cronogramaSheets.Keys | Where-Object { (Remove-Diacritics $_).Contains("cronograma") } | Select-Object -First 1)
if (-not $cronogramaSheet) { $cronogramaSheet = $cronogramaSheets.Keys | Select-Object -First 1 }
$matrix = $cronogramaSheets[$cronogramaSheet]
$headers = $matrix[0] | ForEach-Object { Clean $_ }
$activities = New-Object System.Collections.ArrayList
for ($r = 1; $r -lt $matrix.Count; $r++) {
  $row = $matrix[$r]
  if (-not $row) { continue }
  $obj = @{}
  for ($c = 0; $c -lt $headers.Count; $c++) {
    if ($headers[$c]) { $obj[$headers[$c]] = Clean $row[$c] }
  }
  $bloco = Clean $obj["Bloco"]
  $pavimento = Clean $obj["Pavimento"]
  $pacote = Clean $obj["Pacote"]
  $tarefa = Clean $obj["Tarefa"]
  $discRaw = Clean $obj["Disciplina"]
  if (-not ($bloco -or $tarefa -or $discRaw)) { continue }
  $disc = Discipline-Base $discRaw
  [void]$activities.Add([pscustomobject]@{
    id = "ACT-$r"
    sourceRow = $r + 1
    bloco = $bloco
    pavimento = $pavimento
    pacote = $pacote
    tarefa = $tarefa
    disciplinaRaw = $discRaw
    discipline = $disc
    inicio = Excel-Date $obj["Início"]
    fim = Excel-Date $obj["Fim"]
    duracao = [double]($obj["Duração_dias_úteis"] -replace ',', '.')
    icamento = [double]($obj["Içamento_total"] -replace ',', '.')
    raw = $obj
  })
}

$kitFiles = Get-ChildItem -LiteralPath $Root -Filter "*.xlsx" |
  Where-Object { $_.Name -like "Kit - *" } |
  Sort-Object Name
$kitItems = New-Object System.Collections.ArrayList
foreach ($kitFile in $kitFiles) {
  $kitName = $kitFile.Name
  $sheets = Read-Xlsx $kitFile.FullName
  foreach ($sheetName in $sheets.Keys) {
    $m = $sheets[$sheetName]
    $headerIndex = Find-HeaderRow $m
    if ($headerIndex -lt 0) { continue }
    $h = $m[$headerIndex] | ForEach-Object { Clean $_ }
    $groups = Find-Groups $h
    for ($g = 0; $g -lt $groups.Count; $g++) {
      $group = $groups[$g]
      $kitTitle = Find-NearbyHeader $m $headerIndex $group.code "kit"
      if (-not $kitTitle) { $kitTitle = $sheetName }
      $local = Find-NearbyHeader $m $headerIndex $group.code "local"
      if (-not $local) { $local = Infer-Pavimento "$sheetName $kitTitle" }
      $disc = Infer-Discipline "$kitName $sheetName $kitTitle"
      for ($r = $headerIndex + 1; $r -lt $m.Count; $r++) {
        $row = $m[$r]
        if (-not $row) { continue }
        $code = Clean $row[$group.code]
        $desc = Clean $row[$group.description]
        $qty = Clean $row[$group.quantity]
        if (-not ($code -or $desc -or $qty)) { continue }
        if (-not $code -and $desc -and -not $qty) { continue }
        $dim = if ($group.dimension -ge 0) { Clean $row[$group.dimension] } else { "" }
        $stock = if ($group.stock -ge 0) { Clean $row[$group.stock] } else { "" }
        $order = if ($group.order -ge 0) { Clean $row[$group.order] } else { "" }
        $delivery = if ($group.delivery -ge 0) { Excel-Date $row[$group.delivery] } else { $null }
        [void]$kitItems.Add([pscustomobject]@{
          id = "KIT-$($kitItems.Count + 1)"
          fileName = $kitName
          sheetName = $sheetName
          kitName = $kitTitle
          local = $local
          pavimento = Infer-Pavimento "$local $sheetName $kitTitle"
          discipline = $disc
          code = $code
          description = $desc
          quantity = $qty
          dimension = $dim
          estoque = $stock
          encomendar = $order
          entrega = $delivery
          rowNumber = $r + 1
          incomplete = (-not $code) -or (-not $desc) -or (-not $qty) -or (Needs-Order $order) -or (No-Stock $stock)
        })
      }
    }
  }
}

$documents = New-Object System.Collections.ArrayList
foreach ($pdfFile in (Get-ChildItem -LiteralPath $Root -Filter "*.pdf" | Sort-Object Name)) {
  $pdfName = $pdfFile.Name
  $path = $pdfFile.FullName
  $bytes = [IO.File]::ReadAllBytes($path)
  $text = [Text.Encoding]::GetEncoding(28591).GetString($bytes)
  $pageMatches = [regex]::Matches($text, '/Type\s*/Page\b')
  $imageMatches = [regex]::Matches($text, '/Subtype\s*/Image')
  [void]$documents.Add([pscustomobject]@{
    name = $pdfName
    pages = [Math]::Max(1, $pageMatches.Count)
    text = ""
    terms = @()
    bytes = @()
    needsOcr = ($imageMatches.Count -gt 0)
    ocrStatus = ""
    warning = if ($imageMatches.Count -gt 0) { "PDF escaneado/imagem: texto não extraível na API JSON sem OCR." } else { "Texto de PDF não extraído no pré-processamento." }
  })
}

function Save-Json($Name, $Data) {
  $json = $Data | ConvertTo-Json -Depth 30 -Compress
  [IO.File]::WriteAllText((Join-Path $Api $Name), $json, [Text.UTF8Encoding]::new($false))
}

Save-Json "activities.json" $activities
Save-Json "kit-items.json" $kitItems
Save-Json "documents.json" $documents
Save-Json "manifest.json" ([pscustomobject]@{
  version = 1
  generatedAt = (Get-Date).ToString("s")
  source = "doc"
  endpoints = [pscustomobject]@{
    activities = "activities.json"
    kitItems = "kit-items.json"
    documents = "documents.json"
  }
  counts = [pscustomobject]@{
    activities = $activities.Count
    kitItems = $kitItems.Count
    documents = $documents.Count
  }
})

Write-Output "activities=$($activities.Count)"
Write-Output "kitItems=$($kitItems.Count)"
Write-Output "documents=$($documents.Count)"
Write-Output "api=$Api"
