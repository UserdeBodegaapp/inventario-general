param(
  [string]$ExcelPath = ".\DETALLE PARA CONTEO PAREJAS.xlsx",
  [string]$OutputPath = ".\asignaciones.json"
)
$ErrorActionPreference = "Stop"
$SheetName = "LISTADO"
$TableName = "DETALLE"
$Salt = "INV-GENERAL-2026-V1"
function Normalize-ENumber([object]$Value) {
  if ($null -eq $Value) { return "" }
  return ([string]$Value).Trim().ToUpperInvariant() -replace '[^A-Z0-9]', ''
}
function Get-Sha256([string]$Text) {
  $sha = [System.Security.Cryptography.SHA256]::Create()
  try { return -join ($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)) | ForEach-Object { $_.ToString('x2') }) }
  finally { $sha.Dispose() }
}
$fullExcel = [IO.Path]::GetFullPath($ExcelPath)
$fullOutput = [IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path $fullExcel)) { throw "No se encontró el Excel: $fullExcel" }
$excel = $null; $book = $null; $sheet = $null; $table = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false; $excel.DisplayAlerts = $false
  $book = $excel.Workbooks.Open($fullExcel, 0, $true)
  $sheet = $book.Worksheets.Item($SheetName)
  $table = $sheet.ListObjects.Item($TableName)
  $headers = @{}
  for ($c=1; $c -le $table.ListColumns.Count; $c++) { $headers[[string]$table.HeaderRowRange.Cells.Item(1,$c).Text] = $c }
  foreach ($required in @('E-NUMBER','PAREJA','MESA','LISTADO','UBICACION','PRIORIDAD')) {
    if (-not $headers.ContainsKey($required)) { throw "Falta la columna obligatoria: $required" }
  }
  $best = @{}
  if ($null -ne $table.DataBodyRange) {
    for ($r=1; $r -le $table.DataBodyRange.Rows.Count; $r++) {
      $en = Normalize-ENumber $table.DataBodyRange.Cells.Item($r,$headers['E-NUMBER']).Value2
      $pareja = ([string]$table.DataBodyRange.Cells.Item($r,$headers['PAREJA']).Text).Trim()
      $mesa = ([string]$table.DataBodyRange.Cells.Item($r,$headers['MESA']).Text).Trim()
      $listado = ([string]$table.DataBodyRange.Cells.Item($r,$headers['LISTADO']).Text).Trim()
      $ubicacion = ([string]$table.DataBodyRange.Cells.Item($r,$headers['UBICACION']).Text).Trim()
      $prioText = ([string]$table.DataBodyRange.Cells.Item($r,$headers['PRIORIDAD']).Text).Trim()
      $prio = 999999; [void][int]::TryParse($prioText,[ref]$prio)
      if ([string]::IsNullOrWhiteSpace($en) -or $en -eq 'ENCRIPTADO') { continue }
      if ([string]::IsNullOrWhiteSpace($mesa) -or $prioText -match '^[xX]$') { continue }
      $candidate = [pscustomobject]@{ priority=$prio; id=(Get-Sha256 "$Salt|$en"); pareja=$pareja; mesa=$mesa; listado=$listado; ubicacion=$ubicacion }
      if (-not $best.ContainsKey($en) -or $candidate.priority -lt $best[$en].priority) { $best[$en]=$candidate }
    }
  }
  $out = @($best.Values | Sort-Object id | Select-Object id,pareja,mesa,listado,ubicacion)
  $json = ConvertTo-Json -InputObject $out -Depth 4
  [IO.File]::WriteAllText($fullOutput,$json,(New-Object Text.UTF8Encoding($false)))
  Write-Host "[OK] Archivo generado: $fullOutput" -ForegroundColor Green
  Write-Host "[OK] Asignaciones publicadas: $($out.Count)" -ForegroundColor Green
}
finally {
  if ($book) { $book.Close($false) }
  if ($excel) { $excel.Quit() }
  foreach ($obj in @($table,$sheet,$book,$excel)) { if ($obj) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
