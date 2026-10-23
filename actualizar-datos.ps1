param(
    [string]$ExcelPath = ".\CONSULTA INVENTARIO.xlsx",
    [string]$OutputPath = ".\asignaciones.json"
)

$ErrorActionPreference = "Stop"
$Salt = "INV-GENERAL-2026-V1"
$SheetName = "LISTADO"
$TableName = "DETALLE"

function Normalize-Text([object]$Value) {
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

function Normalize-ENumber([object]$Value) {
    return (Normalize-Text $Value).ToUpperInvariant() -replace '[^A-Z0-9]', ''
}

function Test-Active([object]$Value) {
    $v = (Normalize-Text $Value).ToUpperInvariant()
    return $v -in @('SI','SÍ','YES','TRUE','1','ACTIVO')
}

function Get-Sha256([string]$Text) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return -join ($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)) | ForEach-Object { $_.ToString('x2') })
    }
    finally { $sha.Dispose() }
}

$required = @(
    'ENUMBER',
    'N° PAREJA',
    'MESA',
    'LISTADO',
    'UBICACIÓN A CONTAR',
    'HORARIO DE COLACIÓN',
    'HORARIO DE ALMUERZO',
    'ACTIVO'
)

$fullExcel = [IO.Path]::GetFullPath($ExcelPath)
$fullOutput = [IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path $fullExcel)) { throw "No se encontró el Excel: $fullExcel" }

$excel = $null; $book = $null; $sheet = $null; $table = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($fullExcel, 0, $true)
    $sheet = $book.Worksheets.Item($SheetName)
    $table = $sheet.ListObjects.Item($TableName)

    $headers = @{}
    for ($c = 1; $c -le $table.ListColumns.Count; $c++) {
        $headers[[string]$table.HeaderRowRange.Cells.Item(1,$c).Text] = $c
    }
    foreach ($name in $required) {
        if (-not $headers.ContainsKey($name)) { throw "Falta la columna obligatoria: $name" }
    }

    $seen = @{}
    $output = @()
    $omittedInactive = 0
    $omittedDuplicate = 0
    $omittedEmpty = 0

    if ($null -ne $table.DataBodyRange) {
        for ($r = 1; $r -le $table.DataBodyRange.Rows.Count; $r++) {
            $active = Test-Active $table.DataBodyRange.Cells.Item($r,$headers['ACTIVO']).Value2
            if (-not $active) { $omittedInactive++; continue }

            $en = Normalize-ENumber $table.DataBodyRange.Cells.Item($r,$headers['ENUMBER']).Value2
            if ([string]::IsNullOrWhiteSpace($en) -or $en -eq 'ENCRIPTADO') { $omittedEmpty++; continue }
            if ($seen.ContainsKey($en)) { $omittedDuplicate++; continue }
            $seen[$en] = $true

            $output += [pscustomobject]@{
                id        = Get-Sha256 "$Salt|$en"
                pareja    = Normalize-Text $table.DataBodyRange.Cells.Item($r,$headers['N° PAREJA']).Text
                mesa      = Normalize-Text $table.DataBodyRange.Cells.Item($r,$headers['MESA']).Text
                listado   = Normalize-Text $table.DataBodyRange.Cells.Item($r,$headers['LISTADO']).Text
                ubicacion = Normalize-Text $table.DataBodyRange.Cells.Item($r,$headers['UBICACIÓN A CONTAR']).Text
                colacion  = Normalize-Text $table.DataBodyRange.Cells.Item($r,$headers['HORARIO DE COLACIÓN']).Text
                almuerzo  = Normalize-Text $table.DataBodyRange.Cells.Item($r,$headers['HORARIO DE ALMUERZO']).Text
            }
        }
    }

    $json = ConvertTo-Json -InputObject @($output | Sort-Object id) -Depth 4
    [IO.File]::WriteAllText($fullOutput, $json, (New-Object Text.UTF8Encoding($false)))

    Write-Host "[OK] Archivo generado: $fullOutput" -ForegroundColor Green
    Write-Host "[OK] Registros activos publicados: $($output.Count)" -ForegroundColor Green
    Write-Host "[INFO] Inactivos omitidos: $omittedInactive" -ForegroundColor Cyan
    Write-Host "[INFO] E-Number vacíos/ENCRIPTADO omitidos: $omittedEmpty" -ForegroundColor Cyan
    Write-Host "[INFO] E-Number duplicados omitidos: $omittedDuplicate" -ForegroundColor Cyan
}
finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($obj in @($table,$sheet,$book,$excel)) {
        if ($obj) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
