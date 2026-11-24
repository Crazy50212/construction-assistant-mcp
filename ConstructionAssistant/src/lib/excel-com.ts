import { access, copyFile, mkdtemp, rm, writeFile } from "node:fs/promises";
import { constants } from "node:fs";
import os from "node:os";
import path from "node:path";
import { run } from "./exec.js";
import { inferirFormatosColumnas, type FormatoColumnaExcel } from "./excel-formatos.js";
import { PATRON_HOJA_PREDETERMINADA } from "./excel-hojas.js";
import { crearPlanHojaExcel, type HojaConstruction, type PlanHojaExcel } from "./excel-plan.js";
import {
  FORMULA_CANTIDAD_DESDE_CUBICACION,
  FORMULA_CANTIDAD_DESDE_PRESUPUESTO,
  FORMULA_PRECIO_DESDE_PRESUPUESTO,
} from "./excel-sync-templates.js";

const POWERSHELL_SCRIPT = String.raw`
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
$script:ConstructionProfile = $false
$script:ConstructionTimings = [ordered]@{}
$script:ConstructionStopwatch = [Diagnostics.Stopwatch]::StartNew()

function Mark-ConstructionPhase([string]$name) {
  if (-not $script:ConstructionProfile) { return }
  $script:ConstructionTimings[$name] = $script:ConstructionStopwatch.ElapsedMilliseconds
  $script:ConstructionStopwatch.Restart()
}

function Set-ConstructionCell($targetSheet, [int]$targetRow, [int]$targetColumn, $value) {
  $cell = $targetSheet.Cells.Item($targetRow, $targetColumn)
  if ($null -eq $value) { $cell.Value2 = '' }
  elseif ($value -is [string]) {
    $cell.NumberFormat = '@'
    $text = [string]$value
    if ($text -match '^[=+@-]') { $text = "'" + $text }
    $cell.Value2 = $text
  } else { $cell.Value2 = [double]$value }
}

function Set-ConstructionRows($targetSheet, [int]$firstRow, [object[]]$records, [int]$columnCount) {
  if ($records.Count -eq 0) { return }
  $lastRow = $firstRow + $records.Count - 1
  $matrix = New-Object 'object[,]' $records.Count, $columnCount
  $textColumns = @{}
  for ($row = 0; $row -lt $records.Count; $row++) {
    $record = $records[$row]
    for ($column = 0; $column -lt $columnCount; $column++) {
      $value = if ($column -lt $record.Count) { $record[$column] } else { $null }
      if ($null -eq $value) { $matrix[$row, $column] = '' }
      elseif ($value -is [string]) {
        $text = [string]$value
        if ($text -match '^[=+@-]') { $text = "'" + $text }
        $matrix[$row, $column] = $text
        $textColumns[$column + 1] = $true
      } else { $matrix[$row, $column] = [double]$value }
    }
  }
  foreach ($column in $textColumns.Keys) {
    $textRange = $targetSheet.Range($targetSheet.Cells.Item($firstRow, [int]$column), $targetSheet.Cells.Item($lastRow, [int]$column))
    $textRange.NumberFormat = '@'
  }
  $range = $targetSheet.Range($targetSheet.Cells.Item($firstRow, 1), $targetSheet.Cells.Item($lastRow, $columnCount))
  $range.Value2 = $matrix
}

function Resolve-Formula([string]$template, [hashtable]$values) {
  $formula = $template
  foreach ($key in $values.Keys) { $formula = $formula.Replace("{$key}", [string]$values[$key]) }
  return $formula
}

function Get-ExcelArrayValue($values, [int]$row, [int]$column = 1) {
  if ($values -isnot [System.Array]) { return $values }
  if ($values.Rank -eq 2) {
    return $values.GetValue($values.GetLowerBound(0) + $row - 1, $values.GetLowerBound(1) + $column - 1)
  }
  if ($values.Rank -eq 1) { return $values.GetValue($values.GetLowerBound(0) + $row - 1) }
  return $values
}

function Test-ExcelNumber($value) {
  return ($value -is [double] -or $value -is [single] -or $value -is [int] -or $value -is [long] -or $value -is [decimal])
}

function Get-AuditKeyPart($value) {
  if (Test-ExcelNumber $value) { return ([double]$value).ToString('G17', [Globalization.CultureInfo]::InvariantCulture) }
  return ([string]$value).Trim().ToLowerInvariant()
}

function Ensure-ConstructionTable($targetBook, $targetSheet, [string]$tableName, [string]$lastColumn, [int]$lastDataRow) {
  if ($lastDataRow -lt 3) { throw "La tabla '$tableName' necesita al menos una fila de datos." }
  $table = $null
  foreach ($worksheet in $targetBook.Worksheets) {
    foreach ($candidate in $worksheet.ListObjects) {
      if ($candidate.Name -eq $tableName) {
        if ($worksheet.Name -ne $targetSheet.Name) { throw "El nombre de tabla '$tableName' ya está ocupado en otra hoja." }
        $table = $candidate
        break
      }
    }
  }
  $rangeAddress = 'A2:{0}{1}' -f $lastColumn, $lastDataRow
  $range = $targetSheet.Range($rangeAddress)
  if ($table -eq $null) {
    if ($targetSheet.ListObjects.Count -gt 0) { throw "La hoja '$($targetSheet.Name)' ya contiene una tabla no administrada; no se modificó." }
    $table = $targetSheet.ListObjects.Add(1, $range, $null, 1)
    $table.Name = $tableName
  } else {
    $table.Resize($range) | Out-Null
  }
  $table.TableStyle = 'TableStyleLight1'
  $table.ShowTableStyleRowStripes = $true
  return $table
}

function Set-ConstructionTableFormula($table, [string]$columnName, [string]$formula) {
  $range = $table.ListColumns.Item($columnName).DataBodyRange
  if ($range -eq $null) { return }
  $range.NumberFormat = 'General'
  $range.Formula = $formula
  if ($range.HasFormula -ne $true) { throw "No se pudo establecer la fórmula calculada de la columna '$columnName'." }
}

function New-LinkedFormula([string]$template, [int]$row, [string]$fallbackExpression) {
  return $template.Replace('{fila}', [string]$row).Replace('{fallback}', $fallbackExpression)
}

function Sync-BudgetCubicLinks($budgetSheet, $budgetTable, $cubicTable, [string]$linkTemplate) {
  if ($budgetTable -eq $null -or $cubicTable -eq $null) { return }
  $knownConcepts = @{}
  $conceptValues = $cubicTable.ListColumns.Item('Concepto').DataBodyRange.Value2
  for ($row = 1; $row -le $cubicTable.DataBodyRange.Rows.Count; $row++) {
    $knownConcepts[[string](Get-ExcelArrayValue $conceptValues $row)] = $true
  }
  $budgetValues = $budgetTable.DataBodyRange.Value2
  $quantityFormulas = $budgetTable.ListColumns.Item('Cantidad').DataBodyRange.Formula
  for ($row = 1; $row -le $budgetTable.DataBodyRange.Rows.Count; $row++) {
    $worksheetRow = $row + 2
    $concept = [string](Get-ExcelArrayValue $budgetValues $row 1)
    $unit = ([string](Get-ExcelArrayValue $budgetValues $row 2)).Trim().ToLowerInvariant()
    $isSupported = $unit -in @('m2', 'm²', 'm^2', 'm3', 'm³', 'm^3', 'm3 ajustado', 'm³ ajustado')
    $currentFormula = [string](Get-ExcelArrayValue $quantityFormulas $row)
    $isLinked = $currentFormula -match 'tblCubicaciones'
    if (-not $isLinked -and -not $knownConcepts.ContainsKey($concept)) { continue }
    $quantityCell = $budgetSheet.Cells.Item($worksheetRow, 3)
    if (-not $isSupported) {
      $quantityCell.Formula = '=""'
      continue
    }
    if ($isLinked -and $currentFormula -match ',\s*""\)$') { continue }
    $quantityCell.Formula = New-LinkedFormula $linkTemplate $worksheetRow '""'
  }
}

function Sync-AdvanceBudgetLinks($advanceSheet, $advanceTable, $budgetTable, [string]$quantityTemplate, [string]$priceTemplate) {
  if ($advanceTable -eq $null -or $budgetTable -eq $null) { return }
  $budgetKeys = @{}
  $budgetValues = $budgetTable.DataBodyRange.Value2
  for ($row = 1; $row -le $budgetTable.DataBodyRange.Rows.Count; $row++) {
    $budgetName = [string](Get-ExcelArrayValue $budgetValues $row 1)
    $budgetUnit = Normalize-ConstructionUnit ([string](Get-ExcelArrayValue $budgetValues $row 2))
    $budgetKeys[($budgetName + [char]31 + $budgetUnit)] = $true
  }
  $advanceValues = $advanceTable.DataBodyRange.Value2
  $quantityFormulas = $advanceTable.ListColumns.Item('Cantidad presupuestada').DataBodyRange.Formula
  $priceFormulas = $advanceTable.ListColumns.Item('Precio unitario').DataBodyRange.Formula
  for ($row = 1; $row -le $advanceTable.DataBodyRange.Rows.Count; $row++) {
    $worksheetRow = $row + 2
    $advanceName = [string](Get-ExcelArrayValue $advanceValues $row 1)
    $advanceUnit = Normalize-ConstructionUnit ([string](Get-ExcelArrayValue $advanceValues $row 2))
    $unitMatchesBudget = $budgetKeys.ContainsKey($advanceName + [char]31 + $advanceUnit)
    if (-not $unitMatchesBudget) {
      # Si una versión anterior vinculó por nombre solamente, restaura el
      # respaldo manual guardado al final de la fórmula en vez de conservar
      # una cantidad o un precio de otra unidad.
      foreach ($spec in @(@{ column = 3; formulas = $quantityFormulas }, @{ column = 7; formulas = $priceFormulas })) {
        $currentFormula = [string](Get-ExcelArrayValue $spec.formulas $row)
        if ($currentFormula -notmatch 'tblPresupuesto') { continue }
        $cell = $advanceSheet.Cells.Item($worksheetRow, $spec.column)
        $fallbackMatch = [regex]::Match($currentFormula, ',(-?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)\)$')
        if ($fallbackMatch.Success) {
          $fallbackValue = [double]::Parse($fallbackMatch.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture)
          $cell.Formula = $fallbackValue.ToString('G17', [Globalization.CultureInfo]::InvariantCulture)
        } else {
          $cell.Formula = '0'
        }
      }
      continue
    }
    foreach ($spec in @(@{ column = 3; template = $quantityTemplate; formulas = $quantityFormulas }, @{ column = 7; template = $priceTemplate; formulas = $priceFormulas })) {
      $currentFormula = [string](Get-ExcelArrayValue $spec.formulas $row)
      $cell = $advanceSheet.Cells.Item($worksheetRow, $spec.column)
      if ($currentFormula -match 'tblPresupuesto' -and $currentFormula -match ',\s*""\)$') { continue }
      $cell.Formula = New-LinkedFormula ([string]$spec.template) $worksheetRow '""'
    }
  }
}

function Apply-ConstructionBaseStyle($targetSheet, [string]$lastColumn, [string]$headerAddress, [hashtable]$colors) {
  $used = $targetSheet.UsedRange
  $used.Font.Name = 'Aptos'
  $used.Font.Size = 10
  $used.Font.Color = $colors.text
  $used.VerticalAlignment = -4108
  $used.Borders.LineStyle = 1
  $used.Borders.Color = $colors.border

  $titleAddress = 'A1:{0}1' -f $lastColumn
  $title = $targetSheet.Range($titleAddress)
  $currentTitleMerge = $targetSheet.Range('A1').MergeArea
  $currentTitleMergeAddress = [string]$currentTitleMerge.Address($false, $false)
  if ($currentTitleMergeAddress -ne $titleAddress) {
    $currentTitleMerge.UnMerge() | Out-Null
    $title.Merge() | Out-Null
  }
  $title.Interior.Color = $colors.petrol
  $title.Font.Color = $colors.white
  $title.Font.Bold = $true
  $title.Font.Size = 18
  $title.HorizontalAlignment = -4131
  $title.VerticalAlignment = -4108
  $title.RowHeight = 36
  $title.ShrinkToFit = $true

  $header = $targetSheet.Range($headerAddress)
  $header.Interior.Color = $colors.teal
  $header.Font.Color = $colors.white
  $header.Font.Bold = $true
  $header.HorizontalAlignment = -4108
  $header.VerticalAlignment = -4108
  $header.WrapText = $true
  $header.RowHeight = 40
}

function Set-ConstructionRoleFill($targetSheet, [string]$columnName, [int]$firstRow, [int]$lastRow, [int]$color) {
  if ($lastRow -lt $firstRow) { return }
  $address = '{0}{1}:{0}{2}' -f $columnName, $firstRow, $lastRow
  $range = $targetSheet.Range($address)
  $range.Interior.Color = $color
  $range.Borders.Color = 14934227
}

function Normalize-ConstructionUnit([string]$unit) {
  $normalized = $unit.Trim().ToLowerInvariant().Replace('²', '2').Replace('³', '3').Replace('^', '')
  if ($normalized -eq 'm 2') { return 'm2' }
  if ($normalized -eq 'm 3') { return 'm3' }
  return $normalized
}

function Set-ConstructionNumericValidation($targetSheet, [string]$columnName, [int]$lastRow, [bool]$allowZero, [string]$message) {
  if ($lastRow -lt 3) { return }
  $range = $targetSheet.Range(('{0}3:{0}{1}' -f $columnName, $lastRow))
  try { $range.Validation.Delete() | Out-Null } catch {}
  $operator = if ($allowZero) { 7 } else { 5 }
  $range.Validation.Add(2, 1, $operator, 0) | Out-Null
  $range.Validation.IgnoreBlank = $false
  $range.Validation.ShowError = $true
  $range.Validation.ErrorTitle = 'Dato inválido'
  $range.Validation.ErrorMessage = $message
  $range.Validation.ShowInput = $true
  $range.Validation.InputTitle = 'Dato requerido'
  $range.Validation.InputMessage = 'Ingresa un valor numérico válido para esta columna.'
}

function Set-ConstructionCubicWarnings($targetSheet, [int]$lastRow) {
  if ($lastRow -lt 3) { return }
  foreach ($column in @('C', 'D', 'E', 'H')) {
    $range = $targetSheet.Range(('{0}3:{0}{1}' -f $column, $lastRow))
    foreach ($formula in @(('=NOT(ISNUMBER({0}3))' -f $column), ('={0}3<=0' -f $column))) {
      $exists = $false
      foreach ($condition in $range.FormatConditions) {
        try { if ([string]$condition.Formula1 -eq $formula) { $exists = $true } } catch {}
      }
      if (-not $exists) {
        $condition = $range.FormatConditions.Add(2, 1, $formula)
        $condition.Interior.Color = 13551615
        $condition.Font.Color = 192
      }
    }
  }
  foreach ($column in @('F', 'G', 'I')) {
    $range = $targetSheet.Range(('{0}3:{0}{1}' -f $column, $lastRow))
    $formula = '=NOT(ISNUMBER({0}3))' -f $column
    $exists = $false
    foreach ($condition in $range.FormatConditions) {
      try { if ([string]$condition.Formula1 -eq $formula) { $exists = $true } } catch {}
    }
    if (-not $exists) {
      $condition = $range.FormatConditions.Add(2, 1, $formula)
      $condition.Interior.Color = 13551615
      $condition.Font.Color = 192
    }
  }
}

function Set-ConstructionAdvanceAlerts($targetSheet, [int]$lastRow) {
  if ($lastRow -lt 3) { return }
  $range = $targetSheet.Range(('A3:J{0}' -f $lastRow))
  $hasOverrunRule = $false
  $hasReviewRule = $false
  foreach ($condition in $range.FormatConditions) {
    try {
      $formula = [string]$condition.Formula1
      if ($formula -match 'SOBREEJECUCIÓN') { $hasOverrunRule = $true }
      if ($formula -match 'REVISAR DATOS') { $hasReviewRule = $true }
    } catch {}
  }
  if (-not $hasOverrunRule) {
    $condition = $range.FormatConditions.Add(2, 1, '=$I3="SOBREEJECUCIÓN"')
    $condition.Interior.Color = 13551615
    $condition.Font.Color = 192
  }
  if (-not $hasReviewRule) {
    $condition = $range.FormatConditions.Add(2, 1, '=$I3="REVISAR DATOS"')
    $condition.Interior.Color = 13431551
    $condition.Font.Color = 13107
  }
}

try {
  $file = [Environment]::GetEnvironmentVariable('CONSTRUCTION_EXCEL_PATH')
  $mode = [Environment]::GetEnvironmentVariable('CONSTRUCTION_EXCEL_MODE')
  $script:ConstructionProfile = [Environment]::GetEnvironmentVariable('CONSTRUCTION_EXCEL_PROFILE') -eq '1'
  $payload = [Environment]::GetEnvironmentVariable('CONSTRUCTION_EXCEL_PAYLOAD') | ConvertFrom-Json
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AutomationSecurity = 3
  $book = $excel.Workbooks.Open($file, 0, ($mode -eq 'sheets' -or $mode -eq 'audit'))
  Mark-ConstructionPhase 'openWorkbook'
  if ($mode -eq 'audit') {
    $excel.Calculation = -4105
    $excel.CalculateFullRebuild() | Out-Null
  }
  if ($mode -eq 'append') {
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.DisplayStatusBar = $false
    $excel.Calculation = -4135
    $excel.CalculateBeforeSave = $true
    $excel.AutoCorrect.AutoFillFormulasInLists = $true
    $book.ForceFullCalculation = $true
  }
  if ($mode -eq 'sheets') {
    $names = @()
    foreach ($sheet in $book.Worksheets) { $names += $sheet.Name }
    @{ path = $file; sheets = $names } | ConvertTo-Json -Compress
  } elseif ($mode -eq 'audit') {
    $sheetMap = @{}
    foreach ($candidate in $book.Worksheets) { $sheetMap[[string]$candidate.Name] = $candidate }
    foreach ($required in @('Cubicaciones', 'Presupuesto', 'Avance', 'Resumen')) {
      if (-not $sheetMap.ContainsKey($required)) { throw "Falta la hoja requerida '$required'." }
    }
    $cubicTable = $sheetMap['Cubicaciones'].ListObjects.Item('tblCubicaciones')
    $budgetTable = $sheetMap['Presupuesto'].ListObjects.Item('tblPresupuesto')
    $advanceTable = $sheetMap['Avance'].ListObjects.Item('tblAvance')
    $classCounts = @{
      'VINCULADO' = 0
      'SIN VÍNCULO' = 0
      'UNIDAD INCOMPATIBLE' = 0
      'REGISTRO MANUAL' = 0
      'REGISTRO DUPLICADO' = 0
      'SIN AVANCE' = 0
      'DATOS INVÁLIDOS' = 0
      'FORMULA SOBRESCRITA' = 0
      'SOBREEJECUCIÓN' = 0
    }
    $cubicIssues = @(); $budgetIssues = @(); $advanceItems = @(); $formulaIssues = @(); $formulaErrors = @(); $brokenRefs = @(); $duplicates = @()
    $budgetNames = @{}; $budgetKeys = @{}; $advanceKeys = @{}; $advanceUnitTotals = @{}
    $cubicConceptNames = @{}
    $seenBudgetRecords = @{}; $seenAdvanceRecords = @{}; $seenCubicRecords = @{}
    foreach ($cubicRow in $cubicTable.ListRows) { $cubicConceptNames[[string]$cubicRow.Range.Cells.Item(1, 1).Text] = $true }

    foreach ($budgetRow in $budgetTable.ListRows) {
      $budgetRange = $budgetRow.Range
      $name = [string]$budgetRange.Cells.Item(1, 1).Text
      $unit = Normalize-ConstructionUnit ([string]$budgetRange.Cells.Item(1, 2).Text)
      $budgetNames[$name] = $true
      $budgetKeys[$name + [char]31 + $unit] = $true
      $quantity = $budgetRange.Cells.Item(1, 3).Value2
      $price = $budgetRange.Cells.Item(1, 4).Value2
      $cost = $budgetRange.Cells.Item(1, 5).Value2
      $budgetDuplicateKey = @((Get-AuditKeyPart $name), (Get-AuditKeyPart $unit), (Get-AuditKeyPart $quantity), (Get-AuditKeyPart $price)) -join [char]31
      if ($seenBudgetRecords.ContainsKey($budgetDuplicateKey)) {
        $duplicates += @{ hoja = 'Presupuesto'; partida = $name; unidad = $unit; fila = $budgetRange.Row; duplicadoDeFila = $seenBudgetRecords[$budgetDuplicateKey] }
        $classCounts['REGISTRO DUPLICADO'] = [int]$classCounts['REGISTRO DUPLICADO'] + 1
      } else { $seenBudgetRecords[$budgetDuplicateKey] = $budgetRange.Row }
      $costFormula = $budgetRange.Cells.Item(1, 5).HasFormula
      $quantityFormula = [string]$budgetRange.Cells.Item(1, 3).Formula
      $quantityIsNumber = Test-ExcelNumber $quantity
      $priceIsNumber = Test-ExcelNumber $price
      $costIsNumber = Test-ExcelNumber $cost
      $budgetIsInvalid = (-not $quantityIsNumber) -or (-not $priceIsNumber) -or (-not $costIsNumber)
      if ($quantityIsNumber -and [double]$quantity -le 0) { $budgetIsInvalid = $true }
      if ($priceIsNumber -and [double]$price -lt 0) { $budgetIsInvalid = $true }
      if ($budgetIsInvalid) {
        $budgetIssue = @{ partida = $name; unidad = $unit; row = $budgetRange.Row }
        $budgetIssues += $budgetIssue
        $classCounts['DATOS INVÁLIDOS'] = [int]$classCounts['DATOS INVÁLIDOS'] + 1
      }
      $costFormulaInconsistent = -not $costFormula
      if ($quantityIsNumber -and $priceIsNumber -and $costIsNumber -and [Math]::Abs([double]$cost - ([double]$quantity * [double]$price)) -gt 0.01) { $costFormulaInconsistent = $true }
      if ($costFormulaInconsistent) {
        $formulaIssues += @{ hoja = 'Presupuesto'; partida = $name; row = $budgetRange.Row; columna = 'Costo total'; formulaPresente = [bool]$costFormula; valor = $cost }
        $classCounts['FORMULA SOBRESCRITA'] = [int]$classCounts['FORMULA SOBRESCRITA'] + 1
      }
      $supportedCubicUnit = $unit -in @('m2', 'm²', 'm^2', 'm3', 'm³', 'm^3', 'm3 ajustado', 'm³ ajustado')
      if ($cubicConceptNames.ContainsKey($name) -and $supportedCubicUnit -and $quantityFormula -notmatch 'tblCubicaciones') {
        $formulaIssues += @{ hoja = 'Presupuesto'; partida = $name; row = $budgetRange.Row; columna = 'Cantidad'; issue = 'VÍNCULO DE CUBICACIÓN AUSENTE' }
        $classCounts['FORMULA SOBRESCRITA'] = [int]$classCounts['FORMULA SOBRESCRITA'] + 1
      }
    }

    foreach ($cubicRow in $cubicTable.ListRows) {
      $record = $cubicRow.Range
      $values = @(
        $record.Cells.Item(1, 3).Value2, $record.Cells.Item(1, 4).Value2, $record.Cells.Item(1, 5).Value2,
        $record.Cells.Item(1, 6).Value2, $record.Cells.Item(1, 7).Value2, $record.Cells.Item(1, 8).Value2, $record.Cells.Item(1, 9).Value2
      )
      $cubicDuplicateKey = @((Get-AuditKeyPart $record.Cells.Item(1, 1).Value2), (Get-AuditKeyPart $record.Cells.Item(1, 2).Value2), (Get-AuditKeyPart $values[0]), (Get-AuditKeyPart $values[1]), (Get-AuditKeyPart $values[2]), (Get-AuditKeyPart $values[5])) -join [char]31
      if ($seenCubicRecords.ContainsKey($cubicDuplicateKey)) {
        $duplicates += @{ hoja = 'Cubicaciones'; concepto = [string]$record.Cells.Item(1, 1).Text; fila = $record.Row; duplicadoDeFila = $seenCubicRecords[$cubicDuplicateKey] }
        $classCounts['REGISTRO DUPLICADO'] = [int]$classCounts['REGISTRO DUPLICADO'] + 1
      } else { $seenCubicRecords[$cubicDuplicateKey] = $record.Row }
      $cubicFormulasPresent = $record.Cells.Item(1, 6).HasFormula -and $record.Cells.Item(1, 7).HasFormula -and $record.Cells.Item(1, 9).HasFormula
      $cubeInputNumbers = @((Test-ExcelNumber $values[0]), (Test-ExcelNumber $values[1]), (Test-ExcelNumber $values[2]), (Test-ExcelNumber $values[5]))
      $cubeOutputNumbers = @((Test-ExcelNumber $values[3]), (Test-ExcelNumber $values[4]), (Test-ExcelNumber $values[6]))
      $cubeIsInvalid = ($cubeInputNumbers -contains $false) -or ($cubeOutputNumbers -contains $false)
      if (-not $cubeIsInvalid -and ([double]$values[0] -le 0 -or [double]$values[1] -le 0 -or [double]$values[2] -le 0 -or [double]$values[5] -le 0 -or [double]$values[3] -le 0 -or [double]$values[4] -le 0 -or [double]$values[6] -le 0)) { $cubeIsInvalid = $true }
      if ($cubeIsInvalid) {
        $cubicIssue = @{ concepto = [string]$record.Cells.Item(1, 1).Text; row = $record.Row }
        $cubicIssues += $cubicIssue
        $classCounts['DATOS INVÁLIDOS']++
      }
      $cubicResultMismatch = -not $cubeIsInvalid -and (
        [Math]::Abs([double]$values[3] - ([double]$values[0] * [double]$values[1])) -gt 0.001 -or
        [Math]::Abs([double]$values[4] - ([double]$values[3] * [double]$values[2])) -gt 0.001 -or
        [Math]::Abs([double]$values[6] - ([double]$values[4] * [double]$values[5])) -gt 0.001
      )
      if (-not $cubicFormulasPresent) {
        $formulaIssues += @{ hoja = 'Cubicaciones'; concepto = [string]$record.Cells.Item(1, 1).Text; row = $record.Row; formulaPresente = $false }
        $classCounts['FORMULA SOBRESCRITA'] = [int]$classCounts['FORMULA SOBRESCRITA'] + 1
      } elseif ($cubicResultMismatch) {
        $formulaIssues += @{ hoja = 'Cubicaciones'; concepto = [string]$record.Cells.Item(1, 1).Text; row = $record.Row; issue = 'RESULTADO NO COINCIDE CON LAS DIMENSIONES' }
        $classCounts['FORMULA SOBRESCRITA'] = [int]$classCounts['FORMULA SOBRESCRITA'] + 1
      }
    }

    foreach ($advanceRow in $advanceTable.ListRows) {
      $advanceRange = $advanceRow.Range
      $auditAdvanceName = [string]$advanceRange.Cells.Item(1, 1).Text
      $auditDisplayUnit = [string]$advanceRange.Cells.Item(1, 2).Text
      $auditAdvanceUnit = Normalize-ConstructionUnit $auditDisplayUnit
      $auditAdvanceKey = $auditAdvanceName + [char]31 + $auditAdvanceUnit
      $advanceKeys[$auditAdvanceKey] = $true
      $auditBudgeted = $advanceRange.Cells.Item(1, 3).Value2
      $auditExecuted = $advanceRange.Cells.Item(1, 4).Value2
      $auditPercent = $advanceRange.Cells.Item(1, 5).Value2
      $auditPending = $advanceRange.Cells.Item(1, 6).Value2
      $auditPrice = $advanceRange.Cells.Item(1, 7).Value2
      $auditExecutedCost = $advanceRange.Cells.Item(1, 8).Value2
      $auditStatus = [string]$advanceRange.Cells.Item(1, 9).Text
      $auditExcess = $advanceRange.Cells.Item(1, 10).Value2
      $advanceDuplicateKey = @((Get-AuditKeyPart $auditAdvanceName), (Get-AuditKeyPart $auditAdvanceUnit), (Get-AuditKeyPart $auditBudgeted), (Get-AuditKeyPart $auditExecuted), (Get-AuditKeyPart $auditPrice)) -join [char]31
      $auditDuplicate = $false
      if ($seenAdvanceRecords.ContainsKey($advanceDuplicateKey)) {
        $auditDuplicate = $true
        $duplicates += @{ hoja = 'Avance'; partida = $auditAdvanceName; unidad = $auditAdvanceUnit; fila = $advanceRange.Row; duplicadoDeFila = $seenAdvanceRecords[$advanceDuplicateKey] }
        $classCounts['REGISTRO DUPLICADO'] = [int]$classCounts['REGISTRO DUPLICADO'] + 1
      } else { $seenAdvanceRecords[$advanceDuplicateKey] = $advanceRange.Row }
      if (-not $advanceUnitTotals.ContainsKey($auditDisplayUnit)) { $advanceUnitTotals[$auditDisplayUnit] = @{ budgeted = 0.0; executed = 0.0; pending = 0.0 } }
      if (Test-ExcelNumber $auditBudgeted) { $advanceUnitTotals[$auditDisplayUnit].budgeted += [double]$auditBudgeted }
      if (Test-ExcelNumber $auditExecuted) { $advanceUnitTotals[$auditDisplayUnit].executed += [double]$auditExecuted }
      if (Test-ExcelNumber $auditPending) { $advanceUnitTotals[$auditDisplayUnit].pending += [double]$auditPending }
      $auditQuantityFormula = [string]$advanceRange.Cells.Item(1, 3).Formula
      $auditPriceFormula = [string]$advanceRange.Cells.Item(1, 7).Formula
      $advanceFormulaFields = @(
        @{ column = 'Avance físico (%)'; cell = $advanceRange.Cells.Item(1, 5) },
        @{ column = 'Cantidad pendiente'; cell = $advanceRange.Cells.Item(1, 6) },
        @{ column = 'Costo ejecutado'; cell = $advanceRange.Cells.Item(1, 8) },
        @{ column = 'Estado de avance'; cell = $advanceRange.Cells.Item(1, 9) },
        @{ column = 'Exceso ejecutado'; cell = $advanceRange.Cells.Item(1, 10) }
      )
      $missingAdvanceFormulas = @()
      foreach ($formulaField in $advanceFormulaFields) { if (-not $formulaField.cell.HasFormula) { $missingAdvanceFormulas += [string]$formulaField.column } }
      if ($missingAdvanceFormulas.Count -gt 0) {
        $formulaIssues += @{ hoja = 'Avance'; partida = $auditAdvanceName; row = $advanceRange.Row; columnasSinFormula = $missingAdvanceFormulas }
      }
      $auditQuantityLinked = $auditQuantityFormula -match 'tblPresupuesto'
      $auditPriceLinked = $auditPriceFormula -match 'tblPresupuesto'
      $auditBudgetedIsNumber = Test-ExcelNumber $auditBudgeted
      $auditExecutedIsNumber = Test-ExcelNumber $auditExecuted
      $auditPriceIsNumber = Test-ExcelNumber $auditPrice
      $auditCostIsNumber = Test-ExcelNumber $auditExecutedCost
      $auditPendingIsNumber = Test-ExcelNumber $auditPending
      $auditInvalid = (-not $auditBudgetedIsNumber) -or (-not $auditExecutedIsNumber) -or (-not $auditPriceIsNumber) -or (-not $auditCostIsNumber) -or (-not $auditPendingIsNumber)
      if ($auditBudgetedIsNumber -and [double]$auditBudgeted -le 0) { $auditInvalid = $true }
      if ($auditExecutedIsNumber -and [double]$auditExecuted -lt 0) { $auditInvalid = $true }
      if ($auditPriceIsNumber -and [double]$auditPrice -lt 0) { $auditInvalid = $true }
      if ($auditPendingIsNumber -and [double]$auditPending -lt 0) { $auditInvalid = $true }
      $advanceFormulaMismatch = $false
      if (-not $auditInvalid) {
        $expectedAdvance = [double]$auditExecuted / [double]$auditBudgeted
        $expectedPending = [Math]::Max([double]$auditBudgeted - [double]$auditExecuted, 0)
        $expectedExecutedCost = [double]$auditExecuted * [double]$auditPrice
        $expectedExcess = [Math]::Max([double]$auditExecuted - [double]$auditBudgeted, 0)
        $expectedStatus = if ($expectedExcess -gt 0) { 'SOBREEJECUCIÓN' } else { 'OK' }
        $advanceFormulaMismatch = ([Math]::Abs([double]$auditPercent - $expectedAdvance) -gt 0.0000001) -or
          ([Math]::Abs([double]$auditPending - $expectedPending) -gt 0.001) -or
          ([Math]::Abs([double]$auditExecutedCost - $expectedExecutedCost) -gt 0.01) -or
          ([Math]::Abs([double]$auditExcess - $expectedExcess) -gt 0.001) -or
          ($auditStatus -ne $expectedStatus)
      }
      if ($missingAdvanceFormulas.Count -gt 0 -or $advanceFormulaMismatch) {
        $classCounts['FORMULA SOBRESCRITA'] = [int]$classCounts['FORMULA SOBRESCRITA'] + 1
        if ($advanceFormulaMismatch -and $missingAdvanceFormulas.Count -eq 0) {
          $formulaIssues += @{ hoja = 'Avance'; partida = $auditAdvanceName; row = $advanceRange.Row; issue = 'RESULTADOS DE AVANCE NO COINCIDEN CON SUS ENTRADAS' }
        }
      }
      $advanceFormulaMismatch = $false
      if (-not $auditInvalid -and $missingAdvanceFormulas.Count -eq 0) {
        $expectedPercent = [double]$auditExecuted / [double]$auditBudgeted
        $expectedPending = [Math]::Max([double]$auditBudgeted - [double]$auditExecuted, 0)
        $expectedCost = [double]$auditExecuted * [double]$auditPrice
        $expectedExcess = [Math]::Max([double]$auditExecuted - [double]$auditBudgeted, 0)
        $expectedStatus = if ($expectedExcess -gt 0) { 'SOBREEJECUCIÓN' } else { 'OK' }
        $advanceFormulaMismatch = ([Math]::Abs([double]$auditPercent - $expectedPercent) -gt 0.0000001) -or
          ([Math]::Abs([double]$auditPending - $expectedPending) -gt 0.001) -or
          ([Math]::Abs([double]$auditExecutedCost - $expectedCost) -gt 0.01) -or
          ([Math]::Abs([double]$auditExcess - $expectedExcess) -gt 0.001) -or
          ($auditStatus -ne $expectedStatus)
        if ($advanceFormulaMismatch) {
          $formulaIssues += @{ hoja = 'Avance'; partida = $auditAdvanceName; row = $advanceRange.Row; issue = 'RESULTADO DE FÓRMULA NO COINCIDE CON LOS DATOS' }
          $classCounts['FORMULA SOBRESCRITA'] = [int]$classCounts['FORMULA SOBRESCRITA'] + 1
        }
      }
      if ($auditQuantityLinked -and $auditPriceLinked -and -not $budgetKeys.ContainsKey($auditAdvanceKey)) { $auditLinkClass = 'SIN VÍNCULO' }
      elseif (-not $budgetNames.ContainsKey($auditAdvanceName)) { $auditLinkClass = if ($auditQuantityLinked -or $auditPriceLinked) { 'SIN VÍNCULO' } else { 'REGISTRO MANUAL' } }
      elseif (-not $budgetKeys.ContainsKey($auditAdvanceKey)) { $auditLinkClass = 'UNIDAD INCOMPATIBLE' }
      elseif (-not $auditQuantityLinked -or -not $auditPriceLinked) { $auditLinkClass = 'SIN VÍNCULO' }
      else { $auditLinkClass = 'VINCULADO' }
      $classCounts[$auditLinkClass] = [int]$classCounts[$auditLinkClass] + 1
      $auditClassifications = @($auditLinkClass)
      if ($auditDuplicate) { $auditClassifications += 'REGISTRO DUPLICADO' }
      if ($missingAdvanceFormulas.Count -gt 0 -or $advanceFormulaMismatch) { $auditClassifications += 'FORMULA SOBRESCRITA' }
      if ($auditInvalid) { $classCounts['DATOS INVÁLIDOS'] = [int]$classCounts['DATOS INVÁLIDOS'] + 1; $auditClassifications += 'DATOS INVÁLIDOS' }
      if ((Test-ExcelNumber $auditExcess) -and [double]$auditExcess -gt 0) { $classCounts['SOBREEJECUCIÓN'] = [int]$classCounts['SOBREEJECUCIÓN'] + 1; $auditClassifications += 'SOBREEJECUCIÓN' }
      $advanceItem = [PSCustomObject]@{
        partida = $auditAdvanceName
        unidad = $auditAdvanceUnit
        clasificaciones = $auditClassifications
        avance = $auditPercent
        pendiente = $auditPending
        precioUnitario = $auditPrice
        costoEjecutado = $auditExecutedCost
        exceso = $auditExcess
        estadoExcel = $auditStatus
        row = $advanceRange.Row
      }
      $advanceItems += $advanceItem
    }

    $budgetWithoutAdvance = @()
    foreach ($budgetRow in $budgetTable.ListRows) {
      $name = [string]$budgetRow.Range.Cells.Item(1, 1).Text
      $unit = Normalize-ConstructionUnit ([string]$budgetRow.Range.Cells.Item(1, 2).Text)
      if (-not $advanceKeys.ContainsKey($name + [char]31 + $unit)) {
        $classCounts['SIN AVANCE']++
        $budgetWithoutAdvance += @{ partida = $name; unidad = $unit; row = $budgetRow.Range.Row }
      }
    }

    foreach ($auditSheet in $book.Worksheets) {
      try { $badCells = $auditSheet.UsedRange.SpecialCells(16); foreach ($cell in $badCells) { $formulaErrors += ($auditSheet.Name + '!' + $cell.Address($false, $false) + '=' + $cell.Text) } } catch {}
      try { $refCell = $auditSheet.UsedRange.Find('#REF!', $auditSheet.UsedRange.Cells.Item(1, 1), -4123, 2); if ($null -ne $refCell) { $brokenRefs += ($auditSheet.Name + '!' + $refCell.Address($false, $false)) } } catch {}
    }
    $summary = $sheetMap['Resumen']; $summaryMetrics = @{}
    for ($row = 3; $row -le $summary.UsedRange.Rows.Count; $row++) {
      $label = [string]$summary.Cells.Item($row, 1).Text
      if ($label) { $summaryMetrics[$label] = $summary.Cells.Item($row, 2).Value2 }
    }
    $consistencyIssues = @()
    $budgetCostSum = 0.0; $surfaceSum = 0.0; $volumeSum = 0.0; $executedCostSum = 0.0; $overrunRows = 0
    foreach ($budgetRow in $budgetTable.ListRows) { $value = $budgetRow.Range.Cells.Item(1, 5).Value2; if (Test-ExcelNumber $value) { $budgetCostSum += [double]$value } }
    foreach ($cubicRow in $cubicTable.ListRows) {
      $surface = $cubicRow.Range.Cells.Item(1, 6).Value2; $volume = $cubicRow.Range.Cells.Item(1, 7).Value2
      if (Test-ExcelNumber $surface) { $surfaceSum += [double]$surface }
      if (Test-ExcelNumber $volume) { $volumeSum += [double]$volume }
    }
    foreach ($advanceRow in $advanceTable.ListRows) {
      $cost = $advanceRow.Range.Cells.Item(1, 8).Value2; $excess = $advanceRow.Range.Cells.Item(1, 10).Value2
      if (Test-ExcelNumber $cost) { $executedCostSum += [double]$cost }
      if ((Test-ExcelNumber $excess) -and [double]$excess -gt 0) { $overrunRows++ }
    }
    $subtotalRow = 0; $ivaRow = 0; $totalRow = 0
    for ($row = 3; $row -le $sheetMap['Presupuesto'].UsedRange.Rows.Count; $row++) {
      $label = [string]$sheetMap['Presupuesto'].Cells.Item($row, 4).Text
      if ($label -eq 'Subtotal') { $subtotalRow = $row }
      elseif ($label -match '^IVA') { $ivaRow = $row }
      elseif ($label -eq 'Total') { $totalRow = $row }
    }
    if ($subtotalRow -gt 0 -and [Math]::Abs([double]$sheetMap['Presupuesto'].Cells.Item($subtotalRow, 5).Value2 - $budgetCostSum) -gt 0.01) {
      $consistencyIssues += @{ metric = 'Subtotal presupuesto'; expected = $budgetCostSum; actual = $sheetMap['Presupuesto'].Cells.Item($subtotalRow, 5).Value2 }
    }
    if ($ivaRow -gt 0 -and $subtotalRow -gt 0 -and $totalRow -gt 0) {
      $ivaLabelMatch = [regex]::Match([string]$sheetMap['Presupuesto'].Cells.Item($ivaRow, 4).Text, '\(([\d.,]+)%\)')
      if ($ivaLabelMatch.Success) {
        $rateText = $ivaLabelMatch.Groups[1].Value.Replace(',', '.')
        $ivaRate = [double]::Parse($rateText, [Globalization.CultureInfo]::InvariantCulture) / 100
        $expectedIva = $budgetCostSum * $ivaRate
        $actualIva = [double]$sheetMap['Presupuesto'].Cells.Item($ivaRow, 5).Value2
        $expectedTotal = $budgetCostSum + $expectedIva
        $actualTotal = [double]$sheetMap['Presupuesto'].Cells.Item($totalRow, 5).Value2
        if ([Math]::Abs($actualIva - $expectedIva) -gt 0.01) { $consistencyIssues += @{ metric = 'IVA presupuesto'; expected = $expectedIva; actual = $actualIva } }
        if ([Math]::Abs($actualTotal - $expectedTotal) -gt 0.01) { $consistencyIssues += @{ metric = 'Total presupuesto'; expected = $expectedTotal; actual = $actualTotal } }
      } else { $consistencyIssues += @{ metric = 'Etiqueta IVA'; actual = [string]$sheetMap['Presupuesto'].Cells.Item($ivaRow, 4).Text } }
    }
    foreach ($metricSpec in @(@{ label = 'Superficie acumulada (m²)'; expected = $surfaceSum }, @{ label = 'Volumen acumulado (m³)'; expected = $volumeSum }, @{ label = 'Costo ejecutado (CLP)'; expected = $executedCostSum }, @{ label = 'Partidas con sobre-ejecución'; expected = $overrunRows })) {
      if ($summaryMetrics.ContainsKey($metricSpec.label) -and (Test-ExcelNumber $summaryMetrics[$metricSpec.label]) -and [Math]::Abs([double]$summaryMetrics[$metricSpec.label] - [double]$metricSpec.expected) -gt 0.01) {
        $consistencyIssues += @{ metric = $metricSpec.label; expected = $metricSpec.expected; actual = $summaryMetrics[$metricSpec.label] }
      }
    }
    foreach ($unit in $advanceUnitTotals.Keys) {
      $progressLabel = 'Avance físico ponderado (' + $unit + ')'
      $pendingLabel = 'Saldo pendiente (' + $unit + ')'
      $unitTotals = $advanceUnitTotals[$unit]
      if ($unitTotals.budgeted -gt 0 -and $summaryMetrics.ContainsKey($progressLabel) -and [Math]::Abs([double]$summaryMetrics[$progressLabel] - ($unitTotals.executed / $unitTotals.budgeted)) -gt 0.000001) {
        $consistencyIssues += @{ metric = $progressLabel; expected = ($unitTotals.executed / $unitTotals.budgeted); actual = $summaryMetrics[$progressLabel] }
      }
      if ($summaryMetrics.ContainsKey($pendingLabel) -and [Math]::Abs([double]$summaryMetrics[$pendingLabel] - $unitTotals.pending) -gt 0.001) {
        $consistencyIssues += @{ metric = $pendingLabel; expected = $unitTotals.pending; actual = $summaryMetrics[$pendingLabel] }
      }
    }
    $chartAudits = @()
    foreach ($chartObject in $sheetMap['Resumen'].ChartObjects()) {
      if ($chartObject.Name -notin @('ConstructionAssistantBudgetByItem', 'ConstructionAssistantBudgetByItemLog')) { continue }
      $series = $chartObject.Chart.SeriesCollection().Item(1)
      $seriesValues = @($series.Values); $seriesCategories = @($series.XValues)
      $valuesMatch = $seriesValues.Count -eq $budgetTable.ListRows.Count
      $categoriesMatch = $seriesCategories.Count -eq $budgetTable.ListRows.Count
      if ($valuesMatch -or $categoriesMatch) {
        for ($index = 0; $index -lt $budgetTable.ListRows.Count; $index++) {
          if ($valuesMatch) {
            $expectedValue = $budgetTable.ListRows.Item($index + 1).Range.Cells.Item(1, 5).Value2
            if (-not (Test-ExcelNumber $expectedValue) -or [Math]::Abs([double]$seriesValues[$index] - [double]$expectedValue) -gt 0.01) { $valuesMatch = $false }
          }
          if ($categoriesMatch) {
            $expectedCategory = [string]$budgetTable.ListRows.Item($index + 1).Range.Cells.Item(1, 1).Text
            if ([string]$seriesCategories[$index] -cne $expectedCategory) { $categoriesMatch = $false }
          }
        }
      }
      $isLogarithmic = $chartObject.Chart.Axes(2).ScaleType -eq -4133
      $expectsLogarithmic = $chartObject.Name -eq 'ConstructionAssistantBudgetByItemLog'
      if (-not $valuesMatch -or -not $categoriesMatch -or $isLogarithmic -ne $expectsLogarithmic) {
        $consistencyIssues += @{ chart = $chartObject.Name; valuesMatchBudget = $valuesMatch; categoriesMatchBudget = $categoriesMatch; logarithmic = $isLogarithmic; expectedLogarithmic = $expectsLogarithmic }
      }
      $chartAudits += @{ name = $chartObject.Name; points = $series.Points().Count; valuesMatchBudget = $valuesMatch; categoriesMatchBudget = $categoriesMatch; logarithmic = $isLogarithmic; hasDataLabels = $series.HasDataLabels }
    }
    @{ path = $file; sheets = @($book.Worksheets | ForEach-Object { $_.Name }); tables = @{ cubicaciones = $cubicTable.ListRows.Count; presupuesto = $budgetTable.ListRows.Count; avance = $advanceTable.ListRows.Count }; classifications = $classCounts; advances = $advanceItems; budgetWithoutAdvance = $budgetWithoutAdvance; duplicates = $duplicates; invalidCubicaciones = $cubicIssues; invalidBudget = $budgetIssues; formulaIssues = $formulaIssues; consistencyIssues = $consistencyIssues; charts = $chartAudits; summary = $summaryMetrics; formulaErrors = $formulaErrors; brokenReferences = $brokenRefs } | ConvertTo-Json -Depth 8 -Compress
  } elseif ($mode -eq 'append') {
    $baseName = [string]$payload.sheetBase
    $expectedTitle = [string]$payload.rows[0][0]
    $headers = @($payload.rows[1])
    $items = @($payload.items)
    $sheet = $null
    foreach ($candidate in $book.Worksheets) {
      if ($candidate.Name -eq $baseName) { $sheet = $candidate; break }
    }

    $wasExisting = $sheet -ne $null
    if ($wasExisting) {
      if ([string]$sheet.Cells.Item(1, 1).Text -ne $expectedTitle) {
        throw "La hoja '$baseName' ya existe y no parece pertenecer a Construction Assistant; no se modificó para proteger sus datos."
      }
      if ($baseName -eq 'Avance' -and $headers.Count -eq 10) {
        $legacyAdvanceHeaders = $true
        for ($column = 1; $column -le 8; $column++) {
          if ([string]$sheet.Cells.Item(2, $column).Text -ne [string]$headers[$column - 1]) { $legacyAdvanceHeaders = $false; break }
        }
        if ($legacyAdvanceHeaders -and [string]$sheet.Cells.Item(2, 9).Text -eq '' -and [string]$sheet.Cells.Item(2, 10).Text -eq '') {
          $legacyLastRow = $sheet.UsedRange.Rows.Count
          $extensionIsEmpty = $true
          for ($row = 1; $row -le $legacyLastRow; $row++) {
            if ($null -ne $sheet.Cells.Item($row, 9).Value2 -or $null -ne $sheet.Cells.Item($row, 10).Value2 -or
              $sheet.Cells.Item($row, 9).HasFormula -or $sheet.Cells.Item($row, 10).HasFormula) {
              $extensionIsEmpty = $false; break
            }
          }
          if ($extensionIsEmpty) {
            Set-ConstructionCell $sheet 2 9 $headers[8]
            Set-ConstructionCell $sheet 2 10 $headers[9]
          }
        }
      }
      for ($column = 1; $column -le $headers.Count; $column++) {
        if ([string]$sheet.Cells.Item(2, $column).Text -ne [string]$headers[$column - 1]) {
          throw "La hoja '$baseName' tiene encabezados distintos; no se modificó para proteger sus datos."
        }
      }
    } else {
      $sheet = $book.Worksheets.Add()
      $sheet.Name = $baseName
      for ($column = 0; $column -lt $headers.Count; $column++) {
        Set-ConstructionCell $sheet 2 ($column + 1) $headers[$column]
      }
      Set-ConstructionCell $sheet 1 1 $expectedTitle
    }

    $firstNewRow = 3
    if ($baseName -eq 'Presupuesto' -and $wasExisting) {
      $lastUsedRow = $sheet.UsedRange.Rows.Count
      $subtotalRow = 0
      for ($row = 3; $row -le $lastUsedRow; $row++) {
        if ([string]$sheet.Cells.Item($row, 4).Text -match '^Subtotal$') { $subtotalRow = $row; break }
      }
      if ($subtotalRow -lt 4 -or
        [string]$sheet.Cells.Item($subtotalRow + 1, 4).Text -notmatch '^IVA' -or
        [string]$sheet.Cells.Item($subtotalRow + 2, 4).Text -notmatch '^Total$') {
        throw "La hoja '$baseName' no contiene las filas de subtotal, IVA y total esperadas; no se modificó."
      }
      $firstNewRow = $subtotalRow
      if ($items.Count -gt 0) {
        $insertEndRow = $subtotalRow + $items.Count - 1
        $insertRange = $sheet.Range(('A{0}:E{1}' -f $subtotalRow, $insertEndRow))
        $insertRange.Insert(-4121) | Out-Null
      }
    } elseif ($wasExisting) {
      $firstNewRow = [Math]::Max(3, $sheet.UsedRange.Rows.Count + 1)
    }

    Set-ConstructionRows $sheet $firstNewRow $items $headers.Count
    Mark-ConstructionPhase 'writeRows'

    if ($baseName -eq 'Presupuesto') {
      $subtotalRow = $firstNewRow + $items.Count
      $ivaRow = $subtotalRow + 1
      $totalRow = $subtotalRow + 2
      if (-not $wasExisting) {
        Set-ConstructionCell $sheet $subtotalRow 4 'Subtotal'
        Set-ConstructionCell $sheet $totalRow 4 'Total'
      }
      $ivaPercentage = ([double]$payload.tasaIva * 100).ToString('0.##', [Globalization.CultureInfo]::InvariantCulture)
      Set-ConstructionCell $sheet $ivaRow 4 "IVA ($ivaPercentage%)"
      $firstDataRow = 3
      $lastDataRow = $subtotalRow - 1
      $budgetTable = Ensure-ConstructionTable $book $sheet 'tblPresupuesto' 'E' $lastDataRow
       Set-ConstructionTableFormula $budgetTable 'Costo total' '=IF(AND(ISNUMBER([@Cantidad]),[@Cantidad]>0,ISNUMBER([@[Precio unitario]]),[@[Precio unitario]]>=0),IFERROR([@Cantidad]*[@[Precio unitario]],""),"")'
       Set-ConstructionNumericValidation $sheet 'C' $lastDataRow $false 'La cantidad debe ser numérica y mayor que cero.'
       Set-ConstructionNumericValidation $sheet 'D' $lastDataRow $true 'El precio unitario debe ser numérico y no negativo.'
      $totals = $payload.plan.formulasTotales
      $sheet.Cells.Item($subtotalRow, 5).Formula = '=SUM(tblPresupuesto[Costo total])'
      $sheet.Cells.Item($ivaRow, 5).Formula = Resolve-Formula ([string]$totals.iva) @{ subtotalRow = $subtotalRow }
      $sheet.Cells.Item($totalRow, 5).Formula = Resolve-Formula ([string]$totals.total) @{ subtotalRow = $subtotalRow; ivaRow = $ivaRow }
      $lastDataRow = $totalRow
    } else {
      $lastDataRow = $firstNewRow + $items.Count - 1
      if ($baseName -eq 'Cubicaciones') {
        $cubicTable = Ensure-ConstructionTable $book $sheet 'tblCubicaciones' 'I' $lastDataRow
      } else {
        $advanceTable = Ensure-ConstructionTable $book $sheet 'tblAvance' 'J' $lastDataRow
      }
       if ($baseName -eq 'Cubicaciones') {
         Set-ConstructionTableFormula $cubicTable 'Superficie (m²)' '=IF(OR(NOT(ISNUMBER([@[Largo (m)]])),[@[Largo (m)]]<=0,NOT(ISNUMBER([@[Ancho (m)]])),[@[Ancho (m)]]<=0),"",IFERROR([@[Largo (m)]]*[@[Ancho (m)]],""))'
         Set-ConstructionTableFormula $cubicTable 'Volumen (m³)' '=IF(OR(NOT(ISNUMBER([@[Superficie (m²)]])),NOT(ISNUMBER([@[Profundidad (m)]])),[@[Profundidad (m)]]<=0),"",IFERROR([@[Superficie (m²)]]*[@[Profundidad (m)]],""))'
         Set-ConstructionTableFormula $cubicTable 'Volumen ajustado (m³)' '=IF(OR(NOT(ISNUMBER([@[Volumen (m³)]])),NOT(ISNUMBER([@[Factor ajuste]])),[@[Factor ajuste]]<=0),"",IFERROR([@[Volumen (m³)]]*[@[Factor ajuste]],""))'
         foreach ($column in @('C', 'D', 'E', 'H')) {
           Set-ConstructionNumericValidation $sheet $column $lastDataRow $false 'Ingresa un valor numérico mayor que cero.'
         }
       } else {
       Set-ConstructionTableFormula $advanceTable 'Avance físico (%)' '=IFERROR(IF(AND(ISNUMBER([@[Cantidad ejecutada]]),[@[Cantidad ejecutada]]>=0,ISNUMBER([@[Cantidad presupuestada]]),[@[Cantidad presupuestada]]>0),[@[Cantidad ejecutada]]/[@[Cantidad presupuestada]],""),"")'
       Set-ConstructionTableFormula $advanceTable 'Cantidad pendiente' '=IF(AND(ISNUMBER([@[Cantidad presupuestada]]),[@[Cantidad presupuestada]]>0,ISNUMBER([@[Cantidad ejecutada]]),[@[Cantidad ejecutada]]>=0),MAX([@[Cantidad presupuestada]]-[@[Cantidad ejecutada]],0),"")'
       Set-ConstructionTableFormula $advanceTable 'Costo ejecutado' '=IF(AND(ISNUMBER([@[Cantidad ejecutada]]),[@[Cantidad ejecutada]]>=0,ISNUMBER([@[Precio unitario]]),[@[Precio unitario]]>=0),IFERROR([@[Cantidad ejecutada]]*[@[Precio unitario]],""),"")'
         Set-ConstructionTableFormula $advanceTable 'Estado de avance' '=IF(OR(NOT(ISNUMBER([@[Cantidad presupuestada]])),[@[Cantidad presupuestada]]<=0,NOT(ISNUMBER([@[Cantidad ejecutada]])),[@[Cantidad ejecutada]]<0,NOT(ISNUMBER([@[Precio unitario]])),[@[Precio unitario]]<0,NOT(ISNUMBER([@[Costo ejecutado]]))),"REVISAR DATOS",IF([@[Cantidad ejecutada]]>[@[Cantidad presupuestada]],"SOBREEJECUCIÓN","OK"))'
         Set-ConstructionTableFormula $advanceTable 'Exceso ejecutado' '=IF(AND(ISNUMBER([@[Cantidad presupuestada]]),[@[Cantidad presupuestada]]>0,ISNUMBER([@[Cantidad ejecutada]]),[@[Cantidad ejecutada]]>=0),MAX([@[Cantidad ejecutada]]-[@[Cantidad presupuestada]],0),"")'
         Set-ConstructionNumericValidation $sheet 'C' $lastDataRow $false 'La cantidad presupuestada debe ser numérica y mayor que cero.'
         Set-ConstructionNumericValidation $sheet 'D' $lastDataRow $true 'La cantidad ejecutada debe ser numérica y no negativa.'
         Set-ConstructionNumericValidation $sheet 'G' $lastDataRow $true 'El precio unitario debe ser numérico y no negativo.'
      }
    }
    Mark-ConstructionPhase 'primaryTableFormulasAndValidation'

    # Garantiza tablas y aplica vínculos cruzados incluso si las otras hojas
    # ya existían antes de esta llamada del MCP.
    $cubicSheet = $null; $budgetSheet = $null; $advanceSheet = $null
    foreach ($candidate in $book.Worksheets) {
      if ($candidate.Name -eq 'Cubicaciones') { $cubicSheet = $candidate }
      elseif ($candidate.Name -eq 'Presupuesto') { $budgetSheet = $candidate }
      elseif ($candidate.Name -eq 'Avance') { $advanceSheet = $candidate }
    }
    $cubicTable = $null; $budgetTable = $null; $advanceTable = $null
    if ($cubicSheet -ne $null) {
      $cubicDataLastRow = $cubicSheet.UsedRange.Rows.Count
      $cubicTable = Ensure-ConstructionTable $book $cubicSheet 'tblCubicaciones' 'I' $cubicDataLastRow
    }
    if ($budgetSheet -ne $null) {
      $budgetUsedLastRow = $budgetSheet.UsedRange.Rows.Count
      $budgetSubtotal = 0
      for ($row = 3; $row -le $budgetUsedLastRow; $row++) {
        if ([string]$budgetSheet.Cells.Item($row, 4).Text -match '^Subtotal$') { $budgetSubtotal = $row; break }
      }
      if ($budgetSubtotal -lt 4) { throw "No se encontró el subtotal de Presupuesto para vincular las partidas." }
      $budgetTable = Ensure-ConstructionTable $book $budgetSheet 'tblPresupuesto' 'E' ($budgetSubtotal - 1)
    }
    if ($advanceSheet -ne $null) {
      $advanceDataLastRow = $advanceSheet.UsedRange.Rows.Count
      $advanceTable = Ensure-ConstructionTable $book $advanceSheet 'tblAvance' 'J' $advanceDataLastRow
    }
    Mark-ConstructionPhase 'ensureCrossTables'
    if ($cubicTable -ne $null -and $budgetTable -ne $null) {
      Sync-BudgetCubicLinks $budgetSheet $budgetTable $cubicTable ([string]$payload.syncFormulas.budgetQuantity)
    }
    Mark-ConstructionPhase 'syncBudgetCubic'
    if ($budgetTable -ne $null -and $advanceTable -ne $null) {
      Sync-AdvanceBudgetLinks $advanceSheet $advanceTable $budgetTable ([string]$payload.syncFormulas.advanceQuantity) ([string]$payload.syncFormulas.advancePrice)
    }
    Mark-ConstructionPhase 'syncAdvanceBudget'
    Mark-ConstructionPhase 'tablesAndLinks'

    $numberFormats = @{ clp = '$ #.##0'; cantidad = '#.##0,00'; porcentaje = '0,0 %' }
    foreach ($columnFormat in $payload.columnFormats) {
      $columnIndex = [int]$columnFormat.columna
      $columnRange = $sheet.Range($sheet.Cells.Item(3, $columnIndex), $sheet.Cells.Item($lastDataRow, $columnIndex))
      $columnRange.NumberFormatLocal = $numberFormats[[string]$columnFormat.tipo]
      $columnRange.ShrinkToFit = $true
    }

    $sheet.Rows.Item(1).Font.Bold = $true
    $sheet.Rows.Item(2).Font.Bold = $true
    $sheet.Rows.Item(2).WrapText = $true
    $sheet.Rows.Item(2).RowHeight = 32
    $sheet.UsedRange.VerticalAlignment = -4108

    $summary = $null
    foreach ($candidate in $book.Worksheets) { if ($candidate.Name -eq 'Resumen') { $summary = $candidate; break } }
    if ($null -eq $summary) {
      $summary = $book.Worksheets.Add()
      $summary.Name = 'Resumen'
      Set-ConstructionCell $summary 1 1 'Resumen Construction Assistant'
      Set-ConstructionCell $summary 2 1 'Indicador'
      Set-ConstructionCell $summary 2 2 'Valor'
      $summary.Rows.Item(1).Font.Bold = $true
      $summary.Rows.Item(2).Font.Bold = $true
      $summary.Rows.Item(2).WrapText = $true
    } elseif (
      [string]$summary.Cells.Item(1, 1).Text -ne 'Resumen Construction Assistant' -and
      [string]$summary.Cells.Item(1, 1).Text -ne 'CONSTRUCTION ASSISTANT | RESUMEN DE OBRA'
    ) {
      throw "La hoja 'Resumen' ya existe y no pertenece a Construction Assistant; no se modificó para proteger sus datos."
    }

    $summaryMetrics = @()
    if ($baseName -eq 'Cubicaciones') {
      $summaryMetrics += @{ label = 'Superficie acumulada (m²)'; formula = '=IFERROR(SUM(tblCubicaciones[Superficie (m²)]),0)'; format = 'cantidad'; aliases = @('Superficie excavación (m²)') }
      $summaryMetrics += @{ label = 'Volumen acumulado (m³)'; formula = '=IFERROR(SUM(tblCubicaciones[Volumen (m³)]),0)'; format = 'cantidad'; aliases = @('Volumen excavación (m³)') }
      $summaryMetrics += @{ label = 'Cubicaciones por revisar'; formula = '=IFERROR(SUMPRODUCT(--(((tblCubicaciones[Superficie (m²)]="")+(tblCubicaciones[Volumen (m³)]="")+(tblCubicaciones[Volumen ajustado (m³)]="")>0))),0)'; format = 'cantidad'; aliases = @() }
    }
    if ($baseName -eq 'Presupuesto') {
      $summaryMetrics += @{ label = 'Subtotal presupuesto (CLP)'; formula = "='Presupuesto'!E$subtotalRow"; format = 'clp'; aliases = @('Subtotal presupuesto') }
      $summaryMetrics += @{ label = 'IVA presupuesto (CLP)'; formula = "='Presupuesto'!E$ivaRow"; format = 'clp'; aliases = @('IVA 19%') }
      $summaryMetrics += @{ label = 'Total presupuesto (CLP)'; formula = "='Presupuesto'!E$totalRow"; format = 'clp'; aliases = @('Total presupuesto') }
      $summaryMetrics += @{ label = 'Partidas por revisar'; formula = '=COUNTIF(tblPresupuesto[Costo total],"")'; format = 'cantidad'; aliases = @() }
    }
    if ($baseName -eq 'Avance') {
      $units = @()
      for ($row = 1; $row -le $advanceTable.DataBodyRange.Rows.Count; $row++) {
        $unit = [string]$advanceTable.DataBodyRange.Cells.Item($row, 2).Text
        if ($unit -and $units -notcontains $unit) { $units += $unit }
      }
      $legacyAdvanceAliasAvailable = $true
      foreach ($unit in $units) {
        $escapedUnit = $unit.Replace('"', '""')
        $advanceAliases = @()
        $pendingAliases = @()
        if ($legacyAdvanceAliasAvailable) {
          $advanceAliases = @('Avance físico ponderado (%)', 'Avance físico')
          $legacyAdvanceAliasAvailable = $false
        }
        if ($unit -eq 'm³') { $pendingAliases = @('Cantidad pendiente total', 'Saldo pendiente (m³)') }
        $unitCriterion = '"' + $escapedUnit + '"'
        $unitBudget = 'SUMIFS(tblAvance[Cantidad presupuestada],tblAvance[Unidad],' + $unitCriterion + ')'
        $unitExecuted = 'SUMIFS(tblAvance[Cantidad ejecutada],tblAvance[Unidad],' + $unitCriterion + ')'
        $unitPending = 'SUMIFS(tblAvance[Cantidad pendiente],tblAvance[Unidad],' + $unitCriterion + ')'
        $unitCount = 'COUNTIF(tblAvance[Unidad],' + $unitCriterion + ')'
        $advanceFormula = '=IF(' + $unitCount + '=0,"",IFERROR(' + $unitExecuted + '/' + $unitBudget + ',0))'
        $pendingFormula = '=IF(' + $unitCount + '=0,"",IFERROR(' + $unitPending + ',0))'
        $advanceLabel = "Avance físico ponderado ($unit)"
        $pendingLabel = "Saldo pendiente ($unit)"
        $advanceLabelFormula = '=IF(' + $unitCount + '=0,"","' + $advanceLabel.Replace('"', '""') + '")'
        $pendingLabelFormula = '=IF(' + $unitCount + '=0,"","' + $pendingLabel.Replace('"', '""') + '")'
        $summaryMetrics += @{ label = $advanceLabel; labelFormula = $advanceLabelFormula; formula = $advanceFormula; format = 'porcentaje'; aliases = $advanceAliases }
        $summaryMetrics += @{ label = $pendingLabel; labelFormula = $pendingLabelFormula; formula = $pendingFormula; format = 'cantidad'; aliases = $pendingAliases }
       }
       $summaryMetrics += @{ label = 'Costo ejecutado (CLP)'; formula = '=IFERROR(SUM(tblAvance[Costo ejecutado]),0)'; format = 'clp'; aliases = @('Costo ejecutado') }
       $summaryMetrics += @{ label = 'Partidas con sobre-ejecución'; formula = '=COUNTIF(tblAvance[Exceso ejecutado],">0")'; format = 'cantidad'; aliases = @() }
       $summaryMetrics += @{ label = 'Registros por revisar'; formula = '=COUNTIF(tblAvance[Estado de avance],"REVISAR DATOS")'; format = 'cantidad'; aliases = @() }
       $summaryMetrics += @{ label = 'Avances sin vínculo a Presupuesto'; formula = '=IFERROR(SUMPRODUCT(--NOT(ISFORMULA(tblAvance[Cantidad presupuestada]))),0)'; format = 'cantidad'; aliases = @() }
       $summaryMetrics += @{ label = 'Partidas de Presupuesto sin avance (por nombre)'; formula = '=IFERROR(SUMPRODUCT(--(COUNTIF(tblAvance[Partida],tblPresupuesto[Partida])=0)),0)'; format = 'cantidad'; aliases = @() }
    }
    foreach ($metric in $summaryMetrics) {
      $metricRow = 0
      $summaryLastRow = [Math]::Max(2, $summary.UsedRange.Rows.Count)
      for ($row = 3; $row -le $summaryLastRow; $row++) {
        $label = [string]$summary.Cells.Item($row, 1).Text
        $existingLabelFormula = [string]$summary.Cells.Item($row, 1).Formula
        if ($label -eq $metric.label -or $metric.aliases -contains $label -or
          ($metric.labelFormula -and $existingLabelFormula -eq [string]$metric.labelFormula)) { $metricRow = $row; break }
      }
      if ($metricRow -eq 0) {
        $metricRow = $summaryLastRow + 1
        Set-ConstructionCell $summary $metricRow 1 $metric.label
       } elseif (-not $metric.labelFormula -and [string]$summary.Cells.Item($metricRow, 1).Text -ne [string]$metric.label) {
         Set-ConstructionCell $summary $metricRow 1 $metric.label
       }
       if ($metric.labelFormula) {
         $summary.Cells.Item($metricRow, 1).NumberFormat = 'General'
         $summary.Cells.Item($metricRow, 1).Formula = [string]$metric.labelFormula
       }
      $summary.Cells.Item($metricRow, 2).Formula = [string]$metric.formula
      if (-not $summary.Cells.Item($metricRow, 2).HasFormula) {
        throw "No se pudo instalar la fórmula del indicador '$($metric.label)' en Resumen."
      }
      $summary.Cells.Item($metricRow, 2).NumberFormatLocal = $numberFormats[[string]$metric.format]
      $summary.Cells.Item($metricRow, 2).ShrinkToFit = $true
    }
    # Migra resúmenes previos que dejaron etiquetas dinámicas duplicadas
    # como texto de fórmula; limpia solo las celdas A:B repetidas del MCP.
    $seenUnitMetricFormulas = @{}
    $summaryRowsToClear = @()
    $summaryLastRow = $summary.UsedRange.Rows.Count
    for ($row = 3; $row -le $summaryLastRow; $row++) {
      $labelFormula = [string]$summary.Cells.Item($row, 1).Formula
      if ($labelFormula -notmatch '^=IF\(COUNTIF\(tblAvance\[Unidad\],') { continue }
      if ($seenUnitMetricFormulas.ContainsKey($labelFormula)) { $summaryRowsToClear += $row }
      else { $seenUnitMetricFormulas[$labelFormula] = $true }
    }
    foreach ($row in ($summaryRowsToClear | Sort-Object -Descending)) {
      $summary.Range(('A{0}:B{0}' -f $row)).ClearContents() | Out-Null
    }
    Mark-ConstructionPhase 'summary'

    # Excel crea una hoja inicial vacía al crear un libro. Elimina solo
    # nombres predeterminados sin celdas, fórmulas, tablas ni objetos; nunca
    # borra hojas con contenido del usuario ni la última hoja del libro.
    $removedDefaultSheets = @()
    for ($index = $book.Worksheets.Count; $index -ge 1; $index--) {
      if ($book.Worksheets.Count -le 1) { break }
      $candidate = $book.Worksheets.Item($index)
      if (-not [regex]::IsMatch([string]$candidate.Name, [string]$payload.defaultSheetPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) { continue }
      $usedRange = $candidate.UsedRange
      $isEmpty = $usedRange.Cells.Count -eq 1
      if ($isEmpty) {
        $firstCell = $candidate.Cells.Item(1, 1)
        $isEmpty = ($null -eq $firstCell.Value2 -and -not $firstCell.HasFormula)
      }
      if ($isEmpty -and $candidate.Shapes.Count -eq 0 -and $candidate.ListObjects.Count -eq 0) {
        $removedName = [string]$candidate.Name
        $candidate.Delete() | Out-Null
        $removedDefaultSheets += $removedName
      }
    }
    Mark-ConstructionPhase 'styles'

    # Estilo corporativo aplicado únicamente a las hojas administradas.
    # No cambia valores, fórmulas ni formatos numéricos.
    $colors = @{
      petrol = 4865291
      teal = 9206303
      greenTeal = 9411882
      white = 16777215
      lightGray = 16250353
      input = 16118757
      formula = 15791592
      border = 15065301
      text = 3879972
    }
    $book.Windows.Item(1).DisplayGridlines = $false
    $styleSheetNames = @($baseName, 'Resumen')
    if ($baseName -eq 'Cubicaciones') { $styleSheetNames += 'Presupuesto' }
    elseif ($baseName -eq 'Presupuesto') { $styleSheetNames += 'Avance' }
    foreach ($styledSheet in $book.Worksheets) {
      if ($styleSheetNames -notcontains $styledSheet.Name) { continue }
      if ($styledSheet.Name -eq 'Cubicaciones') {
        $last = [Math]::Max(3, $styledSheet.UsedRange.Rows.Count)
        Apply-ConstructionBaseStyle $styledSheet 'I' 'A2:I2' $colors
        Set-ConstructionCubicWarnings $styledSheet $last
        foreach ($column in @('C', 'D', 'E', 'H')) { Set-ConstructionRoleFill $styledSheet $column 3 $last $colors.input }
        foreach ($column in @('F', 'G', 'I')) { Set-ConstructionRoleFill $styledSheet $column 3 $last $colors.formula }
      } elseif ($styledSheet.Name -eq 'Presupuesto') {
        $lastUsed = $styledSheet.UsedRange.Rows.Count
        $visualSubtotalRow = 0
        for ($row = 3; $row -le $lastUsed; $row++) {
          if ([string]$styledSheet.Cells.Item($row, 4).Text -match '^Subtotal$') { $visualSubtotalRow = $row; break }
        }
        $visualTotalRow = if ($visualSubtotalRow -gt 0) { $visualSubtotalRow + 2 } else { $lastUsed }
        $visualLastDataRow = if ($visualSubtotalRow -gt 0) { $visualSubtotalRow - 1 } else { $lastUsed }
        Apply-ConstructionBaseStyle $styledSheet 'E' 'A2:E2' $colors
        Set-ConstructionRoleFill $styledSheet 'D' 3 $visualLastDataRow $colors.input
        $quantityRange = $styledSheet.Range(('C3:C{0}' -f $visualLastDataRow))
        Set-ConstructionRoleFill $styledSheet 'C' 3 $visualLastDataRow $colors.input
        try {
          $formulaQuantities = $quantityRange.SpecialCells(-4123)
          $formulaQuantities.Interior.Color = $colors.formula
          $formulaQuantities.Borders.Color = 14934227
        } catch {}
        Set-ConstructionRoleFill $styledSheet 'E' 3 $visualTotalRow $colors.formula
        if ($visualSubtotalRow -gt 0) {
          $totalAddress = 'D{0}:E{1}' -f $visualSubtotalRow, ($visualTotalRow - 1)
          $styledSheet.Range($totalAddress).Interior.Color = $colors.lightGray
          $styledSheet.Range($totalAddress).Font.Bold = $true
          $grandTotalAddress = 'D{0}:E{0}' -f $visualTotalRow
          $styledSheet.Range($grandTotalAddress).Interior.Color = $colors.petrol
          $styledSheet.Range($grandTotalAddress).Font.Color = $colors.white
          $styledSheet.Range($grandTotalAddress).Font.Bold = $true
        }
      } elseif ($styledSheet.Name -eq 'Avance') {
        $last = [Math]::Max(3, $styledSheet.UsedRange.Rows.Count)
        Apply-ConstructionBaseStyle $styledSheet 'J' 'A2:J2' $colors
        Set-ConstructionRoleFill $styledSheet 'D' 3 $last $colors.input
        foreach ($column in @('E', 'F', 'H', 'I', 'J')) { Set-ConstructionRoleFill $styledSheet $column 3 $last $colors.formula }
        Set-ConstructionAdvanceAlerts $styledSheet $last
        foreach ($column in @('C', 'G')) {
          $inputRange = $styledSheet.Range(('{0}3:{0}{1}' -f $column, $last))
          Set-ConstructionRoleFill $styledSheet $column 3 $last $colors.input
          try {
            $formulaInputs = $inputRange.SpecialCells(-4123)
            $formulaInputs.Interior.Color = $colors.formula
            $formulaInputs.Borders.Color = 14934227
          } catch {}
        }
      } elseif ($styledSheet.Name -eq 'Resumen') {
        $last = [Math]::Max(2, $styledSheet.UsedRange.Rows.Count)
        Apply-ConstructionBaseStyle $styledSheet 'D' 'A2:B2' $colors
        for ($row = 3; $row -le $last; $row++) {
          $metricAddress = 'A{0}:B{0}' -f $row
          $label = [string]$styledSheet.Cells.Item($row, 1).Text
          $styledSheet.Range($metricAddress).Borders.Color = $colors.border
          $styledSheet.Cells.Item($row, 1).Interior.Color = $colors.lightGray
          $styledSheet.Cells.Item($row, 1).Font.Color = $colors.text
          $styledSheet.Cells.Item($row, 1).Font.Bold = $true
          $styledSheet.Cells.Item($row, 2).Interior.Color = $colors.formula
          $styledSheet.Cells.Item($row, 2).Font.Color = $colors.petrol
          $styledSheet.Cells.Item($row, 2).Font.Bold = $true
          $styledSheet.Cells.Item($row, 2).Font.Size = 14
          $styledSheet.Cells.Item($row, 2).HorizontalAlignment = -4152
          if ($label -match 'Total presupuesto|Avance físico|Costo ejecutado|sobre-ejecución|por revisar|sin vínculo|sin avance') {
            $styledSheet.Range($metricAddress).Interior.Color = $colors.greenTeal
            $styledSheet.Range($metricAddress).Font.Color = $colors.white
            $styledSheet.Range($metricAddress).Font.Bold = $true
          }
        }
      }
    }
    Mark-ConstructionPhase 'charts'

    # Gráfico de costos por partida, enlazado a los costos calculados en
    # Presupuesto. En llamadas posteriores se reconstruye para incluir las
    # nuevas filas sin duplicar el gráfico.
    $budgetSheet = $null
    foreach ($candidate in $book.Worksheets) { if ($candidate.Name -eq 'Presupuesto') { $budgetSheet = $candidate; break } }
    if ($budgetSheet -ne $null -and $budgetSheet.ListObjects.Count -gt 0) {
      $budgetTableForChart = $budgetSheet.ListObjects.Item('tblPresupuesto')
      $hasLinearChart = $false
      $hasLogChart = $false
      foreach ($existingChart in $summary.ChartObjects()) {
        if ($existingChart.Name -eq 'ConstructionAssistantBudgetByItem') { $hasLinearChart = $true }
        elseif ($existingChart.Name -eq 'ConstructionAssistantBudgetByItemLog') { $hasLogChart = $true }
      }
      $hasPositiveCost = $false
      foreach ($budgetRow in $budgetTableForChart.ListRows) {
        $quantityCell = $budgetRow.Range.Cells.Item(1, 3)
        $priceCell = $budgetRow.Range.Cells.Item(1, 4)
        $quantityIsPositive = ($quantityCell.Value2 -is [double] -and [double]$quantityCell.Value2 -gt 0) -or ([string]$quantityCell.Formula -match 'tblCubicaciones')
        if ($quantityIsPositive -and $priceCell.Value2 -is [double] -and [double]$priceCell.Value2 -gt 0) { $hasPositiveCost = $true; break }
      }
      $shouldRefreshCharts = ($baseName -eq 'Presupuesto') -or (-not $hasLinearChart) -or ($hasPositiveCost -and -not $hasLogChart) -or (-not $hasPositiveCost -and $hasLogChart) -or ($budgetTableForChart.ListRows.Count -eq 0 -and ($hasLinearChart -or $hasLogChart))
      if ($shouldRefreshCharts) {
        $charts = $summary.ChartObjects()
        for ($index = $charts.Count; $index -ge 1; $index--) {
          $chartObject = $charts.Item($index)
          if ($chartObject.Name -eq 'ConstructionAssistantBudgetMix' -or $chartObject.Name -eq 'ConstructionAssistantBudgetByItem' -or $chartObject.Name -eq 'ConstructionAssistantBudgetByItemLog') {
            $chartObject.Delete() | Out-Null
          }
        }
        if ($budgetTableForChart.ListRows.Count -gt 0) {
          $chartObject = $summary.ChartObjects().Add($summary.Columns.Item(4).Left, $summary.Rows.Item(2).Top, 520, 300)
          $chartObject.Name = 'ConstructionAssistantBudgetByItem'
          $chart = $chartObject.Chart
          $chart.ChartType = 57
          $chart.ChartStyle = 10
          while ($chart.SeriesCollection().Count -gt 0) { $chart.SeriesCollection().Item(1).Delete() | Out-Null }
          $series = $chart.SeriesCollection().NewSeries()
          $seriesFormula = '=SERIES("Costo por partida",tblPresupuesto[Partida],tblPresupuesto[Costo total],1)'
          $series.Formula = $seriesFormula
          $chart.HasTitle = $true
          $chart.ChartTitle.Text = 'Costos por partida (CLP)'
          $chart.ChartTitle.Font.Color = $colors.petrol
          $chart.HasLegend = $false
          try { $series.Format.Fill.ForeColor.RGB = $colors.teal } catch {}
          try { $series.ApplyDataLabels() | Out-Null; $series.DataLabels().NumberFormatLocal = '$ #.##0' } catch {}
          try { $chart.Axes(2).TickLabels.NumberFormatLocal = '$ #.##0' } catch {}

          # Vista complementaria en escala logarítmica, con referencias a las
          # mismas filas de presupuesto; los costos cero se omiten en el eje.
          if ($hasPositiveCost) {
            $logObject = $summary.ChartObjects().Add($summary.Columns.Item(4).Left, $summary.Rows.Item(18).Top, 520, 300)
            $logObject.Name = 'ConstructionAssistantBudgetByItemLog'
            $logChart = $logObject.Chart
            $logChart.ChartType = 57
            $logChart.ChartStyle = 10
            $logSeries = $logChart.SeriesCollection().NewSeries()
            $logSeries.Formula = '=SERIES("Costo por partida",tblPresupuesto[Partida],tblPresupuesto[Costo total],1)'
            $logChart.HasTitle = $true
            $logChart.ChartTitle.Text = 'Costos por partida (escala logarítmica)'
            $logChart.ChartTitle.Font.Color = $colors.petrol
            $logChart.HasLegend = $false
            $logChart.Axes(2).ScaleType = 1
            $logChart.Axes(2).LogBase = 10
            try { $logSeries.Format.Fill.ForeColor.RGB = $colors.greenTeal } catch {}
            try { $logSeries.ApplyDataLabels() | Out-Null; $logSeries.DataLabels().NumberFormatLocal = '$ #.##0' } catch {}
            try { $logChart.Axes(2).TickLabels.NumberFormatLocal = '$ #.##0' } catch {}
          }
        }
      }
    }

    # Reconstruye dependencias y calcula antes de AutoFit para que también
    # se ajusten los anchos a los resultados reales de las fórmulas.
    $excel.Calculation = -4105
    $excel.CalculateBeforeSave = $true
    $book.ForceFullCalculation = $true
    $excel.CalculateFullRebuild() | Out-Null
    foreach ($styledSheet in $book.Worksheets) {
      if ($styleSheetNames -contains $styledSheet.Name) {
        $styledSheet.UsedRange.Columns.AutoFit() | Out-Null
      }
    }
    $book.Save() | Out-Null
    Mark-ConstructionPhase 'calculateAutofitSave'
    $response = @{ path = $file; sheet = $sheet.Name; rows = $items.Count; appended = $wasExisting; removedDefaultSheets = $removedDefaultSheets; formulaDriven = $true }
    if ($baseName -eq 'Avance') {
      $response.appendedItems = @()
      for ($index = 0; $index -lt $items.Count; $index++) {
        $row = $firstNewRow + $index
        $response.appendedItems += @{
          nombre = [string]$sheet.Cells.Item($row, 1).Value2
          unidad = [string]$sheet.Cells.Item($row, 2).Value2
          cantidadPresupuestada = $sheet.Cells.Item($row, 3).Value2
          cantidadEjecutada = $sheet.Cells.Item($row, 4).Value2
          avanceFisico = $sheet.Cells.Item($row, 5).Value2
          cantidadPendiente = $sheet.Cells.Item($row, 6).Value2
          precioUnitario = $sheet.Cells.Item($row, 7).Value2
          costoEjecutado = $sheet.Cells.Item($row, 8).Value2
          estadoAvance = [string]$sheet.Cells.Item($row, 9).Value2
          sobreEjecucion = $sheet.Cells.Item($row, 10).Value2
        }
      }
    }
    if ($script:ConstructionProfile) { $response.profile = $script:ConstructionTimings }
    $response | ConvertTo-Json -Compress
  } else { throw 'Operación COM no reconocida.' }
} catch {
  [Console]::Error.WriteLine($_.Exception.Message + ' ' + $_.InvocationInfo.PositionMessage)
  exit 2
} finally {
  if ($book -ne $null) { $book.Close($false) | Out-Null }
  if ($excel -ne $null) { $excel.Quit() | Out-Null }
  if ($book -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
  if ($excel -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}`;

const POWERSHELL_COPY_SCRIPT = String.raw`
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
try {
  $source = [Environment]::GetEnvironmentVariable('CONSTRUCTION_EXCEL_TEMPLATE')
  $destination = [Environment]::GetEnvironmentVariable('CONSTRUCTION_EXCEL_TEMP_DESTINATION')
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AutomationSecurity = 3
  $book = $excel.Workbooks.Open($source, 0, $true)
  $book.SaveAs($destination, 51) | Out-Null
  @{ source = $source; destination = $destination; copied = $true } | ConvertTo-Json -Compress
} catch {
  [Console]::Error.WriteLine($_.Exception.Message + ' ' + $_.InvocationInfo.PositionMessage)
  exit 2
} finally {
  if ($book -ne $null) { $book.Close($false) | Out-Null }
  if ($excel -ne $null) { $excel.Quit() | Out-Null }
  if ($book -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
  if ($excel -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}`;

export async function crearLibroDesdePlantilla(templatePath: string, destinationPath: string): Promise<{ path: string; templatePath: string; copied: true }> {
  if (process.platform !== "win32") throw new Error("La creación mediante Excel COM solo está disponible en Windows.");
  if (!path.isAbsolute(templatePath) || !path.isAbsolute(destinationPath)) {
    throw new Error("Indica rutas absolutas para la plantilla y el nuevo libro.");
  }
  const absoluteTemplate = path.resolve(templatePath);
  const absoluteDestination = path.resolve(destinationPath);
  if (absoluteTemplate.toLowerCase() === absoluteDestination.toLowerCase()) {
    throw new Error("La ruta de destino debe ser distinta de la plantilla.");
  }
  if (path.extname(absoluteTemplate).toLowerCase() !== ".xlsx") {
    throw new Error("Para conservar íntegramente la plantilla, usa un libro .xlsx.");
  }
  if (path.extname(absoluteDestination).toLowerCase() !== ".xlsx") {
    throw new Error("El nuevo libro debe guardarse con extensión .xlsx.");
  }
  try {
    await access(absoluteTemplate);
  } catch {
    throw new Error("No se encontró la plantilla Excel indicada o no hay permiso para leerla.");
  }
  let destinationExists = false;
  try { await access(absoluteDestination); destinationExists = true; } catch {}
  if (destinationExists) throw new Error("La ruta de destino ya existe; no se sobrescribió ningún archivo.");

  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "construction-assistant-copy-"));
  const temporaryWorkbook = path.join(temporaryDirectory, "template-copy.xlsx");
  let result;
  try {
    await writeFile(path.join(temporaryDirectory, "copy-template.ps1"), `\uFEFF${POWERSHELL_COPY_SCRIPT}`, "utf8");
    const scriptPath = path.join(temporaryDirectory, "copy-template.ps1");
    result = await run(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", scriptPath],
      {
        timeoutMs: 90_000,
        env: {
          CONSTRUCTION_EXCEL_TEMPLATE: absoluteTemplate,
          CONSTRUCTION_EXCEL_TEMP_DESTINATION: temporaryWorkbook,
        },
      },
    );
    if (result.code !== 0) {
      throw new Error(result.stderr.trim() || "Excel COM no pudo crear el libro desde la plantilla.");
    }
    await access(temporaryWorkbook);
    await copyFile(temporaryWorkbook, absoluteDestination, constants.COPYFILE_EXCL);
    return { path: absoluteDestination, templatePath: absoluteTemplate, copied: true };
  } catch (e) {
    if (e instanceof Error && "code" in e && e.code === "EEXIST") {
      throw new Error("La ruta de destino ya existe; no se sobrescribió ningún archivo.");
    }
    throw new Error(`No se pudo crear el libro Excel desde la plantilla: ${e instanceof Error ? e.message : String(e)}`);
  } finally {
    await rm(temporaryDirectory, { recursive: true, force: true });
  }
}

export type ExcelCell = string | number | null;

export async function ejecutarExcel(
  filePath: string,
  mode: "sheets" | "audit" | "append",
  payload: { sheetBase?: HojaConstruction; rows?: ExcelCell[][]; tasaIva?: number } = {},
): Promise<unknown> {
  if (process.platform !== "win32") throw new Error("La conexión COM con Excel solo está disponible en Windows.");
  if (!path.isAbsolute(filePath)) throw new Error("Indica una ruta absoluta al libro Excel.");
  const extension = path.extname(filePath).toLowerCase();
  if (![".xlsx", ".xlsm", ".xls"].includes(extension)) {
    throw new Error("Usa la ruta de un libro Excel existente con extensión .xlsx, .xlsm o .xls.");
  }
  const absolutePath = path.resolve(filePath);
  try {
    await access(absolutePath);
  } catch {
    throw new Error("No se encontró el libro indicado o no hay permiso para acceder. Verifica la ruta.");
  }
  const columnFormats = mode === "append" ? inferirFormatosColumnas(payload.rows?.[1] ?? []) : [];
  const plan: PlanHojaExcel | undefined = mode === "append" && payload.sheetBase
    ? crearPlanHojaExcel(payload.sheetBase, payload.tasaIva)
    : undefined;
  const processPayload = mode === "append"
    ? {
        ...payload,
        items: (payload.rows ?? []).slice(2),
        columnFormats,
        plan,
        defaultSheetPattern: PATRON_HOJA_PREDETERMINADA.source,
        syncFormulas: {
          budgetQuantity: FORMULA_CANTIDAD_DESDE_CUBICACION,
          advanceQuantity: FORMULA_CANTIDAD_DESDE_PRESUPUESTO,
          advancePrice: FORMULA_PRECIO_DESDE_PRESUPUESTO,
        },
      }
    : payload;
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "construction-assistant-"));
  const scriptPath = path.join(temporaryDirectory, "excel-com.ps1");
  let result;
  try {
    // Ejecutar desde un archivo temporal evita el límite de longitud de
    // CreateProcess en Windows cuando el script COM incorpora más fórmulas.
    // El BOM garantiza que PowerShell 5.1 lea correctamente el texto UTF-8.
    await writeFile(scriptPath, `\uFEFF${POWERSHELL_SCRIPT}`, "utf8");
    result = await run(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", scriptPath],
      {
        timeoutMs: 90_000,
        env: {
          CONSTRUCTION_EXCEL_PATH: absolutePath,
          CONSTRUCTION_EXCEL_MODE: mode,
          CONSTRUCTION_EXCEL_PAYLOAD: JSON.stringify(processPayload),
          ...(process.env.CONSTRUCTION_EXCEL_PROFILE === "1" ? { CONSTRUCTION_EXCEL_PROFILE: "1" } : {}),
        },
      },
    );
  } finally {
    await rm(temporaryDirectory, { recursive: true, force: true });
  }
  if (result.code !== 0) {
    const detalle = result.stderr.trim();
    throw new Error(detalle || "No se pudo automatizar Excel. Confirma que Excel esté instalado y que el libro no esté bloqueado.");
  }
  try {
    return JSON.parse(result.stdout.trim());
  } catch {
    const detalle = result.stdout.trim().slice(0, 800);
    throw new Error(`Excel completó la operación, pero no se pudo interpretar su respuesta${detalle ? `: ${detalle}` : "."}`);
  }
}
