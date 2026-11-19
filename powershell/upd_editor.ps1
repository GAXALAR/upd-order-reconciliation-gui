Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Глобальные переменные для путей
$global:pathOrder = ""
$global:pathUpd = ""
$script:configPath = Join-Path (Split-Path $PSScriptRoot -Parent) "config.json"

function Get-NormalizedArticle {
    param($value, $numberFormat)
    if ($null -eq $value) { return "" }
    if ($value -is [string]) { return $value.Trim() }

    $isNumeric = $value -is [byte] -or $value -is [int16] -or
        $value -is [int32] -or $value -is [int64] -or
        $value -is [single] -or $value -is [double] -or $value -is [decimal]
    if ($isNumeric) {
        $text = [Convert]::ToString($value, [Globalization.CultureInfo]::InvariantCulture)
        $numeric = [double]$value
        if ($numeric -eq [Math]::Floor($numeric)) {
            $text = ([decimal]$value).ToString("0", [Globalization.CultureInfo]::InvariantCulture)
            if ($numberFormat -is [string] -and $numberFormat -match '^0+$') {
                $negative = $text.StartsWith("-")
                $digits = if ($negative) { $text.Substring(1) } else { $text }
                $digits = $digits.PadLeft($numberFormat.Length, '0')
                if ($negative) { $text = "-" + $digits } else { $text = $digits }
            }
        }
        return $text
    }
    return ([string]$value).Trim()
}

function Get-RangeArrayValue {
    param($values, [int]$rowOffset)
    if (-not ($values -is [Array])) { return $values }

    $rowIndex = $values.GetLowerBound(0) + $rowOffset
    if ($values.Rank -eq 1) {
        return $values.GetValue($rowIndex)
    }
    return $values.GetValue($rowIndex, $values.GetLowerBound(1))
}

function Test-NumericCellValue {
    param($value)
    return $value -is [byte] -or $value -is [int16] -or
        $value -is [int32] -or $value -is [int64] -or
        $value -is [single] -or $value -is [double] -or $value -is [decimal]
}

function Get-CellNumberFormat {
    param($cells, [int]$row, [int]$column, $rangeNumberFormat)
    if ($rangeNumberFormat -is [string]) { return $rangeNumberFormat }

    $cell = $cells.Item($row, $column)
    try {
        return $cell.NumberFormat
    } finally {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($cell) | Out-Null
    }
}

function ConvertTo-ColumnIndex {
    param([string]$columnName)
    if ($columnName -notmatch '^[A-Za-z]{1,3}$') {
        throw "Некорректное имя столбца: '$columnName'."
    }

    $index = 0
    foreach ($character in $columnName.ToUpperInvariant().ToCharArray()) {
        $index = ($index * 26) + ([int][char]$character - [int][char]'A' + 1)
    }
    if ($index -gt 16384) {
        throw "Столбец '$columnName' выходит за пределы Excel."
    }
    return $index
}

function Get-AppConfig {
    if (-not (Test-Path -LiteralPath $script:configPath -PathType Leaf)) {
        throw "Файл конфигурации не найден: $script:configPath"
    }
    try {
        $config = Get-Content -LiteralPath $script:configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        throw "Не удалось прочитать config.json: $($_.Exception.Message)"
    }
    if (-not $config.order -or -not $config.upd) {
        throw "В config.json должны быть секции order и upd."
    }

    foreach ($section in @($config.order, $config.upd)) {
        foreach ($key in @("sheet_index", "first_data_row")) {
            $value = $section.$key
            if ($null -eq $value -or $value -is [bool] -or $value -lt 1 -or [double]$value -ne [Math]::Floor([double]$value)) {
                throw "Параметр $key в config.json должен быть положительным целым числом."
            }
        }
    }
    if ($null -eq $config.upd.order_number_row_offset -or
        $config.upd.order_number_row_offset -is [bool] -or
        [double]$config.upd.order_number_row_offset -ne [Math]::Floor([double]$config.upd.order_number_row_offset)) {
        throw "Параметр upd.order_number_row_offset должен быть целым числом."
    }

    foreach ($columnName in @(
        $config.order.internal_code_column,
        $config.order.supplier_article_column,
        $config.order.order_number_label_column,
        $config.order.order_number_value_column,
        $config.upd.article_column,
        $config.upd.order_number_column
    )) {
        $null = ConvertTo-ColumnIndex $columnName
    }
    if ([string]::IsNullOrWhiteSpace([string]$config.order.order_number_label) -or
        [string]::IsNullOrWhiteSpace([string]$config.upd.basis_label)) {
        throw "Метки order_number_label и basis_label в config.json не должны быть пустыми."
    }
    return $config
}

function Write-UiLog {
    param([string]$text)
    $script:uiEvents.Enqueue([PSCustomObject]@{
        Type = "log"
        Text = $text
    })
}

function Show-UiMessage {
    param(
        [string]$text,
        [string]$caption,
        [System.Windows.Forms.MessageBoxButtons]$buttons = [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]$icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )

    $completion = New-Object System.Threading.ManualResetEventSlim($false)
    $request = [hashtable]::Synchronized(@{
        Type = "dialog"
        Text = $text
        Caption = $caption
        Buttons = $buttons
        Icon = $icon
        Completion = $completion
        Result = $null
    })
    $script:uiEvents.Enqueue($request)
    $completion.Wait()
    $result = $request.Result
    $completion.Dispose()
    return $result
}

function Confirm-Changes {
    param([string]$text)
    return (Show-UiMessage $text "Подтверждение изменений" `
        ([System.Windows.Forms.MessageBoxButtons]::YesNo) `
        ([System.Windows.Forms.MessageBoxIcon]::Question)) -eq [System.Windows.Forms.DialogResult]::Yes
}

function Start-Processing {
    param($logOutput)

    if (-not $global:pathOrder -or -not $global:pathUpd) {
        Show-UiMessage "Выберите оба файла перед началом!" "Ошибка" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    if (-not (Test-Path -LiteralPath $global:pathOrder -PathType Leaf) -or
        -not (Test-Path -LiteralPath $global:pathUpd -PathType Leaf)) {
        Show-UiMessage "Один из входных файлов не существует." "Ошибка" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }
    if (([IO.Path]::GetFullPath($global:pathOrder)) -ieq ([IO.Path]::GetFullPath($global:pathUpd))) {
        Show-UiMessage "Файл заказа и файл УПД должны быть разными." "Ошибка" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }
    if ([IO.Path]::GetExtension($global:pathOrder) -ine ".xlsx" -or [IO.Path]::GetExtension($global:pathUpd) -ine ".xlsx") {
        Show-UiMessage "Поддерживаются только файлы формата .xlsx." "Ошибка" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }
    try {
        $config = Get-AppConfig
        $orderInternalColumn = ConvertTo-ColumnIndex $config.order.internal_code_column
        $orderArticleColumn = ConvertTo-ColumnIndex $config.order.supplier_article_column
        $orderLabelColumn = ConvertTo-ColumnIndex $config.order.order_number_label_column
        $orderValueColumn = ConvertTo-ColumnIndex $config.order.order_number_value_column
        $updArticleColumn = ConvertTo-ColumnIndex $config.upd.article_column
        $updOrderColumn = ConvertTo-ColumnIndex $config.upd.order_number_column
    } catch {
        Show-UiMessage "Ошибка конфигурации:`r`n$($_.Exception.Message)" "Ошибка" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    $excel = $null
    $workbooks = $null
    $wbOrder = $null
    $wbOrderClosed = $false
    $orderSheets = $null
    $wsOrder = $null
    $orderCells = $null
    $orderCodeRange = $null
    $orderArticleRange = $null
    $orderLabelRange = $null
    $orderValueRange = $null
    $wbUpd = $null
    $wbUpdClosed = $false
    $updSheets = $null
    $wsUpd = $null
    $updCells = $null
    $updArticleRange = $null
    $changeRange = $null
    $wbVerify = $null
    $wbVerifyClosed = $false
    $verifySheets = $null
    $verifySheet = $null
    $verifyCells = $null
    $verifyCell = $null
    $usedOrder = $null
    $usedUpd = $null
    $foundBase = $null
    $findAfterCell = $null
    $basisMatches = @()
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbooks = $excel.Workbooks
    } catch {
        if ($excel) {
            try { $excel.Quit() } catch {}
            try { [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
        }
        Show-UiMessage "Не удалось запустить MS Excel. Убедитесь, что Excel установлен на компьютере." "Ошибка COM-объекта" | Out-Null
        return
    }

    try {
        Write-UiLog "--- Старт обработки: $(Get-Date -Format 'HH:mm:ss') ---`r`n"
        
        # Загрузка заказа
        $wbOrder = $workbooks.Open($global:pathOrder)
        $orderSheets = $wbOrder.Sheets
        if ($config.order.sheet_index -gt $orderSheets.Count) {
            throw "В заказе отсутствует лист $($config.order.sheet_index), указанный в config.json."
        }
        $wsOrder = $orderSheets.Item($config.order.sheet_index)
        $orderCells = $wsOrder.Cells
        $artMap = @{}
        $duplicates = @{}
        $orderNumber = ""

        $usedOrder = $wsOrder.UsedRange
        $firstOrderRow = $usedOrder.Row
        $lastOrderRow = $usedOrder.Row + $usedOrder.Rows.Count - 1
        # Read columns in blocks; query individual formats only when mixed formats
        # prevent preserving numeric identifiers with custom leading-zero formats.
        $orderCodeRange = $wsOrder.Range(("{0}{1}:{0}{2}" -f $config.order.internal_code_column, $firstOrderRow, $lastOrderRow))
        $orderArticleRange = $wsOrder.Range(("{0}{1}:{0}{2}" -f $config.order.supplier_article_column, $firstOrderRow, $lastOrderRow))
        $orderLabelRange = $wsOrder.Range(("{0}{1}:{0}{2}" -f $config.order.order_number_label_column, $firstOrderRow, $lastOrderRow))
        $orderValueRange = $wsOrder.Range(("{0}{1}:{0}{2}" -f $config.order.order_number_value_column, $firstOrderRow, $lastOrderRow))
        $orderCodeValues = $orderCodeRange.Value2
        $orderArticleValues = $orderArticleRange.Value2
        $orderLabelValues = $orderLabelRange.Value2
        $orderNumberValues = $orderValueRange.Value2
        $orderCodeFormats = $orderCodeRange.NumberFormat
        $orderArticleFormats = $orderArticleRange.NumberFormat
        $orderNumberFormats = $orderValueRange.NumberFormat
        for ($i = $firstOrderRow; $i -le $lastOrderRow; $i++) {
            $rowOffset = $i - $firstOrderRow
            $cellBValue = Get-RangeArrayValue $orderCodeValues $rowOffset
            $cellCValue = Get-RangeArrayValue $orderArticleValues $rowOffset
            $orderLabel = Get-NormalizedArticle (Get-RangeArrayValue $orderLabelValues $rowOffset)
            $orderValue = Get-RangeArrayValue $orderNumberValues $rowOffset
            $cellBFormat = $orderCodeFormats
            $cellCFormat = $orderArticleFormats
            $orderValueFormat = $orderNumberFormats
            if (Test-NumericCellValue $cellBValue) {
                $cellBFormat = Get-CellNumberFormat $orderCells $i $orderInternalColumn $orderCodeFormats
            }
            if (Test-NumericCellValue $cellCValue) {
                $cellCFormat = Get-CellNumberFormat $orderCells $i $orderArticleColumn $orderArticleFormats
            }
            if (Test-NumericCellValue $orderValue) {
                $orderValueFormat = Get-CellNumberFormat $orderCells $i $orderValueColumn $orderNumberFormats
            }
            $cellB = Get-NormalizedArticle $cellBValue $cellBFormat
            $cellC = Get-NormalizedArticle $cellCValue $cellCFormat
            $orderValue = Get-NormalizedArticle $orderValue $orderValueFormat
            if ($i -ge $config.order.first_data_row -and $cellB -and $cellC) {
                $article = $cellC
                $internalCode = $cellB
                if ($artMap.ContainsKey($article) -and $artMap[$article] -ne $internalCode) {
                    $duplicates[$article] = $true
                    $artMap.Remove($article)
                } elseif (-not $duplicates.ContainsKey($article)) {
                    $artMap[$article] = $internalCode
                }
            }
            if ($orderLabel.Contains([string]$config.order.order_number_label)) {
                $orderNumber = $orderValue
            }
        }
        $wbOrder.Close($false)
        $wbOrderClosed = $true
        if (-not $orderNumber) {
            Show-UiMessage "Номер заказа не найден в файле. Обработка остановлена." "Ошибка" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            return
        }
        if ($artMap.Count -eq 0) {
            Show-UiMessage "Не найдены сопоставления по структуре, заданной в config.json." "Ошибка" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            return
        }
        Write-UiLog "Загружено артикулов: $($artMap.Count). Номер заказа: $orderNumber`r`n"
        foreach ($article in $duplicates.Keys) {
            Write-UiLog "[ДУБЛИКАТ] Артикул '$article' имеет неоднозначное соответствие.`r`n"
        }

        # Обработка УПД
        $wbUpd = $workbooks.Open($global:pathUpd)
        $updSheets = $wbUpd.Sheets
        if ($config.upd.sheet_index -gt $updSheets.Count) {
            throw "В УПД отсутствует лист $($config.upd.sheet_index), указанный в config.json."
        }
        $wsUpd = $updSheets.Item($config.upd.sheet_index)
        $updCells = $wsUpd.Cells

        $usedUpd = $wsUpd.UsedRange
        $firstUpdRow = [Math]::Max([int]$config.upd.first_data_row, $usedUpd.Row)
        $lastUpdRow = $usedUpd.Row + $usedUpd.Rows.Count - 1
        $updArticleRange = $wsUpd.Range(("{0}{1}:{0}{2}" -f $config.upd.article_column, $firstUpdRow, $lastUpdRow))
        $updArticleValues = $updArticleRange.Value2
        $updArticleFormats = $updArticleRange.NumberFormat
        $basisLabel = [string]$config.upd.basis_label
        $preview = @()
        for ($j = $firstUpdRow; $j -le $lastUpdRow; $j++) {
            $currentValue = Get-RangeArrayValue $updArticleValues ($j - $firstUpdRow)
            $currentFormat = $updArticleFormats
            if (Test-NumericCellValue $currentValue) {
                $currentFormat = Get-CellNumberFormat $updCells $j $updArticleColumn $updArticleFormats
            }
            $currentArt = Get-NormalizedArticle $currentValue $currentFormat
            if (-not $currentArt) { continue }

            if ($artMap.ContainsKey($currentArt)) {
                $preview += [PSCustomObject]@{
                    Row = $j
                    OldArticle = $currentArt
                    NewArticle = $artMap[$currentArt]
                }
            } else {
                Write-UiLog "[НЕ НАЙДЕН] Строка $j`: артикул '$currentArt'`r`n"
            }
        }

        if ($preview.Count -eq 0) {
            Show-UiMessage "Не найдено совпадений. Проверьте столбец и строку начала данных в config.json." "Предпросмотр" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            $wbUpd.Close($false)
            $wbUpdClosed = $true
            return
        }

        # Validate the order-number destination before asking the user to apply changes.
        $findAfterCell = $updCells.Item(1, 1)
        try {
            $foundBase = $updCells.Find(
                [string]$config.upd.basis_label,
                $findAfterCell,
                -4163,
                1,
                1,
                1,
                $true,
                $false,
                $false
            )
        } finally {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($findAfterCell) | Out-Null
            $findAfterCell = $null
        }
        if ($foundBase) {
            $basisMatches += $foundBase
            $firstAddress = $foundBase.Address($false, $false)
            do {
                $foundBase = $updCells.FindNext($foundBase)
                if (-not $foundBase) { break }
                if ($foundBase.Address($false, $false) -eq $firstAddress) { break }
                $basisMatches += $foundBase
            } while ($true)
        }
        $foundBase = $null
        if ($basisMatches.Count -eq 0) {
        Show-UiMessage "Поле `"$basisLabel`" не найдено. Обработка остановлена." "Ошибка" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            $wbUpd.Close($false)
            $wbUpdClosed = $true
            return
        }
        if ($basisMatches.Count -gt 1) {
            Show-UiMessage "Найдено несколько полей `"$basisLabel`". Обработка остановлена." "Ошибка" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            $wbUpd.Close($false)
            $wbUpdClosed = $true
            return
        }

        $targetRow = $basisMatches[0].Row + [int]$config.upd.order_number_row_offset
        if ($targetRow -le 0) {
            Show-UiMessage "Поле `"$basisLabel`" находится в первой строке; невозможно определить строку для номера заказа." "Ошибка" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            $wbUpd.Close($false)
            $wbUpdClosed = $true
            return
        }

        $previewText = ($preview | ForEach-Object { "Строка $($_.Row): $($_.OldArticle) -> $($_.NewArticle)" }) -join "`r`n"
        $targetColumnName = ([string]$config.upd.order_number_column).ToUpperInvariant()
        $previewText += "`r`n`r`n${basisLabel}: номер заказа '$orderNumber' будет записан в ячейку ${targetColumnName}$targetRow."
        if (-not (Confirm-Changes "Будут внесены следующие изменения:`r`n`r`n$previewText`r`n`r`nПродолжить?")) {
            Write-UiLog "Операция отменена пользователем.`r`n"
            $wbUpd.Close($false)
            $wbUpdClosed = $true
            return
        }

        $changeIndex = 0
        while ($changeIndex -lt $preview.Count) {
            $groupStart = $changeIndex
            $groupEnd = $changeIndex
            while ($groupEnd + 1 -lt $preview.Count -and
                $preview[$groupEnd + 1].Row -eq $preview[$groupEnd].Row + 1) {
                $groupEnd++
            }
            $groupCount = $groupEnd - $groupStart + 1
            $firstChangeRow = $preview[$groupStart].Row
            $lastChangeRow = $preview[$groupEnd].Row
            $changeRange = $wsUpd.Range(("{0}{1}:{0}{2}" -f $config.upd.article_column, $firstChangeRow, $lastChangeRow))
            try {
                if ($groupCount -eq 1) {
                    $changeRange.Value2 = $preview[$groupStart].NewArticle
                } else {
                    $changeValues = New-Object 'object[,]' $groupCount, 1
                    for ($offset = 0; $offset -lt $groupCount; $offset++) {
                        $changeValues[$offset, 0] = $preview[$groupStart + $offset].NewArticle
                    }
                    $changeRange.Value2 = $changeValues
                }
            } finally {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($changeRange) | Out-Null
                $changeRange = $null
            }
            for ($offset = 0; $offset -lt $groupCount; $offset++) {
                $change = $preview[$groupStart + $offset]
                Write-UiLog "[OK] Строка $($change.Row): заменён артикул '$($change.OldArticle)' на '$($change.NewArticle)'`r`n"
            }
            $changeIndex = $groupEnd + 1
        }

        $targetCell = $updCells.Item($targetRow, $updOrderColumn)
        try {
            $targetCell.Value2 = $orderNumber
        } finally {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($targetCell) | Out-Null
        }
        Write-UiLog "Номер заказа вписан в строку $targetRow.`r`n"

        $directory = [IO.Path]::GetDirectoryName($global:pathUpd)
        $name = [IO.Path]::GetFileNameWithoutExtension($global:pathUpd)
        $outputPath = Join-Path $directory ($name + "_reconciled.xlsx")
        $counter = 1
        while (Test-Path -LiteralPath $outputPath) {
            $outputPath = Join-Path $directory (($name + "_reconciled_" + $counter) + ".xlsx")
            $counter++
        }
        $wbUpd.SaveAs($outputPath, 51)
        $wbUpd.Close($true)
        $wbUpdClosed = $true

        # Reopen the serialized output and validate every requested cell change.
        $wbVerify = $workbooks.Open($outputPath, 0, $true)
        $verifySheets = $wbVerify.Sheets
        $verifySheet = $verifySheets.Item($config.upd.sheet_index)
        $verifyCells = $verifySheet.Cells
        $verifyFirstRow = $preview[0].Row
        $verifyLastRow = $preview[$preview.Count - 1].Row
        $verifyRange = $verifySheet.Range(("{0}{1}:{0}{2}" -f $config.upd.article_column, $verifyFirstRow, $verifyLastRow))
        try {
            $verifiedValues = $verifyRange.Value2
            $verifiedFormats = $verifyRange.NumberFormat
        } finally {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($verifyRange) | Out-Null
            $verifyRange = $null
        }
        foreach ($change in $preview) {
            $savedValue = Get-RangeArrayValue $verifiedValues ($change.Row - $verifyFirstRow)
            $savedFormat = $verifiedFormats
            if (Test-NumericCellValue $savedValue) {
                $savedFormat = Get-CellNumberFormat $verifyCells $change.Row $updArticleColumn $verifiedFormats
            }
            $savedArticle = Get-NormalizedArticle $savedValue $savedFormat
            if ($savedArticle -cne $change.NewArticle) {
                throw "Проверка результата не пройдена: ячейка $($config.upd.article_column)$($change.Row) содержит '$savedArticle', ожидалось '$($change.NewArticle)'."
            }
        }
        $verifyCell = $verifyCells.Item($targetRow, $updOrderColumn)
        try {
            $savedOrderNumber = [string]$verifyCell.Value2
        } finally {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($verifyCell) | Out-Null
            $verifyCell = $null
        }
        if ($savedOrderNumber -cne $orderNumber) {
            throw "Проверка результата не пройдена: ячейка ${targetColumnName}$targetRow содержит '$savedOrderNumber', ожидалось '$orderNumber'."
        }
        $wbVerify.Close($false)
        $wbVerifyClosed = $true
        Write-UiLog "--- Успешно завершено! ---`r`n"
        Show-UiMessage "Обработка завершена успешно!`r`nФайл сохранён:`r`n$outputPath" "Успех" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null

    } catch {
        Write-UiLog "ОШИБКА: $($_.Exception.Message)`r`n"
        Show-UiMessage "Произошла ошибка:`r`n$($_.Exception.Message)" "Ошибка" ([System.Windows.Forms.MessageBoxButtons]::OK) ([System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } finally {
        if ($wbUpd -and -not $wbUpdClosed) { try { $wbUpd.Close($false) } catch {} }
        if ($wbVerify -and -not $wbVerifyClosed) { try { $wbVerify.Close($false) } catch {} }
        if ($wbOrder -and -not $wbOrderClosed) { try { $wbOrder.Close($false) } catch {} }
        foreach ($object in @($basisMatches) + @($verifyCell, $verifyRange, $verifyCells, $verifySheet, $verifySheets, $wbVerify, $findAfterCell, $foundBase, $changeRange, $updArticleRange, $orderCodeRange, $orderArticleRange, $orderLabelRange, $orderValueRange, $usedUpd, $usedOrder, $updCells, $orderCells, $wsUpd, $wsOrder, $updSheets, $orderSheets, $wbUpd, $wbOrder, $workbooks)) {
            if ($null -ne $object) {
                try { [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($object) | Out-Null } catch {}
            }
        }
        if ($excel) { try { $excel.Quit() } catch {} }
        if ($excel) { try { [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {} }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# --- Интерфейс ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Редактор УПД (Комус)"
$form.Size = New-Object System.Drawing.Size(600, 450)
$form.StartPosition = "CenterScreen"

$btnOrder = New-Object System.Windows.Forms.Button
$btnOrder.Text = "1. Выбрать файл ЗАКАЗА"
$btnOrder.Location = New-Object System.Drawing.Point(20, 20)
$btnOrder.Size = New-Object System.Drawing.Size(200, 30)

$btnUpd = New-Object System.Windows.Forms.Button
$btnUpd.Text = "2. Выбрать файл УПД"
$btnUpd.Location = New-Object System.Drawing.Point(20, 60)
$btnUpd.Size = New-Object System.Drawing.Size(200, 30)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "ЗАПУСТИТЬ ОБРАБОТКУ"
$btnRun.Location = New-Object System.Drawing.Point(350, 20)
$btnRun.Size = New-Object System.Drawing.Size(200, 70)
$btnRun.BackColor = [System.Drawing.Color]::LightGreen

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Location = New-Object System.Drawing.Point(20, 110)
$txtLog.Size = New-Object System.Drawing.Size(540, 280)
$txtLog.ReadOnly = $true

$script:uiEvents = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
$script:workerRunspace = $null
$script:workerPowerShell = $null
$script:workerAsync = $null
$script:workerTimer = New-Object System.Windows.Forms.Timer
$script:workerTimer.Interval = 50
$script:workerTimer.Add_Tick({
    $request = $null
    while ($script:uiEvents.TryDequeue([ref]$request)) {
        if ($request.Type -eq "log") {
            $txtLog.AppendText($request.Text)
        } elseif ($request.Type -eq "dialog") {
            $request.Result = [System.Windows.Forms.MessageBox]::Show(
                $form,
                $request.Text,
                $request.Caption,
                $request.Buttons,
                $request.Icon
            )
            $request.Completion.Set()
        }
        $request = $null
    }

    if ($script:workerAsync -and $script:workerAsync.IsCompleted) {
        try {
            $null = $script:workerPowerShell.EndInvoke($script:workerAsync)
            if ($script:workerPowerShell.HadErrors) {
                $workerErrors = ($script:workerPowerShell.Streams.Error | Out-String)
                [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    $workerErrors,
                    "Ошибка обработки",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                ) | Out-Null
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                $form,
                "Не удалось завершить обработку:`r`n$($_.Exception.Message)",
                "Ошибка",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally {
            $script:workerPowerShell.Dispose()
            $script:workerRunspace.Close()
            $script:workerRunspace.Dispose()
            $script:workerPowerShell = $null
            $script:workerRunspace = $null
            $script:workerAsync = $null
            $btnOrder.Enabled = $true
            $btnUpd.Enabled = $true
            $btnRun.Enabled = $true
            $script:workerTimer.Stop()
        }
    }
})

$form.Add_FormClosing({
    param($sender, $eventArgs)
    if ($script:workerAsync -and -not $script:workerAsync.IsCompleted) {
        $eventArgs.Cancel = $true
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            "Дождитесь завершения обработки перед закрытием окна.",
            "Обработка",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
})

# События кнопок
$btnOrder.Add_Click({
    $fd = New-Object System.Windows.Forms.OpenFileDialog
    $fd.Filter = "Excel Files (*.xlsx)|*.xlsx"
    if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { 
        $global:pathOrder = $fd.FileName 
        $btnOrder.Text = "Заказ: " + $fd.SafeFileName 
    }
})

$btnUpd.Add_Click({
    $fd = New-Object System.Windows.Forms.OpenFileDialog
    $fd.Filter = "Excel Files (*.xlsx)|*.xlsx"
    if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { 
        $global:pathUpd = $fd.FileName 
        $btnUpd.Text = "УПД: " + $fd.SafeFileName 
    }
})

$btnRun.Add_Click({
    $btnOrder.Enabled = $false
    $btnUpd.Enabled = $false
    $btnRun.Enabled = $false

    try {
        $functionNames = @(
            "Get-NormalizedArticle",
            "ConvertTo-ColumnIndex",
            "Get-AppConfig",
            "Write-UiLog",
            "Show-UiMessage",
            "Confirm-Changes",
            "Start-Processing"
        )
        $functionDefinitions = foreach ($functionName in $functionNames) {
            $functionBody = (Get-Command $functionName -CommandType Function).ScriptBlock.ToString()
            "function $functionName {`r`n$functionBody`r`n}"
        }
        $workerScript = @"
param(`$orderPath, `$updPath, `$events, `$configPath)
Add-Type -AssemblyName System.Windows.Forms
`$script:uiEvents = `$events
`$script:configPath = `$configPath
`$global:pathOrder = `$orderPath
`$global:pathUpd = `$updPath
$($functionDefinitions -join "`r`n")
Start-Processing `$null
"@

        $script:workerRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:workerRunspace.ApartmentState = [System.Threading.ApartmentState]::STA
        $script:workerRunspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $script:workerRunspace.Open()
        $script:workerPowerShell = [System.Management.Automation.PowerShell]::Create()
        $script:workerPowerShell.Runspace = $script:workerRunspace
        $null = $script:workerPowerShell.AddScript($workerScript)
        $null = $script:workerPowerShell.AddArgument($global:pathOrder)
        $null = $script:workerPowerShell.AddArgument($global:pathUpd)
        $null = $script:workerPowerShell.AddArgument($script:uiEvents)
        $null = $script:workerPowerShell.AddArgument($script:configPath)
        $script:workerAsync = $script:workerPowerShell.BeginInvoke()
        $script:workerTimer.Start()
    } catch {
        if ($script:workerPowerShell) {
            $script:workerPowerShell.Dispose()
            $script:workerPowerShell = $null
        }
        if ($script:workerRunspace) {
            $script:workerRunspace.Close()
            $script:workerRunspace.Dispose()
            $script:workerRunspace = $null
        }
        $script:workerAsync = $null
        $btnOrder.Enabled = $true
        $btnUpd.Enabled = $true
        $btnRun.Enabled = $true
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            "Не удалось запустить фоновую обработку:`r`n$($_.Exception.Message)",
            "Ошибка",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})

$form.Controls.AddRange(@($btnOrder, $btnUpd, $btnRun, $txtLog))
$form.ShowDialog() | Out-Null