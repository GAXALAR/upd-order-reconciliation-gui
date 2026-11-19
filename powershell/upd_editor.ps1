Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Глобальные переменные для путей
$global:pathOrder = ""
$global:pathUpd = ""

function Start-Processing {
    param($logOutput)

    if (-not $global:pathOrder -or -not $global:pathUpd) {
        [System.Windows.Forms.MessageBox]::Show("Выберите оба файла перед началом!", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    if (-not (Test-Path -LiteralPath $global:pathOrder -PathType Leaf) -or
        -not (Test-Path -LiteralPath $global:pathUpd -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show("Один из входных файлов не существует.", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if (([IO.Path]::GetFullPath($global:pathOrder)) -ieq ([IO.Path]::GetFullPath($global:pathUpd))) {
        [System.Windows.Forms.MessageBox]::Show("Файл заказа и файл УПД должны быть разными.", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ([IO.Path]::GetExtension($global:pathOrder) -ine ".xlsx" -or [IO.Path]::GetExtension($global:pathUpd) -ine ".xlsx") {
        [System.Windows.Forms.MessageBox]::Show("Поддерживаются только файлы формата .xlsx.", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $excel = $null
    $wbOrder = $null
    $wsOrder = $null
    $wbUpd = $null
    $wsUpd = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Не удалось запустить MS Excel. Убедитесь, что Excel установлен на компьютере.", "Ошибка COM-объекта")
        return
    }

    try {
        $logOutput.AppendText("--- Старт обработки: $(Get-Date -Format 'HH:mm:ss') ---`r`n")
        
        # Загрузка заказа
        $wbOrder = $excel.Workbooks.Open($global:pathOrder)
        $wsOrder = $wbOrder.Sheets.Item(1)
        $artMap = @{}
        $duplicates = @{}
        $orderNumber = ""

        $usedOrder = $wsOrder.UsedRange
        $firstOrderRow = $usedOrder.Row
        $lastOrderRow = $usedOrder.Row + $usedOrder.Rows.Count - 1
        for ($i = $firstOrderRow; $i -le $lastOrderRow; $i++) {
            $cellB = [string]$wsOrder.Cells.Item($i, 2).Text
            $cellC = [string]$wsOrder.Cells.Item($i, 3).Text
            if ($cellB.Trim() -and $cellC.Trim()) {
                $article = $cellC.Trim()
                $internalCode = $cellB.Trim()
                if ($artMap.ContainsKey($article) -and $artMap[$article] -ne $internalCode) {
                    $duplicates[$article] = $true
                    $artMap.Remove($article)
                } elseif (-not $duplicates.ContainsKey($article)) {
                    $artMap[$article] = $internalCode
                }
            }
            if ($cellB -like "*Заказ №*") { $orderNumber = [string]$wsOrder.Cells.Item($i, 3).Text }
        }
        $wbOrder.Close($false)
        $logOutput.AppendText("Загружено артикулов: $($artMap.Count). Номер заказа: $orderNumber`r`n")
        foreach ($article in $duplicates.Keys) {
            $logOutput.AppendText("[ДУБЛИКАТ] Артикул '$article' имеет неоднозначное соответствие.`r`n")
        }

        # Обработка УПД
        $wbUpd = $excel.Workbooks.Open($global:pathUpd)
        $wsUpd = $wbUpd.Sheets.Item(1)

        $usedUpd = $wsUpd.UsedRange
        $firstUpdRow = [Math]::Max(17, $usedUpd.Row)
        $lastUpdRow = $usedUpd.Row + $usedUpd.Rows.Count - 1
        $preview = @()
        for ($j = $firstUpdRow; $j -le $lastUpdRow; $j++) {
            $currentArt = [string]$wsUpd.Cells.Item($j, 2).Text.Trim()
            if (-not $currentArt) { continue }

            if ($artMap.ContainsKey($currentArt)) {
                $preview += [PSCustomObject]@{
                    Row = $j
                    OldArticle = $currentArt
                    NewArticle = $artMap[$currentArt]
                }
            } else {
                $logOutput.AppendText("[НЕ НАЙДЕН] Строка $j`: артикул '$currentArt'`r`n")
            }
        }

        if ($preview.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Не найдено ни одного однозначного артикула для замены.", "Предпросмотр", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $wbUpd.Close($false)
            $wbUpd = $null
            return
        }

        $previewText = ($preview | ForEach-Object { "Строка $($_.Row): $($_.OldArticle) -> $($_.NewArticle)" }) -join "`r`n"
        $confirmation = [System.Windows.Forms.MessageBox]::Show("Будут внесены следующие изменения:`r`n`r`n$previewText`r`n`r`nПродолжить?", "Подтверждение изменений", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
            $logOutput.AppendText("Операция отменена пользователем.`r`n")
            $wbUpd.Close($false)
            $wbUpd = $null
            return
        }

        foreach ($change in $preview) {
            $wsUpd.Cells.Item($change.Row, 2).Value2 = $change.NewArticle
            $logOutput.AppendText("[OK] Строка $($change.Row): заменён артикул '$($change.OldArticle)' на '$($change.NewArticle)'`r`n")
        }

        # Запись номера заказа
        $foundBase = $wsUpd.Cells.Find("Основание передачи")
        if ($foundBase) {
            $wsUpd.Cells.Item($foundBase.Row - 1, 20).Value2 = $orderNumber
            $logOutput.AppendText("Номер заказа вписан в строку $($foundBase.Row - 1).`r`n")
        }

        $directory = [IO.Path]::GetDirectoryName($global:pathUpd)
        $name = [IO.Path]::GetFileNameWithoutExtension($global:pathUpd)
        $outputPath = Join-Path $directory ($name + "_reconciled.xlsx")
        $wbUpd.SaveAs($outputPath, 51)
        $wbUpd.Close($true)
        $wbUpd = $null
        $logOutput.AppendText("--- Успешно завершено! ---`r`n")
        [System.Windows.Forms.MessageBox]::Show("Обработка завершена успешно!`r`nФайл сохранён:`r`n$outputPath", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    } catch {
        $logOutput.AppendText("ОШИБКА: $($_.Exception.Message)`r`n")
        [System.Windows.Forms.MessageBox]::Show("Произошла ошибка:`r`n$($_.Exception.Message)", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        if ($wbUpd) { try { $wbUpd.Close($false) } catch {} }
        if ($wbOrder) { try { $wbOrder.Close($false) } catch {} }
        foreach ($object in @($wsUpd, $wsOrder, $wbUpd, $wbOrder, $excel)) {
            if ($null -ne $object) {
                try { [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($object) | Out-Null } catch {}
            }
        }
        if ($excel) { try { $excel.Quit() } catch {} }
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

# События кнопок
$btnOrder.Add_Click({
    $fd = New-Object System.Windows.Forms.OpenFileDialog
    $fd.Filter = "Excel Files (*.xlsx; *.xls)|*.xlsx;*.xls"
    if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { 
        $global:pathOrder = $fd.FileName 
        $btnOrder.Text = "Заказ: " + $fd.SafeFileName 
    }
})

$btnUpd.Add_Click({
    $fd = New-Object System.Windows.Forms.OpenFileDialog
    $fd.Filter = "Excel Files (*.xlsx; *.xls)|*.xlsx;*.xls"
    if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { 
        $global:pathUpd = $fd.FileName 
        $btnUpd.Text = "УПД: " + $fd.SafeFileName 
    }
})

$btnRun.Add_Click({ Start-Processing $txtLog })

$form.Controls.AddRange(@($btnOrder, $btnUpd, $btnRun, $txtLog))
$form.ShowDialog() | Out-Null