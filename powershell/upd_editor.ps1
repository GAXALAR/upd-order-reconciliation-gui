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
        $orderNumber = ""

        $rowsCount = $wsOrder.UsedRange.Rows.Count
        for ($i = 1; $i -le $rowsCount; $i++) {
            $cellB = [string]$wsOrder.Cells.Item($i, 2).Text
            $cellC = [string]$wsOrder.Cells.Item($i, 3).Text
            if ($cellB.Trim() -and $cellC.Trim()) { $artMap[$cellC.Trim()] = $cellB.Trim() }
            if ($cellB -like "*Заказ №*") { $orderNumber = [string]$wsOrder.Cells.Item($i, 3).Text }
        }
        $wbOrder.Close($false)
        $logOutput.AppendText("Загружено артикулов: $($artMap.Count). Номер заказа: $orderNumber`r`n")

        # Обработка УПД
        $wbUpd = $excel.Workbooks.Open($global:pathUpd)
        $wsUpd = $wbUpd.Sheets.Item(1)

        $updRows = $wsUpd.UsedRange.Rows.Count
        for ($j = 17; $j -le $updRows; $j++) {
            $currentArt = [string]$wsUpd.Cells.Item($j, 2).Text.Trim()
            if (-not $currentArt) { continue }

            if ($artMap.ContainsKey($currentArt)) {
                $wsUpd.Cells.Item($j, 2).Value2 = $artMap[$currentArt]
                $logOutput.AppendText("[OK] Строка $($j): Заменен номер '$currentArt'`r`n")
            }
        }

        # Запись номера заказа
        $foundBase = $wsUpd.Cells.Find("Основание передачи")
        if ($foundBase) {
            $wsUpd.Cells.Item($foundBase.Row - 1, 20).Value2 = $orderNumber
            $logOutput.AppendText("Номер заказа вписан в строку $($foundBase.Row - 1).`r`n")
        }

        $wbUpd.Save()
        $wbUpd.Close($true)
        $logOutput.AppendText("--- Успешно завершено! ---`r`n")
        [System.Windows.Forms.MessageBox]::Show("Обработка завершена успешно!", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    } catch {
        $logOutput.AppendText("ОШИБКА: $($_.Exception.Message)`r`n")
    } finally {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
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