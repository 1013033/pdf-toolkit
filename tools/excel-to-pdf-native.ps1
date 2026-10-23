param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PassedArgs
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.Windows.Forms

# 判斷傳入參數（拖曳或命令列傳入檔案）
$targetFiles = @()
if ($PassedArgs -and $PassedArgs.Count -gt 0) {
    $targetFiles = @($PassedArgs)
}

# 若無傳入參數（使用者雙擊執行 .bat），彈出原生多選檔案視窗
if ($targetFiles.Count -eq 0) {
    Write-Host "提示：未直接傳入檔案，正在開啟檔案選取對話框..." -ForegroundColor Cyan
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Multiselect = $true
    $openDialog.Filter = "Excel 試算表 (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|所有檔案 (*.*)|*.*"
    $openDialog.Title = "請選取要轉換為 PDF 的 Excel 檔案 (可按住 Ctrl 多選)"

    $result = $openDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $targetFiles = @($openDialog.FileNames)
    } else {
        Write-Host "使用者已取消選取檔案。" -ForegroundColor Yellow
        exit 0
    }
}

Write-Host "已選取 $($targetFiles.Count) 個檔案，正在初始化本機 Microsoft Excel 引擎..." -ForegroundColor Yellow

$excelApp = $null
$successCount = 0
$failCount = 0

try {
    # 建立本機 Excel COM Automation 物件
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
} catch {
    Write-Host "錯誤：無法啟動本機 Microsoft Excel 應用程式。" -ForegroundColor Red
    Write-Host "請確認您的 Windows 電腦已正確安裝 Microsoft Office Excel。" -ForegroundColor Red
    Write-Host "例外資訊: $($_.Exception.Message)" -ForegroundColor DarkRed
    exit 1
}

foreach ($filePath in $targetFiles) {
    if (-not (Test-Path -LiteralPath $filePath)) {
        Write-Host "略過不存在的檔案: $filePath" -ForegroundColor DarkGray
        continue
    }

    $fileItem = Get-Item -LiteralPath $filePath
    $ext = $fileItem.Extension.ToLower()

    if ($ext -ne ".xlsx" -and $ext -ne ".xls" -and $ext -ne ".csv") {
        Write-Host "略過不支援的檔案副檔名: $($fileItem.Name)" -ForegroundColor DarkGray
        continue
    }

    $outPdfPath = [System.IO.Path]::ChangeExtension($fileItem.FullName, ".pdf")
    Write-Host "正在轉換: $($fileItem.Name) -> $($fileItem.BaseName).pdf ..." -NoNewline

    $workbook = $null
    try {
        # 唯讀模式開啟試算表，避免鎖定檔案或更新外部連結
        $workbook = $excelApp.Workbooks.Open($fileItem.FullName, 0, $true)

        # 0 代表 xlTypePDF (ExportAsFixedFormat)
        $workbook.ExportAsFixedFormat(0, $outPdfPath)

        Write-Host " [成功]" -ForegroundColor Green
        $successCount++
    } catch {
        Write-Host " [失敗]" -ForegroundColor Red
        Write-Host "  -> 錯誤細節: $($_.Exception.Message)" -ForegroundColor DarkRed
        $failCount++
    } finally {
        if ($workbook -ne $null) {
            try {
                $workbook.Close($false)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
            } catch {}
            $workbook = $null
        }
    }
}

if ($excelApp -ne $null) {
    try {
        $excelApp.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
    } catch {}
    $excelApp = $null
}

[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "======================================================================"
Write-Host "轉檔統計：成功 $successCount 個檔案，失敗 $failCount 個檔案" -ForegroundColor Cyan
Write-Host "產出的 PDF 檔案已儲存於原檔案相同之目錄中。" -ForegroundColor Cyan
Write-Host "======================================================================"
