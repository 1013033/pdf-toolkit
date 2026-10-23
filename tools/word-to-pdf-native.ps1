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

# 若無傳入參數（使用者直接執行），彈出原生多選檔案視窗
if ($targetFiles.Count -eq 0) {
    Write-Host "提示：未直接傳入檔案，正在開啟檔案選取對話框..." -ForegroundColor Cyan
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Multiselect = $true
    $openDialog.Filter = "Word 文件 (*.docx;*.doc)|*.docx;*.doc|所有檔案 (*.*)|*.*"
    $openDialog.Title = "請選取要轉換為 PDF 的 Word 文件 (可按住 Ctrl 多選)"

    $result = $openDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $targetFiles = @($openDialog.FileNames)
    } else {
        Write-Host "使用者已取消選取檔案。" -ForegroundColor Yellow
        exit 0
    }
}

Write-Host "已選取 $($targetFiles.Count) 個檔案，正在初始化本機 Microsoft Word 引擎..." -ForegroundColor Yellow

$wordApp = $null
$successCount = 0
$failCount = 0

try {
    # 建立本機 Word COM Automation 物件
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0 # 隱藏警告對話框 (wdAlertsNone)

    # WdSaveFormat.wdFormatPDF = 17
    $wdFormatPDF = 17

    foreach ($filePath in $targetFiles) {
        if (-not (Test-Path -LiteralPath $filePath)) {
            Write-Host "找不到檔案: $filePath" -ForegroundColor Red
            $failCount++
            continue
        }

        $fileItem = Get-Item -LiteralPath $filePath
        if ($fileItem.Extension -notmatch '^\.(docx|doc)$') {
            Write-Host "略過非 Word 檔案: $($fileItem.Name)" -ForegroundColor DarkGray
            continue
        }

        $dirName = $fileItem.DirectoryName
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileItem.Name)
        $outPdfPath = [System.IO.Path]::Combine($dirName, "$baseName.pdf")

        Write-Host "正在轉換: $($fileItem.Name) ..." -ForegroundColor Cyan

        try {
            # 打開 Word 文件 (唯讀模式開啟以保護原檔)
            $doc = $wordApp.Documents.Open($fileItem.FullName, $false, $true, $false)
            
            # 原生匯出 PDF (wdFormatPDF = 17)
            $doc.SaveAs([ref]$outPdfPath, [ref]$wdFormatPDF)
            
            # 關閉文件不儲存修改
            $doc.Close([ref]$false)

            Write-Host " -> [100% 原生向量轉換成功] 已儲存至: $outPdfPath" -ForegroundColor Green
            $successCount++
        } catch {
            Write-Host " -> [轉換失敗]: $($_.Exception.Message)" -ForegroundColor Red
            $failCount++
        }
    }
} catch {
    Write-Host ""
    Write-Host "[錯誤] 無法啟動本機 Microsoft Word COM 元件！" -ForegroundColor Red
    Write-Host "請確認您的電腦是否已安裝 Microsoft Word (Office 2010 或以上版本)。" -ForegroundColor Yellow
    Write-Host "錯誤訊息: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($wordApp -ne $null) {
        $wordApp.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

Write-Host ""
Write-Host "==================== 轉檔作業統計 ====================" -ForegroundColor Cyan
Write-Host " 總計選取: $($targetFiles.Count) 份" -ForegroundColor White
Write-Host " 成功轉換: $successCount 份 (100% Word 原生向量品質)" -ForegroundColor Green
if ($failCount -gt 0) {
    Write-Host " 失敗數量: $failCount 份" -ForegroundColor Red
}
Write-Host "======================================================" -ForegroundColor Cyan
