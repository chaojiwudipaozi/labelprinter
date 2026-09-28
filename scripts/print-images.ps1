# 静默打印图片到指定打印机（不弹出任何对话框）
# 用法: powershell -File print-images.ps1 -Printer "打印机名" -List "文件列表.txt" -WidthMm 100 -HeightMm 60
#   文件列表.txt：每行一个待打印图片的绝对路径（按打印顺序，可重复同一文件实现多份）
param(
  [Parameter(Mandatory = $true)][string]$Printer,
  [Parameter(Mandatory = $true)][string]$List,
  [double]$WidthMm = 100,
  [double]$HeightMm = 60
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

if (-not (Test-Path $List)) { Write-Output '{"ok":false,"message":"打印列表文件不存在"}'; exit 3 }
$files = @(Get-Content -LiteralPath $List -Encoding UTF8 | Where-Object { $_ -and (Test-Path $_) })
if ($files.Count -eq 0) { Write-Output '{"ok":false,"message":"没有可打印的图片"}'; exit 3 }

$doc = New-Object System.Drawing.Printing.PrintDocument
$doc.PrinterSettings.PrinterName = $Printer
if (-not $doc.PrinterSettings.IsValid) {
  Write-Output ('{"ok":false,"message":"打印机不可用或未安装: ' + $Printer.Replace('"', '') + '"}')
  exit 2
}

# 纸张尺寸按标签实际尺寸设置（单位 1/100 英寸）
$w100 = [int][Math]::Round($WidthMm * 100 / 25.4)
$h100 = [int][Math]::Round($HeightMm * 100 / 25.4)
try {
  $paper = New-Object System.Drawing.Printing.PaperSize('LabelSize', $w100, $h100)
  $doc.DefaultPageSettings.PaperSize = $paper
  $doc.DefaultPageSettings.Margins = New-Object System.Drawing.Printing.Margins(0, 0, 0, 0)
} catch { }
$doc.OriginAtMargins = $false
# 关键：使用 StandardPrintController 实现静默打印，不弹任何对话框
$doc.PrintController = New-Object System.Drawing.Printing.StandardPrintController

$script:fileList = $files
$script:idx = 0
$script:failed = 0

$handler = {
  param($sender, $e)
  if ($script:idx -ge $script:fileList.Count) { $e.HasMorePages = $false; return }
  $img = $null
  try {
    $img = [System.Drawing.Image]::FromFile($script:fileList[$script:idx])
    $rect = New-Object System.Drawing.Rectangle 0, 0, $e.PageBounds.Width, $e.PageBounds.Height
    $e.Graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $e.Graphics.DrawImage($img, $rect)
  } catch {
    $script:failed = $script:failed + 1
  } finally {
    if ($img) { $img.Dispose() }
  }
  $script:idx = $script:idx + 1
  $e.HasMorePages = ($script:idx -lt $script:fileList.Count)
}
$doc.add_PrintPage($handler)

$err = ''
try {
  $doc.Print()
} catch {
  $err = $_.Exception.Message
} finally {
  $doc.Dispose()
}

if ($err) {
  Write-Output ('{"ok":false,"message":"' + $err.Replace('\', '\\').Replace('"', '\"') + '"}')
  exit 4
}
Write-Output ('{"ok":true,"printed":' + $script:idx + ',"failed":' + $script:failed + ',"printer":"' + $Printer.Replace('"', '') + '"}')
