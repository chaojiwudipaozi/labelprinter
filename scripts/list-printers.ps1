# 列出本机已安装打印机及其状态（JSON 输出，UTF-8）
# 性能设计：仅做 3 次批量查询（Win32_Printer / Win32_PrintJob / Get-Printer），
#           避免逐台逐队列查询导致离线打印机把脚本挂死。
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'SilentlyContinue'

# 1) 一次性取回所有 Win32_Printer：离线标记 / 默认打印机 / 错误码
$cimMap = @{}
foreach ($c in @(Get-CimInstance -ClassName Win32_Printer -ErrorAction SilentlyContinue)) {
  if ($c.Name) { $cimMap[[string]$c.Name] = $c }
}

# 2) 一次性取回所有打印任务，按打印机名聚合队列长度（Win32_PrintJob.Name 形如 "打印机名, 作业号"）
$jobMap = @{}
foreach ($j in @(Get-CimInstance -ClassName Win32_PrintJob -ErrorAction SilentlyContinue)) {
  $pn = ([string]$j.Name -split ',')[0].Trim()
  if ($pn) { $jobMap[$pn] = 1 + [int]$jobMap[$pn] }
}

$out = @()
$gp = @(Get-Printer -ErrorAction SilentlyContinue)

if ($gp.Count -gt 0) {
  foreach ($p in $gp) {
    $nm = [string]$p.Name
    $cim = $cimMap[$nm]
    $offline = $false; $isDefault = $false; $errState = 0
    if ($cim) {
      $offline = [bool]$cim.WorkOffline
      $isDefault = [bool]$cim.Default
      $errState = [int]$cim.DetectedErrorState
    }
    $journal = 0
    if ($jobMap.ContainsKey($nm)) { $journal = [int]$jobMap[$nm] }

    $out += [PSCustomObject]@{
      name      = $nm
      driver    = [string]$p.DriverName
      port      = [string]$p.PortName
      status    = [string]$p.PrinterStatus
      jobs      = $journal
      offline   = $offline
      default   = $isDefault
      shared    = [bool]$p.Shared
      errorCode = $errState
      source    = 'Get-Printer'
    }
  }
} else {
  foreach ($c in $cimMap.Values) {
    $nm = [string]$c.Name
    $journal = 0
    if ($jobMap.ContainsKey($nm)) { $journal = [int]$jobMap[$nm] }
    $out += [PSCustomObject]@{
      name      = $nm
      driver    = [string]$c.DriverName
      port      = [string]$c.PortName
      status    = [string]$c.PrinterStatus
      jobs      = $journal
      offline   = [bool]$c.WorkOffline
      default   = [bool]$c.Default
      shared    = [bool]$c.Shared
      errorCode = [int]$c.DetectedErrorState
      source    = 'Win32_Printer'
    }
  }
}

ConvertTo-Json -InputObject @($out) -Depth 3 -Compress
