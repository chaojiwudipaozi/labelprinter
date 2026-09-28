<#
  添加打印机（USB 已识别 / 网络打印机）到 Windows，使其能被标签打印工具站识别
  需以【管理员】身份运行。

  常用用法：
    # 1) 查看本机已安装的打印驱动（挑一个用于 -Driver）
    powershell -ExecutionPolicy Bypass -File add-printer.ps1 -ListDrivers
    powershell -ExecutionPolicy Bypass -File add-printer.ps1 -ListDrivers -Filter Zebra

    # 2) 添加网络打印机（自动创建标准 TCP/IP 端口，默认 9100）
    powershell -ExecutionPolicy Bypass -File add-printer.ps1 -Name "Zebra GX430t" -Ip 192.168.1.27 -Driver "ZDesigner GX430t (ZPL)"

    # 3) 添加 USB 打印机（先用 -ListPorts 看 USB 端口，或直接指定 -Port USB001）
    powershell -ExecutionPolicy Bypass -File add-printer.ps1 -Name "Zebra GX430t" -Port USB010 -Driver "ZDesigner GX430t (ZPL)"

    # 4) 删除本脚本添加的打印机
    powershell -ExecutionPolicy Bypass -File add-printer.ps1 -Name "Zebra GX430t" -Remove
#>
param(
  [string]$Name,                       # 打印机显示名称
  [string]$Ip,                         # 网络打印机 IP（与 -Port 二选一）
  [int]$RawPort = 9100,                # TCP/IP 端口号，默认 9100(RAW)
  [string]$Port,                       # 直接指定端口名（如 USB001 / 192.168.1.27）
  [string]$Driver,                     # 驱动名（须为本机已安装驱动）
  [string]$Filter,                     # -ListDrivers 时的关键字过滤，如 Zebra
  [switch]$ListDrivers,
  [switch]$ListPorts,
  [switch]$Remove
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'

$isAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($ListDrivers -or $ListPorts) {
  if ($ListPorts) {
    Write-Host '===== 本机打印机端口 =====' -ForegroundColor Cyan
    Get-PrinterPort | Select-Object Name, Description, PrinterHostAddress | Format-Table -AutoSize
  }
  if ($ListDrivers) {
    $f = $Filter
    Write-Host '===== 本机已安装打印驱动 =====' -ForegroundColor Cyan
    $list = Get-PrinterDriver | Sort-Object Manufacturer, Name
    if ($f) { $list = $list | Where-Object { $_.Name -match $f -or $_.Manufacturer -match $f } }
    $list | Select-Object Name, Manufacturer | Format-Table -AutoSize
    Write-Host '提示：标签机驱动通常名为 ZDesigner <型号> (ZPL)；若列表中没有目标型号，请先到厂商官网下载安装驱动。' -ForegroundColor Yellow
  }
  exit 0
}

if (-not $isAdmin) {
  Write-Host '[X] 需要管理员权限。请右键 PowerShell -> 以管理员身份运行，然后重新执行。' -ForegroundColor Red
  exit 1
}

if (-not $Name) { Write-Host '[X] 请用 -Name 指定打印机名称。' -ForegroundColor Red; exit 1 }

# 删除模式
if ($Remove) {
  if (Get-Printer -Name $Name -ErrorAction SilentlyContinue) {
    Remove-Printer -Name $Name
    Write-Host "[√] 已删除打印机：$Name" -ForegroundColor Yellow
  } else {
    Write-Host "[i] 未找到打印机：$Name" -ForegroundColor Yellow
  }
  exit 0
}

# 端口：优先使用 -Port，否则根据 -Ip 创建/复用 TCP/IP 端口
if (-not $Port) {
  if (-not $Ip) { Write-Host '[X] 请用 -Ip 指定打印机 IP，或用 -Port 指定端口名。' -ForegroundColor Red; exit 1 }
  $Port = $Ip
  if (-not (Get-PrinterPort -Name $Port -ErrorAction SilentlyContinue)) {
    Write-Host "[i] 正在创建标准 TCP/IP 端口：$Port (RAW:$RawPort)" -ForegroundColor Cyan
    Add-PrinterPort -Name $Port -PrinterHostAddress $Ip -PortNumber $RawPort
  } else {
    Write-Host "[i] 端口已存在，直接复用：$Port" -ForegroundColor Cyan
  }
} elseif (-not (Get-PrinterPort -Name $Port -ErrorAction SilentlyContinue)) {
  Write-Host "[X] 端口不存在：$Port（可用 -ListPorts 查看现有端口）" -ForegroundColor Red
  exit 1
}

# 驱动校验
if (-not $Driver) { Write-Host '[X] 请用 -Driver 指定驱动名（可用 -ListDrivers 查看）' -ForegroundColor Red; exit 1 }
$drv = Get-PrinterDriver -Name $Driver -ErrorAction SilentlyContinue
if (-not $drv) {
  Write-Host "[X] 本机未安装该驱动：$Driver" -ForegroundColor Red
  Write-Host '    可用驱动（标签机相关）：' -ForegroundColor Yellow
  Get-PrinterDriver | Where-Object { $_.Name -match 'Zebra|ZDesigner|ZPL|TSC|Godex|标签' } |
    Select-Object Name, Manufacturer | Format-Table -AutoSize
  Write-Host '    请先到厂商官网下载安装对应驱动后重试（斑马：ZDesigner 驱动）。' -ForegroundColor Yellow
  exit 2
}

# 已存在同名打印机则先移除，避免冲突
if (Get-Printer -Name $Name -ErrorAction SilentlyContinue) {
  Write-Host "[i] 已存在同名打印机，先删除：$Name" -ForegroundColor Cyan
  Remove-Printer -Name $Name
}

Write-Host "[i] 正在添加打印机：$Name  驱动=$Driver  端口=$Port" -ForegroundColor Cyan
Add-Printer -Name $Name -DriverName $Driver -PortName $Port

$p = Get-Printer -Name $Name
Write-Host "[√] 添加成功：$($p.Name)  驱动=$($p.DriverName)  端口=$($p.PortName)  状态=$($p.PrinterStatus)" -ForegroundColor Green
$cim = Get-CimInstance Win32_Printer -Filter ("Name='" + $Name.Replace("'", "''") + "'")
if ($cim) {
  Write-Host ("    脱机(WorkOffline)=" + $cim.WorkOffline + "   默认打印机=" + $cim.Default) -ForegroundColor Gray
  if ($cim.WorkOffline) {
    Write-Host '    提示：若设备已连接却显示脱机，请在“设备和打印机”中取消勾选「脱机使用打印机」。' -ForegroundColor Yellow
  }
}
Write-Host '现在刷新标签打印工具站页面，即可在打印机下拉中选择该打印机。' -ForegroundColor Green
