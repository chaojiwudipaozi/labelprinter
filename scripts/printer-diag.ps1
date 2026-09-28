<#
  打印机诊断（只读）：用于排查「某台打印机在工具里看不到」的原因
  输出 JSON：
    drivers      已安装的打印机驱动
    ports        打印机端口（含被哪个队列使用）
    unusedTcpPorts  已创建但没有任何队列使用的 TCP/IP 端口（通常是队列被删/未创建）
    devices      系统里 class=Printer/PrintQueue 的即插即用设备
    orphanDevices 已连接（有硬件设备）但没有对应打印机队列的设备
  用法: powershell -File printer-diag.ps1
#>
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'SilentlyContinue'

# 1) 打印队列
$queues = @(Get-Printer -ErrorAction SilentlyContinue)
$queueNames = @($queues | ForEach-Object { [string]$_.Name })

# 2) 驱动
$drivers = @(Get-PrinterDriver -ErrorAction SilentlyContinue | ForEach-Object {
  [PSCustomObject]@{
    name        = [string]$_.Name
    manufacturer = [string]$_.Manufacturer
    environment = [string]$_.PrinterEnvironment
    zebra       = ([string]$_.Manufacturer -match 'Zebra') -or ([string]$_.Name -match 'Zebra|ZDesigner|ZPL')
  }
})

# 3) 端口 + 使用情况
$usedMap = @{}
foreach ($q in $queues) {
  $p = [string]$q.PortName
  if (-not $usedMap.ContainsKey($p)) { $usedMap[$p] = @() }
  $usedMap[$p] += [string]$q.Name
}
$ports = @(Get-PrinterPort -ErrorAction SilentlyContinue | ForEach-Object {
  $n = [string]$_.Name
  $host_ = [string]$_.PrinterHostAddress
  $used = @()
  if ($usedMap.ContainsKey($n)) { $used = $usedMap[$n] }
  [PSCustomObject]@{
    name        = $n
    description = [string]$_.Description
    host        = $host_
    usedBy      = $used
    isTcpIp     = ([string]$_.Description -match 'TCP/IP') -or ($host_ -ne '')
  }
})

# 4) 即插即用设备（打印机类）
$devices = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
  $_.Class -eq 'Printer' -or $_.Class -eq 'PrintQueue' -or $_.InstanceId -match 'USBPRINT|PRINTENUM'
} | ForEach-Object {
  [PSCustomObject]@{
    friendlyName = [string]$_.FriendlyName
    className    = [string]$_.Class
    status       = [string]$_.Status
    instanceId   = [string]$_.InstanceId
  }
})

# 5) 已连接硬件但没有打印队列的设备（关键诊断项）
#    注意：USB 打印机的 PnP 名称常等于"驱动名"（如 Epson ESC/P-R V4 Class Driver），
#    这类设备其实已有队列，需排除，避免误报
$driverNames = @($drivers | ForEach-Object { [string]$_.name })
$orphans = @($devices | Where-Object {
  $_.className -eq 'Printer' -and $_.friendlyName `
    -and ($queueNames -notcontains $_.friendlyName) `
    -and ($driverNames -notcontains $_.friendlyName)
} | ForEach-Object {
  $extra = @{}
  if ($_.instanceId -match 'VID_([0-9A-F]{4})') { $extra.vid = $Matches[1] }
  [PSCustomObject]@{ friendlyName = $_.friendlyName; instanceId = $_.instanceId; status = $_.status; vid = $extra.vid }
})

# 6) 已安装但没有被任何打印机使用的驱动（若含 Zebra，说明"驱动已就绪，只差添加打印机"）
$queueDrivers = @($queues | ForEach-Object { [string]$_.DriverName })
$unusedDrivers = @($drivers | Where-Object { $queueDrivers -notcontains [string]$_.name } | ForEach-Object {
  [PSCustomObject]@{ name = [string]$_.name; manufacturer = [string]$_.manufacturer; zebra = $_.zebra }
})

$payload = [PSCustomObject]@{
  queues         = $queueNames
  drivers        = $drivers
  zebraDrivers   = @($drivers | Where-Object { $_.zebra } | ForEach-Object { $_.name })
  ports          = $ports
  unusedTcpPorts = @($ports | Where-Object { $_.isTcpIp -and $_.usedBy.Count -eq 0 })
  devices        = $devices
  orphanDevices  = $orphans
  unusedDrivers  = $unusedDrivers
}
ConvertTo-Json -InputObject $payload -Depth 5 -Compress
