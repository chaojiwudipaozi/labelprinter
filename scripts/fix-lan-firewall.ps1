<#
  局域网访问修复脚本（需以【管理员】身份运行）
  作用：
    1. 为标签打印工具站放行入站 TCP 端口（默认 11235，仅限本地子网，更安全）
    2. 可选：禁用系统/安全软件自动生成的 node.exe 入站「阻止」规则
       （Windows 防火墙中「阻止」规则优先级高于「允许」规则，且按程序拦截、与端口无关，
         不禁用则换任何端口都无法访问）
  用法：
    powershell -ExecutionPolicy Bypass -File fix-lan-firewall.ps1 -UnblockNode
    powershell -ExecutionPolicy Bypass -File fix-lan-firewall.ps1 -Port 9000 -AnyRemote
    powershell -ExecutionPolicy Bypass -File fix-lan-firewall.ps1 -Remove
#>
param(
  [int]$Port = 11235,
  [switch]$UnblockNode,   # 同时禁用 node.exe 的入站阻止规则
  [switch]$AnyRemote,     # 允许来自任意远程地址（默认仅本地子网）
  [switch]$Remove         # 反向操作：删除本脚本添加的放行规则
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'

$isAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
  Write-Host '[X] 需要管理员权限。请右键 PowerShell -> 以管理员身份运行，然后重新执行本脚本。' -ForegroundColor Red
  exit 1
}

$ruleName = "标签打印工具站 入站 TCP $Port"

if ($Remove) {
  Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue | Remove-NetFirewallRule
  Write-Host "[√] 已删除放行规则：$ruleName" -ForegroundColor Yellow
  exit 0
}

Write-Host '===== 1) 当前网络类别 =====' -ForegroundColor Cyan
Get-NetConnectionProfile | Select-Object InterfaceAlias, NetworkCategory | Format-Table -AutoSize

Write-Host '===== 2) 放行入站 TCP 端口 =====' -ForegroundColor Cyan
Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue | Remove-NetFirewallRule
$remote = if ($AnyRemote) { 'Any' } else { 'LocalSubnet' }
New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow `
  -Protocol TCP -LocalPort $Port -Profile Any -RemoteAddress $remote | Out-Null
Write-Host "[√] 已添加放行规则：$ruleName（远程地址范围：$remote）" -ForegroundColor Green

Write-Host '===== 3) 检查 node.exe 入站阻止规则 =====' -ForegroundColor Cyan
$blockRules = @(Get-NetFirewallRule -Direction Inbound -Enabled True -ErrorAction SilentlyContinue | Where-Object {
  $_.Action -eq 'Block' -and (($_ | Get-NetFirewallApplicationFilter).Program -like '*node.exe')
})
if ($blockRules.Count -eq 0) {
  Write-Host '[√] 未发现 node.exe 的入站阻止规则' -ForegroundColor Green
} else {
  Write-Host ("[!] 发现 $($blockRules.Count) 条 node.exe 阻止规则（Windows 中阻止优先于允许，会导致放行无效）：")
  $blockRules | Select-Object DisplayName, Profile, Enabled | Format-Table -AutoSize
  if ($UnblockNode) {
    $blockRules | Disable-NetFirewallRule
    Write-Host "[√] 已禁用上述阻止规则（可用 Enable-NetFirewallRule 恢复）" -ForegroundColor Green
  } else {
    Write-Host '[i] 如需禁用，请追加参数 -UnblockNode 重新执行' -ForegroundColor Yellow
  }
}

Write-Host '===== 4) 可用访问地址 =====' -ForegroundColor Cyan
$urls = Get-NetIPAddress -AddressFamily IPv4 |
  Where-Object { $_.IPAddress -notlike '127.*' -and $_.PrefixOrigin -ne 'WellKnown' } |
  ForEach-Object { "http://$($_.IPAddress):$Port" }
$urls | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
Write-Host '  提示：局域网其他主机请使用上面地址访问（不要用 localhost）' -ForegroundColor Gray
