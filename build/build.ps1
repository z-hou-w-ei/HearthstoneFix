# 打包脚本：把项目编译成单文件 HearthstoneFix.exe
# 依赖：Windows 自带的 .NET Framework 4.x csc.exe（无需安装任何东西）
# 用法：powershell -ExecutionPolicy Bypass -File build.ps1

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $PSCommandPath
$root = Split-Path -Parent $here

$csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $csc)) { $csc = 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe' }
if (-not (Test-Path -LiteralPath $csc)) { throw '找不到 csc.exe（需要 .NET Framework 4.x）' }

$out = Join-Path $here 'HearthstoneFix.exe'
if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }

$srcPs1 = Join-Path $root 'HearthstoneFix.ps1'
$srcBat = Join-Path $root '一键修复.bat'
$srcIni = Join-Path $root 'config.ini'
$srcMd  = Join-Path $root 'README.md'
$srcIco = Join-Path $root 'icon\HearthstoneFix.ico'

$cscArgs = @(
    '/nologo', '/target:exe', '/codepage:65001', '/optimize',
    ('/win32icon:' + $srcIco),
    ('/resource:' + $srcPs1 + ',res_ps1'),
    ('/resource:' + $srcBat + ',res_bat'),
    ('/resource:' + $srcIni + ',res_ini'),
    ('/resource:' + $srcMd + ',res_md'),
    ('/resource:' + $srcIco + ',res_ico'),
    ('/out:' + $out),
    (Join-Path $here 'launcher.cs')
)

& $csc $cscArgs
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $out)) { throw '编译失败' }

$size = (Get-Item -LiteralPath $out).Length
Write-Host ('生成成功: ' + $out)
Write-Host ('大小: ' + [math]::Round($size / 1KB, 1) + ' KB')
