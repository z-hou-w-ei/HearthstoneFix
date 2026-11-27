<#
  炉石一键修复 + 自动启动（通用版）
  ------------------------------------------------------------------
  流程：
    1. 结束炉石/战网/盒子残留进程
    2. 检测内核级卡死残留（这种进程只能重启清除），提示按键确认后自动重启
    3. 关闭 Windows 游戏叠加层（降低盒子"拔线"后闪退概率）
    4. 清理超期/过大的炉石日志
    5. 重启战网后台服务
    6. 启动战网 -> 自动识别界面上的「进入游戏」按钮并模拟点击 -> 启动炉石
    7. 校验游戏真的起来了（防"闪一下即退出"被误判成功）-> 启动炉石盒子
    8. 全部完成后由 bat 自动关窗

  跨电脑可用性说明：
    - 不写死任何屏幕坐标：按钮位置靠「OCR 识别文字 + 自适应色块 + 颜色兜底」三重方式，
      在任意分辨率、缩放、窗口尺寸下都会重新定位。
    - 路径不写死：优先读同目录 config.ini，其次查注册表与常见安装位置，找不到才报错。
    - 点击前会自动最小化本脚本所在的命令行窗口，避免它挡住战网按钮。

  用法：
    .\一键修复.bat                    完整流程
    .\一键修复.bat -DryRun            只看会做什么
    .\一键修复.bat -NoLaunch          只修复不启动
    .\一键修复.bat -KeepBox           不结束盒子进程
    .\一键修复.bat -Retry             启动失败自动再修一次
    .\一键修复.bat -LaunchTimeout 180 等待游戏窗口秒数（默认 120）
    .\一键修复.bat -TestClick         只测试按钮识别与点击
    .\一键修复.bat -KeepConsole       点击时不最小化命令行窗口（调试用）
    .\一键修复.bat -NoShortcut        跳过“首次运行自动创建桌面快捷方式”
    .\一键修复.bat -DiagOnly          卡死时只采集现场证据，不结束任何进程（分析用）
  说明：第一次真实运行（非 -DryRun）时，会在桌面创建「炉石一键修复」快捷方式，
        图标使用本目录 icon\HearthstoneFix.ico；已存在则不动。
#>
[CmdletBinding()]
param(
    [switch]$DryRun,
    [switch]$KeepBox,
    [switch]$ResetAgent,
    [switch]$RestoreOverlay,
    [switch]$SkipOverlay,
    [switch]$SkipLogs,
    [switch]$NoLaunch,
    [switch]$Retry,
    [switch]$TestReboot,
    [switch]$TestClick,
    [switch]$KeepConsole,
    [switch]$NoShortcut,
    [switch]$DiagOnly,
    [string]$SimKeys = '',
    [int]$KeepDays = 3,
    [int]$LaunchTimeout = 120
)

# 调试用：-SimKeys 预置按键（s=空格、e=回车、x=其它键），仅配合 -TestReboot 验证倒计时逻辑
$script:SimKeysQueue = @()
if ($SimKeys) {
    foreach ($ch in $SimKeys.ToCharArray()) {
        switch ($ch) {
            's' { $script:SimKeysQueue += 32 }
            'e' { $script:SimKeysQueue += 13 }
            'x' { $script:SimKeysQueue += 65 }
        }
    }
}

$ScriptDir = Split-Path -Parent $PSCommandPath
# 诊断模式写独立日志，避免把上一次真跑（或这次真跑）的日志覆盖掉，两份可以对照着看
$script:LogFile = Join-Path $ScriptDir $(if ($DiagOnly) { 'hs-diag-log.txt' } else { 'hs-fix-log.txt' })
try { if (Test-Path -LiteralPath $script:LogFile) { Remove-Item -LiteralPath $script:LogFile -Force -ErrorAction Stop } } catch {}
function Log($m) {
    Write-Host $m
    try { Add-Content -LiteralPath $script:LogFile -Value $m -Encoding UTF8 -ErrorAction Stop }
    catch {
        if (-not $script:LogWarned) { $script:LogWarned = $true; Write-Host ('  [日志写入失败] ' + $_.Exception.Message) }
    }
}
function Finish([int]$Code) { exit $Code }

# 首次运行时把快捷方式建到桌面（图标用项目自带的 icon\HearthstoneFix.ico）
function Ensure-Shortcut {
    $lnkName = '炉石一键修复.lnk'
    $desktop = [Environment]::GetFolderPath('Desktop')
    if (-not $desktop) { Log '  [快捷方式] 取不到桌面路径，跳过。'; return $false }
    $lnkPath = Join-Path $desktop $lnkName
    $batPath = Join-Path $ScriptDir '一键修复.bat'
    $icoPath = Join-Path $ScriptDir 'icon\HearthstoneFix.ico'
    if (Test-Path -LiteralPath $lnkPath) {
        # 已存在：如果指向的正是本目录的 bat，就不动；指向别处（比如旧的解压目录）则改指到当前目录
        try {
            $ex = (New-Object -ComObject WScript.Shell).CreateShortcut($lnkPath)
            $tgt = $ex.TargetPath
            if ($tgt) {
                $tgtDir = Split-Path -Parent $tgt
                if ($tgtDir -eq $ScriptDir -and (Test-Path -LiteralPath $tgt)) { return $true }
                Log ('  [快捷方式] 已有快捷方式指向 ' + $tgt + '，改为指向当前目录')
            }
        } catch {}
    }
    if (-not (Test-Path -LiteralPath $batPath)) { Log '  [快捷方式] 找不到 一键修复.bat，跳过。'; return $false }
    try {
        $wsh = New-Object -ComObject WScript.Shell
        $lnk = $wsh.CreateShortcut($lnkPath)
        $lnk.TargetPath = $batPath
        $lnk.WorkingDirectory = $ScriptDir
        $lnk.Description = '炉石一键修复：清理卡死残留并自动启动战网、炉石与盒子'
        if (Test-Path -LiteralPath $icoPath) { $lnk.IconLocation = $icoPath + ',0' }
        $lnk.WindowStyle = 1
        $lnk.Save()
        Log ('  [快捷方式] 已在桌面创建：' + $lnkPath)
        return $true
    } catch {
        Log ('  [快捷方式] 创建失败：' + $_.Exception.Message)
        return $false
    }
}

function Human($bytes) {
    if ($bytes -ge 1GB) { return ('{0:N2} GB' -f ($bytes / 1GB)) }
    if ($bytes -ge 1MB) { return ('{0:N1} MB' -f ($bytes / 1MB)) }
    return ('{0:N0} KB' -f ($bytes / 1KB))
}

# ================= 路径探测（换电脑不用改脚本） =================
function Test-ExePath([string]$p) {
    if (-not $p) { return $null }
    $p = $p.Trim().Trim('"')
    if (Test-Path -LiteralPath $p -PathType Leaf) { return $p }
    return $null
}
function Get-UninstallDirs([string[]]$names) {
    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $res = @()
    foreach ($r in $roots) {
        foreach ($k in @(Get-ItemProperty -Path $r -ErrorAction SilentlyContinue)) {
            if (-not $k.DisplayName) { continue }
            foreach ($n in $names) {
                if ($k.DisplayName -like "*$n*" -and $k.InstallLocation) { $res += [string]$k.InstallLocation; break }
            }
        }
    }
    return $res
}
function Get-DriveRoots {
    return @(Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
             Where-Object { $_.Root -match '^[A-Za-z]:\\$' } | ForEach-Object { $_.Root })
}
function Find-ExeUnder([string]$dir, [string]$exeName, [int]$depth = 2) {
    if (-not $dir) { return $null }
    if (-not (Test-Path -LiteralPath $dir)) { return $null }
    $direct = Join-Path $dir $exeName
    if (Test-Path -LiteralPath $direct -PathType Leaf) { return $direct }
    try {
        $f = Get-ChildItem -LiteralPath $dir -Filter $exeName -File -Recurse -Depth $depth -ErrorAction SilentlyContinue |
             Select-Object -First 1
        if ($f) { return $f.FullName }
    } catch {}
    return $null
}
function Find-ExeByShortcut([string]$pattern, [string]$exclude = '') {
    try {
        $sh = New-Object -ComObject WScript.Shell
        $roots = @("$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
                   "$env:APPDATA\Microsoft\Windows\Start Menu\Programs",
                   "$env:USERPROFILE\Desktop", "$env:PUBLIC\Desktop")
        foreach ($rt in $roots) {
            if (-not (Test-Path -LiteralPath $rt)) { continue }
            foreach ($f in @(Get-ChildItem -LiteralPath $rt -Filter '*.lnk' -Recurse -ErrorAction SilentlyContinue)) {
                if ($f.Name -notmatch $pattern) { continue }
                if ($exclude -and $f.Name -match $exclude) { continue }
                $t = $null
                try { $t = $sh.CreateShortcut($f.FullName).TargetPath } catch {}
                if (Test-ExePath $t) { return $t }
            }
        }
    } catch {}
    return $null
}
function Find-ExeByProcess([string]$procName) {
    $p = Get-Process -Name $procName -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($p) { try { return (Test-ExePath $p.Path) } catch {} }
    return $null
}
function Find-ExeByDiskSearch([string]$exeName, [int]$seconds = 30, [int]$depth = 4) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    foreach ($d in @(Get-DriveRoots)) {
        if ($sw.Elapsed.TotalSeconds -gt $seconds) { break }
        try {
            $hit = Get-ChildItem -LiteralPath $d -Filter $exeName -File -Recurse -Depth $depth -ErrorAction SilentlyContinue |
                   Select-Object -First 1
            if ($hit) { return $hit.FullName }
        } catch {}
    }
    return $null
}
function Resolve-Paths {
    $cfg = @{
        BattleNetExe = $null
        GameUid      = 'WTCG'
        GameUidAlt   = 'hs_beta'
        GameExe      = $null
        BoxExe       = $null
        GameLauncherExe = $null
        LogRoot      = $null
    }
    # 1) 同目录 config.ini 优先
    $ini = Join-Path $ScriptDir 'config.ini'
    if (Test-Path -LiteralPath $ini) {
        foreach ($line in @(Get-Content -LiteralPath $ini -Encoding UTF8 -ErrorAction SilentlyContinue)) {
            if ($line -match '^\s*([A-Za-z][A-Za-z0-9]*)\s*=\s*(.+?)\s*$') {
                $v = $Matches[2].Trim().Trim('"')
                if ($v) { $cfg[$Matches[1]] = $v }
            }
        }
        Log ('  已读取配置: ' + $ini)
    }
    # 2) 战网（按可靠性依次尝试：配置文件 / 运行进程 / 快捷方式 / 注册表 / 常见目录 / 浅层搜索）
    if (-not (Test-ExePath $cfg.BattleNetExe)) {
        $bn = $null
        # 2.1 战网自己的 Battle.net.config 里记录了安装目录
        $bcfg = Join-Path $env:APPDATA 'Battle.net\Battle.net.config'
        if (Test-Path -LiteralPath $bcfg) {
            $raw = Get-Content -LiteralPath $bcfg -Raw -ErrorAction SilentlyContinue
            if ($raw -and $raw -match '"Path"\s*:\s*"((?:[^"\\]|\\.)*)"') {
                $ip = $Matches[1] -replace '\\\\', '\'
                $bn = Find-ExeUnder $ip 'Battle.net.exe' 2
            }
        }
        # 2.2 正在运行的战网进程的真实路径
        if (-not $bn) {
            $rp = Get-Process -Name 'Battle.net' -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($rp) { try { $bn = Test-ExePath $rp.Path } catch {} }
        }
        # 2.3 开始菜单 / 桌面快捷方式
        if (-not $bn) { $bn = Find-ExeByShortcut 'Battle|战网' }
        # 2.4 卸载项 / 注册表安装路径
        if (-not $bn) {
            foreach ($d in @(Get-UninstallDirs @('Battle.net'))) { if (-not $bn) { $bn = Find-ExeUnder $d 'Battle.net.exe' 3 } }
        }
        if (-not $bn) {
            foreach ($rk in @('HKLM:\SOFTWARE\WOW6432Node\Blizzard Entertainment\Battle.net',
                              'HKLM:\SOFTWARE\Blizzard Entertainment\Battle.net')) {
                $ip2 = (Get-ItemProperty -Path $rk -Name 'InstallPath' -ErrorAction SilentlyContinue).InstallPath
                if ($ip2) { $bn = Find-ExeUnder $ip2 'Battle.net.exe' 2 }
            }
        }
        # 2.5 各盘常见位置
        if (-not $bn) {
            foreach ($d in @(Get-DriveRoots)) {
                foreach ($rel in @('Battle.net\Battle.net.exe', 'Program Files\Battle.net\Battle.net.exe',
                                   'Program Files (x86)\Battle.net\Battle.net.exe', 'Games\Battle.net\Battle.net.exe',
                                   'Blizzard\Battle.net\Battle.net.exe', 'battle.net\Battle.net.exe')) {
                    if (-not $bn) { $bn = Test-ExePath (Join-Path $d $rel) }
                }
            }
        }
        # 2.6 最后手段：浅层全盘搜索，限时 30 秒
        if (-not $bn) { $bn = Find-ExeByDiskSearch 'Battle.net.exe' 30 4 }
        $cfg.BattleNetExe = $bn
    }
    # 3) 炉石主程序
    if (-not (Test-ExePath $cfg.GameExe)) {
        $hs = $null
        foreach ($d in @(Get-UninstallDirs @('Hearthstone', '炉石传说')) ) { if (-not $hs) { $hs = Find-ExeUnder $d 'Hearthstone.exe' 2 } }
        if (-not $hs) {
            $rk = 'HKLM:\SOFTWARE\WOW6432Node\Blizzard Entertainment\Hearthstone'
            $ip = (Get-ItemProperty -Path $rk -Name 'InstallLocation' -ErrorAction SilentlyContinue).InstallLocation
            if ($ip) { $hs = Find-ExeUnder $ip 'Hearthstone.exe' 1 }
        }
        # 3.25 战网目录内的炉石：战网“内置安装”时炉石不在独立目录，而在 <战网目录>\Hearthstone 下
        if (-not $hs -and (Test-ExePath $cfg.BattleNetExe)) {
            $bnDir = Split-Path -Parent $cfg.BattleNetExe
            $bnParent = Split-Path -Parent $bnDir
            if ($bnDir) { $hs = Find-ExeUnder $bnDir 'Hearthstone.exe' 2 }
            if (-not $hs -and $bnParent) { $hs = Test-ExePath (Join-Path $bnParent 'Hearthstone\Hearthstone.exe') }
        }
        if (-not $hs) {
            foreach ($d in @(Get-DriveRoots)) {
                foreach ($rel in @('Hearthstone\Hearthstone.exe', 'Program Files\Hearthstone\Hearthstone.exe',
                                   'Program Files (x86)\Hearthstone\Hearthstone.exe', 'Games\Hearthstone\Hearthstone.exe')) {
                    if (-not $hs) { $hs = Test-ExePath (Join-Path $d $rel) }
                }
            }
        }
        # 3.4 正在运行的炉石进程
        if (-not $hs) { $hs = Find-ExeByProcess 'Hearthstone' }
        # 3.5 开始菜单 / 桌面快捷方式（国服快捷方式多指向 Hearthstone Beta Launcher.exe，取同目录的 Hearthstone.exe）
        if (-not $hs) {
            $lnk = Find-ExeByShortcut '炉石传说|Hearthstone' 'DeckTracker|HDT|Tracker|Uninstall|uninst'
            if ($lnk) {
                if ($lnk -match 'Hearthstone\.exe$') { $hs = $lnk }
                else {
                    $ldir = Split-Path -Parent $lnk
                    $hs = Find-ExeUnder $ldir 'Hearthstone.exe' 1
                    if (-not $hs) { $hs = Find-ExeUnder (Split-Path -Parent $ldir) 'Hearthstone.exe' 2 }
                }
            }
        }
        # 3.6 最后手段：浅层全盘搜索
        if (-not $hs) { $hs = Find-ExeByDiskSearch 'Hearthstone.exe' 30 4 }
        $cfg.GameExe = $hs
    }
    # 3.5 炉石独立启动器（国服有 Hearthstone Beta Launcher.exe，可绕过直启 exe 的授权问题）
    if ($cfg.GameExe) {
        $launcher = Find-ExeUnder (Split-Path -Parent $cfg.GameExe) 'Hearthstone Beta Launcher.exe' 0
        if ($launcher) { $cfg.GameLauncherExe = $launcher }
    }
    # 4) 炉石盒子
    if (-not (Test-ExePath $cfg.BoxExe)) {
        $bx = $null
        foreach ($d in @(Get-UninstallDirs @('炉石传说盒子', 'HSAng', '炉石盒子'))) { if (-not $bx) { $bx = Find-ExeUnder $d 'HSAng.exe' 2 } }
        if (-not $bx) {
            foreach ($d in @(Get-DriveRoots)) {
                foreach ($rel in @('Program Files (x86)\炉石传说盒子\HSAng.exe', 'Program Files\炉石传说盒子\HSAng.exe',
                                   '炉石传说盒子\HSAng.exe')) {
                    if (-not $bx) { $bx = Test-ExePath (Join-Path $d $rel) }
                }
            }
        }
        # 4.4 正在运行的盒子进程
        if (-not $bx) { $bx = Find-ExeByProcess 'HSAng' }
        # 4.5 开始菜单 / 桌面快捷方式
        if (-not $bx) { $bx = Find-ExeByShortcut '炉石传说盒子|HSAng|炉石盒子' 'Uninstall|uninst' }
        # 4.6 最后手段：浅层全盘搜索
        if (-not $bx) { $bx = Find-ExeByDiskSearch 'HSAng.exe' 20 4 }
        $cfg.BoxExe = $bx
    }
    # 5) 日志目录：游戏目录下的 Logs
    if (-not $cfg.LogRoot -and (Test-ExePath $cfg.GameExe)) {
        $cfg.LogRoot = Join-Path (Split-Path -Parent $cfg.GameExe) 'Logs'
    }
    if ($cfg.LogRoot -and -not (Test-Path -LiteralPath $cfg.LogRoot)) { $cfg.LogRoot = $null }
    return $cfg
}
$cfg = Resolve-Paths
Log ('  战网: ' + $(if ($cfg.BattleNetExe) { $cfg.BattleNetExe } else { '未找到' }))
Log ('  炉石: ' + $(if ($cfg.GameExe) { $cfg.GameExe } else { '未找到' }))
Log ('  盒子: ' + $(if ($cfg.BoxExe) { $cfg.BoxExe } else { '未找到' }))
if ($cfg.GameLauncherExe) { Log ('  启动器: ' + $cfg.GameLauncherExe) }
Log ('')

# ================= 进程状态 =================
function Wait-Window([string]$procName, [int]$timeoutSec) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $timeoutSec) {
        $p = Get-Process -Name $procName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
        if ($p) { return $p }
        Start-Sleep -Milliseconds 800
    }
    return $null
}
function Kill-Game {
    foreach ($n in @('Hearthstone', 'Hearthstone_x64')) {
        Get-Process -Name $n -ErrorAction SilentlyContinue | ForEach-Object {
            try { Stop-Process -Id $_.Id -Force -ErrorAction Stop } catch {}
        }
    }
}
function Get-GameProcs {
    $res = @()
    foreach ($n in @('Hearthstone', 'Hearthstone_x64')) {
        foreach ($p in @(Get-Process -Name $n -ErrorAction SilentlyContinue)) {
            $mem = 0; $hasWin = $false
            try { $mem = [int64]$p.WorkingSet64 } catch {}
            try { $hasWin = ($p.MainWindowHandle -ne 0) } catch {}
            $res += [pscustomobject]@{ Pid = $p.Id; Name = $n; Mem = $mem; HasWindow = $hasWin }
        }
    }
    return $res
}
function Test-GameAlive {
    $ps = Get-GameProcs
    return [bool](@($ps | Where-Object { $_.Mem -ge (20MB) -or $_.HasWindow }).Count -gt 0)
}
function Get-ZombieGame {
    # 真正的「卡死残留」有两种形态，判据不能只看内存大小：
    #   A. 线程已被清空（Threads=0）但进程对象仍不释放（HasExited=False）—— TerminateProcess 已是空转，
    #      内核里还有人引用这个进程对象，所以杀不掉。常见于闪退后残留的启动子进程。
    #   B. 内存极小且没有窗口 —— 老判据，保留兜底。
    $out = @()
    foreach ($n in @('Hearthstone', 'Hearthstone_x64')) {
        foreach ($p in @(Get-Process -Name $n -ErrorAction SilentlyContinue)) {
            $t = -1; $hx = $null; $mem = 0; $win = $false
            try { $t = @($p.Threads).Count } catch {}
            try { $hx = $p.HasExited } catch {}
            try { $mem = [int64]$p.WorkingSet64 } catch {}
            try { $win = ($p.MainWindowHandle -ne 0) } catch {}
            if ($hx -eq $true) { continue }
            $hit = ($t -eq 0) -or ($mem -lt 20MB -and -not $win)
            if ($hit) {
                $out += [pscustomobject]@{ Pid = $p.Id; Name = $n; Mem = $mem; HasWindow = $win; Threads = $t; HasExited = $hx }
            }
        }
    }
    return @($out)
}
# 第一步杀完之后仍然活着的炉石进程，一律算卡死残留（不再按内存大小筛，2GB 卡死的也算）
function Get-StuckGame {
    return @(Get-GameProcs)
}

function Get-ProcParent([int]$TargetPid) {
    try {
        $c = Get-CimInstance Win32_Process -Filter ('ProcessId=' + $TargetPid) -ErrorAction Stop
        if ($c) { return [int]$c.ParentProcessId }
    } catch {}
    return 0
}

# ================= 卡死诊断（只读采集，不改任何状态） =================
# 卡死时最有价值的三组证据：线程在等什么（WaitReason）、内核时间还涨不涨、进程里有没有第三方注入的模块
function Get-ThreadWaitSummary($Proc) {
    if (-not $Proc) { return '' }
    try {
        $g = @($Proc.Threads | Group-Object -Property {
                $st = ''; $wr = ''
                try { $st = [string]$_.ThreadState } catch {}
                try { $wr = [string]$_.WaitReason } catch {}
                if (-not $st) { $st = '?' }
                if (-not $wr) { $wr = '?' }
                $st + '/' + $wr
            } | Sort-Object Count -Descending | Select-Object -First 4)
        return (($g | ForEach-Object { $_.Name + '×' + $_.Count }) -join ', ')
    } catch { return '（线程信息不可读）' }
}

function Get-ProcSnapshot([int]$TargetPid) {
    $p = Get-Process -Id $TargetPid -ErrorAction SilentlyContinue
    if (-not $p) { return $null }
    $o = [pscustomobject]@{
        Pid = $TargetPid; Name = $p.ProcessName; Mem = [int64]0
        Threads = 0; Handles = 0; Responding = '?'; HasExited = ''
        KernelTime = [double]-1; UserTime = [double]-1
        StartTime = ''; Path = ''; Waits = ''
    }
    try { $o.HasExited = [string]$p.HasExited } catch { $o.HasExited = '不可读' }
    try { $o.Mem = [int64]$p.WorkingSet64 } catch {}
    try { $o.Threads = @($p.Threads).Count } catch {}
    try { $o.Handles = $p.HandleCount } catch {}
    try {
        if ($p.MainWindowHandle -ne 0) { $o.Responding = [string]$p.Responding } else { $o.Responding = '无窗口' }
    } catch { $o.Responding = '?' }
    try { $o.KernelTime = [math]::Round([double]$p.PrivilegedProcessorTime.TotalSeconds, 3) } catch {}
    try { $o.UserTime = [math]::Round([double]$p.UserProcessorTime.TotalSeconds, 3) } catch {}
    try { $o.StartTime = $p.StartTime.ToString('MM-dd HH:mm:ss') } catch {}
    try { $o.Path = $p.Path } catch {}
    $o.Waits = Get-ThreadWaitSummary $p
    return $o
}

# 只在「Windows 目录之外」+「进程自己目录之外」的模块 = 真正被第三方注入进来的
function Get-NonSystemModules([int]$TargetPid) {
    $p = Get-Process -Id $TargetPid -ErrorAction SilentlyContinue
    if (-not $p) { return @() }
    $ownDir = ''
    try { if ($p.Path) { $ownDir = (Split-Path -Parent $p.Path) } } catch {}
    try {
        return @($p.Modules | Where-Object {
                $f = $_.FileName
                if (-not $f) { return $false }
                if ($f -match '(?i)^[A-Z]:\\Windows\\') { return $false }
                if ($ownDir -and ((Split-Path -Parent $f) -eq $ownDir)) { return $false }
                return $true
            } | ForEach-Object { $_.ModuleName } | Sort-Object -Unique | Select-Object -First 40)
    } catch { return @('（模块读取失败：' + $_.Exception.Message + '）') }
}

# 一条进程的完整画像，含"内核时间还涨不涨"的冻结判定
function Write-ProcReport([int]$TargetPid, [string]$Tag) {
    $s = Get-ProcSnapshot $TargetPid
    if (-not $s) { return }
    Log ('  [{0}] {1} PID={2} 内存={3} 线程={4} 句柄={5} 响应={6}' -f $Tag, $s.Name, $s.Pid, (Human $s.Mem), $s.Threads, $s.Handles, $s.Responding)
    Log ('          启动={0} 内核时间={1}s 用户时间={2}s' -f $s.StartTime, $s.KernelTime, $s.UserTime)
    if ($s.Path) { Log ('          路径=' + $s.Path) }
    if ($s.Waits) { Log ('          线程等待分布=' + $s.Waits) }
    Start-Sleep -Milliseconds 1200
    $s2 = Get-ProcSnapshot $TargetPid
    if ($s2) {
        $d = [math]::Round(($s2.KernelTime + $s2.UserTime) - ($s.KernelTime + $s.UserTime), 3)
        if ($d -le 0.05) { Log '          活动度=1.2 秒内无 CPU 消耗（可能空闲等待，也可能卡死，要结合上面线程等待看）' }
        else { Log ('          活动度=1.2 秒内消耗 {0}s CPU（仍在跑）' -f $d) }
    }
}

# 用 SYSTEM 身份的临时计划任务去 taskkill：内核挂起的进程有时只有 SYSTEM 能收掉
function Invoke-SystemKill([int[]]$TargetPids) {
    $list = @($TargetPids | Where-Object { $_ -gt 0 })
    if ($list.Count -eq 0) { return $false }
    if (-not $isAdmin) { Log '  [SYSTEM 强杀] 当前不是管理员，跳过。'; return $false }
    $sys32 = Join-Path $env:SystemRoot 'System32'
    $taskkillExe = Join-Path $sys32 'taskkill.exe'
    $schtasksExe = Join-Path $sys32 'schtasks.exe'
    foreach ($tp in $list) {
        $tn = 'HearthstoneFix_Kill_' + $tp
        $arg = ('/F /T /PID ' + $tp)
        try {
            $action = New-ScheduledTaskAction -Execute $taskkillExe -Argument $arg
            $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
            Register-ScheduledTask -TaskName $tn -Action $action -Principal $principal -Force -ErrorAction Stop | Out-Null
            Start-ScheduledTask -TaskName $tn -ErrorAction Stop
            Log ('  [SYSTEM 强杀] 已提交 PID {0} 的结束请求（SYSTEM 身份）' -f $tp)
        } catch {
            try {
                & $schtasksExe /Create /TN $tn /TR ('"' + $taskkillExe + '" ' + $arg) /SC ONCE /ST 23:59 /RU SYSTEM /RL HIGHEST /F 2>&1 | Out-Null
                & $schtasksExe /Run /TN $tn 2>&1 | Out-Null
                Log ('  [SYSTEM 强杀] 已提交 PID {0} 的结束请求（schtasks 方式）' -f $tp)
            } catch {
                Log ('  [SYSTEM 强杀] PID {0} 提交失败：{1}' -f $tp, $_.Exception.Message)
                continue
            }
        }
        Start-Sleep -Seconds 4
        & $schtasksExe /Delete /TN $tn /F 2>&1 | Out-Null
    }
    Start-Sleep -Seconds 2
    $left = @(Get-StuckGame | Where-Object { $list -contains $_.Pid })
    if ($left.Count -gt 0) {
        foreach ($lp in $left) {
            $s = Get-ProcSnapshot $lp.Pid
            if ($s) { Log ('  [SYSTEM 强杀] PID {0} 仍存活：线程={1} 句柄={2} 内核时间={3}s 等待={4}' -f $s.Pid, $s.Threads, $s.Handles, $s.KernelTime, $s.Waits) }
        }
    }
    return ($left.Count -eq 0)
}

# ================= 战网界面：窗口控制 + 截图 + 按钮识别 =================
$script:BnUiReady = $false
function Initialize-BnUi {
    if ($script:BnUiReady) { return $true }
    try {
        $code = @'
using System;
using System.Runtime.InteropServices;
public static class BnUi {
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern int GetSystemMetrics(int n);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, UIntPtr extra);
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    public static RECT RectOf(IntPtr h) { RECT r; GetWindowRect(h, out r); return r; }
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(300);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);   // left down
        System.Threading.Thread.Sleep(120);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);   // left up
    }
}
'@
        Add-Type -TypeDefinition $code -ErrorAction Stop
        [BnUi]::SetProcessDPIAware() | Out-Null
        $script:BnUiReady = $true
    } catch { Log ('  [提示] 界面自动化初始化失败: ' + $_.Exception.Message) }
    return $script:BnUiReady
}
$script:OwnConsoleHwnd = [IntPtr]::Zero
function Minimize-OwnConsole {
    if (-not $script:BnUiReady) { return }
    try {
        if ($script:OwnConsoleHwnd -eq [IntPtr]::Zero) { $script:OwnConsoleHwnd = [BnUi]::GetConsoleWindow() }
        if ($script:OwnConsoleHwnd -ne [IntPtr]::Zero) { [BnUi]::ShowWindow($script:OwnConsoleHwnd, 6) | Out-Null; Start-Sleep -Milliseconds 400 }
    } catch {}
}
function Restore-OwnConsole {
    if (-not $script:BnUiReady) { return }
    try {
        if ($script:OwnConsoleHwnd -eq [IntPtr]::Zero) { $script:OwnConsoleHwnd = [BnUi]::GetConsoleWindow() }
        if ($script:OwnConsoleHwnd -ne [IntPtr]::Zero) { [BnUi]::ShowWindow($script:OwnConsoleHwnd, 9) | Out-Null }
    } catch {}
}
function Wait-BnWindow([int]$Timeout = 90) {
    $deadline = (Get-Date).AddSeconds($Timeout)
    while ((Get-Date) -lt $deadline) {
        $p = Get-Process -Name 'Battle.net' -ErrorAction SilentlyContinue |
             Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
        if ($p) { return $p }
        Start-Sleep -Milliseconds 800
    }
    return $null
}
function Get-BnCapture($rect) {
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
    $sw = [BnUi]::GetSystemMetrics(0); $sh = [BnUi]::GetSystemMetrics(1)
    $l = [Math]::Max($rect.Left, 0); $t = [Math]::Max($rect.Top, 0)
    $r = [Math]::Min($rect.Right, $sw); $b = [Math]::Min($rect.Bottom, $sh)
    $w = $r - $l; $h = $b - $t
    if ($w -lt 200 -or $h -lt 200) { return $null }
    $bmp = New-Object System.Drawing.Bitmap($w, $h, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($l, $t, 0, 0, $bmp.Size)
    $g.Dispose()
    $bd = $bmp.LockBits((New-Object System.Drawing.Rectangle(0, 0, $w, $h)),
                        [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $stride = $bd.Stride
    $raw = New-Object byte[] ($stride * $h)
    [System.Runtime.InteropServices.Marshal]::Copy($bd.Scan0, $raw, 0, $raw.Length)
    $bmp.UnlockBits($bd)
    return @{ Bmp = $bmp; Left = $l; Top = $t; Width = $w; Height = $h; Stride = $stride; Raw = $raw }
}
function Get-PackedBgra($cap) {
    $w = $cap.Width; $h = $cap.Height; $stride = $cap.Stride; $raw = $cap.Raw
    $line = $w * 4
    if ($stride -eq $line) { return $raw }
    $packed = New-Object byte[] ($line * $h)
    for ($y = 0; $y -lt $h; $y++) { [Array]::Copy($raw, $y * $stride, $packed, $y * $line, $line) }
    return $packed
}

# ---- 策略 1：OCR 识别按钮文字（跨分辨率/主题最可靠）----
$script:OcrEngine = $null
$script:OcrTried = $false
$script:AsTaskMI = $null
function Initialize-Ocr {
    if ($script:OcrTried) { return $script:OcrEngine }
    $script:OcrTried = $true
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
        [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
        [Windows.Graphics.Imaging.BitmapPixelFormat, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
        [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
        [Windows.Globalization.Language, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
        $script:AsTaskMI = ([System.WindowsRuntimeSystemExtensions].GetMethods() |
            Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
                           $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
        $eng = $null
        foreach ($tag in @('zh-Hans-CN', 'zh-Hans', 'zh-CN')) {
            try {
                $lang = New-Object Windows.Globalization.Language($tag)
                $eng = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
            } catch {}
            if ($eng) { break }
        }
        if (-not $eng) { $eng = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
        $script:OcrEngine = $eng
        if ($eng) { Log ('  [识别] 文字识别引擎就绪（' + $eng.RecognizerLanguage.LanguageTag + '）') }
        else { Log '  [识别] 系统未安装文字识别语言包，改用画面识别' }
    } catch { Log ('  [识别] 引擎初始化失败: ' + $_.Exception.Message) }
    return $script:OcrEngine
}
function Await-WinRt($op, [Type]$type) {
    $task = $script:AsTaskMI.MakeGenericMethod($type).Invoke($null, @($op))
    if (-not $task.Wait(30000)) { throw 'WinRT 调用超时' }
    return $task.Result
}
function Find-BnButtonByOcr($cap) {
    $eng = Initialize-Ocr
    if (-not $eng) { return $null }
    try {
        $bytes = Get-PackedBgra $cap
        $buf = [System.Runtime.InteropServices.WindowsRuntime.WindowsRuntimeBufferExtensions]::AsBuffer($bytes)
        $sb = [Windows.Graphics.Imaging.SoftwareBitmap]::CreateCopyFromBuffer(
                  $buf, [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8, $cap.Width, $cap.Height)
        $res = Await-WinRt ($eng.RecognizeAsync($sb)) ([Windows.Media.Ocr.OcrResult])
        $sb.Dispose()
    } catch { Log ('  [识别] 文字识别出错: ' + $_.Exception.Message); return $null }

    $kw = @('进入游戏', '开始游戏', '进行游戏', '启动游戏', '开始对局', '进入对局', 'Play')
    $best = $null; $bestScore = -1.0
    foreach ($line in @($res.Lines)) {
        $words = @($line.Words)
        if ($words.Count -eq 0) { continue }
        $text = (($words | ForEach-Object { $_.Text }) -join '')
        $hit = $false
        foreach ($k in $kw) { if ($text -like "*$k*") { $hit = $true; break } }
        if (-not $hit) { continue }
        $minX = $cap.Width; $minY = $cap.Height; $maxX = -1; $maxY = -1
        foreach ($wd in $words) {
            $br = $wd.BoundingRect
            if ($br.X -lt $minX) { $minX = $br.X }
            if ($br.Y -lt $minY) { $minY = $br.Y }
            if (($br.X + $br.Width) -gt $maxX) { $maxX = $br.X + $br.Width }
            if (($br.Y + $br.Height) -gt $maxY) { $maxY = $br.Y + $br.Height }
        }
        $cx = [int](($minX + $maxX) / 2); $cy = [int](($minY + $maxY) / 2)
        # 位于窗口左下区域更可能是主按钮
        $score = ($cy / [double]$cap.Height) - ($cx / [double]$cap.Width) * 0.4
        if ($text -notmatch '更新|安装|下载') { $score += 0.6 }
        if ($score -gt $bestScore) {
            $bestScore = $score
            $best = @{ X = $cap.Left + $cx; Y = $cap.Top + $cy; Text = $text; How = '文字识别' }
        }
    }
    if ($best) { Log ('  [识别] 文字命中「' + $best.Text + '」，屏幕坐标 (' + $best.X + ',' + $best.Y + ')') }
    else {
        $sample = @(@($res.Lines) | Select-Object -First 8 | ForEach-Object { (($_.Words) | ForEach-Object { $_.Text }) -join '' })
        if ($sample.Count -gt 0) { Log ('  [识别] 未见启动字样，识别到: ' + ($sample -join ' / ')) }
    }
    return $best
}

# ---- 策略 2：自适应色块（不写死颜色，换主题也能用）----
function Find-BnButtonByColor($cap) {
    $w = $cap.Width; $h = $cap.Height; $stride = $cap.Stride; $raw = $cap.Raw
    $step = 2
    $yStart = [int]($h * 0.45)
    $bestLen = 0; $bestStart = -1; $bestEnd = -1; $bestY = 0
    for ($y = $yStart; $y -lt $h; $y += $step) {
        $rowOff = $y * $stride
        $run = 0; $runStart = -1
        for ($x = 0; $x -lt $w; $x += $step) {
            $i = $rowOff + $x * 4
            $b = $raw[$i]; $g = $raw[$i + 1]; $r = $raw[$i + 2]
            $mx = [Math]::Max($r, [Math]::Max($g, $b))
            $mn = [Math]::Min($r, [Math]::Min($g, $b))
            if (($mx - $mn) -ge 45 -and $mx -ge 70 -and $mn -le 205) {
                if ($run -eq 0) { $runStart = $x }
                $run++
                if ($run -gt $bestLen) { $bestLen = $run; $bestStart = $runStart; $bestEnd = $x; $bestY = $y }
            } else { $run = 0 }
        }
    }
    if ($bestLen -lt 10) { return $null }                       # 至少 20px 宽
    $bw = $bestEnd - $bestStart
    if ($bw -lt ($w * 0.10) -or $bw -gt ($w * 0.70)) { return $null }
    # 以该行为种子，上下扩展同宽度区域
    $top = $bestY; $bot = $bestY
    for ($y = $bestY - $step; $y -ge $yStart; $y -= $step) {
        if (Test-RowBand $raw $y $stride $bestStart $bestEnd $step) { $top = $y } else { break }
    }
    for ($y = $bestY + $step; $y -lt $h; $y += $step) {
        if (Test-RowBand $raw $y $stride $bestStart $bestEnd $step) { $bot = $y } else { break }
    }
    $bh = $bot - $top
    if ($bh -lt 22) { return $null }
    $cx = [int](($bestStart + $bestEnd) / 2); $cy = [int](($top + $bot) / 2)
    Log ('  [识别] 画面中找到高亮条状区域，尺寸 ' + $bw + 'x' + $bh + '，屏幕坐标 (' + ($cap.Left + $cx) + ',' + ($cap.Top + $cy) + ')')
    return @{ X = $cap.Left + $cx; Y = $cap.Top + $cy; Text = ''; How = '画面色块' }
}
function Test-RowBand($raw, [int]$y, [int]$stride, [int]$x1, [int]$x2, [int]$step) {
    $rowOff = $y * $stride
    $hit = 0; $total = 0
    for ($x = $x1; $x -le $x2; $x += $step) {
        $i = $rowOff + $x * 4
        $b = $raw[$i]; $g = $raw[$i + 1]; $r = $raw[$i + 2]
        $mx = [Math]::Max($r, [Math]::Max($g, $b)); $mn = [Math]::Min($r, [Math]::Min($g, $b))
        if (($mx - $mn) -ge 45 -and $mx -ge 70 -and $mn -le 205) { $hit++ }
        $total++
    }
    if ($total -eq 0) { return $false }
    return ([double]$hit / $total -ge 0.6)
}

# ---- 策略 3：固定蓝色判据（最后兜底）----
function Find-BnButtonByBlue($cap) {
    $w = $cap.Width; $h = $cap.Height; $stride = $cap.Stride; $raw = $cap.Raw
    $x1 = 20; $x2 = [Math]::Min(720, $w - 1)
    $y1 = [Math]::Max($h - 260, 0); $y2 = $h - 20
    $minX = 99999; $maxX = -1; $minY = 99999; $maxY = -1; $cnt = 0
    for ($y = $y1; $y -lt $y2; $y += 2) {
        $rowOff = $y * $stride
        for ($x = $x1; $x -lt $x2; $x += 2) {
            $i = $rowOff + $x * 4
            $b = $raw[$i]; $g = $raw[$i + 1]; $r = $raw[$i + 2]
            if ($b -gt 150 -and $b -lt 235 -and $r -lt 60 -and $g -gt 80 -and $g -lt 170) {
                $cnt++
                if ($x -lt $minX) { $minX = $x }
                if ($x -gt $maxX) { $maxX = $x }
                if ($y -lt $minY) { $minY = $y }
                if ($y -gt $maxY) { $maxY = $y }
            }
        }
    }
    $bw = $maxX - $minX; $bh = $maxY - $minY
    if ($cnt -lt 300 -or $bw -lt 120 -or $bw -gt 700 -or $bh -lt 24) { return $null }
    $cx = [int](($minX + $maxX) / 2); $cy = [int](($minY + $maxY) / 2)
    Log ('  [识别] 蓝色按钮兜底命中，尺寸 ' + $bw + 'x' + $bh + '，屏幕坐标 (' + ($cap.Left + $cx) + ',' + ($cap.Top + $cy) + ')')
    return @{ X = $cap.Left + $cx; Y = $cap.Top + $cy; Text = ''; How = '颜色兜底' }
}

function Invoke-BnEnterGame([int]$WaitSec = 20, [int]$Retry = 3) {
    if (-not (Initialize-BnUi)) { return $false }
    $p = Wait-BnWindow 60
    if (-not $p) { Log '  [点击] 等了 60 秒战网窗口仍未出现'; return $false }
    $h = [IntPtr]$p.MainWindowHandle
    if (-not $KeepConsole) { Minimize-OwnConsole }      # 关键：自己的命令行窗口不能挡住按钮
    [BnUi]::ShowWindow($h, 3) | Out-Null
    Start-Sleep -Milliseconds 600
    [BnUi]::BringWindowToTop($h) | Out-Null
    [BnUi]::SetForegroundWindow($h) | Out-Null
    Start-Sleep -Milliseconds 1500
    $rect = [BnUi]::RectOf($h)
    Log ('  [窗口] 战网 rect=' + $rect.Left + ',' + $rect.Top + ',' + $rect.Right + ',' + $rect.Bottom)
    for ($attempt = 1; $attempt -le $Retry; $attempt++) {
        [BnUi]::BringWindowToTop($h) | Out-Null
        [BnUi]::SetForegroundWindow($h) | Out-Null
        Start-Sleep -Milliseconds 400
        $cap = Get-BnCapture $rect
        if (-not $cap) { Log '  [点击] 截图失败'; Start-Sleep -Seconds 3; continue }
        $hit = Find-BnButtonByOcr $cap
        if (-not $hit) { $hit = Find-BnButtonByColor $cap }
        if (-not $hit) { $hit = Find-BnButtonByBlue $cap }
        try { $cap.Bmp.Dispose() } catch {}
        if ($hit) {
            Log ('  [点击] 第 {0} 次用「{1}」定位到 ({2},{3})' -f $attempt, $hit.How, $hit.X, $hit.Y)
            [BnUi]::Click($hit.X, $hit.Y)
            for ($i = 0; $i -lt $WaitSec; $i++) {
                Start-Sleep -Seconds 1
                if (Test-GameAlive) { return $true }
            }
            return (Test-GameAlive)
        }
        Log ('  [点击] 第 {0} 次未定位到按钮' -f $attempt)
        if ($attempt -lt $Retry) { Start-Sleep -Seconds 5 }
    }
    if (-not $KeepConsole) { Restore-OwnConsole }
    return $false
}

# 读最近的崩溃事件，判断是哪一种崩法——它直接决定"杀不掉"是不是必然结果
function Get-LastCrash([int]$Minutes = 360) {
    try {
        $ev = Get-WinEvent -FilterHashtable @{ LogName = 'Application'; Id = 1000; StartTime = (Get-Date).AddMinutes(-$Minutes) } -ErrorAction Stop |
              Where-Object { $_.Message -match 'Hearthstone\.exe' } | Select-Object -First 1
        if (-not $ev) { return $null }
        $m = $ev.Message
        $mod = ''; $code = ''
        if ($m -match '(?:错误模块名称|Faulting module name)[:：]\s*([^\s,，]+)') { $mod = $matches[1] }
        if ($m -match '(?:异常代码|Exception code)[:：]\s*(0x[0-9a-fA-F]+)') { $code = $matches[1] }
        $mean = ''
        switch ($code) {
            '0xc0000374' { $mean = '堆损坏（HEAP_CORRUPTION）：进程退出时清理堆会自己卡死，所以怎么杀都杀不掉，这是必然' }
            '0xc0000005' { $mean = '访问违例（ACCESS_VIOLATION）：多半是注入进游戏的叠加层/记牌器踩了野指针' }
            '0xc0000409' { $mean = '快速失败（栈保护或安全 cookie 被破坏）' }
            default { if ($code) { $mean = '异常码 ' + $code } }
        }
        return [pscustomobject]@{ Time = $ev.TimeCreated; Module = $mod; Code = $code; Mean = $mean }
    } catch { return $null }
}

function Show-CrashEvidence {
    $log = "$env:USERPROFILE\AppData\LocalLow\Blizzard Entertainment\Hearthstone\Player.log"
    if (Test-Path -LiteralPath $log) {
        Log '  最后 12 行游戏日志:'
        Get-Content -LiteralPath $log -Tail 12 -ErrorAction SilentlyContinue | ForEach-Object { Log ('    ' + $_) }
    }
    $ev = Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddMinutes(-10)} -ErrorAction SilentlyContinue |
          Where-Object { $_.Message -match 'Hearthstone' -and $_.Id -eq 1000 } | Select-Object -First 1
    if ($ev) { Log ('  最近崩溃: ' + (($ev.Message -split "`n")[1..2] -join ' / ')) }
}
# 询问并执行重启：空格 / 回车 = 确认，倒计时中再按可每次提前 1 秒，其它键 = 取消
function Get-KeyCode {
    # 非阻塞取一个按键码；0 表示当前没有按键
    if ($script:SimKeysQueue.Count -gt 0) {
        $c = $script:SimKeysQueue[0]
        $script:SimKeysQueue = @($script:SimKeysQueue | Select-Object -Skip 1)
        return [int]$c
    }
    try {
        if ([Console]::KeyAvailable) {
            $k = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            return [int]$k.VirtualKeyCode
        }
    } catch {}
    return 0
}
function Invoke-RebootAsk([int]$Countdown = 8) {
    Log ''
    Log '  是否立即重启电脑来清除这些卡死进程？'
    Log '    注意：重启会关闭所有程序，未保存的工作会丢失。'
    Log '    按 [空格] 或 [回车] 确认；按其他键取消（只退出，不重启）'
    $confirm = $false
    if ([Console]::IsInputRedirected -and $script:SimKeysQueue.Count -eq 0) {
        $ans = Read-Host '    输入 y 后回车确认重启，直接回车或其它字符取消'
        $confirm = ($ans -eq 'y' -or $ans -eq 'Y')
    } else {
        $deadline = (Get-Date).AddSeconds(60)
        while ((Get-Date) -lt $deadline) {
            $code = Get-KeyCode
            if ($code -ne 0) {
                if ($code -eq 32 -or $code -eq 13) { $confirm = $true }
                break
            }
            Start-Sleep -Milliseconds 150
        }
    }
    if (-not $confirm) { Log '  已取消重启。'; return 'cancel' }
    Log ('  开始重启倒计时 {0} 秒：再按 [空格] / [回车] 可每次提前 1 秒，按其它键取消...' -f $Countdown)
    $remain = $Countdown
    while ($remain -gt 0) {
        Write-Host ('    剩余 {0} 秒' -f $remain)
        $elapsed = 0
        $pressed = $false
        while ($elapsed -lt 1000 -and -not $pressed -and $remain -gt 0) {
            $code = Get-KeyCode
            if ($code -ne 0) {
                if ($code -eq 32 -or $code -eq 13) {
                    $remain--
                    $pressed = $true
                    Log ('  [提前] 已加速 1 秒，剩余 {0} 秒' -f ([Math]::Max($remain, 0)))
                } else {
                    Log '  已取消重启。'
                    return 'cancel'
                }
            }
            Start-Sleep -Milliseconds 100
            $elapsed += 100
        }
        if (-not $pressed -and $remain -gt 0) { $remain-- }
    }
    if ($TestReboot) {
        Log '  [测试] 倒计时结束，已命中"执行重启"分支（测试模式不真的重启）'
        return 'cancel'
    }
    Log '  正在重启...'
    $ok = $false
    try {
        $r = & shutdown.exe /r /t 0 /f 2>&1
        if ($LASTEXITCODE -eq 0) { $ok = $true }
        else { Log ('  shutdown.exe 返回 ' + $LASTEXITCODE + '：' + ($r -join ' ')) }
    } catch { Log ('  shutdown.exe 调用失败: ' + $_.Exception.Message) }
    if (-not $ok) {
        try { Restart-Computer -Force -ErrorAction Stop; $ok = $true }
        catch { Log ('  Restart-Computer 也失败: ' + $_.Exception.Message) }
    }
    if (-not $ok) { Log '  >>> 自动重启失败，请手动重启电脑（开始菜单 → 重启）。'; return 'fail' }
    return 'ok'
}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Log ('模式: ' + $(if ($DryRun) { '预览' } else { '执行' }) + '   管理员: ' + $(if ($isAdmin) { '有' } else { '无' }))
Log ''
if (-not $DryRun -and -not $NoShortcut -and -not $DiagOnly) { $null = Ensure-Shortcut }

if ($TestReboot) {
    Log '=== 测试模式：仅验证按键与重启询问 ==='
    $null = Invoke-RebootAsk 3
    Log '=== 测试结束 ==='
    Finish 0
}
if ($TestClick) {
    Log '=== 测试模式：识别并点击战网「进入游戏」 ==='
    $ok = Invoke-BnEnterGame 25 3
    Log ('结果: ' + $ok)
    Finish 0
}

# ---------- 0. 诊断模式 ----------
if ($DiagOnly) {
    Log '=== 诊断模式：只采集现场，不结束任何进程、不启动任何程序 ==='
    $probe = @()
    foreach ($n in @('Hearthstone', 'Hearthstone_x64', 'Battle.net', 'Battle.net Helper', 'Agent', 'BlizzardBrowser', 'HSAng')) {
        $probe += @(Get-Process -Name $n -ErrorAction SilentlyContinue)
    }
    if (@($probe).Count -eq 0) {
        Log '  当前没有相关进程在跑（游戏没开，或已经退干净了）。'
        Log '=== 诊断结束 ==='
        Finish 0
    }
    $parentMap = @{}
    try {
        Get-CimInstance Win32_Process -ErrorAction Stop | ForEach-Object {
            $parentMap[[int]$_.ProcessId] = [pscustomobject]@{ Parent = [int]$_.ParentProcessId; Cmd = $_.CommandLine }
        }
    } catch {}
    foreach ($p in @($probe)) {
        Write-ProcReport $p.Id '现场'
        if ($parentMap.ContainsKey([int]$p.Id)) {
            $ln = $parentMap[[int]$p.Id]
            if ($parentMap.ContainsKey($ln.Parent)) { Log ('          父进程PID=' + $ln.Parent + '（仍在）') }
            else { Log ('          父进程PID=' + $ln.Parent + '（已退出）') }
            if ($ln.Cmd) { Log ('          命令行=' + $ln.Cmd) }
        }
        $mods = Get-NonSystemModules $p.Id
        if (@($mods).Count -gt 0) { Log ('          非系统模块(' + @($mods).Count + ')=' + ($mods -join ', ')) }
    }
    Log ''
    Log '=== 诊断结束：未结束任何进程，直接把 hs-fix-log.txt 拿来分析即可 ==='
    Finish 0
}

# ---------- 1. 结束残留进程 ----------
Log '=== 1/5 结束残留进程 ==='

# 顺序很关键：先把「注入方」退干净，最后才退被注入的炉石。
# 炉石盒子（HSAng）会把 DLL 注入进炉石进程；注入方先死、钩子先解，被注入方才有可能从等待里出来。
# 反过来先杀炉石，等于把还在等注入方的线程直接钉死，之后再怎么提权也收不掉。
$groups = @()
if (-not $KeepBox) { $groups += ,@('HSAng') }
$groups += ,@('BlizzardBrowser', 'BlizzardError')
$groups += ,@('Battle.net Helper')
$groups += ,@('Agent')
$groups += ,@('Battle.net')
$groups += ,@('Hearthstone', 'Hearthstone_x64')

$killed = 0
$script:BypassBn = $false
$gi = 0
foreach ($g in $groups) {
    $gi++
    $targets = @()
    foreach ($n in $g) { $targets += @(Get-Process -Name $n -ErrorAction SilentlyContinue) }
    if (@($targets).Count -eq 0) { continue }
    # 组内按内存从大到小处理：真正在跑的游戏（父进程、大内存）先死，它留下的空壳子进程才有机会被连带回收
    $targets = @($targets | Sort-Object -Property { try { [int64]$_.WorkingSet64 } catch { [int64]0 } } -Descending)
    Log ('  -- 第 {0}/{1} 步：{2} --' -f $gi, @($groups).Count, (@($g) -join ' / '))

    foreach ($p in @($targets)) {
        $tp = $p.Id
        $bp = Get-ProcSnapshot $tp
        if (-not $bp) { continue }
        Log ('     [处理] {0} PID={1} 内存={2} 线程={3} 内核时间={4}s 等待={5}' -f $bp.Name, $tp, (Human $bp.Mem), $bp.Threads, $bp.KernelTime, $bp.Waits)

        if ($DryRun) { Log '           （预览，不实际结束）'; continue }

        try { Stop-Process -Id $tp -Force -ErrorAction Stop }
        catch { Log ('           [结束请求被拒] ' + $_.Exception.Message) }

        $sw = [Diagnostics.Stopwatch]::StartNew()
        $gone = $false
        while ($sw.Elapsed.TotalSeconds -lt 5) {
            if (-not (Get-Process -Id $tp -ErrorAction SilentlyContinue)) { $gone = $true; break }
            Start-Sleep -Milliseconds 250
        }
        if ($gone) {
            $killed++
            Log ('           [已结束] 用时 {0:N2}s' -f $sw.Elapsed.TotalSeconds)
        } else {
            $ap = Get-ProcSnapshot $tp
            Log '           [未结束] 5 秒内仍存活'
            if ($ap) { Log ('                    残留：线程={0} 句柄={1} 内核时间={2}s 等待={3}' -f $ap.Threads, $ap.Handles, $ap.KernelTime, $ap.Waits) }
        }
    }

    if (-not $DryRun) {
        if (@($g) -contains 'HSAng') {
            # 盒子退完后单独观察炉石：如果此时炉石的线程等待状态松了，就说明"注入方先行"这条思路对路
            Start-Sleep -Milliseconds 1500
            foreach ($hp in @(Get-Process -Name 'Hearthstone', 'Hearthstone_x64' -ErrorAction SilentlyContinue)) {
                $hs = Get-ProcSnapshot $hp.Id
                if ($hs) { Log ('     [观察] 盒子退出后炉石 PID={0}：线程={1} 内核时间={2}s 等待={3}' -f $hs.Pid, $hs.Threads, $hs.KernelTime, $hs.Waits) }
            }
        }
        elseif (-not (@($g) -contains 'Hearthstone')) { Start-Sleep -Milliseconds 800 }
    }
}
if (-not $DryRun -and $killed -eq 0) { Log '  没有需要结束的进程' }

# ---------- 1.5 卡死残留（含第 1 步没杀掉的） ----------
if ($DryRun) { $zombies = @(Get-ZombieGame) } else { $zombies = @(Get-StuckGame) }
# 先看崩溃事件：如果是堆损坏，"杀不掉"就是必然结果，直接决定后面走绕过而不是硬杀
$crash = if ($DryRun) { $null } else { Get-LastCrash }
$heapCrash = $false
if ($crash) {
    Log ('  最近一次崩溃：{0}  模块={1}  代码={2}' -f $crash.Time.ToString('MM-dd HH:mm:ss'), $crash.Module, $crash.Code)
    if ($crash.Mean) { Log ('               ' + $crash.Mean) }
    if ($crash.Code -eq '0xc0000374') { $heapCrash = $true }
}
if (@($zombies).Count -gt 0) {
    Log ('  卡死残留：' + (@($zombies | ForEach-Object { 'PID=' + $_.Pid + '(' + [int]($_.Mem / 1KB) + 'KB)' }) -join ' '))
    foreach ($z in @($zombies)) {
        $s = Get-ProcSnapshot $z.Pid
        if ($s) {
            Log ('    [卡死画像] PID={0} 线程={1} 句柄={2} 已退出={3} 内核时间={4}s 等待={5}' -f $s.Pid, $s.Threads, $s.Handles, $s.HasExited, $s.KernelTime, $s.Waits)
            $par = Get-ProcParent $z.Pid
            if ($par -gt 0) {
                $pp = Get-Process -Id $par -ErrorAction SilentlyContinue
                if ($pp) { Log ('               父进程 PID={0}（{1}，仍在跑）' -f $par, $pp.ProcessName) }
                else { Log ('               父进程 PID={0}（已退出）' -f $par) }
            }
            $mods = Get-NonSystemModules $z.Pid
            if (@($mods).Count -gt 0) { Log ('               非系统模块(' + @($mods).Count + ')=' + ($mods -join ', ')) }
        }
    }
    $shellOnly = (@($zombies).Count -gt 0 -and @($zombies | Where-Object { $_.HasExited -eq $true }).Count -eq @($zombies).Count)
    if (-not $DryRun) {
        Kill-Game; Start-Sleep -Seconds 2; $zombies = @(Get-StuckGame)
        $shellOnly = (@($zombies).Count -gt 0 -and @($zombies | Where-Object { $_.HasExited -eq $true }).Count -eq @($zombies).Count)
        if (@($zombies).Count -gt 0) {
            if ($shellOnly) {
                Log '  这些是已终止的空壳（HasExited=True）：终止请求对它们本来就是空转，跳过 SYSTEM 强杀。'
                if ($heapCrash) { Log '  崩溃码是 0xc0000374 堆损坏：进程退出时清理堆会自己卡死，用户态没有手段把它推完，只能绕过。' }
            } else {
                Log '  普通结束无效，改用 SYSTEM 身份再试一次（这一步不重启也有机会清掉）...'
                $null = Invoke-SystemKill @($zombies | ForEach-Object { $_.Pid })
                $zombies = @(Get-StuckGame)
                if (@($zombies).Count -eq 0) { Log '  [已清除] SYSTEM 强杀成功，无需重启，继续自动启动。' }
            }
        } else {
            Log '  [已清除] 再杀一次就干净了。'
        }
    }
    if (@($zombies).Count -gt 0) {
        Log '  [阻断] 上面的进程仍然活着：普通结束、提权结束、SYSTEM 身份都收不掉。'
        if ($shellOnly) {
            Log '  [判别] 它们是「已终止但内核对象未回收」的空壳：线程=0 且 HasExited=True。'
            Log '         杀是空转（不是权限问题），但它们的互斥体/文件锁在终止那一刻已释放，游戏自己不会再拦。'
            Log '         真正拦路的只是战网客户端的「正在游戏中」判断。'
            if (-not $NoLaunch -and -not $DryRun) {
                Log '  [降级] 先不重启：改走「绕过战网、直接启动游戏主程序」。'
                $script:BypassBn = $true
            } else {
                Log '         >>> （预览或 -NoLaunch）真实执行时会走绕过战网直接启动的降级路线。'
            }
        } else {
            Log '  [判别] 有进程尚未真正退出（HasExited=False 或线程仍在），属于内核态等待卡死。'
            Log '         若父进程仍在跑，先结束父进程再等 2 秒；父进程也没了就只能重启。'
            if (-not $NoLaunch -and -not $DryRun) {
                $rb = Invoke-RebootAsk 8
                if ($rb -eq 'ok') { Log '  系统即将重启，重启后重新双击「一键修复」即可全自动。'; Finish 0 }
                elseif ($rb -eq 'cancel') { Log ''; Log '=== 已取消：未重启，按需再双击运行 ==='; Finish 2 }
                else { Log ''; Log '=== 自动重启失败：请手动重启电脑后重试 ==='; Finish 3 }
            } else {
                Log '         >>> （预览模式）实际执行时会尝试 SYSTEM 强杀，仍失败才提示按键重启。'
            }
        }
    }
}
Log ''

# ---------- 2. 叠加层 ----------
Log '=== 2/5 游戏叠加层 ==='
if ($SkipOverlay) { Log '  跳过' }
else {
    $jobs = if ($RestoreOverlay) {
        @(
            @{ K = 'HKCU:\System\GameConfigStore'; N = 'GameDVR_Enabled'; V = 1 },
            @{ K = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR'; N = 'AppCaptureEnabled'; V = 1 }
        )
    } else {
        @(
            @{ K = 'HKCU:\System\GameConfigStore'; N = 'GameDVR_Enabled'; V = 0 },
            @{ K = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR'; N = 'AppCaptureEnabled'; V = 0 },
            @{ K = 'HKCU:\Software\Microsoft\GameBar'; N = 'ShowStartupPanel'; V = 0 },
            @{ K = 'HKCU:\Software\Microsoft\GameBar'; N = 'UseNexusForGameBarEnabled'; V = 0 }
        )
    }
    foreach ($j in $jobs) {
        if (-not (Test-Path $j.K)) { if (-not $DryRun) { New-Item -Path $j.K -Force | Out-Null } }
        $old = (Get-ItemProperty -Path $j.K -Name $j.N -ErrorAction SilentlyContinue).($j.N)
        if ($DryRun) { Log ('  [预览] {0}\{1}: {2} -> {3}' -f $j.K, $j.N, $old, $j.V) }
        else {
            try { Set-ItemProperty -Path $j.K -Name $j.N -Value $j.V -Type DWord -ErrorAction Stop; Log ('  [已设置] {0}: {1} -> {2}' -f $j.N, $old, $j.V) }
            catch { Log ('  [失败] ' + $j.N) }
        }
    }
}
Log ''

# ---------- 3. 日志 ----------
Log '=== 3/5 日志清理 ==='
if ($SkipLogs) { Log '  跳过' }
elseif (-not $cfg.LogRoot) { Log '  未找到日志目录' }
else {
    $running = [bool](Get-Process -Name Hearthstone -ErrorAction SilentlyContinue)
    if ($running) { Log ('  游戏在运行' + $(if ($DryRun) { '，实际执行会跳过' } else { '，跳过清理' })) }
    $cutoff = (Get-Date).AddDays(-[Math]::Abs($KeepDays))
    $freed = 0
    if (-not ($running -and -not $DryRun)) {
        Log ('  保留 {0} 天，早于 {1} 的会话将删除' -f $KeepDays, $cutoff.ToString('yyyy-MM-dd HH:mm'))
        $dirs = Get-ChildItem -LiteralPath $cfg.LogRoot -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like 'Hearthstone_*' -and $_.LastWriteTime -lt $cutoff }
        foreach ($d in $dirs) {
            $size = [int64]((Get-ChildItem -LiteralPath $d.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum)
            if ($DryRun) { Log ('  [预览] 删除 {0} ({1})' -f $d.Name, (Human $size)) }
            else {
                try { Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction Stop; $freed += $size; Log ('  [已删除] {0} ({1})' -f $d.Name, (Human $size)) }
                catch { Log ('  [跳过] ' + $d.Name) }
            }
        }
        if (@($dirs).Count -eq 0) { Log '  没有超期日志' }
        $big = Get-ChildItem -LiteralPath $cfg.LogRoot -Directory -ErrorAction SilentlyContinue |
               Where-Object { $_.LastWriteTime -ge $cutoff }
        foreach ($d in $big) {
            $pw = Join-Path $d.FullName 'Power.log'
            if (Test-Path -LiteralPath $pw) {
                $len = (Get-Item -LiteralPath $pw).Length
                if ($len -gt 50MB) {
                    if ($DryRun) { Log ('  [预览] 清空 {0}\Power.log ({1})' -f $d.Name, (Human $len)) }
                    else {
                        try { Set-Content -LiteralPath $pw -Value '' -Force -ErrorAction Stop; Log ('  [已清空] {0}\Power.log ({1})' -f $d.Name, (Human $len)) }
                        catch { Log '  [跳过] Power.log 被占用' }
                    }
                }
            }
        }
        if (-not $DryRun) { Log ('  释放: ' + (Human $freed)) }
    }
}
Log ''

# ---------- 4. 战网服务 ----------
Log '=== 4/5 战网服务 ==='
$svc = Get-Service battlenet_helpersvc -ErrorAction SilentlyContinue
if ($DryRun) { Log ('  [预览] 重启 battlenet_helpersvc（当前 ' + $svc.Status + '）') }
elseif (-not $isAdmin) { Log '  需要管理员权限，跳过' }
elseif (-not $svc) { Log '  未安装 battlenet_helpersvc，跳过' }
else {
    try { Stop-Service battlenet_helpersvc -Force -ErrorAction SilentlyContinue; Start-Service battlenet_helpersvc -ErrorAction Stop; Log '  [已重启]' }
    catch { Log ('  [失败] ' + $_.Exception.Message) }
}
if ($ResetAgent) {
    $ad = 'C:\ProgramData\Battle.net\Agent\Agent.dat'
    if (Test-Path -LiteralPath $ad) {
        $bak = $ad + '.bak.' + (Get-Date -Format 'yyyyMMddHHmmss')
        if ($DryRun) { Log ('  [预览] 备份 Agent.dat -> ' + $bak) }
        else { try { Move-Item -LiteralPath $ad -Destination $bak -Force -ErrorAction Stop; Log ('  [已备份] ' + $bak) } catch { Log '  [失败] Agent.dat 备份' } }
    }
}
Log ''

# ---------- 5. 启动 ----------
Log '=== 5/5 启动战网 / 炉石 / 盒子 ==='
function Start-All {
    # 5.1 战网
    $bn = Get-Process -Name 'Battle.net' -ErrorAction SilentlyContinue
    if ($bn) {
        Log '  战网已在运行，等窗口就绪...'
        $w = Wait-BnWindow 60
        if ($w) { Log ('  战网窗口就绪 PID=' + $w.Id) } else { Log '  战网窗口 60 秒未出现，继续尝试' }
    }
    else {
        if (-not $cfg.BattleNetExe) { Log '  [失败] 未找到战网安装路径，请在 config.ini 里填写 BattleNetExe'; return $false }
        Log ('  启动战网: ' + $cfg.BattleNetExe)
        Start-Process -FilePath $cfg.BattleNetExe | Out-Null
        $w = Wait-BnWindow 120
        if ($w) { Log ('  战网窗口就绪 PID=' + $w.Id) } else { Log '  战网窗口 120 秒未出现（可能需要登录或首次初始化）' }
    }
    # 5.1.5 等后台 Agent 就绪
    $agentOk = $false
    for ($i = 0; $i -lt 90; $i++) {
        if (Get-Process -Name 'Agent' -ErrorAction SilentlyContinue) { $agentOk = $true; break }
        Start-Sleep -Seconds 1
    }
    if ($agentOk) { Log '  战网服务 Agent 已就绪' } else { Log '  [警告] 战网 Agent 未就绪，拉起游戏可能失败' }
    # 5.2 炉石：先在战网界面点「进入游戏」
    Kill-Game
    Start-Sleep -Seconds 2
    Log '  请求战网启动炉石...'
    $launched = Test-GameAlive
    if ($launched) { Log '  游戏已在运行' }
    if (-not $launched) {
        if ($script:BypassBn) {
            Log '  [降级] 战网此时必然一直显示"正在游戏中"，跳过界面点击，直接启动游戏主程序'
        } else {
            Log '  方式1: 识别战网界面并点击「进入游戏」'
            if (Invoke-BnEnterGame 25 3) { $launched = $true; Log '  [成功] 点击「进入游戏」后游戏已启动' }
            else { Log '  点击未成功，改用兜底方式' }
        }
    }
    if (-not $launched) {
        $exeTries = @(
            @{ Kind = 'exe'; Value = $cfg.GameExe;         Desc = '直接启动游戏主程序（绕过战网）' },
            @{ Kind = 'exe'; Value = $cfg.GameLauncherExe; Desc = '炉石独立启动器' }
        )
        $bnTries = @(
            @{ Kind = 'arg'; Value = ('--uid=' + $cfg.GameUid);          Desc = ('战网 --uid=' + $cfg.GameUid) },
            @{ Kind = 'uri'; Value = ('battlenet://' + $cfg.GameUid);    Desc = ('战网 URI battlenet://' + $cfg.GameUid) },
            @{ Kind = 'arg'; Value = ('--uid=' + $cfg.GameUidAlt);       Desc = ('战网 --uid=' + $cfg.GameUidAlt) },
            @{ Kind = 'uri'; Value = ('battlenet://' + $cfg.GameUidAlt); Desc = ('战网 URI battlenet://' + $cfg.GameUidAlt) }
        )
        # 降级模式下把"直接启动游戏"提到最前：战网已经被空壳挡住，先试不经过战网的路
        if ($script:BypassBn) { $tries = @($exeTries) + @($bnTries) } else { $tries = @($bnTries) + @($exeTries) }
        foreach ($t in $tries) {
            if (Test-GameAlive) { $launched = $true; break }
            Log ('  兜底尝试: ' + $t.Desc)
            try {
                switch ($t.Kind) {
                    'arg' { Start-Process -FilePath $cfg.BattleNetExe -ArgumentList $t.Value -ErrorAction Stop }
                    'uri' { Start-Process -FilePath $t.Value -ErrorAction Stop }
                    'exe' {
                        if (-not $t.Value -or -not (Test-Path -LiteralPath $t.Value)) { Log '  [跳过] 找不到游戏主程序'; continue }
                        Start-Process -FilePath $t.Value -WorkingDirectory (Split-Path $t.Value) -ErrorAction Stop | Out-Null
                    }
                }
            } catch { Log ('  [失败] ' + $_.Exception.Message); continue }
            for ($i = 0; $i -lt 8; $i++) {
                Start-Sleep -Seconds 1
                if (Test-GameAlive) { $launched = $true; break }
            }
            if ($launched) { Log ('  [成功] 已由"' + $t.Desc + '"拉起'); break }
        }
    }
    if (-not $launched) {
        Log '  [警告] 所有启动方式都未拉起游戏'
        Log '         请确认：战网已登录、停在炉石页面，且没有别的窗口挡住按钮。'
        try {
            $bnWin = Get-Process -Name 'Battle.net' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
            if ($bnWin) {
                Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
                [Microsoft.VisualBasic.Interaction]::AppActivate($bnWin.Id) | Out-Null
                Log '  已把战网窗口切到前台，可在战网点一次「进入游戏」'
            }
        } catch {}
    }
    # 5.3 等待游戏窗口
    Log ('  等待游戏窗口（最多 {0} 秒）...' -f $LaunchTimeout)
    $hs = Wait-Window 'Hearthstone' $LaunchTimeout
    if (-not $hs) {
        if (Test-GameAlive) { Log '  [部分成功] 游戏进程已运行，但窗口尚未就绪（可能仍在加载）' }
        else {
            Log '  [失败] 游戏窗口未出现'
            Show-CrashEvidence
            return $false
        }
    }
    if ($hs) {
        Start-Sleep -Seconds 6
        if (-not (Test-GameAlive)) {
            Log '  [失败] 游戏窗口出现后立刻退出了（多半是战网没完成授权）'
            Show-CrashEvidence
            return $false
        }
        Log ('  [成功] 游戏窗口就绪 PID={0} 标题"{1}"' -f $hs.Id, $hs.MainWindowTitle)
    }
    # 5.4 盒子
    Start-Sleep -Seconds 3
    $box = Get-Process -Name 'HSAng' -ErrorAction SilentlyContinue
    if ($box) { Log ('  盒子已在运行 PID=' + $box.Id) }
    elseif ($cfg.BoxExe -and (Test-Path -LiteralPath $cfg.BoxExe)) {
        Log '  启动炉石盒子...'
        try {
            Start-Process -FilePath $cfg.BoxExe -WorkingDirectory (Split-Path $cfg.BoxExe) | Out-Null
            Start-Sleep -Seconds 3
            $b = Get-Process -Name 'HSAng' -ErrorAction SilentlyContinue
            if ($b) { Log ('  [成功] 盒子已启动 PID=' + $b.Id) } else { Log '  [警告] 盒子进程未出现' }
        }
        catch { Log ('  [失败] 盒子启动: ' + $_.Exception.Message) }
    } else { Log '  [跳过] 未找到盒子程序' }
    return $true
}

if ($NoLaunch) { Log '  已跳过（-NoLaunch）'; Finish 0 }
elseif ($DryRun) {
    Log ('  [预览] 启动战网: ' + $cfg.BattleNetExe)
    Log ('  [预览] 识别并点击「进入游戏」，失败再用 --uid / battlenet:// / 直接启动兜底')
    Log ('  [预览] 等待窗口 ' + $LaunchTimeout + ' 秒')
    Log ('  [预览] 启动盒子: ' + $cfg.BoxExe)
    Finish 0
} else {
    $res = Start-All
    if (-not $res -and $Retry) {
        Log ''
        Log '=== 自动重试 ==='
        Kill-Game
        Start-Sleep -Seconds 3
        $res = Start-All
    }
    if ($res) { Log ''; Log '=== 全部完成：游戏与盒子已就绪 ==='; Finish 0 }
    else {
        Log ''; Log '=== 启动未成功，请查看上方诊断信息 ==='
        Show-CrashEvidence
        if ($script:BypassBn) {
            Log ''
            Log '  绕过战网直接启动也没能进游戏，剩下只有重启能清掉那两个空壳。'
            Log '  （如果你看到的是"请通过战网启动"之类提示，说明这条路对当前版本不成立，请告诉我）'
            $rb = Invoke-RebootAsk 8
            if ($rb -eq 'ok') { Log '  系统即将重启，重启后重新双击「一键修复」即可全自动。'; Finish 0 }
            elseif ($rb -eq 'cancel') { Log ''; Log '=== 已取消：未重启 ==='; Finish 2 }
            else { Log ''; Log '=== 自动重启失败：请手动重启电脑后重试 ==='; Finish 3 }
        }
        Finish 1
    }
}
