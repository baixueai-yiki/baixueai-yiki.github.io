# =========「晴小姐启动器与守护程序」=============
# 自动启动并监视 晴小姐.exe 或 qing.ps1
# 如果主程序意外退出，自动重新启动
# 兼容 exe 打包，控制台隐藏，互斥体保护，完全静默

# -----------------------
# 获取脚本目录（兼容 exe 和 ps1）
# -----------------------
if ($PSScriptRoot) { 
    $scriptDir = $PSScriptRoot 
} else { 
    $scriptDir = [System.AppDomain]::CurrentDomain.BaseDirectory
}

# 判断管理员权限 - 判断当前进程是否为管理员权限。
function Test-IsAdmin {
    try {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

# =========「输出风格模块」=============
# 可爱输出 - 统一控制台输出风格（前缀 + 柔和色）。
function Show-ToastMessage {
    param([string]$Text)
    try {
        if ([System.Windows.Forms.MessageBox] -ne $null) {
            [System.Windows.Forms.MessageBox]::Show($Text, "芝芝", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } else {
            Write-Host $Text -ForegroundColor Magenta
        }
    } catch {
        Write-Host $Text -ForegroundColor Magenta
    }
}

function Write-CuteHost {
    param(
        [Parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Object,
        [ConsoleColor]$ForegroundColor,
        [switch]$NoNewline
    )
    $text = ($Object | ForEach-Object { $_ }) -join " "
    $prefix = "【晴】♡ "
    $showText = $prefix + $text
    if ($NoNewline) {
        Write-Host $showText -NoNewline
        return
    }
    Show-ToastMessage -Text $showText
}

# -----------------------
# 设置控制台图标（ht.ico）
# -----------------------
$iconPath = Join-Path $scriptDir "ht.ico"
if (Test-Path $iconPath) {
    Add-Type -Namespace Win32 -Name NativeMethods -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern System.IntPtr GetConsoleWindow();
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)]
public static extern System.IntPtr SendMessage(System.IntPtr hWnd, int msg, System.IntPtr wParam, System.IntPtr lParam);
"@
    Add-Type -AssemblyName System.Drawing
    $hIcon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath).Handle
    $hWnd = [Win32.NativeMethods]::GetConsoleWindow()
    if ($hWnd -ne [IntPtr]::Zero -and $hIcon -ne [IntPtr]::Zero) {
        $WM_SETICON = 0x80
        [Win32.NativeMethods]::SendMessage($hWnd, $WM_SETICON, [IntPtr]1, $hIcon) | Out-Null
        [Win32.NativeMethods]::SendMessage($hWnd, $WM_SETICON, [IntPtr]0, $hIcon) | Out-Null
    }
}


# -----------------------
# 互斥体保护，防止重复启动
# -----------------------
$mutex = New-Object System.Threading.Mutex($false, "QingMissSystem.Guard")
if (-not $mutex.WaitOne(0, $false)) { exit }

# -----------------------
# 主程序路径（仅 exe）
# -----------------------
$exePath = Join-Path $scriptDir "晴小姐.exe"

if (Test-Path $exePath) {
    $targetPath = $exePath
} else {
    exit
}

# -----------------------
# 计划任务名称（主程序）
# -----------------------
$appTaskName = "守护晴小姐"
$restartAudioDir = Join-Path $scriptDir "sounds"
$global:AudioPath = $restartAudioDir
$restartAudioFiles = @("run_zhizhi1.wav","run_zhizhi2.wav","run_zhizhi3.wav")

# -----------------------
# 注册主程序计划任务 - 注册 晴小姐.exe 的计划任务（登录时启动）
# -----------------------
function Register-AppTask {
    param([string]$TaskName, [string]$ExePath)
    try {
        if (-not (Test-IsAdmin)) { return $false }
        if (-not (Test-Path $ExePath)) { return $false }
        $taskCmd = "`"$ExePath`""
        & schtasks /Create /F /SC ONLOGON /TN $TaskName /TR $taskCmd | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Write-CuteHost "主程序计划任务创建失败（返回码 $LASTEXITCODE）" -ForegroundColor Yellow
            return $false
        }
        return $true
    } catch {
        Write-CuteHost "主程序计划任务创建失败: $_" -ForegroundColor Yellow
        return $false
    }
}

# 启动主程序计划任务 - 触发计划任务立即运行
function Start-AppTask {
    param([string]$TaskName)
    try {
        if (-not (Test-IsAdmin)) { return $false }
        & schtasks /Query /TN $TaskName | Out-Null
        if ($LASTEXITCODE -ne 0) { return $false }
        & schtasks /Run /TN $TaskName | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Write-CuteHost "主程序计划任务启动失败（返回码 $LASTEXITCODE）" -ForegroundColor Yellow
            return $false
        }
        # 删除任务（已禁用）：保持任务长期存在，避免下次无法触发
        # Start-Sleep -Seconds 60
        # & schtasks /Delete /TN $TaskName /F | Out-Null
        return $true
    } catch {
        Write-CuteHost "主程序计划任务启动失败: $_" -ForegroundColor Yellow
        return $false
    }
}

# 播放音频 - 播放音频文件，支持同步/异步。
function Invoke-Audio {
    param([string]$FileName, [switch]$Async)
    try {
        $fullPath = Join-Path $global:AudioPath $FileName
        
        if (-not (Test-Path $fullPath)) {
            Write-CuteHost "音频文件不存在: $fullPath" -ForegroundColor Yellow
            return $null
        }
        
        $player = New-Object System.Media.SoundPlayer
        $player.SoundLocation = $fullPath
        
        if ($Async) {
            $player.Play()
        } else {
            $player.PlaySync()
        }

        return $player
    } catch {
        Write-CuteHost "播放音频时出错: $_" -ForegroundColor Red
        return $null
    }
}

# 播放自动重启语音 - 仅在触发自动拉起时播放
function Invoke-RestartAudio {
    param([string]$AudioDir, [string[]]$AudioFiles)
    try {
        if ($null -eq $AudioFiles -or $AudioFiles.Count -eq 0) { return }
        $existing = @()
        foreach ($f in $AudioFiles) {
            $p = Join-Path $AudioDir $f
            if (Test-Path $p) { $existing += $p }
        }
        if ($existing.Count -eq 0) {
            if (-not $global:RestartAudioWarned) {
                Write-CuteHost "提示：未找到重启语音文件，请检查 sounds 目录。" -ForegroundColor Yellow
                $global:RestartAudioWarned = $true
            }
            return
        }
        $audioPath = Get-Random -InputObject $existing
        $fileName = [System.IO.Path]::GetFileName($audioPath)
        $global:AudioPlayer = Invoke-Audio -FileName $fileName -Async
    } catch {}
}

# -----------------------
# 注册并启动主程序计划任务
# -----------------------
 $appTaskReady = $false
 $appTaskErrorShown = $false
try { $appTaskReady = Register-AppTask -TaskName $appTaskName -ExePath $targetPath } catch {}
try {
    $proc = Get-Process -ErrorAction SilentlyContinue | Where-Object { 
        $_.Path -eq $targetPath 
    }
    if (-not $proc) {
        if ($appTaskReady) {
            Start-AppTask -TaskName $appTaskName | Out-Null
        } elseif (-not $appTaskErrorShown) {
            Write-CuteHost "提示：主程序计划任务不可用（可能缺少管理员权限）。" -ForegroundColor Yellow
            $appTaskErrorShown = $true
        }
    }
} catch {}

while ($true) {
    try {
        # 每秒检测主程序，不在则通过计划任务拉起
        $proc = Get-Process -ErrorAction SilentlyContinue | Where-Object { 
            $_.Path -eq $targetPath 
        }

        if (-not $proc) {
            Invoke-RestartAudio -AudioDir $restartAudioDir -AudioFiles $restartAudioFiles
            if ($appTaskReady) {
                Start-AppTask -TaskName $appTaskName | Out-Null
            } elseif (-not $appTaskErrorShown) {
                Write-CuteHost "提示：主程序计划任务不可用（可能缺少管理员权限）。" -ForegroundColor Yellow
                $appTaskErrorShown = $true
            }
        }

        Start-Sleep -Seconds 1
    } catch {
        # 异常时休息 10 秒再重试，防止循环报错占用 CPU
        Start-Sleep -Seconds 10
    }
}
