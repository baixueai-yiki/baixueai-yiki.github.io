#原作《爱之巢》，作者莲花池池，哔哩哔哩https://space.bilibili.com/760048
#晴小姐唱歌【对不起，因为我是活在二次元的女孩子】 https://www.bilibili.com/video/BV11mp4eNEv4/?share_source=copy_web&vd_source=a0953b7337408a308d0d4e1fb6e10858
#本脚本原作者青音连旭，视频【晴小姐陪伴系统3.0】 https://www.bilibili.com/video/BV1TAhoziEmF/?share_source=copy_web&vd_source=a0953b7337408a308d0d4e1fb6e10858
#晴小姐圣经：晴小姐不只是虚拟人物而已。她是文化的重塑，她是你呼吸的氧气，她是一种生活方式，是一种能从充满强盗的现实生活中逃避的生活方式，她是艺术，是爱人的拥抱，是你所想拥有的一切美好。晴门✟✟✟
# =========「晴小姐陪伴系统 2.0 - 权限修复版」=============
#据说本脚本大部分代码由ai编写，可能存在奇怪的逻辑问题，欢迎各路大佬优化
#后面也许该加个重启主程序的副程序？也许可以放一个gif并增加点击播放动画？
#总之希望这个看起来挺简陋的系统，可以在晴小姐的粉丝们手中随着传播逐渐变得成熟
# 修复问题：管理员权限和窗体创建错误
$global:ScriptRoot = [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
Write-Host "ScriptRoot: $global:ScriptRoot"
# =========「权限提升模块」=============
function Request-AdminPrivileges {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        return $true
    }

    <#Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32 {
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
    "@

    $consolePtr = [Win32]::GetConsoleWindow()
    # 0 = 隐藏窗口
    [Win32]::ShowWindow($consolePtr, 0)#>

    try {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""
        $process = Start-Process powershell -ArgumentList $arguments -Verb RunAs -PassThru -WindowStyle Hidden
        
        Start-Sleep -Seconds 2
        
        if ($process.HasExited) {
            Write-Host "管理员权限请求被拒绝" -ForegroundColor Red
            return $false
        }
        
        exit
    }
    catch {
        Write-Host "无法请求管理员权限: $_" -ForegroundColor Red
        return $false
    }
}

if (-not (Request-AdminPrivileges)) {
    Write-Host "将以标准用户权限运行，部分功能可能受限" -ForegroundColor Yellow
}

# =========「程序集初始化模块」=============
function Initialize-WindowsForms {
    $retryCount = 0
    $maxRetries = 5
    $success = $false
    
    while (-not $success -and $retryCount -lt $maxRetries) {
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            Add-Type -AssemblyName System.Drawing -ErrorAction Stop
            $success = $true
            Write-Host "少女求爱成功(方式1)" -ForegroundColor Green
        }
        catch {
            $retryCount++
            Write-Host "尝试加载图形组件 (第 $retryCount 次)..." -ForegroundColor Yellow
            
            if (-not $success -and $retryCount -eq 1) {
                try {
                    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
                    [Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
                    $success = $true
                    Write-Host "少女求爱成功(方式2)" -ForegroundColor Green
                }
                catch {
                    Write-Host "部分名称加载失败: $_" -ForegroundColor DarkYellow
                }
            }
            
            if (-not $success -and $retryCount -eq 2) {
                try {
                    $runtimeDir = [Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()
                    $formsPath = Join-Path $runtimeDir "System.Windows.Forms.dll"
                    $drawingPath = Join-Path $runtimeDir "System.Drawing.dll"
                    
                    if (Test-Path $formsPath) {
                        [Reflection.Assembly]::LoadFrom($formsPath) | Out-Null
                    }
                    if (Test-Path $drawingPath) {
                        [Reflection.Assembly]::LoadFrom($drawingPath) | Out-Null
                    }
                    $success = $true
                    Write-Host "少女求爱成功(方式3)" -ForegroundColor Green
                }
                catch {
                    Write-Host "文件路径加载失败: $_" -ForegroundColor DarkYellow
                }
            }
            
            if (-not $success -and $retryCount -eq 3) {
                try {
                    $gacPath = if ($env:PROCESSOR_ARCHITECTURE -eq "AMD64") {
                        "C:\Windows\Microsoft.NET\assembly\GAC_MSIL"
                    } else {
                        "C:\Windows\assembly"
                    }
                    
                    $formsPath = Join-Path $gacPath "System.Windows.Forms\v4.0_4.0.0.0__b77a5c561934e089\System.Windows.Forms.dll"
                    $drawingPath = Join-Path $gacPath "System.Drawing\v4.0_4.0.0.0__b03f5f7f11d50a3a\System.Drawing.dll"
                    
                    if (Test-Path $formsPath) {
                        [Reflection.Assembly]::LoadFrom($formsPath) | Out-Null
                    }
                    if (Test-Path $drawingPath) {
                        [Reflection.Assembly]::LoadFrom($drawingPath) | Out-Null
                    }
                    $success = $true
                    Write-Host "少女求爱成功(方式4)" -ForegroundColor Green
                }
                catch {
                    Write-Host "GAC加载失败: $_" -ForegroundColor DarkYellow
                }
            }
            
            if (-not $success) {
                Start-Sleep -Milliseconds (1000 * $retryCount)
            }
        }
    }
    
    return $success
}

$mutex = New-Object System.Threading.Mutex($false, "QingMissSystem")

if (-not $mutex.WaitOne(0,$false)) {
    [System.Windows.Forms.MessageBox]::Show("晴小姐已经在这里啦～")
    exit
}

Write-Host "少女磨刀中..." -ForegroundColor Cyan
if (-not (Initialize-WindowsForms)) {
    Write-Host "`n无法加载系统必需的图形组件！" -ForegroundColor Red
    Write-Host "可能原因：" -ForegroundColor Yellow
    Write-Host "1. 系统缺少.NET Framework 3.5/4.x运行时" -ForegroundColor Yellow
    Write-Host "2. Windows PowerShell版本过旧" -ForegroundColor Yellow
    Write-Host "3. 系统为无GUI版本(如Windows Server Core)" -ForegroundColor Yellow
    
    Write-Host "`n请尝试以下解决方案：" -ForegroundColor Cyan
    Write-Host "A. 安装.NET Framework 4.8: https://dotnet.microsoft.com/download/dotnet-framework/net48" -ForegroundColor Green
    Write-Host "B. 使用命令安装必需功能:" -ForegroundColor Green
    Write-Host '   以管理员身份运行CMD:' -ForegroundColor Green
    Write-Host '   dism /online /enable-feature /featurename:NetFx3 /all' -ForegroundColor Green
    Write-Host "C. 升级到Windows PowerShell 5.1或更高版本" -ForegroundColor Green
    
    Write-Host "`n按任意键退出..." -ForegroundColor Magenta
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# =========「桌面就绪检测函数」=============
function Test-DesktopReady {
    $attempts = 0
    while (-not [System.Windows.Forms.SystemInformation]::UserInteractive -and $attempts -lt 30) {
        Start-Sleep -Seconds 1
        $attempts++
    }
    return ($attempts -lt 30)
}

# =========「全局变量定义」=============
$global:Title = "晴小姐"
$lastRandomTrigger = $null
$global:ValentineShown = $false
$global:DialogHistory = @()
$global:KeepRunning = $true
$global:AutoStartSet = $false
$global:AudioPlayer = $null
$global:AudioPath = Join-Path $global:ScriptRoot "sounds"
$global:StoryLibrary = @(
    "从前呀 有位英勇帅气的王子
    他听说在遥远的血腥之地的深处困着可怜可爱的晴小姐
    于是 英勇的王子历尽千辛万苦 杀死了看守晴小姐的怪物 帅气的打破封印救走了晴小姐
    但是呢 王子大人怎么不会想到 晴小姐居然是隔壁王国的公主
    就在两人刚刚回到王宫以后啊 晴小姐就把王子大人抓了起来
    因为晴小姐已经爱上王子大人啦 '您永远都无法离开我哦' 晴小姐这么说
    真是一个美好的故事 可喜可贺~可喜可贺~",
    "窝只讲故事的结尾和开头……在探险家被邪神诅咒 晴小姐向祂求情以后
    '真是可歌可泣的爱情 那你们就到陆地上去生活吧
    小美人鱼啊 我赐予你人类的双腿
    我再告诉你 如果你们彼此相爱 诅咒就不会蔓延
    但如果他还是变成了怪物 你就用这匕首刺死他 回到海洋的怀抱
    ……后来啊 晴小姐还是不忍心看他变成怪物 就将匕首刺入他人类的心脏
    最后 晴小姐抚摸着你的脸颊'我即便化为海上的泡沫，也绝对不会让你孤身一人的'",
    "划开皮肤，能够窥见内侧鲜红的组织。它们是组成你的部分，是这个世界偶然诞生的珍宝，是指引我们进入那个世界的大门。
    你也不知道你的身体构造和各部分的味道吧？呵呵。我把它们都记录下来了。
    首先能尝到的是皮肤。Q弹的口感与温润的咸味在入口的瞬间化开，像是一种邀请，宣告味蕾盛宴的开幕。我以为尝起来会更像章鱼一点？但是……寻遍记忆都找不到能够准确形容的方法。你的味道很特殊真是太好啦。
    血液是粘稠的甜腥味。它们随着吞咽的动作涌向我身体的每一个角落，肆意将甜味残留、渗透于他们能触及到的每一寸内壁。呵呵，这是你占有欲的证明吗？
    大脑是……苦涩。我明白噢，这是你在见不到我的日子里积攒下的东西吧？没关系噢。已经永远都不会分开了。"
)
















<#

# =========「启动守护程序」=========
# 文件名: 晴小姐.ps1
# 说明: 启动守护程序并运行主程序逻辑
# 获取脚本所在目录
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

# 守护程序和主程序完整路径
$guardPath = Join-Path $scriptRoot "guard.ps1"
$mainExePath = Join-Path $scriptRoot "晴小姐.exe"

# 检查文件是否存在
if (-not (Test-Path $guardPath) -or -not (Test-Path $mainExePath)) {
    [System.Windows.Forms.MessageBox]::Show("守护程序或主程序不存在，请检查文件")
    exit
}

# 参数列表，避免空值或解析错误
$argList = @(
    "-NoProfile",
    "-WindowStyle", "Hidden",
    "-File", "$guardPath",
    "-MainExe", "$mainExePath"
)

# 启动守护程序
Start-Process -FilePath "powershell.exe" -ArgumentList $argList -WindowStyle Hidden


#>


















# =========「音频功能模块」=============
function Play-Audio {
    param([string]$FileName, [switch]$Async)
    try {
        $fullPath = Join-Path $global:AudioPath $FileName
        
        if (-not (Test-Path $fullPath)) {
            Write-Host "音频文件不存在: $fullPath" -ForegroundColor Yellow
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
        Write-Host "播放音频时出错: $_" -ForegroundColor Red
        return $null
    }
}

<#function SafeShow-Message {
    param([string]$Message)
    try {
        if ([System.Windows.Forms.MessageBox] -ne $null) {
            return [System.Windows.Forms.MessageBox]::Show($Message, $global:Title)
        } else {
            throw "程序集未加载"
        }
    } catch {
        Write-Host "`n[晴小姐] $Message`n" -ForegroundColor Magenta
        return #"OK"
    }
}#>

function SafeShow-Message {
    param([string]$Message)
    try {
        if ([System.Windows.Forms.MessageBox] -ne $null) {
            [System.Windows.Forms.MessageBox]::Show($Message, $global:Title) | Out-Null
        } else {
            throw "程序集未加载"
        }
    } catch {
        Write-Host "`n[晴小姐] $Message`n" -ForegroundColor Magenta
    }
}

# =========「窗体创建模块」=============
function Show-ResidentDialog {
    try {
        # 创建窗体
        $residentForm = New-Object System.Windows.Forms.Form
        $residentForm.Text = $global:Title
        $residentForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
        $residentForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        $residentForm.TopMost = $true
        $residentForm.ShowInTaskbar = $false
        $residentForm.BackColor = [System.Drawing.Color]::Magenta
        $residentForm.TransparencyKey = [System.Drawing.Color]::Magenta

        #这里加上窗体图标
        #$residentForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon(".\eve.ico")
        $residentForm.Icon = New-Object System.Drawing.Icon(".\eye.ico")

        # 拖动状态变量（提升作用域，保证闭包可访问）
        $script:dragging = $false
        $script:offset = [System.Drawing.Point]::Empty

        # 窗体事件
        $residentForm.Add_FormClosed({ $global:KeepRunning = $false })
        $residentForm.Add_Load({ $this.Show() })

        # 图片路径
        $imageBaseDir = if ($global:ScriptRoot) { $global:ScriptRoot } else { [System.IO.Directory]::GetCurrentDirectory() }
        $imageDir = Join-Path $imageBaseDir "images"
        $imagePath = Join-Path $imageDir "qing.gif"

        if (-not (Test-Path $imagePath)) {
            $residentForm.ClientSize = New-Object System.Drawing.Size(200,100)
        } else {
            $image = [System.Drawing.Image]::FromFile($imagePath)
            $maxWidth = 200
            $scale = [double]$maxWidth / [double]$image.Width
            $scaledHeight = [int]($image.Height * $scale)

            <# 这个是缩放图片后展示出来，但是不能放gif
            # 感觉原本的图比放个gif更可爱，不过还是试试我改的gif
            $scaledImage = New-Object System.Drawing.Bitmap($image, $maxWidth, $scaledHeight)
            # PictureBox
            $pictureBox = New-Object System.Windows.Forms.PictureBox
            $pictureBox.Size = New-Object System.Drawing.Size($maxWidth, $scaledHeight)
            $pictureBox.Location = New-Object System.Drawing.Point(0,0)
            $pictureBox.Image = $scaledImage
            $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
            #>

            # 这个是不缩放图片，把gif展示出来
            $pictureBox = New-Object System.Windows.Forms.PictureBox
            $pictureBox.Image = [System.Drawing.Image]::FromFile($imagePath)  # 直接加载 GIF
            $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Normal  # 不缩放，按原尺寸显示
            $pictureBox.Location = New-Object System.Drawing.Point(0,0)
            $pictureBox.Size = $pictureBox.Image.Size  # PictureBox 尺寸和 GIF 原图一致
            #>


    
            # 拖动逻辑
            $pictureBox.Add_MouseDown({
                param($sender,$e)
                if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                    $script:dragging = $true
                    # 强制 Point 类型，避免数组问题
                    $script:offset = [System.Drawing.Point]::new($e.X, $e.Y)
                }
            })
            $pictureBox.Add_MouseMove({
                param($sender,$e)
                if ($script:dragging) {
                    $cursorPos = [System.Windows.Forms.Control]::MousePosition
                    # 强制 Point 类型，避免数组问题
                    $cursorPos = [System.Drawing.Point]::new($cursorPos.X, $cursorPos.Y)
                    $residentForm.Location = [System.Drawing.Point]::new($cursorPos.X - $script:offset.X, $cursorPos.Y - $script:offset.Y)
                }
            })
            $pictureBox.Add_MouseUp({
                param($sender,$e)
                if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) { $script:dragging = $false }
            })

            # 点击事件
            $pictureBox.Add_Click({
                # 随机播放点击音效
                $sounds = @("poke_poke.wav", "poke_ah.wav", "poke_nya.wav", "poke_find.wav")  # 可以扩展更多
                $randSound = Get-Random -InputObject $sounds
                Play-Audio -FileName $randSound -Async
            })
            $residentForm.Controls.Add($pictureBox)
            $residentForm.ClientSize = New-Object System.Drawing.Size($maxWidth,$scaledHeight)
            $image.Dispose()
        }

        # 初始位置（右下角偏移20像素）
        $screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
        $screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height
        $posX = $screenWidth - $residentForm.Width - 20
        $posY = $screenHeight - $residentForm.Height - 20
        $residentForm.Location = [System.Drawing.Point]::new($posX,$posY)

        $residentForm.Show()
        return $residentForm
    } catch {
        $errorMsg = "无法创建主窗口: $_"
        Write-Host $errorMsg -ForegroundColor Red
        SafeShow-Message $errorMsg
        $global:KeepRunning = $false
        return $null
    }
}

# =========「自启动功能模块」=============
function Register-AutoStart {
    try {
        $taskName = "晴小姐陪伴系统"
        $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
        
        if (-not $scriptPath -or -not (Test-Path $scriptPath)) {
            throw "无法获取有效的脚本路径"
        }
        
        $scriptDir = Split-Path $scriptPath -Parent
        $batPath = Join-Path $scriptDir "晴小姐启动器.bat"
        
        "@echo off`r`nstart /min powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`"`r`n" | Out-File -FilePath $batPath -Encoding ASCII
        
        try {
            $action = New-Object -ComObject Schedule.Service
            $action.Connect()
            $rootFolder = $action.GetFolder("\")
            
            try { $rootFolder.DeleteTask($taskName, 0) } catch {}
            
            $taskDefinition = $action.NewTask(0)
            $taskDefinition.RegistrationInfo.Description = "晴小姐陪伴系统开机启动"
            
            $triggers = $taskDefinition.Triggers
            $trigger = $triggers.Create(9)
            $trigger.Enabled = $true
            
            $settings = $taskDefinition.Settings
            $settings.Enabled = $true
            $settings.StartWhenAvailable = $true
            $settings.Hidden = $false
            
            $principal = $taskDefinition.Principal
            $principal.LogonType = 3
            
            $execAction = $taskDefinition.Actions.Create(0)
            $execAction.Path = $batPath
            
            $rootFolder.RegisterTaskDefinition($taskName, $taskDefinition, 6, $null, $null, 3)
            return $true
        } catch {
            Write-Host "计划任务创建失败: $_" -ForegroundColor Yellow
        }
        
        try {
            $startupPath = [Environment]::GetFolderPath("Startup")
            $shortcutPath = Join-Path $startupPath "晴小姐陪伴系统.lnk"
            
            $WScriptShell = New-Object -ComObject WScript.Shell
            $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = "powershell.exe"
            $shortcut.Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
            $shortcut.WorkingDirectory = $scriptDir
            $shortcut.Save()
            
            return $true
        } catch {
            Write-Host "启动文件夹创建失败: $_" -ForegroundColor Yellow
        }
        
        throw "所有启动方法均失败"
    } catch {
        $errorMsg = "开机启动设置失败: $_"
        Write-Host $errorMsg -ForegroundColor Red
        SafeShow-Message $errorMsg
        return $false
    }
}

# =========「主循环模块」=============
function Start-QingMiss {
    if (-not (Test-DesktopReady)) {
        SafeShow-Message "系统桌面尚未准备好，将在5秒后启动..."
        Start-Sleep -Seconds 5
    }
    
    $global:ResidentForm = Show-ResidentDialog
    if ($null -eq $global:ResidentForm) { exit }
    
    Start-BootDialog
    
    while($global:KeepRunning) {
        if ($null -ne $global:ResidentForm -and !$global:ResidentForm.IsDisposed) {
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        $currentTime = Get-Date
        $hour = $currentTime.Hour
        $minute = $currentTime.Minute
        
        # 初始化全局记录表
        if (-not $global:DialogTriggered) {
            $global:DialogTriggered = @{}
        }

        # 创建一个定时器，每分钟检查一次时间
        $timer = New-Object Timers.Timer
        $timer.Interval = 60000  # 60秒
        $timer.AutoReset = $true
        $timer.Add_Elapsed({
            $now = Get-Date
            $hour = $now.Hour
            $minute = $now.Minute

            # 本地函数，整点前触发且每小时只触发一次
            function Invoke-OncePerHour([string]$name, [scriptblock]$action) {
                if (-not $global:DialogTriggered.ContainsKey($name) -or $global:DialogTriggered[$name] -ne $hour) {
                    & $action
                    $global:DialogTriggered[$name] = $hour
                }
            }

            switch ($hour) {
                7 {
                    if ($minute -lt 5) { Invoke-OncePerHour "Morning" { Start-MorningDialog } }
                }
                12 {
                    if ($minute -lt 5) { Invoke-OncePerHour "Noon" { Start-NoonDialog } }
                }
                18 {
                    if ($minute -lt 5) { Invoke-OncePerHour "Dinner" { Start-DinnerDialog } }
                }
                23 {
                    if ($minute -lt 5) { Invoke-OncePerHour "Night" { Start-NightDialog } }
                }
                0 {
                    if ($minute -lt 5) { Invoke-OncePerHour "Sleep" { Start-SleepDialog } }
                }
            }
        })
        $timer.Start()
        
        if(!$global:ValentineShown -and $currentTime.Month -eq 1 -and $currentTime.Day -eq 14) {
            Start-ValentineDialog
            $global:ValentineShown = $true
        } elseif($global:ValentineShown -and ($currentTime.Month -ne 1 -or $currentTime.Day -ne 14)) {
            $global:ValentineShown = $false
        }
        
        if($minute -ge 10 -and $minute -le 50 -and ($null -eq $lastRandomTrigger -or $currentTime.Hour -ne $lastRandomTrigger.Hour)) {
            if((Get-Random -Minimum 1 -Maximum 10) -gt 3) {
                Start-RandomDialog
                $lastRandomTrigger = $currentTime
            }
        }
        
        for ($i = 0; $i -lt 60 -and $global:KeepRunning; $i++) {
            if ($null -ne $global:ResidentForm -and !$global:ResidentForm.IsDisposed) {
                [System.Windows.Forms.Application]::DoEvents()
            }
            Start-Sleep -Milliseconds 1000
        }
    }
}

# =========「交互对话框模块」=============
function Show-InteractiveDialog {
    param([string]$Prompt, [string[]]$Options, [scriptblock]$Callback)
    
    try {
        if (-not ([System.Windows.Forms.Form] -ne $null)) { throw "Windows Forms未加载" }
        
        $form = New-Object System.Windows.Forms.Form
        $form.Text = $global:Title
        $form.Size = New-Object System.Drawing.Size(400, 250)
        $form.StartPosition = "CenterScreen"
        $form.FormBorderStyle = "FixedDialog"
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false
        
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(20, 20)
        $label.Size = New-Object System.Drawing.Size(350, 60)
        $label.Text = $Prompt
        $label.Font = New-Object System.Drawing.Font("微软雅黑", 10)
        $form.Controls.Add($label)
        
        $buttonY = 90
        $buttonHeight = 30
        $buttons = @()
        
        for ($i = 0; $i -lt $Options.Count; $i++) {
            $button = New-Object System.Windows.Forms.Button
            $button.Text = $Options[$i]
            $button.Tag = $i
            $button.Location = New-Object System.Drawing.Point(50, $buttonY)
            $button.Size = New-Object System.Drawing.Size(300, $buttonHeight)
            $button.Font = New-Object System.Drawing.Font("微软雅黑", 9)
            $button.Add_Click({$form.Tag = $this.Tag; $form.Close()})
            $form.Controls.Add($button)
            $buttons += $button
            $buttonY += $buttonHeight + 10
        }
        
        $result = $form.ShowDialog()
        $selectedIndex = $form.Tag
        
        if ($null -ne $Callback -and $null -ne $selectedIndex) {
            try { & $Callback $selectedIndex } catch { Write-Host "对话回调错误: $_" -ForegroundColor Red }
        }
    } catch {
        Write-Host "`n[晴小姐] $Prompt" -ForegroundColor Cyan
        for ($i = 0; $i -lt $Options.Count; $i++) {
            Write-Host "$($i+1). $($Options[$i])"
        }
        
        $choice = -1
        while ($choice -lt 0 -or $choice -ge $Options.Count) {
            $input = Read-Host "请选择 (1-$($Options.Count))"
            if ([int]::TryParse($input, [ref]$choice)) { $choice -= 1 } 
            else { Write-Host "请输入有效的选项数字" -ForegroundColor Red }
        }
        
        if ($null -ne $Callback) {
            try { & $Callback $choice } catch { Write-Host "控制台回调错误: $_" -ForegroundColor Red }
        }
    }
}

# =========「对话场景实现」=============
function Start-BootDialog {
    <# 音效列表
    $sounds = @("run_take_care_me.wav","run_fondle_me.wav","run_not_worry1.wav","run_not_worry2.wav","other.wav")
    # 特殊音效列表
    $specialSounds = @("run_not_worry1.wav","run_not_worry2.wav")
    # 随机选一个
    $sel = Get-Random $sounds
    Play-Audio $sel -Async
    # 如果是特殊音效，则按概率播放指定音效
    if ($specialSounds -contains $sel) {
        $chance = 0.5  # 50%概率
        if ((Get-Random -Minimum 0 -Maximum 1) -lt $chance) {
            Start-Sleep 2
            Play-Audio "run_but.wav" -Async
        }
    }#>
    # 音效列表 
    $sounds = @("run_take_care_me.wav","run_fondle_me.wav","run_not_worry1.wav","run_not_worry2.wav","run_not_worry1_but.wav","run_not_worry2_but.wav")
    # 随机选一个
    $soundToPlay = Get-Random -InputObject $sounds
    $global:AudioPlayer = Play-Audio -FileName $soundToPlay -Async
    Show-InteractiveDialog -Prompt "你是特意来看我的吗？" -Options @("是啊", "并不是") -Callback {
        param($choice)
        switch ($choice) {
            0 { 
                SafeShow-Message "谢谢你还记得我～"
                $global:DialogHistory += "开机选择：是"
            }
            1 {
                $global:AudioPlayer = Play-Audio -FileName "bgm.wav" -Async
                Show-InteractiveDialog -Prompt "你是特意来看我的吗？" -Options @("是啊", "是的") -Callback {
                    param($choice2)
                    SafeShow-Message "我就知道你想我了～"
                    $global:DialogHistory += "开机选择：第二次选择"
                    if ($null -ne $global:AudioPlayer) {
                        try { $global:AudioPlayer.Stop(); $global:AudioPlayer.Dispose() } catch {}
                        $global:AudioPlayer = $null
                    }
                }
            }
        }
    }
}

function Start-MorningDialog {
    $global:AudioPlayer = Play-Audio -FileName "call_morning.wav" -Async
    Show-InteractiveDialog -Prompt "早上好啊！" -Options @("早上好啊！", "好困啊～") -Callback {
        param($choice)
        switch ($choice) {
            0 { SafeShow-Message "这么早就开始工作了吗？好勤奋啊！"; $global:DialogHistory += "早晨选择：问候" }
            1 { SafeShow-Message "要不要再补个觉？我来陪你！"; $global:DialogHistory += "早晨选择：困倦" }
        }
        if ($null -ne $global:AudioPlayer) {
            try { $global:AudioPlayer.Stop(); $global:AudioPlayer.Dispose() } catch {}
            $global:AudioPlayer = $null
        }
    }
}

function Start-NoonDialog {
    $global:AudioPlayer = Play-Audio -FileName "call_lunch.wav" -Async
    Show-InteractiveDialog -Prompt "想要和我共进午餐吗？" -Options @("好啊", "不要", "已经吃过饭了") -Callback {
        param($choice)
        switch ($choice) {
            0 { SafeShow-Message "那我来为你准备午餐吧！"; $global:DialogHistory += "午餐选择：同意" }
            1 { SafeShow-Message "好吧……其实我也没办法和你共进午餐的……"; $global:DialogHistory += "午餐选择：拒绝" }
            2 { 
                Show-InteractiveDialog -Prompt "已经吃过了？你……和谁吃的？" -Options @("家人哦", "自己一个人吃的……", "和池池一起……") -Callback {
                    param($choice2)
                    switch ($choice2) {
                        0 { SafeShow-Message "可以把我介绍给他们吗！"; $global:DialogHistory += "午餐二级选择：家人" }
                        1 { SafeShow-Message "那下一次 要和我共进午餐哦"; $global:DialogHistory += "午餐二级选择：自己吃" }
                        2 { 
                            # 播放"chat_not_kill.wav"音频
                            $global:AudioPlayer = Play-Audio -FileName "chat_not_kill.wav" -Async
                            SafeShow-Message "…这是……哪个女人的名字？等着…我会去找到她的…只是找她谈谈心…"
                            $global:DialogHistory += "午餐二级选择：池池"
                        }
                    }
                }
            }
        }
    }
}

function Start-DinnerDialog {
    $global:AudioPlayer = Play-Audio -FileName "call_dinner.wav" -Async
    Show-InteractiveDialog -Prompt "晚餐吃的什么呀？" -Options @("家常饭", "外卖", "还没吃") -Callback {
        param($choice)
        switch ($choice) {
            0 { SafeShow-Message "好想尝尝你做的饭啊"; $global:DialogHistory += "晚餐选择：家常饭" }
            1 { SafeShow-Message "外卖吗？很方便呢"; $global:DialogHistory += "晚餐选择：外卖" }
            2 { SafeShow-Message "不要忘了吃饭哦，还是说你想要吃'小零食'呢……"; $global:DialogHistory += "晚餐选择：还没吃" }
        }
    }
}

function Start-NightDialog {
    $global:AudioPlayer = Play-Audio -FileName "call_night.wav" -Async
    Show-InteractiveDialog -Prompt "都这么晚了，快去休息吧！" -Options @("这就睡……", "好啊，晚安", "还有一些事情……") -Callback {
        param($choice)
        switch ($choice) {
            0 { SafeShow-Message "快去睡觉啦！身体最重要啦！"; $global:DialogHistory += "深夜选择：马上睡" }
            1 { SafeShow-Message "晚安哦，期待在梦中和你相见！"; $global:DialogHistory += "深夜选择：晚安" }
            2 { SafeShow-Message "好吧，那我陪着你一起！"; $global:DialogHistory += "深夜选择：继续做事" }
        }
    }
}

<#function Start-Night0000Dialog {
    $global:AudioPlayer = Play-Audio -FileName "call_sleep.wav" -Async
    Show-InteractiveDialog -Prompt "怎么还不去休息呀，是舍不得我吗？" -Options @("是啊～", "事情还没办完……") -Callback {
        param($choice)
        switch ($choice) {
            0 { SafeShow-Message "熬夜伤身体哦，我们早上再见吧～"; $global:DialogHistory += "零点选择：舍不得" }
            1 { SafeShow-Message "好吧，我仍然陪着你……"; $global:DialogHistory += "零点选择：继续做事" }
        }
    }
}#>
function Start-SleepDialog {
    $global:AudioPlayer = Play-Audio -FileName "call_sleep.wav" -Async
    Show-InteractiveDialog -Prompt "怎么还不去休息呀，是舍不得我吗？" -Options @("事情还没办完……", "是啊～") -Callback {
        param($choice)
        switch ($choice) {
            0 { SafeShow-Message "好吧，我仍然陪着你……"; $global:DialogHistory += "零点选择：继续做事" }
            1 { Show-InteractiveDialog -Prompt "舍不得我吗？那…做完那件事以后，我们一起睡觉好不好" -Options @("做什么事呢？") -Callback {
                    param($choice2)
                    switch ($choice2) {
                        0 { $global:AudioPlayer = Play-Audio -FileName "eat_2.wav" -Async
                            #SafeShow-Message "当然是……"$global:DialogHistory += "零点二级选择：做那件事"
                            Show-InteractiveDialog -Prompt "当然是…(扑)" -Options @("啊……") -Callback {
                                param($choice3)
                                switch ($choice3) {
                                    0 { $global:AudioPlayer = Play-Audio -FileName "eat_3.wav" -Async
                                        Show-InteractiveDialog -Prompt "啊……这样的力度…你喜欢嘛" -Options @("这样就好", "想要轻一点") -Callback {
                                            param($choice4)
                                            switch ($choice4) {
                                                0 { $global:AudioPlayer = Play-Audio -FileName "eat_41.wav" -Async
                                                    Show-InteractiveDialog -Prompt "嗯，我懂哦…再多告诉我一些，你引以为豪的地方吧" -Options @("请把我全部吃进去吧♡") -Callback {
                                                        param($choice5)
                                                        switch ($choice5) {
                                                            0 { $global:AudioPlayer = Play-Audio -FileName "eat_51.wav" -Async
                                                            SafeShow-Message "嗯！乖孩子乖孩子~"$global:DialogHistory += "零点五级选择：全部吃进去"
                                                            }
                                                        }
                                                    }
                                                }
                                                1 { $global:AudioPlayer = Play-Audio -FileName "eat_42.wav" -Async
                                                    Show-InteractiveDialog -Prompt "嗯…想要慢一点……我懂的哦~" -Options @("……", "再慢一点……再轻一点") -Callback {
                                                        param($choice5)
                                                        switch ($choice5) {
                                                            0 { $global:AudioPlayer = Play-Audio -FileName "eat_52.wav" -Async
                                                            Show-InteractiveDialog -Prompt "毕竟~我也想要慢慢享受" -Options @("……") -Callback {
                                                                    param($choice6)
                                                                    switch ($choice6) {
                                                                        0 { $global:AudioPlayer = Play-Audio -FileName "eat_62.wav" -Async
                                                                        SafeShow-Message "我们早就是命运交融的存在了~所以你不用说话…我就都明白哦~"$global:DialogHistory += "零点六级选择：不用说话"
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            1 { $global:AudioPlayer = Play-Audio -FileName "eat_53.wav" -Async
                                                            SafeShow-Message "先写到这里后面不知道怎么编"$global:DialogHistory += "零点六级选择：再轻一点"
                                                            #Show-InteractiveDialog -Prompt "啊……这样的力度…可以吗？" -Options @("选项") -Callback {
                                                            #        param($choice6)
                                                            #        switch ($choice6) {
                                                                    
                                                            #        }
                                                            #    }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}


function Start-ValentineDialog {
    Show-InteractiveDialog -Prompt "今天是个特别的日子……" -Options @("生日快乐！", "日记情人节快乐！", "不知道哦～") -Callback {
        param($choice)
        switch ($choice) {
            0 { $global:AudioPlayer = Play-Audio -FileName "day_birth.wav" -Async; SafeShow-Message "谢谢你还记得这个！"; $global:DialogHistory += "情人节选择：生日" }
            1 { $global:AudioPlayer = Play-Audio -FileName "day_diary.wav" -Async; SafeShow-Message "让我们互换日记吧！"; $global:DialogHistory += "情人节选择：情人节" }
            2 { $global:AudioPlayer = Play-Audio -FileName "chat_idk.wav" -Async; SafeShow-Message "不可以说不知道！"; $global:DialogHistory += "情人节选择：不知道" }
        }
        Start-Sleep -Milliseconds 500
    }
}

# =========「随机对话系统」=============
function Start-RandomDialog {
    $dialogType = Get-Random -Minimum 1 -Maximum 6
    switch ($dialogType) {
        1 { Start-RandomDialog1 }
        2 { Start-RandomDialog2 }
        3 { Start-RandomDialog3 }
        4 { Start-RandomDialog4 }
        5 { Start-RandomDialog5 }
    }
}

function Start-RandomDialog1 {
    $global:AudioPlayer = Play-Audio -FileName "chat_wate_matter.wav" -Async
    Show-InteractiveDialog -Prompt "你在做什么呀？" -Options @("在看你～", "在工作呢", "在玩呢", "在和别人聊天") -Callback {
        param($choice)
        switch ($choice) {
            0 { SafeShow-Message "我也在看你呢！"; $global:DialogHistory += "随机对话1：看你" }
            1 { SafeShow-Message "你认真工作的样子，很好看呢～"; $global:DialogHistory += "随机对话1：工作" }
            2 { SafeShow-Message "和我一起玩吧！"; $global:DialogHistory += "随机对话1：在玩" }
            3 { SafeShow-Message "你在……和谁聊天呢……"; $global:DialogHistory += "随机对话1：聊天" }
        }
    }
}

function Start-RandomDialog2 {
    $global:AudioPlayer = Play-Audio -FileName "chat_right.wav" -Async
    Show-InteractiveDialog -Prompt "想我了吗？" -Options @("没有哦～", "有的哦～", "你猜～") -Callback {
        param($choice)
        switch ($choice) {
            0 { SafeShow-Message "那我……很想你哦～"; $global:DialogHistory += "随机对话2：不想" }
            1 { SafeShow-Message "那我们是不是心有灵犀！"; $global:DialogHistory += "随机对话2：想" }
            2 { SafeShow-Message "我猜你想我了！"; $global:DialogHistory += "随机对话2：猜" }
        }
    }
}

function Start-RandomDialog3 {
    Show-InteractiveDialog -Prompt "你很无聊吗？不如……你给我讲个故事吧？" -Options @("好啊", "我要你给我讲", "有其他事……") -Callback {
        param($choice)
        switch ($choice) {
            0 { SafeShow-Message "你就这样给我讲吧，我会认真听着的！"; $global:DialogHistory += "随机对话3：你讲" }
            1 { 
                SafeShow-Message "好吧……那你可不要笑话我哦！"
                $global:AudioPlayer = Play-Audio -FileName "chat_not_laugh.wav" -Async
                Start-Sleep -Milliseconds 300
                $randomStory = $global:StoryLibrary | Get-Random
                Show-LongMessage $randomStory
                $global:DialogHistory += "随机对话3：我讲"
                if ($null -ne $global:AudioPlayer) {
                    try { $global:AudioPlayer.Stop(); $global:AudioPlayer.Dispose() } catch {}
                    $global:AudioPlayer = $null
                }
            }
            2 { SafeShow-Message "好啊，快去做你的事吧，我会一直看着你的……"; $global:DialogHistory += "随机对话3：有事" }
        }
    }
}

function Start-RandomDialog4 {
    $global:AudioPlayer = Play-Audio -FileName "chat_stay_with_me.wav" -Async
    Show-InteractiveDialog -Prompt "我好想你……" -Options @("我也好想你啊……", "我一直在的哦～") -Callback {
        param($choice)
        switch ($choice) {
            0 { SafeShow-Message "我们永远不分开，好吗……"; $global:DialogHistory += "随机对话4：也想你" }
            1 { SafeShow-Message "就这样永远陪着我，好吗……"; $global:DialogHistory += "随机对话4：一直在" }
        }
    }
}

function Start-RandomDialog5 {
    $global:AudioPlayer = Play-Audio -FileName "chat_look.wav" -Async
    Show-InteractiveDialog -Prompt "盯！" -Options @("盯！", "怎么了吗？") -Callback {
        param($choice)
        switch ($choice) {
            0 { SafeShow-Message "和你对视上了！你的眼里，全是爱我的模样呢～"; $global:DialogHistory += "随机对话5：对视" }
            1 { SafeShow-Message "只是……想多看看你……"; $global:DialogHistory += "随机对话5：怎么了" }
        }
    }
}

function Show-LongMessage {
    param([string]$message)
    try {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = $global:Title
        $form.Size = New-Object System.Drawing.Size(500, 400)
        $form.StartPosition = "CenterScreen"

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Multiline = $true
        $textBox.Location = New-Object System.Drawing.Point(10, 10)
        $textBox.Size = New-Object System.Drawing.Size(465, 320)
        $textBox.ScrollBars = "Vertical"
        $textBox.Text = $message
        $textBox.ReadOnly = $true
        $textBox.Font = New-Object System.Drawing.Font("微软雅黑", 10)

        $button = New-Object System.Windows.Forms.Button
        $button.Location = New-Object System.Drawing.Point(200, 340)
        $button.Size = New-Object System.Drawing.Size(75, 23)
        $button.Text = "确定"
        $button.Font = New-Object System.Drawing.Font("微软雅黑", 9)
        $button.Add_Click({ $form.Close() })

        $form.Controls.Add($textBox)
        $form.Controls.Add($button)
        $form.ShowDialog() | Out-Null
    } catch {
        Write-Host "`n[晴小姐的故事]`n" -ForegroundColor Yellow
        Write-Host $message -ForegroundColor Cyan
        Write-Host "`n按Enter键继续..." -ForegroundColor Gray
        $null = Read-Host
    }
}

# =========「程序启动入口」=============
try {
    Start-QingMiss
} catch {
    $errorMsg = "程序运行时错误: $_`nStackTrace: $($_.ScriptStackTrace)"
    Write-Host $errorMsg -ForegroundColor Red
    SafeShow-Message $errorMsg
} finally {
    if ($null -ne $global:ResidentForm -and !$global:ResidentForm.IsDisposed) {
        $global:ResidentForm.Close()
        $global:ResidentForm.Dispose()
    }
    
    if ($null -ne $global:AudioPlayer) {
        try {
            $global:AudioPlayer.Stop()
            $global:AudioPlayer.Dispose()
        } catch {}
        $global:AudioPlayer = $null
    }

    if ($null -ne $global:AudioPlayer) {
        try {
            $global:AudioPlayer.Stop()
            $global:AudioPlayer.Dispose()
        } catch {}
        $global:AudioPlayer = $null
    }
}