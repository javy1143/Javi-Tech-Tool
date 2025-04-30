Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Add Windows API for setting foreground window
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@

# Initialize ToolTip Object
$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.AutoPopDelay = 10000
$tooltip.InitialDelay = 500
$tooltip.ReshowDelay = 100
$tooltip.ShowAlways = $true

# Define global services hashtable
$global:services = @{
    "AssignedAccessManagerSvc"="Disabled"; "BDESVC"="Disabled"; "BthServ"="Automatic"; "DiagTrack"="Disabled"; 
    "DmEnrollmentSvc"="Manual"; "DPS"="Manual"; "Fax"="Disabled"; "FontCache"="Manual"; "HotspotService"="Disabled"; 
    "lfsvc"="Disabled"; "MapsBroker"="Manual"; "stisvc"="Manual"; "SysMain"="Manual"; "TermService"="Disabled"; 
    "TrkWks"="Manual"; "WlanSvc"="Manual"; "WManSvc"="Manual"; "wuauserv"="Manual"; "WSearch"="Manual"; 
    "XblAuthManager"="Disabled"; "XblGameSave"="Disabled"; "XboxNetApiSvc"="Disabled"; "RetailDemo"="Disabled"; 
    "RemoteRegistry"="Disabled"; "PhoneSvc"="Disabled"; "wisvc"="Disabled"; "icssvc"="Disabled"; "seclogon"="Disabled"
}

# Modified New-Checkbox Function to Support Tooltips
function New-Checkbox($text, $yPos, $panel, $tooltipText) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Location = New-Object System.Drawing.Point(10, $yPos)
    $cb.Size = New-Object System.Drawing.Size(380, 25)
    $cb.Text = $text
    $cb.ForeColor = [System.Drawing.Color]::White
    $cb.BackColor = $form.BackColor
    $cb.Font = New-Object System.Drawing.Font("Segoe Rodrigo", 9)
    $cb.Anchor = "Left,Right"
    if ($tooltipText) { $tooltip.SetToolTip($cb, $tooltipText) }
    $panel.Controls.Add($cb)
    return $cb
}

# Form Setup
$form = New-Object System.Windows.Forms.Form
$form.Text = "Win11 Optimizer"
$form.Size = New-Object System.Drawing.Size(700, 750)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$form.ForeColor = [System.Drawing.Color]::White

# Tab Control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$tabControl.Font = New-Object System.Drawing.Font("Segoe Rodrigo", 10)
$tabControl.BackColor = $form.BackColor
$tabControl.ForeColor = $form.ForeColor
$form.Controls.Add($tabControl)

# Function to safely disable and stop a service if it exists
function Disable-ServiceIfExists {
    param (
        [string]$ServiceName
    )
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($svc) {
        try {
            Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
            Set-Service -Name $ServiceName -StartupType Disabled -ErrorAction SilentlyContinue
            Write-Host "Disabled and stopped service: $ServiceName"
        } catch {
            Write-Warning "Failed to disable service: $ServiceName"
        }
    } else {
        Write-Host "Service not found: $ServiceName"
    }
}

# Utility Function: Add Tab
function Add-Tab($name) {
    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = $name
    $tab.BackColor = $form.BackColor
    $tab.ForeColor = $form.ForeColor
    $tabControl.TabPages.Add($tab)
    return $tab
}

# Common UI Elements Creation
function New-StatusLabel($yPos) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(10, $yPos)
    $lbl.AutoSize = $true
    $lbl.Height = 30
    $lbl.Text = "Status: Ready"
    $lbl.ForeColor = [System.Drawing.Color]::White
    $lbl.BackColor = $form.BackColor
    $lbl.Font = New-Object System.Drawing.Font("Segoe Rodrigo", 9)
    $lbl.TextAlign = "MiddleLeft"
    return $lbl
}

function New-ProgressBar($dock) {
    $pb = New-Object System.Windows.Forms.ProgressBar
    $pb.Dock = $dock
    $pb.Height = 20
    $pb.Minimum = 0
    $pb.Maximum = 100
    $pb.Value = 0
    $pb.Visible = $false
    $pb.BackColor = $form.BackColor
    $pb.ForeColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
    return $pb
}

function New-Button($text, $xPos, $width, $panel) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Location = New-Object System.Drawing.Point($xPos, 5)
    $btn.Size = New-Object System.Drawing.Size($width, 30)
    $btn.Text = $text
    $btn.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.Font = New-Object System.Drawing.Font("Segoe Rodrigo", 9, [System.Drawing.FontStyle]::Bold)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    $panel.Controls.Add($btn)
    return $btn
}

# System Info Tab
$tabSystemInfo = Add-Tab "System Info"
$containerSystemInfo = New-Object System.Windows.Forms.Panel
$containerSystemInfo.Dock = "Fill"
$containerSystemInfo.BackColor = $form.BackColor
$tabSystemInfo.Controls.Add($containerSystemInfo)

# Logo for System Info Tab from URL
$logoBox = New-Object System.Windows.Forms.PictureBox
$logoBox.Size = New-Object System.Drawing.Size(150, 150)
$logoBox.Location = New-Object System.Drawing.Point((($form.Width - $logoBox.Width) / 2), 20)
$logoBox.SizeMode = "StretchImage"
$logoBox.BackColor = $form.BackColor
try {
    $logoUrl = "https://i.ibb.co/zTJs6M7Y/Chat-GPT-Image-Apr-25-2025-10-02-08-AM.png"
    $response = Invoke-WebRequest -Uri $logoUrl -Method Get -ErrorAction Stop
    if ($response.Headers['Content-Type'] -notlike 'image/*') {
        throw "URL did not return an image. Content-Type: $($response.Headers['Content-Type'])"
    }
    $memoryStream = New-Object System.IO.MemoryStream(,$response.Content)
    $logoBox.Image = [System.Drawing.Image]::FromStream($memoryStream)
    $memoryStream.Dispose()
} catch {
    Write-Host "Failed to load logo from URL '$logoUrl': $($_.Exception.Message)"
}
$containerSystemInfo.Controls.Add($logoBox)

# Header Label
$lblWelcome = New-Object System.Windows.Forms.Label
$lblWelcome.Location = New-Object System.Drawing.Point(0, 180)
$lblWelcome.Size = New-Object System.Drawing.Size($form.Width, 50)
$lblWelcome.Text = "Optimize and manage your Windows 11 experience"
$lblWelcome.ForeColor = [System.Drawing.Color]::White
$lblWelcome.Font = New-Object System.Drawing.Font("Segoe Rodrigo", 18, [System.Drawing.FontStyle]::Bold)
$lblWelcome.TextAlign = "MiddleCenter"
$containerSystemInfo.Controls.Add($lblWelcome)

# System Info Panel
$panelSysInfo = New-Object System.Windows.Forms.Panel
$panelSysInfo.Size = New-Object System.Drawing.Size(540, 350)
$panelSysInfo.Location = New-Object System.Drawing.Point((($form.Width - $panelSysInfo.Width) / 2), 240)
$panelSysInfo.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$panelSysInfo.AutoScroll = $true
$containerSystemInfo.Controls.Add($panelSysInfo)

$sysInfoText = New-Object System.Windows.Forms.Label
$sysInfoText.Location = New-Object System.Drawing.Point(10, 10)
$sysInfoText.Size = New-Object System.Drawing.Size(510, 330)
$sysInfoText.ForeColor = [System.Drawing.Color]::White
$sysInfoText.Font = New-Object System.Drawing.Font("Segoe Rodrigo", 9)
$panelSysInfo.Controls.Add($sysInfoText)

function Load-SystemInfo {
    $os = Get-CimInstance Win32_OperatingSystem
    $computer = Get-CimInstance Win32_ComputerSystem
    $processor = Get-CimInstance Win32_Processor
    $disk = Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DeviceID -eq 'C:' }
    $memory = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
    $freeSpace = [math]::Round($disk.FreeSpace / 1GB, 2)
    $totalSpace = [math]::Round($disk.Size / 1GB, 2)

    $info = @"
System Information
=================
OS Name: $($os.Caption)
OS Version: $($os.Version)
OS Build: $($os.BuildNumber)
System Name: $($computer.Name)
Manufacturer: $($computer.Manufacturer)
Model: $($computer.Model)
Processor: $($processor.Name)
Total Memory: $memory GB
Disk C: Free Space: $freeSpace GB / Total: $totalSpace GB
Last Boot: $($os.LastBootUpTime)
=================
"@
    $sysInfoText.Text = $info
}

Load-SystemInfo

# Speed Up PC Tab
$tabSpeed = Add-Tab "Speed Up PC"
$containerSpeed = New-Object System.Windows.Forms.Panel
$containerSpeed.Dock = "Fill"
$containerSpeed.BackColor = $form.BackColor
$tabSpeed.Controls.Add($containerSpeed)

$panelSpeed = New-Object System.Windows.Forms.Panel
$panelSpeed.Location = New-Object System.Drawing.Point(10, 80)
$panelSpeed.Size = New-Object System.Drawing.Size(540, 450)
$panelSpeed.AutoScroll = $true
$panelSpeed.BackColor = $form.BackColor
$containerSpeed.Controls.Add($panelSpeed)

$yPos = 10
$speedCheckboxes = @(
    @{ Text = "Delete Temp Files"; Tooltip = "Removes temporary files from %TEMP% to free up disk space." },
    @{ Text = "Disable Consumer Features"; Tooltip = "Disables Windows consumer experiences like ads and tips." },
    @{ Text = "Disable Telemetry"; Tooltip = "Turns off data collection for Microsoft telemetry services." },
    @{ Text = "Disable Activity History"; Tooltip = "Prevents Windows from storing activity history." },
    @{ Text = "Disable GameDVR"; Tooltip = "Disables Xbox Game DVR and game bar features." },
    @{ Text = "Disable Hibernation"; Tooltip = "Turns off hibernation to save disk space." },
    @{ Text = "Disable Location Tracking"; Tooltip = "Prevents Windows from tracking your location." },
    @{ Text = "Disable Storage Sense"; Tooltip = "Turns off automatic storage cleanup features." },
    @{ Text = "Disable Wifi-Sense"; Tooltip = "Disables automatic Wi-Fi network sharing." },
    @{ Text = "Disable Recall"; Tooltip = "Disables Windows AI recall features." },
    @{ Text = "Debloat Edge"; Tooltip = "Removes first-run prompts and bloat from Microsoft Edge." },
    @{ Text = "Disable Background Applications"; Tooltip = "Prevents apps from running in the background." },
    @{ Text = "Clear Delivery Optimization Cache"; Tooltip = "Clears cache used for Windows update delivery." },
    @{ Text = "Set Services"; Tooltip = "Modifies startup types for services:`n" + ($global:services.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Key): $($_.Value)" } | Out-String).Trim() }
)
foreach ($check in $speedCheckboxes) {
    New-Checkbox $check.Text $yPos $panelSpeed $check.Tooltip | Out-Null
    $yPos += 28
}
$panelSpeed.VerticalScroll.Value = $panelSpeed.VerticalScroll.Maximum
$panelSpeed.PerformLayout()

$lblStatusSpeed = New-StatusLabel 10
$progressBarSpeed = New-ProgressBar "Bottom"
$btnPanelSpeed = New-Object System.Windows.Forms.Panel
$btnPanelSpeed.Location = New-Object System.Drawing.Point(10, 40)
$btnPanelSpeed.Size = New-Object System.Drawing.Size(540, 40)
$btnPanelSpeed.BackColor = $form.BackColor
$containerSpeed.Controls.AddRange(@($progressBarSpeed, $lblStatusSpeed, $btnPanelSpeed))

$btnSelectAllSpeed = New-Button "Select All" 0 120 $btnPanelSpeed
$tooltip.SetToolTip($btnSelectAllSpeed, "Selects all tweaks in the Speed Up PC tab.")
$btnSelectAllSpeed.Add_Click({ foreach ($control in $panelSpeed.Controls) { if ($control -is [System.Windows.Forms.CheckBox]) { $control.Checked = $true } } })

$btnRunSpeed = New-Button "Apply Tweaks" 130 120 $btnPanelSpeed
$tooltip.SetToolTip($btnRunSpeed, "Applies all selected tweaks to optimize system performance.")
$btnRunSpeed.Add_Click({
    $selectedTweaks = @($panelSpeed.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] -and $_.Checked })
    $totalTweaks = $selectedTweaks.Count
    if ($totalTweaks -eq 0) { $lblStatusSpeed.Text = "Status: No tweaks selected"; return }
    
    $progressBarSpeed.Visible = $true
    $progressBarSpeed.Maximum = $totalTweaks
    $progressBarSpeed.Value = 0
    $lblStatusSpeed.Text = "Status: Applying tweaks (0/$totalTweaks)..."
    $currentTweak = 0

    foreach ($control in $selectedTweaks) {
        $currentTweak++
        $tweak = $control.Text
        $lblStatusSpeed.Text = "Status: Applying tweak $currentTweak of $totalTweaks ($tweak)..."
        switch ($tweak) {
            "Delete Temp Files" { Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue }
            "Disable Consumer Features" { Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableWindowsConsumerFeatures" -Value 1 -ErrorAction SilentlyContinue }
            "Disable Telemetry" { Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0 -ErrorAction SilentlyContinue }
            "Disable Activity History" { Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -Name "EnableActivityFeed" -Value 0 -ErrorAction SilentlyContinue }
            "Disable GameDVR" { Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR" -Name "AllowGameDVR" -Value 0 -ErrorAction SilentlyContinue }
            "Disable Hibernation" { powercfg /hibernate off }
            "Disable Location Tracking" { Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors" -Name "DisableLocation" -Value 1 -ErrorAction SilentlyContinue }
            "Disable Storage Sense" { Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\StorageSense" -Name "AllowStorageSenseGlobal" -Value 0 -ErrorAction SilentlyContinue }
            "Disable Wifi-Sense" { Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config" -Name "AutoConnectAllowedOEM" -Value 0 -ErrorAction SilentlyContinue }
            "Disable Recall" { Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" -Name "DisableRecall" -Value 1 -ErrorAction SilentlyContinue }
            "Debloat Edge" { Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "HideFirstRunExperience" -Value 1 -ErrorAction SilentlyContinue }
            "Clear Delivery Optimization Cache" { Remove-Item -Path "C:\Windows\SoftwareDistribution\DeliveryOptimization" -Recurse -Force -ErrorAction SilentlyContinue }
            "Disable Background Applications" { 
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsRunInBackground" -Value 2 -ErrorAction SilentlyContinue
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications\Microsoft.OneDrive" -Name "Disabled" -Value 0 -ErrorAction SilentlyContinue
            }
            "Set Services" { 
				$safeList = @(
					"SysMain",            # SysMain (Superfetch/Memory Prefetcher)
					"WSearch",            # Windows Search
					"dmwappushservice",   # WAP Push Message Routing Service
					"Fax",                # Fax
					"XblGameSave",        # Xbox Live Game Save
					"XboxNetApiSvc",      # Xbox Live Networking Service
					"RetailDemo",         # Retail Demo Service
					"XblAuthManager",     # Xbox Live Auth Manager
					"PhoneSvc",           # Phone Service
					"wisvc",              # Windows Insider Service
					"icssvc",             # Internet Connection Sharing (ICS)
					"MapsBroker"          # Downloaded Maps Manager
				)
				foreach ($svc in $global:services.GetEnumerator()) {
					if ($safeList -contains $svc.Key) {
						Disable-ServiceIfExists -ServiceName $svc.Key
					} else {
						try {
							Set-Service -Name $svc.Key -StartupType $svc.Value -ErrorAction SilentlyContinue
						} catch {
							Write-Warning "Could not set service '$($svc.Key)' to '$($svc.Value)'"
						}
					}
				}
			}
        }
        $progressBarSpeed.Value = $currentTweak
        $form.Refresh()
    }
    $progressBarSpeed.Visible = $false
    $lblStatusSpeed.Text = "Status: Done"
})

# Preferences Tab
$tabPrefs = Add-Tab "Preferences"
$containerPrefs = New-Object System.Windows.Forms.Panel
$containerPrefs.Dock = "Fill"
$containerPrefs.BackColor = $form.BackColor
$tabPrefs.Controls.Add($containerPrefs)

$panelPrefs = New-Object System.Windows.Forms.Panel
$panelPrefs.Location = New-Object System.Drawing.Point(10, 80)
$panelPrefs.Size = New-Object System.Drawing.Size(540, 400)
$panelPrefs.AutoScroll = $true
$panelPrefs.BackColor = $form.BackColor
$containerPrefs.Controls.Add($panelPrefs)

$yPos = 10
$prefsCheckboxes = @(
    @{ Text = "Dark Theme for Windows"; Tooltip = "Enables dark mode for Windows apps and settings." },
    @{ Text = "Disable Bing Search in Start Menu"; Tooltip = "Removes Bing search results from the Start Menu." },
    @{ Text = "Disable Recommendations in Start Menu"; Tooltip = "Turns off suggested apps and content in Start Menu." },
    @{ Text = "Search Button in Taskbar"; Tooltip = "Shows the search button on the taskbar." },
    @{ Text = "Disable Widgets in Taskbar"; Tooltip = "Removes the Widgets button from the taskbar." },
    @{ Text = "Disable Animations"; Tooltip = "Turns off visual animations for faster performance." },
    @{ Text = "Turn Off Taskbar Transparency"; Tooltip = "Disables transparency effects on the taskbar." },
    @{ Text = "Disable Snap Assist"; Tooltip = "Turns off window snapping features." },
    @{ Text = "Add and Activate Ultimate Performance Profile"; Tooltip = "Adds and enables the Ultimate Performance power plan." },
    @{ Text = "Improve Visual Effects for Performance"; Tooltip = "Optimizes visual effects for better performance while keeping essential effects enabled." }
)
foreach ($check in $prefsCheckboxes) {
    New-Checkbox $check.Text $yPos $panelPrefs $check.Tooltip | Out-Null
    $yPos += 30
}

$lblStatusPrefs = New-StatusLabel 10
$progressBarPrefs = New-ProgressBar "Bottom"
$btnPanelPrefs = New-Object System.Windows.Forms.Panel
$btnPanelPrefs.Location = New-Object System.Drawing.Point(10, 40)
$btnPanelPrefs.Size = New-Object System.Drawing.Size(540, 40)
$btnPanelPrefs.BackColor = $form.BackColor
$containerPrefs.Controls.AddRange(@($progressBarPrefs, $lblStatusPrefs, $btnPanelPrefs))

$btnSelectAllPrefs = New-Button "Select All" 0 120 $btnPanelPrefs
$tooltip.SetToolTip($btnSelectAllPrefs, "Selects all preferences in the Preferences tab.")
$btnSelectAllPrefs.Add_Click({ foreach ($control in $panelPrefs.Controls) { if ($control -is [System.Windows.Forms.CheckBox]) { $control.Checked = $true } } })

$btnRunPrefs = New-Button "Apply Tweaks" 130 120 $btnPanelPrefs
$tooltip.SetToolTip($btnRunPrefs, "Applies all selected preferences to customize Windows.")
$btnRunPrefs.Add_Click({
    $selectedTweaks = @($panelPrefs.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] -and $_.Checked })
    $totalTweaks = $selectedTweaks.Count
    if ($totalTweaks -eq 0) { $lblStatusPrefs.Text = "Status: No tweaks selected"; return }

    $progressBarPrefs.Visible = $true
    $progressBarPrefs.Maximum = $totalTweaks
    $progressBarPrefs.Value = 0
    $lblStatusPrefs.Text = "Status: Applying tweaks (0/$totalTweaks)..."
    $currentTweak = 0

    foreach ($control in $selectedTweaks) {
        $currentTweak++
        $tweak = $control.Text
        $lblStatusPrefs.Text = "Status: Applying tweak $currentTweak of $totalTweaks ($tweak)..."

        switch ($tweak) {
            "Dark Theme for Windows" {
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0 -ErrorAction SilentlyContinue
            }
            "Disable Bing Search in Start Menu" {
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled" -Value 0 -ErrorAction SilentlyContinue
            }
            "Disable Recommendations in Start Menu" {
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Start_IrisRecommendations" -Value 0 -ErrorAction SilentlyContinue
            }
            "Search Button in Taskbar" {
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 1 -ErrorAction SilentlyContinue
            }
            "Disable Widgets in Taskbar" {
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Widgets" -Name "AllowWidgets" -Value 0 -ErrorAction SilentlyContinue
            }
            "Disable Animations" {
                if (-not ([System.Management.Automation.PSTypeName]'SysParams').Type) {
                    Add-Type @"
                    using System;
                    using System.Runtime.InteropServices;
                    public class SysParams {
                        [DllImport("user32.dll", SetLastError = true)]
                        public static extern bool SystemParametersInfo(int uAction, int uParam, ref bool lpvParam, int fuWinIni);
                    }
"@
                }
                $SPI_SETANIMATION = 0x0049
                $SPIF_UPDATEINIFILE = 0x01
                $SPIF_SENDCHANGE = 0x02
                $animation = $false
                [SysParams]::SystemParametersInfo($SPI_SETANIMATION, 0, [ref]$animation, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE) | Out-Null
            }
            "Turn Off Taskbar Transparency" {
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "EnableTransparency" -Value 0
            }
            "Disable Snap Assist" {
                Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WindowArrangementActive" -Value 0
            }
            "Add and Activate Ultimate Performance Profile" {
                powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61
                $guid = (powercfg -l | Where-Object { $_ -match "Ultimate Performance" } | ForEach-Object { $_.Split()[3] })
                if ($guid) { powercfg -s $guid }
            }
            "Improve Visual Effects for Performance" {
                if (-not ([System.Management.Automation.PSTypeName]'SysParamsAdv').Type) {
                    Add-Type @"
                    using System;
                    using System.Runtime.InteropServices;
                    public class SysParamsAdv {
                        [DllImport("user32.dll", SetLastError = true)]
                        public static extern bool SystemParametersInfo(int uAction, int uParam, ref bool lpvParam, int fuWinIni);
                    }
"@
                }

                $SPIF_UPDATEINIFILE = 0x01
                $SPIF_SENDCHANGE = 0x02

                function Set-SPIBool($action, $enable) {
                    [void][SysParamsAdv]::SystemParametersInfo($action, 0, [ref]$enable, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)
                }

                $allEffects = @{
                    0x1001 = $false  # Animate controls
                    0x1003 = $false  # Animate windows
                    0x1012 = $false  # Enable Peek
                    0x1013 = $true   # ✅ Animations in the taskbar
                    0x1014 = $false  # Fade/slide menus
                    0x1015 = $false  # Fade/slide ToolTips
                    0x1016 = $false  # Fade out menu items
                    0x1017 = $false  # Save taskbar previews
                    0x1018 = $false  # Show shadow under mouse
                    0x1019 = $false  # Show shadows under windows
                    0x101A = $true   # ✅ Show thumbnails instead of icons
                    0x101B = $false  # Translucent selection rectangle
                    0x101C = $false  # Show contents while dragging
                    0x101D = $false  # Slide open combo boxes
                    0x101E = $true   # ✅ Smooth edges of screen fonts
                    0x101F = $false  # Smooth-scroll list boxes
                    0x1020 = $false  # Drop shadows for desktop icons
                }

                foreach ($action in $allEffects.Keys) {
                    Set-SPIBool -action $action -enable $allEffects[$action]
                }
            }
        }

        $progressBarPrefs.Value = $currentTweak
        $form.Refresh()
    }

    $progressBarPrefs.Visible = $false
    $lblStatusPrefs.Text = "Status: Done"
})


# Required types
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.ComponentModel

# --- Windows Updates Tab with Async Install, Selection & Restart ---
$tabUpdates = Add-Tab "Windows Updates"
$containerUpdates = New-Object System.Windows.Forms.Panel
$containerUpdates.Dock = "Fill"
$containerUpdates.BackColor = $form.BackColor
$tabUpdates.Controls.Add($containerUpdates)

$listUpdates = New-Object System.Windows.Forms.CheckedListBox
$listUpdates.Location = New-Object System.Drawing.Point(10, 140)
$listUpdates.Size = New-Object System.Drawing.Size(540, 340)
$listUpdates.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$listUpdates.ForeColor = [System.Drawing.Color]::White
$listUpdates.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$listUpdates.CheckOnClick = $true
$containerUpdates.Controls.Add($listUpdates)

$lblStatusUpdates = New-Object System.Windows.Forms.Label
$lblStatusUpdates.Location = New-Object System.Drawing.Point(10, 10)
$lblStatusUpdates.Size = New-Object System.Drawing.Size(540, 25)
$lblStatusUpdates.Text = "Status: Ready"
$lblStatusUpdates.ForeColor = [System.Drawing.Color]::White
$containerUpdates.Controls.Add($lblStatusUpdates)

$progressBarUpdates = New-Object System.Windows.Forms.ProgressBar
$progressBarUpdates.Dock = "Bottom"
$progressBarUpdates.Height = 20
$progressBarUpdates.Minimum = 0
$progressBarUpdates.Maximum = 100
$progressBarUpdates.Visible = $false
$containerUpdates.Controls.Add($progressBarUpdates)

$btnPanelUpdates = New-Object System.Windows.Forms.Panel
$btnPanelUpdates.Location = New-Object System.Drawing.Point(10, 40)
$btnPanelUpdates.Size = New-Object System.Drawing.Size(540, 90)
$btnPanelUpdates.BackColor = $form.BackColor
$containerUpdates.Controls.Add($btnPanelUpdates)

function New-UpdateButton($text, $x, $y) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Location = New-Object System.Drawing.Point($x, $y)
    $btn.Size = New-Object System.Drawing.Size(150, 30)
    $btn.Text = $text
    $btn.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    return $btn
}

$btnLoadAll = New-UpdateButton "Load All Updates" 0 5
$btnInstallAll = New-UpdateButton "Install All Updates" 160 5
$btnLoadSecurity = New-UpdateButton "Load Security Updates" 0 45
$btnInstallSecurity = New-UpdateButton "Install Security Updates" 160 45
$btnRestart = New-UpdateButton "Restart Now" 320 5
$btnInstallSelected = New-UpdateButton "Install Selected" 320 45

$btnPanelUpdates.Controls.AddRange(@($btnLoadAll, $btnInstallAll, $btnLoadSecurity, $btnInstallSecurity, $btnRestart, $btnInstallSelected))

$script:allUpdates = @()
$script:securityUpdates = @()

$btnLoadAll.Add_Click({
    $lblStatusUpdates.Text = "Status: Loading all updates..."
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -Confirm:$false
    }
    Import-Module PSWindowsUpdate -Force
    $listUpdates.Items.Clear()
    $script:allUpdates = Get-WindowsUpdate -MicrosoftUpdate -ErrorAction SilentlyContinue
    if ($script:allUpdates) {
        foreach ($update in $script:allUpdates) {
            $listUpdates.Items.Add("[$($update.KB)] $($update.Title)") | Out-Null
        }
        $lblStatusUpdates.Text = "Status: Loaded all updates"
    } else {
        $lblStatusUpdates.Text = "Status: No updates found"
    }
})

$btnLoadSecurity.Add_Click({
    $lblStatusUpdates.Text = "Status: Loading security updates..."
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -Confirm:$false
    }
    Import-Module PSWindowsUpdate -Force
    $listUpdates.Items.Clear()
    $script:securityUpdates = Get-WindowsUpdate -MicrosoftUpdate -Category "Security Updates" -ErrorAction SilentlyContinue
    if ($script:securityUpdates) {
        foreach ($update in $script:securityUpdates) {
            $listUpdates.Items.Add("[$($update.KB)] $($update.Title)") | Out-Null
        }
        $lblStatusUpdates.Text = "Status: Loaded security updates"
    } else {
        $lblStatusUpdates.Text = "Status: No security updates found"
    }
})

function Install-UpdatesAsync($updates) {
    $progressBarUpdates.Visible = $true
    $progressBarUpdates.Maximum = $updates.Count
    $progressBarUpdates.Value = 0

    $worker = New-Object System.ComponentModel.BackgroundWorker
    $worker.WorkerReportsProgress = $true

    $worker.add_DoWork({
        for ($i = 0; $i -lt $updates.Count; $i++) {
            $update = $updates[$i]
            Install-WindowsUpdate -MicrosoftUpdate -KBArticleID $update.KB -AcceptAll -IgnoreReboot -ErrorAction SilentlyContinue
            $worker.ReportProgress($i + 1, "Installed KB$($update.KB)")
        }
    })

    $worker.add_ProgressChanged({
	    param($s, $e)
	    $progressBarUpdates.Value = $e.ProgressPercentage
	    $percent = [math]::Round(($e.ProgressPercentage / $progressBarUpdates.Maximum) * 100)
	    $lblStatusUpdates.Text = "Status: $($e.UserState) ($percent%)"
	    $listUpdates.Items.Add("✅ $($e.UserState) ($percent%)") | Out-Null
	    $form.Refresh()
	})

    $worker.add_RunWorkerCompleted({
        $lblStatusUpdates.Text = "Status: Updates installed"
        $progressBarUpdates.Visible = $false
        $listUpdates.Items.Clear()
        $script:allUpdates = @()
        $script:securityUpdates = @()
    })

    $worker.RunWorkerAsync()
}

$btnInstallAll.Add_Click({
    if (-not $script:allUpdates -or $script:allUpdates.Count -eq 0) {
        $lblStatusUpdates.Text = "Status: No updates to install"
        return
    }
    Install-UpdatesAsync $script:allUpdates
})

$btnInstallSecurity.Add_Click({
    if (-not $script:securityUpdates -or $script:securityUpdates.Count -eq 0) {
        $lblStatusUpdates.Text = "Status: No security updates to install"
        return
    }
    Install-UpdatesAsync $script:securityUpdates
})

$btnInstallSelected.Add_Click({
    if (-not $script:allUpdates -or $script:allUpdates.Count -eq 0) {
        $lblStatusUpdates.Text = "Status: No updates loaded"
        return
    }
    $selected = @()
    for ($i = 0; $i -lt $listUpdates.Items.Count; $i++) {
        if ($listUpdates.GetItemChecked($i)) {
            $title = $listUpdates.Items[$i]
            $kb = ($title -split '\[|\]')[1]
            $update = $script:allUpdates | Where-Object { $_.KB -eq $kb }
            if ($update) { $selected += $update }
        }
    }
    if ($selected.Count -eq 0) {
        $lblStatusUpdates.Text = "Status: No updates selected"
        return
    }
    Install-UpdatesAsync $selected
})

$btnRestart.Add_Click({
    if (([System.Windows.Forms.MessageBox]::Show("Are you sure you want to restart now?", "Restart Confirmation", "YesNo", "Question")) -eq "Yes") {
        Restart-Computer -Force
    }
})

### Startup Manager Tab ###
$tabStartup = Add-Tab "Startup Manager"
$containerStartup = New-Object System.Windows.Forms.Panel
$containerStartup.Dock = "Fill"
$containerStartup.BackColor = $form.BackColor
$tabStartup.Controls.Add($containerStartup)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Location = New-Object System.Drawing.Point(10, 80)
$lblFilter.Size = New-Object System.Drawing.Size(60, 25)
$lblFilter.Text = "Filter:"
$lblFilter.ForeColor = [System.Drawing.Color]::White
$containerStartup.Controls.Add($lblFilter)

$comboFilter = New-Object System.Windows.Forms.ComboBox
$comboFilter.Location = New-Object System.Drawing.Point(70, 80)
$comboFilter.Size = New-Object System.Drawing.Size(180, 25)
$comboFilter.Items.AddRange(@("All", "Current User", "Local Machine", "Startup Folder", "Enabled"))
$comboFilter.SelectedIndex = 0
$comboFilter.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$comboFilter.ForeColor = [System.Drawing.Color]::White
$comboFilter.DropDownStyle = "DropDownList"
$containerStartup.Controls.Add($comboFilter)

$lvStartup = New-Object System.Windows.Forms.ListView
$lvStartup.Location = New-Object System.Drawing.Point(10, 110)
$lvStartup.Size = New-Object System.Drawing.Size(650, 250)
$lvStartup.View = 'Details'
$lvStartup.CheckBoxes = $true
$lvStartup.FullRowSelect = $true
$lvStartup.GridLines = $true
$lvStartup.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$lvStartup.ForeColor = [System.Drawing.Color]::White
$lvStartup.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lvStartup.Columns.Add("Name", 180) | Out-Null
$lvStartup.Columns.Add("Hive", 100) | Out-Null
$lvStartup.Columns.Add("Enabled", 80) | Out-Null
$lvStartup.Columns.Add("Command", 280) | Out-Null
$containerStartup.Controls.Add($lvStartup)

$lblStartupStatus = New-Object System.Windows.Forms.Label
$lblStartupStatus.Location = New-Object System.Drawing.Point(10, 370)
$lblStartupStatus.Size = New-Object System.Drawing.Size(640, 25)
$lblStartupStatus.Text = "Status: Ready"
$lblStartupStatus.ForeColor = [System.Drawing.Color]::White
$containerStartup.Controls.Add($lblStartupStatus)

function Load-StartupItems {
    $lvStartup.Items.Clear()
    $script:startupItems = @()
    $approved = @{}

    $approvedPaths = @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
    )
    foreach ($path in $approvedPaths) {
        if (Test-Path $path) {
            $items = Get-ItemProperty $path
            foreach ($prop in $items.PSObject.Properties) {
                $val = $items.$($prop.Name)
                if ($val.Length -ge 1) { $approved[$prop.Name] = ($val[0] -eq 2) }
            }
        }
    }

    $regSources = @(
        @{ Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"; Hive = "HKCU" },
        @{ Path = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"; Hive = "HKLM" }
    )
    foreach ($src in $regSources) {
        if (Test-Path $src.Path) {
            $props = Get-ItemProperty $src.Path
            foreach ($name in $props.PSObject.Properties.Name) {
                $cmd = $props.$name
                $enabled = if ($approved.ContainsKey($name)) { $approved[$name] } else { $true }
                $script:startupItems += [PSCustomObject]@{
                    Name = $name; Hive = $src.Hive; Enabled = $enabled; Command = $cmd
                }
            }
        }
    }

    $startupFolders = @(
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup",
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
    )
    $shell = New-Object -ComObject Shell.Application
    foreach ($folder in $startupFolders) {
        if (Test-Path $folder) {
            $fobj = $shell.NameSpace($folder)
            $items = $fobj.Items()
            for ($i = 0; $i -lt $items.Count; $i++) {
                $item = $items.Item($i)
                if ($item.Name -like "*.lnk") {
                    $link = $item.GetLink()
                    $target = $link.Path
                    $name = [System.IO.Path]::GetFileNameWithoutExtension($item.Name)
                    $script:startupItems += [PSCustomObject]@{
                        Name = $name; Hive = "StartupFolder"; Enabled = $true; Command = $target
                    }
                }
            }
        }
    }

    $criterion = $comboFilter.SelectedItem
    $filtered = switch ($criterion) {
        "Current User"   { $script:startupItems | Where-Object { $_.Hive -eq "HKCU" } }
        "Local Machine"  { $script:startupItems | Where-Object { $_.Hive -eq "HKLM" } }
        "Startup Folder" { $script:startupItems | Where-Object { $_.Hive -eq "StartupFolder" } }
        "Enabled"        { $script:startupItems | Where-Object { $_.Enabled -eq $true } }
        default          { $script:startupItems }
    }

    foreach ($item in $filtered) {
        $entry = New-Object System.Windows.Forms.ListViewItem($item.Name)
        $entry.SubItems.Add($item.Hive) | Out-Null
        $entry.SubItems.Add(($item.Enabled) ? "Yes" : "No") | Out-Null
        $entry.SubItems.Add($item.Command) | Out-Null
        $entry.Checked = $item.Enabled
        $lvStartup.Items.Add($entry) | Out-Null
    }

    $lblStartupStatus.Text = "Status: Loaded $($filtered.Count) startup item(s)"
}

function Toggle-StartupApproval {
    foreach ($entry in $lvStartup.Items) {
        $item = $script:startupItems | Where-Object { $_.Name -eq $entry.Text } | Select-Object -First 1
        if ($item -and $item.Hive -ne "StartupFolder") {
            $hiveKey = if ($item.Hive -eq "HKCU") {
                "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
            } else {
                "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
            }
            $value = if ($entry.Checked) { 2 } else { 3 }
            $bytes = ,$value + (,0 * 11)  # 12-byte structure
            Set-ItemProperty -Path $hiveKey -Name $item.Name -Value $bytes -ErrorAction SilentlyContinue
        }
    }
    Load-StartupItems
}

$comboFilter.Add_SelectedIndexChanged({ Load-StartupItems })
$lvStartup.add_ItemChecked({ Toggle-StartupApproval })
Load-StartupItems

# Show Form
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
