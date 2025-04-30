# Win11Optimizer.ps1 - Fixed and Optimized

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Windows API for setting foreground window
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@

# Initialize ToolTip
$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.AutoPopDelay = 10000
$tooltip.InitialDelay = 500
$tooltip.ReshowDelay = 100
$tooltip.ShowAlways = $true

# Services hashtable
$global:services = @{
    "AssignedAccessManagerSvc"="Disabled"; "BDESVC"="Disabled"; "BthServ"="Automatic"; "DiagTrack"="Disabled"; 
    "DmEnrollmentSvc"="Manual"; "DPS"="Manual"; "Fax"="Disabled"; "FontCache"="Manual"; "HotspotService"="Disabled"; 
    "lfsvc"="Disabled"; "MapsBroker"="Manual"; "stisvc"="Manual"; "SysMain"="Manual"; "TermService"="Disabled"; 
    "TrkWks"="Manual"; "WlanSvc"="Manual"; "WManSvc"="Manual"; "wuauserv"="Manual"; "WSearch"="Manual"; 
    "XblAuthManager"="Disabled"; "XblGameSave"="Disabled"; "XboxNetApiSvc"="Disabled"; "RetailDemo"="Disabled"; 
    "RemoteRegistry"="Disabled"; "PhoneSvc"="Disabled"; "wisvc"="Disabled"; "icssvc"="Disabled"; "seclogon"="Disabled"
}

# New-Checkbox Function
function New-Checkbox($text, $yPos, $panel, $tooltipText) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Location = New-Object System.Drawing.Point(10, $yPos)
    $cb.Size = New-Object System.Drawing.Size(520, 25)
    $cb.Text = $text
    $cb.ForeColor = [System.Drawing.Color]::White
    $cb.BackColor = $panel.BackColor
    $cb.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cb.AutoSize = $true
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
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$tabControl.BackColor = $form.BackColor
$tabControl.ForeColor = $form.ForeColor
$form.Controls.Add($tabControl)

# Utility Functions
function Add-Tab($name) {
    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = $name
    $tab.BackColor = $form.BackColor
    $tab.ForeColor = $form.ForeColor
    $tabControl.TabPages.Add($tab)
    return $tab
}

function New-StatusLabel($yPos) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(10, $yPos)
    $lbl.AutoSize = $true
    $lbl.Text = "Status: Ready"
    $lbl.ForeColor = [System.Drawing.Color]::White
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
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
    $pb.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    return $pb
}

function New-Button($text, $xPos, $width, $panel) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Location = New-Object System.Drawing.Point($xPos, 5)
    $btn.Size = New-Object System.Drawing.Size($width, 30)
    $btn.Text = $text
    $btn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    if ($panel) { $panel.Controls.Add($btn) }
    return $btn
}

# System Info Tab
$tabSystemInfo = Add-Tab "System Info"
$containerSystemInfo = New-Object System.Windows.Forms.Panel
$containerSystemInfo.Dock = "Fill"
$containerSystemInfo.BackColor = $form.BackColor
$tabSystemInfo.Controls.Add($containerSystemInfo)

# Logo
$logoBox = New-Object System.Windows.Forms.PictureBox
$logoBox.Size = New-Object System.Drawing.Size(150, 150)
$logoBox.Location = New-Object System.Drawing.Point((($form.Width - $logoBox.Width) / 2), 20)
$logoBox.SizeMode = "StretchImage"
$containerSystemInfo.Controls.Add($logoBox)

try {
    $logoUrl = "https://i.ibb.co/zTJs6M7/Chat-GPT-Image-Apr-25-2025-10-02-08-AM.png"
    $client = New-Object System.Net.WebClient
    $imageBytes = $client.DownloadData($logoUrl)
    $memoryStream = New-Object System.IO.MemoryStream($imageBytes)
    $logoBox.Image = [System.Drawing.Image]::FromStream($memoryStream)
    $memoryStream.Dispose()
} catch {
    Write-Host "Failed to load logo (URL may be invalid or inaccessible): $($_.Exception.Message)"
}

# Header Label
$lblWelcome = New-Object System.Windows.Forms.Label
$lblWelcome.Location = New-Object System.Drawing.Point(0, 180)
$lblWelcome.Size = New-Object System.Drawing.Size($form.Width, 50)
$lblWelcome.Text = "Optimize Windows 11"
$lblWelcome.ForeColor = [System.Drawing.Color]::White
$lblWelcome.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
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
$sysInfoText.Font = New-Object System.Drawing.Font("Segoe UI", 9)
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
OS: $($os.Caption)
Version: $($os.Version)
Build: $($os.BuildNumber)
Name: $($computer.Name)
Manufacturer: $($computer.Manufacturer)
Model: $($computer.Model)
Processor: $($processor.Name)
Memory: $memory GB
Disk C: $freeSpace GB / $totalSpace GB
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
$panelSpeed.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$containerSpeed.Controls.Add($panelSpeed)

$yPos = 10
$speedCheckboxes = @(
    @{ Text = "Delete Temp Files"; Tooltip = "Clears %TEMP% directory." },
    @{ Text = "Disable Consumer Features"; Tooltip = "Disables ads and tips." },
    @{ Text = "Disable Telemetry"; Tooltip = "Stops Microsoft telemetry." },
    @{ Text = "Disable Activity History"; Tooltip = "Stops activity tracking." },
    @{ Text = "Disable GameDVR"; Tooltip = "Disables Xbox Game DVR." },
    @{ Text = "Disable Hibernation"; Tooltip = "Disables hibernation." },
    @{ Text = "Disable Location Tracking"; Tooltip = "Stops location tracking." },
    @{ Text = "Disable Storage Sense"; Tooltip = "Disables auto storage cleanup." },
    @{ Text = "Disable Wifi-Sense"; Tooltip = "Disables Wi-Fi sharing." },
    @{ Text = "Disable Recall"; Tooltip = "Disables AI recall features." },
    @{ Text = "Debloat Edge"; Tooltip = "Removes Edge first-run prompts." },
    @{ Text = "Disable Background Apps"; Tooltip = "Stops background apps." },
    @{ Text = "Set Services"; Tooltip = "Sets service startup types.`n" + ($global:services.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Key): $($_.Value)" } | Out-String).Trim() }
)
foreach ($check in $speedCheckboxes) {
    New-Checkbox $check.Text $yPos $panelSpeed $check.Tooltip | Out-Null
    $yPos += 30
}

$lblStatusSpeed = New-StatusLabel 10
$progressBarSpeed = New-ProgressBar "Bottom"
$btnPanelSpeed = New-Object System.Windows.Forms.Panel
$btnPanelSpeed.Location = New-Object System.Drawing.Point(10, 40)
$btnPanelSpeed.Size = New-Object System.Drawing.Size(540, 40)
$btnPanelSpeed.BackColor = $form.BackColor
$containerSpeed.Controls.AddRange(@($progressBarSpeed, $lblStatusSpeed, $btnPanelSpeed))

$btnSelectAllSpeed = New-Button "Select All" 0 120 $btnPanelSpeed
$tooltip.SetToolTip($btnSelectAllSpeed, "Selects all tweaks.")
$btnSelectAllSpeed.Add_Click({ $panelSpeed.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] } | ForEach-Object { $_.Checked = $true } })

$btnRunSpeed = New-Button "Apply Tweaks" 130 120 $btnPanelSpeed
$tooltip.SetToolTip($btnRunSpeed, "Applies selected tweaks.")
$btnRunSpeed.Add_Click({
    $selectedTweaks = $panelSpeed.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] -and $_.Checked }
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
        $lblStatusSpeed.Text = "Status: Applying tweak $currentTweak/$totalTweaks ($tweak)..."
        try {
            switch ($tweak) {
                "Delete Temp Files" { Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction Stop }
                "Disable Consumer Features" { 
                    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Force | Out-Null
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableWindowsConsumerFeatures" -Value 1 
                }
                "Disable Telemetry" { 
                    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Force | Out-Null
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0 
                }
                "Disable Activity History" { 
                    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -Force | Out-Null
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -Name "EnableActivityFeed" -Value 0 
                }
                "Disable GameDVR" { 
                    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR" -Force | Out-Null
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR" -Name "AllowGameDVR" -Value 0 
                }
                "Disable Hibernation" { powercfg /hibernate off }
                "Disable Location Tracking" { 
                    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors" -Force | Out-Null
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors" -Name "DisableLocation" -Value 1 
                }
                "Disable Storage Sense" { 
                    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\StorageSense" -Force | Out-Null
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\StorageSense" -Name "AllowStorageSenseGlobal" -Value 0 
                }
                "Disable Wifi-Sense" { 
                    New-Item -Path "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config" -Force | Out-Null
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config" -Name "AutoConnectAllowedOEM" -Value 0 
                }
                "Disable Recall" { 
                    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" -Force | Out-Null
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" -Name "DisableRecall" -Value 1 
                }
                "Debloat Edge" { 
                    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Force | Out-Null
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "HideFirstRunExperience" -Value 1 
                }
                "Disable Background Apps" { 
                    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Force | Out-Null
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsRunInBackground" -Value 2 
                }
                "Set Services" { 
                    foreach ($svc in $global:services.GetEnumerator()) { 
                        Set-Service -Name $svc.Key -StartupType $svc.Value -ErrorAction Stop 
                    }
                }
            }
        } catch {
            $lblStatusSpeed.Text = "Status: Error on $tweak - $($_.Exception.Message)"
            $progressBarSpeed.Visible = $false
            return
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
$panelPrefs.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$containerPrefs.Controls.Add($panelPrefs)

$yPos = 10
$prefsCheckboxes = @(
    @{ Text = "Dark Theme"; Tooltip = "Enables dark mode." },
    @{ Text = "Disable Bing in Start"; Tooltip = "Removes Bing from Start Menu." },
    @{ Text = "Disable Start Recommendations"; Tooltip = "Stops suggested apps." },
    @{ Text = "Show Search Button"; Tooltip = "Shows search button on taskbar." },
    @{ Text = "Disable Widgets"; Tooltip = "Removes Widgets button." },
    @{ Text = "Disable Animations"; Tooltip = "Disables animations." },
    @{ Text = "Disable Taskbar Transparency"; Tooltip = "Disables taskbar transparency." },
    @{ Text = "Disable Snap Assist"; Tooltip = "Disables window snapping." },
    @{ Text = "Ultimate Performance Profile"; Tooltip = "Enables Ultimate Performance plan." },
    @{ Text = "Optimize Visual Effects"; Tooltip = "Optimizes visual effects." }
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
$tooltip.SetToolTip($btnSelectAllPrefs, "Selects all preferences.")
$btnSelectAllPrefs.Add_Click({ $panelPrefs.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] } | ForEach-Object { $_.Checked = $true } })

$btnRunPrefs = New-Button "Apply Tweaks" 130 120 $btnPanelPrefs
$tooltip.SetToolTip($btnRunPrefs, "Applies selected preferences.")
$btnRunPrefs.Add_Click({
    $selectedTweaks = $panelPrefs.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] -and $_.Checked }
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
        $lblStatusPrefs.Text = "Status: Applying tweak $currentTweak/$totalTweaks ($tweak)..."
        try {
            switch ($tweak) {
                "Dark Theme" { 
                    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0 -ErrorAction Stop 
                }
                "Disable Bing in Start" { 
                    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled" -Value 0 -ErrorAction Stop 
                }
                "Disable Start Recommendations" { 
                    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Start_IrisRecommendations" -Value 0 -ErrorAction Stop 
                }
                "Show Search Button" { 
                    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 1 -ErrorAction Stop 
                }
                "Disable Widgets" { 
                    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Widgets" -Force | Out-Null
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Widgets" -Name "AllowWidgets" -Value 0 -ErrorAction Stop 
                }
                "Disable Animations" { 
                    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "UserPreferencesMask" -Value ([byte[]](0x90,0x12,0x03,0x80,0x10,0x00,0x00,0x00)) -ErrorAction Stop
                    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop\WindowMetrics" -Name "MinAnimate" -Value 0 -ErrorAction Stop
                }
                "Disable Taskbar Transparency" { 
                    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "EnableTransparency" -Value 0 -ErrorAction Stop 
                }
                "Disable Snap Assist" { 
                    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WindowArrangementActive" -Value 0 -ErrorAction Stop 
                }
                "Ultimate Performance Profile" { 
                    powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61
                    $guid = (powercfg -l | Where-Object { $_ -match "Ultimate Performance" } | ForEach-Object { $_.Split()[3] })
                    if ($guid) { powercfg -s $guid }
                }
                "Optimize Visual Effects" {
                    $mask = [byte[]](0x90, 0x12, 0x07, 0x82, 0x10, 0x01, 0x00, 0x00)
                    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "UserPreferencesMask" -Value $mask -ErrorAction Stop
                    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop\WindowMetrics" -Name "MinAnimate" -Value 1 -ErrorAction Stop
                    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "FontSmoothing" -Value 2 -ErrorAction Stop
                }
            }
        } catch {
            $lblStatusPrefs.Text = "Status: Error on $tweak - $($_.Exception.Message)"
            $progressBarPrefs.Visible = $false
            return
        }
        $progressBarPrefs.Value = $currentTweak
        $form.Refresh()
    }
    $progressBarPrefs.Visible = $false
    $lblStatusPrefs.Text = "Status: Done"
})

# Win Updates Tab
$tabUpdates = Add-Tab "Win Updates"
$containerUpdates = New-Object System.Windows.Forms.Panel
$containerUpdates.Dock = "Fill"
$containerUpdates.BackColor = $form.BackColor
$tabUpdates.Controls.Add($containerUpdates)

$listUpdates = New-Object System.Windows.Forms.ListBox
$listUpdates.Location = New-Object System.Drawing.Point(10, 120)
$listUpdates.Size = New-Object System.Drawing.Size(540, 360)
$listUpdates.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$listUpdates.ForeColor = [System.Drawing.Color]::White
$listUpdates.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$containerUpdates.Controls.Add($listUpdates)

$lblStatusUpdates = New-StatusLabel 10
$progressBarUpdates = New-ProgressBar "Bottom"
$btnPanelUpdates = New-Object System.Windows.Forms.Panel
$btnPanelUpdates.Location = New-Object System.Drawing.Point(10, 40)
$btnPanelUpdates.Size = New-Object System.Drawing.Size(540, 80)
$btnPanelUpdates.BackColor = $form.BackColor
$containerUpdates.Controls.AddRange(@($listUpdates, $progressBarUpdates, $lblStatusUpdates, $btnPanelUpdates))

$btnLoadSecurity = New-Button "Load Security Updates" 0 150 $btnPanelUpdates
$btnLoadSecurity.Location = New-Object System.Drawing.Point(0, 5)
$btnInstallSecurity = New-Button "Install Security Updates" 160 150 $btnPanelUpdates
$btnInstallSecurity.Location = New-Object System.Drawing.Point(160, 5)
$btnLoadAll = New-Button "Load All Updates" 0 150 $btnPanelUpdates
$btnLoadAll.Location = New-Object System.Drawing.Point(0, 45)
$btnInstallAll = New-Button "Install All Updates" 160 150 $btnPanelUpdates
$btnInstallAll.Location = New-Object System.Drawing.Point(160, 45)

$script:securityUpdates = $null
$script:allUpdates = $null

$btnLoadSecurity.Add_Click({
    $lblStatusUpdates.Text = "Status: Loading security updates..."
    try {
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            $lblStatusUpdates.Text = "Status: Installing PSWindowsUpdate..."
            Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -Confirm:$false -ErrorAction Stop
        }
        Import-Module PSWindowsUpdate -ErrorAction Stop
        $listUpdates.Items.Clear()
        $script:securityUpdates = Get-WindowsUpdate -MicrosoftUpdate -Category "Security Updates" -ErrorAction Stop
        if ($script:securityUpdates) {
            $totalUpdates = $script:securityUpdates.Count
            $progressBarUpdates.Visible = $true
            $progressBarUpdates.Maximum = $totalUpdates
            $progressBarUpdates.Value = 0
            foreach ($update in $script:securityUpdates) {
                $progressBarUpdates.Value++
                $listUpdates.Items.Add("[$($update.KB)] $($update.Title)")
                $lblStatusUpdates.Text = "Status: Loading update $progressBarUpdates.Value/$totalUpdates..."
                $form.Refresh()
            }
            $lblStatusUpdates.Text = "Status: $totalUpdates security updates loaded"
        } else {
            $lblStatusUpdates.Text = "Status: No security updates found"
        }
    } catch {
        $lblStatusUpdates.Text = "Status: Error - $($_.Exception.Message)"
    }
    $progressBarUpdates.Visible = $false
})

$btnInstallSecurity.Add_Click({
    if (-not $script:securityUpdates) { $lblStatusUpdates.Text = "Status: No updates to install"; return }
    $totalUpdates = $script:securityUpdates.Count
    $progressBarUpdates.Visible = $true
    $progressBarUpdates.Maximum = $totalUpdates
    $progressBarUpdates.Value = 0
    $lblStatusUpdates.Text = "Status: Installing updates (0/$totalUpdates)..."
    try {
        foreach ($update in $script:securityUpdates) {
            $progressBarUpdates.Value++
            $lblStatusUpdates.Text = "Status: Installing update $progressBarUpdates.Value/$totalUpdates (KB$($update.KB))..."
            Install-WindowsUpdate -MicrosoftUpdate -KBArticleID $update.KB -AcceptAll -IgnoreReboot -ErrorAction Stop
            $form.Refresh()
        }
        $lblStatusUpdates.Text = "Status: Installation complete"
        $listUpdates.Items.Clear()
        $script:securityUpdates = $null
    } catch {
        $lblStatusUpdates.Text = "Status: Error - $($_.Exception.Message)"
    }
    $progressBarUpdates.Visible = $false
})

$btnLoadAll.Add_Click({
    $lblStatusUpdates.Text = "Status: Loading all updates..."
    try {
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            $lblStatusUpdates.Text = "Status: Installing PSWindowsUpdate..."
            Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -Confirm:$false -ErrorAction Stop
        }
        Import-Module PSWindowsUpdate -ErrorAction Stop
        $listUpdates.Items.Clear()
        $script:allUpdates = Get-WindowsUpdate -MicrosoftUpdate -ErrorAction Stop
        if ($script:allUpdates) {
            $totalUpdates = $script:allUpdates.Count
            $progressBarUpdates.Visible = $true
            $progressBarUpdates.Maximum = $totalUpdates
            $progressBarUpdates.Value = 0
            foreach ($update in $script:allUpdates) {
                $progressBarUpdates.Value++
                $listUpdates.Items.Add("[$($update.KB)] $($update.Title)")
                $lblStatusUpdates.Text = "Status: Loading update $progressBarUpdates.Value/$totalUpdates..."
                $form.Refresh()
            }
            $lblStatusUpdates.Text = "Status: $totalUpdates updates loaded"
        } else {
            $lblStatusUpdates.Text = "Status: No updates found"
        }
    } catch {
        $lblStatusUpdates.Text = "Status: Error - $($_.Exception.Message)"
    }
    $progressBarUpdates.Visible = $false
})

$btnInstallAll.Add_Click({
    if (-not $script:allUpdates) { $lblStatusUpdates.Text = "Status: No updates to install"; return }
    $totalUpdates = $script:allUpdates.Count
    $progressBarUpdates.Visible = $true
    $progressBarUpdates.Maximum = $totalUpdates
    $progressBarUpdates.Value = 0
    $lblStatusUpdates.Text = "Status: Installing updates (0/$totalUpdates)..."
    try {
        foreach ($update in $script:allUpdates) {
            $progressBarUpdates.Value++
            $lblStatusUpdates.Text = "Status: Installing update $progressBarUpdates.Value/$totalUpdates (KB$($update.KB))..."
            Install-WindowsUpdate -MicrosoftUpdate -KBArticleID $update.KB -AcceptAll -IgnoreReboot -ErrorAction Stop
            $form.Refresh()
        }
        $lblStatusUpdates.Text = "Status: Installation complete"
        $listUpdates.Items.Clear()
        $script:allUpdates = $null
    } catch {
        $lblStatusUpdates.Text = "Status: Error - $($_.Exception.Message)"
    }
    $progressBarUpdates.Visible = $false
})

# Startup Manager Tab
$tabStartup = Add-Tab "Startup Manager"
$containerStartup = New-Object System.Windows.Forms.Panel
$containerStartup.Dock = "Fill"
$containerStartup.BackColor = $form.BackColor
$tabStartup.Controls.Add($containerStartup)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Location = New-Object System.Drawing.Point(10, 80)
$lblFilter.Size = New-Object System.Drawing.Size(60, 25)
$lblFilter.Text = "Filter by:"
$lblFilter.ForeColor = [System.Drawing.Color]::White
$lblFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$containerStartup.Controls.Add($lblFilter)

$comboFilter = New-Object System.Windows.Forms.ComboBox
$comboFilter.Location = New-Object System.Drawing.Point(70, 80)
$comboFilter.Size = New-Object System.Drawing.Size(120, 25)
$comboFilter.Items.AddRange(@("All", "Current User", "Local Machine", "Enabled"))
$comboFilter.SelectedIndex = 0
$comboFilter.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$comboFilter.ForeColor = [System.Drawing.Color]::White
$comboFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$comboFilter.DropDownStyle = "DropDownList"
$tooltip.SetToolTip($comboFilter, "Filter startup items.")
$containerStartup.Controls.Add($comboFilter)

$clbStartup = New-Object System.Windows.Forms.CheckedListBox
$clbStartup.Size = New-Object System.Drawing.Size(540, 220)
$clbStartup.Location = New-Object System.Drawing.Point(10, 110)
$clbStartup.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$clbStartup.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$clbStartup.ForeColor = [System.Drawing.Color]::White
$clbStartup.CheckOnClick = $true
$containerStartup.Controls.Add($clbStartup)

$txtStartupDetails = New-Object System.Windows.Forms.TextBox
$txtStartupDetails.Location = New-Object System.Drawing.Point(10, 340)
$txtStartupDetails.Size = New-Object System.Drawing.Size(540, 150)
$txtStartupDetails.Multiline = $true
$txtStartupDetails.ReadOnly = $true
$txtStartupDetails.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$txtStartupDetails.ForeColor = [System.Drawing.Color]::White
$txtStartupDetails.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtStartupDetails.Text = "Select a startup item."
$tooltip.SetToolTip($txtStartupDetails, "Shows startup item details.")
$containerStartup.Controls.Add($txtStartupDetails)

$lblStartupStatus = New-StatusLabel 10
$btnPanelStartup = New-Object System.Windows.Forms.Panel
$btnPanelStartup.Location = New-Object System.Drawing.Point(10, 40)
$btnPanelStartup.Size = New-Object System.Drawing.Size(540, 40)
$btnPanelStartup.BackColor = $form.BackColor
$containerStartup.Controls.Add($btnPanelStartup)

$btnReloadStartup = New-Button "Reload" 0 140 $btnPanelStartup
$btnApplyStartupChanges = New-Button "Apply Changes" 150 140 $btnPanelStartup
$containerStartup.Controls.AddRange(@($clbStartup, $txtStartupDetails, $lblStartupStatus, $btnPanelStartup))

$tooltip.SetToolTip($btnApplyStartupChanges, "Applies startup changes.")
$tooltip.SetToolTip($btnReloadStartup, "Reloads startup items.")

$script:startupItems = @()
function Load-StartupItems {
    $clbStartup.Items.Clear()
    $script:startupItems = @()
    $registryPaths = @(
        @{ Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"; Hive = "HKCU" },
        @{ Path = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"; Hive = "HKLM" }
    )
    foreach ($reg in $registryPaths) {
        if (Test-Path $reg.Path) {
            $items = Get-ItemProperty -Path $reg.Path -ErrorAction SilentlyContinue
            foreach ($name in $items.PSObject.Properties.Name) {
                $command = $items.$name
                $publisher = "Unknown"
                $filePath = "Unknown"
                try {
                    if ($command -match '^"([^"]+)"' -or $command -match '^([^\s]+)') {
                        $exePath = $matches[1]
                        if (Test-Path $exePath -ErrorAction SilentlyContinue) {
                            $fileInfo = Get-Item $exePath -ErrorAction SilentlyContinue
                            $publisher = if ($fileInfo.VersionInfo.CompanyName) { $fileInfo.VersionInfo.CompanyName } else { "Unknown" }
                            $filePath = $exePath
                        }
                    }
                } catch {
                    $publisher = "Unknown"
                    $filePath = "Unknown"
                }
                $script:startupItems += [PSCustomObject]@{
                    Name = $name
                    Enabled = $true
                    Location = $reg.Path
                    Hive = $reg.Hive
                    Command = $command
                    Publisher = $publisher
                    FilePath = $filePath
                }
            }
        }
    }

    $filterCriterion = $comboFilter.SelectedItem
    $filteredItems = switch ($filterCriterion) {
        "Current User" { $script:startupItems | Where-Object { $_.Hive -eq "HKCU" } }
        "Local Machine" { $script:startupItems | Where-Object { $_.Hive -eq "HKLM" } }
        "Enabled" { $script:startupItems | Where-Object { $_.Enabled -eq $true } }
        default { $script:startupItems }
    }

    $filteredItems = $filteredItems | Sort-Object Name
    foreach ($item in $filteredItems) {
        $status = if ($item.Enabled) { "Enabled" } else { "Disabled" }
        $scope = if ($item.Hive -eq "HKCU") { "Current User" } else { "Local Machine" }
        $displayName = "$($item.Name) [$status, $scope]"
        $clbStartup.Items.Add($displayName, $item.Enabled) | Out-Null
    }
    $lblStartupStatus.Text = "Status: $($filteredItems.Count) startup items loaded"
}

$btnReloadStartup.Add_Click({Load-StartupItems})

$comboFilter.Add_SelectedIndexChanged({ Load-StartupItems })

$clbStartup.Add_SelectedIndexChanged({
    $selectedIndex = $clbStartup.SelectedIndex
    if ($selectedIndex -ge 0) {
        $filterCriterion = $comboFilter.SelectedItem
        $filteredItems = switch ($filterCriterion) {
            "Current User" { $script:startupItems | Where-Object { $_.Hive -eq "HKCU" } }
            "Local Machine" { $script:startupItems | Where-Object { $_.Hive -eq "HKLM" } }
            "Enabled" { $script:startupItems | Where-Object { $_.Enabled -eq $true } }
            default { $script:startupItems }
        }
        $filteredItems = $filteredItems | Sort-Object Name
        $item = $filteredItems[$selectedIndex]
        $details = @"
Name: $($item.Name)
Enabled: $($item.Enabled)
Location: $($item.Location)
Command: $($item.Command)
Publisher: $($item.Publisher)
File Path: $($item.FilePath)
"@
        $txtStartupDetails.Text = $details
    } else {
        $txtStartupDetails.Text = "Select a startup item."
    }
})

$btnApplyStartupChanges.Add_Click({
    $changesMade = $false
    $filterCriterion = $comboFilter.SelectedItem
    $filteredItems = switch ($filterCriterion) {
        "Current User" { $script:startupItems | Where-Object { $_.Hive -eq "HKCU" } }
        "Local Machine" { $script:startupItems | Where-Object { $_.Hive -eq "HKLM" } }
        "Enabled" { $script:startupItems | Where-Object { $_.Enabled -eq $true } }
        default { $script:startupItems }
    }
    $filteredItems = $filteredItems | Sort-Object Name
    for ($i = 0; $i -lt $clbStartup.Items.Count; $i++) {
        $item = $filteredItems[$i]
        $isChecked = $clbStartup.GetItemChecked($i)
        if (-not $isChecked -and $item.Enabled) {
            try {
                Remove-ItemProperty -Path $item.Location -Name $item.Name -ErrorAction Stop
                $item.Enabled = $false
                $changesMade = $true
            } catch {
                $lblStartupStatus.Text = "Status: Error disabling $($item.Name)"
            }
        } elseif ($isChecked -and -not $item.Enabled) {
            try {
                Set-ItemProperty -Path $item.Location -Name $item.Name -Value $item.Command -ErrorAction Stop
                $item.Enabled = $true
                $changesMade = $true
            } catch {
                $lblStartupStatus.Text = "Status: Error enabling $($item.Name)"
            }
        }
    }
    if ($changesMade) {
        $lblStartupStatus.Text = "Status: Changes applied"
        Load-StartupItems
    } else {
        $lblStartupStatus.Text = "Status: No changes made"
    }
})

Load-StartupItems

# Resize Event
$form.Add_Resize({
    $logoBox.Location = New-Object System.Drawing.Point((($form.Width - $logoBox.Width) / 2), 20)
    $lblWelcome.Size = New-Object System.Drawing.Size($form.Width, 50)
    $panelSysInfo.Location = New-Object System.Drawing.Point((($form.Width - $panelSysInfo.Width) / 2), 240)
})

# Show Form
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()