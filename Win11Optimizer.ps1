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
    @{ Text = "Disable Homegroup"; Tooltip = "Disables legacy Homegroup networking services." },
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
            "Disable Homegroup" { Set-Service -Name "HomeGroupListener", "HomeGroupProvider" -StartupType Disabled -ErrorAction SilentlyContinue }
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
                foreach ($svc in $global:services.GetEnumerator()) { 
                    Set-Service -Name $svc.Key -StartupType $svc.Value -ErrorAction SilentlyContinue 
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
            "Dark Theme for Windows" { Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0 -ErrorAction SilentlyContinue }
            "Disable Bing Search in Start Menu" { Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled" -Value 0 -ErrorAction SilentlyContinue }
            "Disable Recommendations in Start Menu" { Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Start_IrisRecommendations" -Value 0 -ErrorAction SilentlyContinue }
            "Search Button in Taskbar" { Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 1 -ErrorAction SilentlyContinue }
            "Disable Widgets in Taskbar" { Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Widgets" -Name "AllowWidgets" -Value 0 -ErrorAction SilentlyContinue }
            "Disable Animations" { 
                Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "UserPreferencesMask" -Value ([byte[]](0x90,0x12,0x03,0x80,0x10,0x00,0x00,0x00)) -ErrorAction SilentlyContinue
                Set-ItemProperty -Path "HKCU:\Control Panel\Desktop\WindowMetrics" -Name "MinAnimate" -Value 0
            }
            "Turn Off Taskbar Transparency" { Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "EnableTransparency" -Value 0 }
            "Disable Snap Assist" { Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WindowArrangementActive" -Value 0 }
            "Add and Activate Ultimate Performance Profile" { 
                powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61
                $guid = (powercfg -l | Where-Object { $_ -match "Ultimate Performance" } | ForEach-Object { $_.Split()[3] })
                if ($guid) { powercfg -s $guid }
            }
            "Improve Visual Effects for Performance" {
                # Adjust for best performance with specific effects enabled/disabled
                # - Animate windows when minimizing and maximizing: Enabled
                # - Show shadows under mouse pointer: Disabled
                # - Show thumbnails instead of icons: Enabled
                # - Show translucent selection rectangle: Disabled
                # - Smooth edges of screen fonts: Enabled
                $mask = [byte[]](0x90, 0x12, 0x07, 0x82, 0x10, 0x01, 0x00, 0x00)
                Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "UserPreferencesMask" -Value $mask -ErrorAction SilentlyContinue
                # Ensure MinAnimate is set to 1 for window animations
                Set-ItemProperty -Path "HKCU:\Control Panel\Desktop\WindowMetrics" -Name "MinAnimate" -Value 1 -ErrorAction SilentlyContinue
                # Ensure font smoothing is enabled
                Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "FontSmoothing" -Value 2 -ErrorAction SilentlyContinue
            }
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
$listUpdates.BorderStyle = "None"
$listUpdates.Font = New-Object System.Drawing.Font("Segoe Rodrigo", 9)
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
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        $lblStatusUpdates.Text = "Status: Installing PSWindowsUpdate module..."
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -Confirm:$false -ErrorAction SilentlyContinue
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) { $lblStatusUpdates.Text = "Status: Failed to install module"; return }
    }
    Import-Module PSWindowsUpdate
    $listUpdates.Items.Clear()
    $script:securityUpdates = Get-WindowsUpdate -MicrosoftUpdate -Category "Security Updates" -ErrorAction SilentlyContinue
    if ($script:securityUpdates) {
        $totalUpdates = $script:securityUpdates.Count
        $progressBarUpdates.Visible = $true
        $progressBarUpdates.Maximum = $totalUpdates
        $progressBarUpdates.Value = 0
        $currentUpdate = 0
        foreach ($update in $script:securityUpdates) {
            $currentUpdate++
            $lblStatusUpdates.Text = "Status: Loading update $currentUpdate of $totalUpdates..."
            $listUpdates.Items.Add("[$($update.KB)] $($update.Title)")
            $progressBarUpdates.Value = $currentUpdate
            $form.Refresh()
        }
        $lblStatusUpdates.Text = "Status: $totalUpdates security updates loaded"
    } else {
        $lblStatusUpdates.Text = "Status: No security updates found"
    }
    $progressBarUpdates.Visible = $false
})

$btnInstallSecurity.Add_Click({
    if ($script:securityUpdates) {
        $totalUpdates = $script:securityUpdates.Count
        $progressBarUpdates.Visible = $true
        $progressBarUpdates.Maximum = $totalUpdates
        $progressBarUpdates.Value = 0
        $currentUpdate = 0
        $lblStatusUpdates.Text = "Status: Installing updates (0/$totalUpdates)..."
        foreach ($update in $script:securityUpdates) {
            $currentUpdate++
            $lblStatusUpdates.Text = "Status: Installing update $currentUpdate of $totalUpdates (KB$($update.KB))..."
            Install-WindowsUpdate -MicrosoftUpdate -KBArticleID $update.KB -AcceptAll -IgnoreReboot -ErrorAction SilentlyContinue
            $progressBarUpdates.Value = $currentUpdate
            $form.Refresh()
        }
        $lblStatusUpdates.Text = "Status: Installation complete"
        $listUpdates.Items.Clear()
        $script:securityUpdates = $null
    } else {
        $lblStatusUpdates.Text = "Status: No security updates to install"
    }
    $progressBarUpdates.Visible = $false
})

$btnLoadAll.Add_Click({
    $lblStatusUpdates.Text = "Status: Loading all updates..."
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        $lblStatusUpdates.Text = "Status: Installing PSWindowsUpdate module..."
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -Confirm:$false -ErrorAction SilentlyContinue
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) { $lblStatusUpdates.Text = "Status: Failed to install module"; return }
    }
    Import-Module PSWindowsUpdate
    $listUpdates.Items.Clear()
    $script:allUpdates = Get-WindowsUpdate -MicrosoftUpdate -ErrorAction SilentlyContinue
    if ($script:allUpdates) {
        $totalUpdates = $script:allUpdates.Count
        $progressBarUpdates.Visible = $true
        $progressBarUpdates.Maximum = $totalUpdates
        $progressBarUpdates.Value = 0
        $currentUpdate = 0
        foreach ($update in $script:allUpdates) {
            $currentUpdate++
            $lblStatusUpdates.Text = "Status: Loading update $currentUpdate of $totalUpdates..."
            $listUpdates.Items.Add("[$($update.KB)] $($update.Title)")
            $progressBarUpdates.Value = $currentUpdate
            $form.Refresh()
        }
        $lblStatusUpdates.Text = "Status: $totalUpdates updates loaded"
    } else {
        $lblStatusUpdates.Text = "Status: No updates found"
    }
    $progressBarUpdates.Visible = $false
})

$btnInstallAll.Add_Click({
    if ($script:allUpdates) {
        $totalUpdates = $script:allUpdates.Count
        $progressBarUpdates.Visible = $true
        $progressBarUpdates.Maximum = $totalUpdates
        $progressBarUpdates.Value = 0
        $currentUpdate = 0
        $lblStatusUpdates.Text = "Status: Installing updates (0/$totalUpdates)..."
        foreach ($update in $script:allUpdates) {
            $currentUpdate++
            $lblStatusUpdates.Text = "Status: Installing update $currentUpdate of $totalUpdates (KB$($update.KB))..."
            Install-WindowsUpdate -MicrosoftUpdate -KBArticleID $update.KB -AcceptAll -IgnoreReboot -ErrorAction SilentlyContinue
            $progressBarUpdates.Value = $currentUpdate
            $form.Refresh()
        }
        $lblStatusUpdates.Text = "Status: Installation complete"
        $listUpdates.Items.Clear()
        $script:allUpdates = $null
    } else {
        $lblStatusUpdates.Text = "Status: No updates to install"
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
$lblFilter.Font = New-Object System.Drawing.Font("Segoe Rodrigo", 9)
$containerStartup.Controls.Add($lblFilter)

$comboFilter = New-Object System.Windows.Forms.ComboBox
$comboFilter.Location = New-Object System.Drawing.Point(70, 80)
$comboFilter.Size = New-Object System.Drawing.Size(120, 25)
$comboFilter.Items.AddRange(@("Current User", "Local Machine", "Enabled"))
$comboFilter.SelectedIndex = 0
$comboFilter.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$comboFilter.ForeColor = [System.Drawing.Color]::White
$comboFilter.Font = New-Object System.Drawing.Font("Segoe Rodrigo", 9)
$comboFilter.DropDownStyle = "DropDownList"
$tooltip.SetToolTip($comboFilter, "Filter startup items by Current User (HKCU), Local Machine (HKLM), or Enabled status.")
$containerStartup.Controls.Add($comboFilter)

$clbStartup = New-Object System.Windows.Forms.CheckedListBox
$clbStartup.Size = New-Object System.Drawing.Size(540, 220)
$clbStartup.Location = New-Object System.Drawing.Point(10, 110)
$clbStartup.Font = New-Object System.Drawing.Font("Segoe Rodrigo", 9)
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
$txtStartupDetails.Font = New-Object System.Drawing.Font("Segoe Rodrigo", 9)
$txtStartupDetails.Text = "Select a startup item to view details."
$tooltip.SetToolTip($txtStartupDetails, "Displays details for the selected startup item.")
$containerStartup.Controls.Add($txtStartupDetails)

$lblStartupStatus = New-StatusLabel 10
$btnPanelStartup = New-Object System.Windows.Forms.Panel
$btnPanelStartup.Location = New-Object System.Drawing.Point(10, 40)
$btnPanelStartup.Size = New-Object System.Drawing.Size(540, 40)
$btnPanelStartup.BackColor = $form.BackColor
$containerStartup.Controls.AddRange(@($clbStartup, $txtStartupDetails, $lblStartupStatus, $btnPanelStartup))

$btnApplyStartupChanges = New-Button "Apply Changes" 0 140 $btnPanelStartup
$tooltip.SetToolTip($btnApplyStartupChanges, "Applies changes to enable or disable selected startup items.")

$script:startupItems = @()
function Load-StartupItems {
    $clbStartup.Items.Clear()
    $script:startupItems = @()
    $registryPaths = @(
        @{ Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"; Hive = "HKCU" },
        @{ Path = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"; Hive = "HKLM" },
        @{ Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"; Hive = "HKCU" },
        @{ Path = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"; Hive = "HKLM" }
    )
    $approvedApps = @{}
    foreach ($reg in $registryPaths) {
        if ($reg.Path -like "*StartupApproved*") {
            if (Test-Path $reg.Path) {
                $items = Get-ItemProperty -Path $reg.Path
                foreach ($prop in $items.PSObject.Properties) {
                    $value = $items.$($prop.Name)
                    if ($value.Length -ge 1) { $approvedApps[$prop.Name] = ($value[0] -eq 2) }
                }
            }
        }
    }
    foreach ($reg in $registryPaths) {
        if ($reg.Path -notlike "*StartupApproved*") {
            if (Test-Path $reg.Path) {
                $items = Get-ItemProperty -Path $reg.Path | Select-Object -Property * -ExcludeProperty PSPath, PSParentPath, PSChildName, PSDrive, PSProvider
                foreach ($name in $items.PSObject.Properties.Name) {
                    $command = $items.$name
                    $publisher = "Unknown"
                    $filePath = "Unknown"
                    try {
                        if ($command -match '^"([^"]+)"' -or $command -match '^([^\s]+)') {
                            $exePath = $matches[1]
                            if (Test-Path $exePath) {
                                $fileInfo = Get-Item $exePath
                                $publisher = if ((Get-ItemProperty $exePath -ErrorAction SilentlyContinue).VersionInfo.CompanyName) { (Get-ItemProperty $exePath).VersionInfo.CompanyName } else { "Unknown" }
                                $filePath = $exePath
                            }
                        }
                    } catch {
                        $publisher = "Error retrieving publisher"
                        $filePath = "Error retrieving file path"
                    }
                    $enabled = if ($approvedApps.ContainsKey($name)) { $approvedApps[$name] } else { $true }
                    $script:startupItems += [PSCustomObject]@{
                        Name = $name
                        Enabled = $enabled
                        Location = $reg.Path
                        Hive = $reg.Hive
                        Command = $command
                        OriginalCommand = $command
                        Publisher = $publisher
                        FilePath = $filePath
                    }
                }
            }
        }
    }

    # Filter items based on selected filter criterion
    $filterCriterion = $comboFilter.SelectedItem
    $filteredItems = switch ($filterCriterion) {
        "Current User" { $script:startupItems | Where-Object { $_.Hive -eq "HKCU" } }
        "Local Machine" { $script:startupItems | Where-Object { $_.Hive -eq "HKLM" } }
        "Enabled" { $script:startupItems | Where-Object { $_.Enabled -eq $true } }
        default { $script:startupItems | Where-Object { $_.Hive -eq "HKCU" } }
    }

    # Sort filtered items by Name for consistency
    $filteredItems = $filteredItems | Sort-Object Name

    # Populate CheckedListBox
    foreach ($item in $filteredItems) {
        $clbStartup.Items.Add($item.Name, $item.Enabled) | Out-Null
    }
    $lblStartupStatus.Text = "Status: $($filteredItems.Count) startup items loaded"
}

$comboFilter.Add_SelectedIndexChanged({
    Load-StartupItems
})

$clbStartup.Add_SelectedIndexChanged({
    $selectedIndex = $clbStartup.SelectedIndex
    if ($selectedIndex -ge 0) {
        $filterCriterion = $comboFilter.SelectedItem
        $filteredItems = switch ($filterCriterion) {
            "Current User" { $script:startupItems | Where-Object { $_.Hive -eq "HKCU" } }
            "Local Machine" { $script:startupItems | Where-Object { $_.Hive -eq "HKLM" } }
            "Enabled" { $script:startupItems | Where-Object { $_.Enabled -eq $true } }
            default { $script:startupItems | Where-Object { $_.Hive -eq "HKCU" } }
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
        $txtStartupDetails.Text = "Select a startup item to view details."
    }
})

$btnApplyStartupChanges.Add_Click({
    $changesMade = $false
    $filterCriterion = $comboFilter.SelectedItem
    $filteredItems = switch ($filterCriterion) {
        "Current User" { $script:startupItems | Where-Object { $_.Hive -eq "HKCU" } }
        "Local Machine" { $script:startupItems | Where-Object { $_.Hive -eq "HKLM" } }
        "Enabled" { $script:startupItems | Where-Object { $_.Enabled -eq $true } }
        default { $script:startupItems | Where-Object { $_.Hive -eq "HKCU" } }
    }
    $filteredItems = $filteredItems | Sort-Object Name
    for ($i = 0; $i -lt $clbStartup.Items.Count; $i++) {
        $item = $filteredItems[$i]
        $isChecked = $clbStartup.GetItemChecked($i)
        if (-not $isChecked -and $item.Enabled) {
            Remove-ItemProperty -Path $item.Location -Name $item.Name -ErrorAction SilentlyContinue
            $item.Enabled = $false
            $changesMade = $true
        } elseif ($isChecked -and -not $item.Enabled -and $item.OriginalCommand) {
            try {
                Set-ItemProperty -Path $item.Location -Name $item.Name -Value $item.OriginalCommand -ErrorAction Stop
                $item.Enabled = $true
                $changesMade = $true
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to re-enable '$($item.Name)': $_", "Error")
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
    $btnRunSpeed.Location = New-Object System.Drawing.Point(($btnPanelSpeed.Width - $btnRunSpeed.Width - 10), 5)
    $btnSelectAllSpeed.Location = New-Object System.Drawing.Point(($btnPanelSpeed.Width - $btnRunSpeed.Width - $btnSelectAllSpeed.Width - 20), 5)
    $btnRunPrefs.Location = New-Object System.Drawing.Point(($btnPanelPrefs.Width - $btnRunPrefs.Width - 10), 5)
    $btnSelectAllPrefs.Location = New-Object System.Drawing.Point(($btnPanelPrefs.Width - $btnRunPrefs.Width - $btnSelectAllPrefs.Width - 20), 5)
    $btnApplyStartupChanges.Location = New-Object System.Drawing.Point(($btnPanelStartup.Width - $btnApplyStartupChanges.Width - 10), 5)
})

# Show Form
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
