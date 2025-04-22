# Win11Optimizer.ps1 - Optimized Version with Non-Overlapping Status Labels and Consistent Button Layout

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

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
function New-Checkbox($text, $yPos, $panel) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Location = New-Object System.Drawing.Point(10, $yPos)
    $cb.Size = New-Object System.Drawing.Size(380, 25)
    $cb.Text = $text
    $cb.ForeColor = [System.Drawing.Color]::White
    $cb.BackColor = $form.BackColor
    $cb.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cb.Anchor = "Left,Right"
    $panel.Controls.Add($cb)
    return $cb
}

function New-StatusLabel($yPos) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(10, $yPos)
    $lbl.AutoSize = $true
    $lbl.Height = 30
    $lbl.Text = "Status: Ready"
    $lbl.ForeColor = [System.Drawing.Color]::White
    $lbl.BackColor = $form.BackColor
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
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
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    $panel.Controls.Add($btn)
    return $btn
}

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
$speedCheckboxes = @("Delete Temp Files", "Disable Consumer Features", "Disable Telemetry", "Disable Activity History", "Disable GameDVR", "Disable Hibernation", "Disable Homegroup", "Disable Location Tracking", "Disable Storage Sense", "Disable Wifi-Sense", "Disable Recall", "Debloat Edge", "Disable Background Applications", "Clear Delivery Optimization Cache", "Set Services")
foreach ($check in $speedCheckboxes) {
    New-Checkbox $check $yPos $panelSpeed | Out-Null
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
$btnSelectAllSpeed.Add_Click({ foreach ($control in $panelSpeed.Controls) { if ($control -is [System.Windows.Forms.CheckBox]) { $control.Checked = $true } } })

$btnRunSpeed = New-Button "Apply Tweaks" 130 120 $btnPanelSpeed
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
                $services = @{
                    "AssignedAccessManagerSvc"="Disabled"; "BDESVC"="Disabled"; "BthServ"="Automatic"; "DiagTrack"="Disabled"; "DmEnrollmentSvc"="Manual"; "DPS"="Manual"; "Fax"="Disabled"; 
                    "FontCache"="Manual"; "HotspotService"="Disabled"; "lfsvc"="Disabled"; "MapsBroker"="Manual"; "stisvc"="Manual"; "SysMain"="Manual"; "TermService"="Disabled"; 
                    "TrkWks"="Manual"; "WlanSvc"="Manual"; "WManSvc"="Manual"; "wuauserv"="Manual"; "WSearch"="Manual"; "XblAuthManager"="Disabled"; "XblGameSave"="Disabled"; 
                    "XboxNetApiSvc"="Disabled"; "RetailDemo"="Disabled"; "RemoteRegisty"="Disabled"; "PhoneSvc"="Disabled"; "wisvc"="Disabled"; "icssvc"="Disabled"; "seclogon"="Disabled"
                }
                foreach ($svc in $services.GetEnumerator()) { Set-Service -Name $svc.Key -StartupType $svc.Value -ErrorAction SilentlyContinue }
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
$prefsCheckboxes = @("Dark Theme for Windows", "Disable Bing Search in Start Menu", "Disable Recommendations in Start Menu", "Search Button in Taskbar", "Disable Widgets in Taskbar", "Disable Animations", "Turn Off Taskbar Transparency", "Disable Snap Assist", "Add and Activate Ultimate Performance Profile")
foreach ($check in $prefsCheckboxes) {
    New-Checkbox $check $yPos $panelPrefs | Out-Null
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
$btnSelectAllPrefs.Add_Click({ foreach ($control in $panelPrefs.Controls) { if ($control -is [System.Windows.Forms.CheckBox]) { $control.Checked = $true } } })

$btnRunPrefs = New-Button "Apply Tweaks" 130 120 $btnPanelPrefs
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
        }
        $progressBarPrefs.Value = $currentTweak
        $form.Refresh()
    }
    $progressBarPrefs.Visible = $false
    $lblStatusPrefs.Text = "Status: Done"
})

# Uninstall Apps Tab
$tabUninstall = Add-Tab "Uninstall Apps"
$containerUninstall = New-Object System.Windows.Forms.Panel
$containerUninstall.Dock = "Fill"
$containerUninstall.BackColor = $form.BackColor
$tabUninstall.Controls.Add($containerUninstall)

$clbUninstallApps = New-Object System.Windows.Forms.CheckedListBox
$clbUninstallApps.Size = New-Object System.Drawing.Size(540, 300)
$clbUninstallApps.Location = New-Object System.Drawing.Point(10, 80)
$clbUninstallApps.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$clbUninstallApps.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$clbUninstallApps.ForeColor = [System.Drawing.Color]::White
$clbUninstallApps.CheckOnClick = $true
$containerUninstall.Controls.Add($clbUninstallApps)

$lblStatusUninstall = New-StatusLabel 10
$btnPanelUninstall = New-Object System.Windows.Forms.Panel
$btnPanelUninstall.Location = New-Object System.Drawing.Point(10, 40)
$btnPanelUninstall.Size = New-Object System.Drawing.Size(540, 40)
$btnPanelUninstall.BackColor = $form.BackColor
$containerUninstall.Controls.AddRange(@($lblStatusUninstall, $btnPanelUninstall))

$btnRefresh = New-Button "Refresh List" 0 120 $btnPanelUninstall
$btnUninstall = New-Button "Uninstall Selected" 130 150 $btnPanelUninstall

$script:uninstallItems = @()
function Load-UninstallApps {
    $clbUninstallApps.Items.Clear()
    $script:uninstallItems = @()
    $apps = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
        Where-Object { $_.DisplayName -and $_.UninstallString } |
        Sort-Object DisplayName
    foreach ($app in $apps) {
        $displayText = "$($app.DisplayName) (Version: $($app.DisplayVersion))"
        $script:uninstallItems += [PSCustomObject]@{
            Name = $app.DisplayName
            UninstallString = $app.UninstallString
            DisplayText = $displayText
        }
        $clbUninstallApps.Items.Add($displayText, $false) | Out-Null
    }
    $lblStatusUninstall.Text = "Status: $($script:uninstallItems.Count) apps loaded"
}

$btnRefresh.Add_Click({ Load-UninstallApps })

$btnUninstall.Add_Click({
    $checkedItems = $clbUninstallApps.CheckedItems
    if ($checkedItems.Count -eq 0) {
        $lblStatusUninstall.Text = "Status: No apps selected"
        return
    }
    $totalApps = $checkedItems.Count
    $currentApp = 0
    foreach ($item in $checkedItems) {
        $currentApp++
        $app = $script:uninstallItems | Where-Object { $_.DisplayText -eq $item }
        if ($app) {
            $lblStatusUninstall.Text = "Status: Uninstalling app $currentApp of $totalApps ($($app.Name))..."
            try {
                Start-Process "cmd.exe" "/c $($app.UninstallString)" -Verb RunAs -Wait -ErrorAction Stop
                $lblStatusUninstall.Text = "Status: Uninstalled $($app.Name)"
            } catch {
                $lblStatusUninstall.Text = "Status: Error uninstalling $($app.Name)"
            }
            $form.Refresh()
        }
    }
    $lblStatusUninstall.Text = "Status: Uninstall complete"
    Load-UninstallApps
})

# Auto-load uninstall apps when tab is selected
$tabControl.Add_SelectedIndexChanged({
    if ($tabControl.SelectedTab -eq $tabUninstall) {
        Load-UninstallApps
    }
})

# Install Apps Tab
$tabInstall = Add-Tab "Install Apps"
$containerInstall = New-Object System.Windows.Forms.Panel
$containerInstall.Dock = "Fill"
$containerInstall.BackColor = $form.BackColor
$tabInstall.Controls.Add($containerInstall)

$clbInstallApps = New-Object System.Windows.Forms.CheckedListBox
$clbInstallApps.Location = New-Object System.Drawing.Point(10, 80)
$clbInstallApps.Size = New-Object System.Drawing.Size(540, 400)
$clbInstallApps.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$clbInstallApps.ForeColor = [System.Drawing.Color]::White
$clbInstallApps.BorderStyle = "None"
$clbInstallApps.CheckOnClick = $true
$clbInstallApps.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lblStatusInstall = New-StatusLabel 10
$progressBarInstall = New-ProgressBar "Bottom"
$btnPanelInstall = New-Object System.Windows.Forms.Panel
$btnPanelInstall.Location = New-Object System.Drawing.Point(10, 40)
$btnPanelInstall.Size = New-Object System.Drawing.Size(540, 40)
$btnPanelInstall.BackColor = $form.BackColor
$containerInstall.Controls.AddRange(@($clbInstallApps, $progressBarInstall, $lblStatusInstall, $btnPanelInstall))

$appList = @(
    @{ Name = "7-Zip"; Url = "https://www.7-zip.org/a/7z2301-x64.exe" },
    @{ Name = "RingCentral"; Url = "https://app.ringcentral.com/download/RingCentral.exe?V=20138739841159500" },
    @{ Name = "Google Chrome"; Url = "https://dl.google.com/chrome/install/ChromeSetup.exe" },
    @{ Name = "Microsoft Teams"; Url = "https://go.microsoft.com/fwlink/?linkid=2281613&clcid=0x409&culture=en-us&country=us" }
)
foreach ($app in $appList) { $clbInstallApps.Items.Add($app.Name) | Out-Null }

$btnSelectAllInstall = New-Button "Select All" 0 120 $btnPanelInstall
$btnSelectAllInstall.Add_Click({ for ($i = 0; $i -lt $clbInstallApps.Items.Count; $i++) { $clbInstallApps.SetItemChecked($i, $true) } })

$btnInstallSelected = New-Button "Install Selected" 130 120 $btnPanelInstall
$btnInstallSelected.Add_Click({
    $selectedApps = $clbInstallApps.CheckedItems
    $totalApps = $selectedApps.Count
    if ($totalApps -eq 0) { $lblStatusInstall.Text = "Status: No apps selected"; return }

    $progressBarInstall.Visible = $true
    $progressBarInstall.Maximum = $totalApps
    $progressBarInstall.Value = 0
    $lblStatusInstall.Text = "Status: Installing apps (0/$totalApps)..."
    $currentApp = 0

    foreach ($appName in $selectedApps) {
        $currentApp++
        $app = $appList | Where-Object { $_.Name -eq $appName }
        if ($app) {
            $lblStatusInstall.Text = "Status: Installing app $currentApp of $totalApps ($appName)..."
            $downloadPath = "$env:TEMP\$($app.Name)_installer.exe"
            try {
                Invoke-WebRequest -Uri $app.Url -OutFile $downloadPath -ErrorAction Stop
                Start-Process -FilePath $downloadPath -ArgumentList "/S" -Wait -NoNewWindow -ErrorAction Stop
                Remove-Item $downloadPath -Force -ErrorAction SilentlyContinue
            } catch {
                $lblStatusInstall.Text = "Status: Error installing $appName"
            }
        }
        $progressBarInstall.Value = $currentApp
        $form.Refresh()
    }
    $progressBarInstall.Visible = $false
    $lblStatusInstall.Text = "Status: Done"
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
$listUpdates.Font = New-Object System.Drawing.Font("Segoe UI", 9)

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
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction SilentlyContinue
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
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction SilentlyContinue
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

$clbStartup = New-Object System.Windows.Forms.CheckedListBox
$clbStartup.Size = New-Object System.Drawing.Size(540, 300)
$clbStartup.Location = New-Object System.Drawing.Point(10, 80)
$clbStartup.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$clbStartup.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$clbStartup.ForeColor = [System.Drawing.Color]::White
$clbStartup.CheckOnClick = $true

$lblStartupStatus = New-StatusLabel 10
$btnPanelStartup = New-Object System.Windows.Forms.Panel
$btnPanelStartup.Location = New-Object System.Drawing.Point(10, 40)
$btnPanelStartup.Size = New-Object System.Drawing.Size(540, 40)
$btnPanelStartup.BackColor = $form.BackColor
$containerStartup.Controls.AddRange(@($clbStartup, $lblStartupStatus, $btnPanelStartup))

$btnApplyStartupChanges = New-Button "Apply Changes" 0 140 $btnPanelStartup

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
                $items = Get-ItemProperty $reg.Path | Select-Object -Property * -ExcludeProperty PSPath, PSParentPath, PSChildName, PSDrive, PSProvider
                foreach ($name in $items.PSObject.Properties.Name) {
                    $enabled = if ($approvedApps.ContainsKey($name)) { $approvedApps[$name] } else { $true }
                    $script:startupItems += [PSCustomObject]@{ Name = $name; Enabled = $enabled; Location = $reg.Path; Hive = $reg.Hive }
                    $clbStartup.Items.Add($name, $enabled) | Out-Null
                }
            }
        }
    }
    $lblStartupStatus.Text = "Status: $($script:startupItems.Count) startup items loaded"
}

$btnApplyStartupChanges.Add_Click({
    for ($i = 0; $i -lt $clbStartup.Items.Count; $i++) {
        $item = $script:startupItems[$i]
        $isChecked = $clbStartup.GetItemChecked($i)
        if (-not $isChecked -and $item.Enabled) {
            Remove-ItemProperty -Path $item.Location -Name $item.Name -ErrorAction SilentlyContinue
            $item.Enabled = $false
        } elseif ($isChecked -and -not $item.Enabled) {
            [System.Windows.Forms.MessageBox]::Show("Cannot re-enable '$($item.Name)' without knowing the original command.", "Warning")
        }
    }
    $lblStartupStatus.Text = "Status: Changes applied"
})

Load-StartupItems

# Resize Event
$form.Add_Resize({
    $btnRunSpeed.Location = New-Object System.Drawing.Point(($btnPanelSpeed.Width - $btnRunSpeed.Width - 10), 5)
    $btnSelectAllSpeed.Location = New-Object System.Drawing.Point(($btnPanelSpeed.Width - $btnRunSpeed.Width - $btnSelectAllSpeed.Width - 20), 5)
    $btnRunPrefs.Location = New-Object System.Drawing.Point(($btnPanelPrefs.Width - $btnRunPrefs.Width - 10), 5)
    $btnSelectAllPrefs.Location = New-Object System.Drawing.Point(($btnPanelPrefs.Width - $btnRunPrefs.Width - $btnSelectAllPrefs.Width - 20), 5)
    $btnUninstall.Location = New-Object System.Drawing.Point(($btnPanelUninstall.Width - $btnUninstall.Width - 10), 5)
    $btnRefresh.Location = New-Object System.Drawing.Point(($btnPanelUninstall.Width - $btnUninstall.Width - $btnRefresh.Width - 20), 5)
    $btnInstallSelected.Location = New-Object System.Drawing.Point(($btnPanelInstall.Width - $btnInstallSelected.Width - 10), 5)
    $btnSelectAllInstall.Location = New-Object System.Drawing.Point(($btnPanelInstall.Width - $btnInstallSelected.Width - $btnSelectAllInstall.Width - 20), 5)
    $btnInstallSecurity.Location = New-Object System.Drawing.Point(($btnPanelUpdates.Width - $btnInstallSecurity.Width - 10), 5)
    $btnLoadSecurity.Location = New-Object System.Drawing.Point(($btnPanelUpdates.Width - $btnInstallSecurity.Width - $btnLoadSecurity.Width - 20), 5)
    $btnInstallAll.Location = New-Object System.Drawing.Point(($btnPanelUpdates.Width - $btnInstallAll.Width - 10), 45)
    $btnLoadAll.Location = New-Object System.Drawing.Point(($btnPanelUpdates.Width - $btnInstallAll.Width - $btnLoadAll.Width - 20), 45)
    $btnApplyStartupChanges.Location = New-Object System.Drawing.Point(($btnPanelStartup.Width - $btnApplyStartupChanges.Width - 10), 5)
})

# Show Form
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()