# Win11Optimizer.ps1
# Ensure assemblies are loaded for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Win11 Optimizer"
$form.Size = New-Object System.Drawing.Size(600,700)
$form.MinimumSize = New-Object System.Drawing.Size(500,600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38) # Dark background
$form.ForeColor = [System.Drawing.Color]::White

# Create tab control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"
$tabControl.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$tabControl.ForeColor = [System.Drawing.Color]::White
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$tabControl.Padding = New-Object System.Drawing.Point(10, 5)
$form.Controls.Add($tabControl)

# Speed Up PC tab
$tabSpeed = New-Object System.Windows.Forms.TabPage
$tabSpeed.Text = "Speed Up PC"
$tabSpeed.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)

$panelSpeed = New-Object System.Windows.Forms.Panel
$panelSpeed.Dock = "Top"
$panelSpeed.Height = 450
$panelSpeed.AutoScroll = $true
$panelSpeed.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$tabSpeed.Controls.Add($panelSpeed)

$yPos = 10
$speedCheckboxes = @("Delete Temp Files", "Disable Consumer Features", "Disable Telemetry", "Disable Activity History", "Disable GameDVR", "Disable Hibernation", "Disable Homegroup", "Disable Location Tracking", "Disable Storage Sense", "Disable Wifi-Sense", "Disable Recall", "Debloat Edge", "Disable Background Applications","Clear Delivery Optimization Cache", "Set Services")
foreach ($check in $speedCheckboxes) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Location = New-Object System.Drawing.Point(10, $yPos)
    $cb.Size = New-Object System.Drawing.Size(380, 25)
    $cb.Text = $check
    $cb.ForeColor = [System.Drawing.Color]::White
    $cb.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
    $cb.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $cb.Anchor = "Left,Right"
    $panelSpeed.Controls.Add($cb)
    $yPos += 28
}

# Force scroll to the bottom to make "Set Services" visible
$panelSpeed.VerticalScroll.Value = $panelSpeed.VerticalScroll.Maximum
$panelSpeed.PerformLayout()

$lblStatusSpeed = New-Object System.Windows.Forms.Label
$lblStatusSpeed.Dock = "Top"
$lblStatusSpeed.Height = 30
$lblStatusSpeed.Text = "Status: Ready"
$lblStatusSpeed.ForeColor = [System.Drawing.Color]::White
$lblStatusSpeed.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$lblStatusSpeed.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$lblStatusSpeed.TextAlign = "MiddleLeft"
$tabSpeed.Controls.Add($lblStatusSpeed)

# Add a ProgressBar below the status label
$progressBarSpeed = New-Object System.Windows.Forms.ProgressBar
$progressBarSpeed.Dock = "Top"
$progressBarSpeed.Height = 20
$progressBarSpeed.Minimum = 0
$progressBarSpeed.Maximum = 100
$progressBarSpeed.Value = 0
$progressBarSpeed.Visible = $false
$progressBarSpeed.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$progressBarSpeed.ForeColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$tabSpeed.Controls.Add($progressBarSpeed)

$btnPanelSpeed = New-Object System.Windows.Forms.Panel
$btnPanelSpeed.Dock = "Bottom"
$btnPanelSpeed.Height = 40
$btnPanelSpeed.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$tabSpeed.Controls.Add($btnPanelSpeed)

$btnSelectAllSpeed = New-Object System.Windows.Forms.Button
$btnSelectAllSpeed.Location = New-Object System.Drawing.Point(10,5)
$btnSelectAllSpeed.Size = New-Object System.Drawing.Size(120,30)
$btnSelectAllSpeed.Text = "Select All"
$btnSelectAllSpeed.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$btnSelectAllSpeed.ForeColor = [System.Drawing.Color]::White
$btnSelectAllSpeed.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSelectAllSpeed.FlatStyle = "Flat"
$btnSelectAllSpeed.FlatAppearance.BorderSize = 0
$btnSelectAllSpeed.Add_Click({ foreach ($control in $panelSpeed.Controls) { if ($control -is [System.Windows.Forms.CheckBox]) { $control.Checked = $true } } })
$btnPanelSpeed.Controls.Add($btnSelectAllSpeed)

$btnRunSpeed = New-Object System.Windows.Forms.Button
$btnRunSpeed.Location = New-Object System.Drawing.Point(290,5)
$btnRunSpeed.Size = New-Object System.Drawing.Size(120,30)
$btnRunSpeed.Text = "Apply Tweaks"
$btnRunSpeed.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$btnRunSpeed.ForeColor = [System.Drawing.Color]::White
$btnRunSpeed.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnRunSpeed.FlatStyle = "Flat"
$btnRunSpeed.FlatAppearance.BorderSize = 0
$btnRunSpeed.Anchor = "Right"
$btnRunSpeed.Add_Click({ 
    # Count the number of selected tweaks
    $selectedTweaks = @($panelSpeed.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] -and $_.Checked })
    $totalTweaks = $selectedTweaks.Count

    if ($totalTweaks -eq 0) {
        $lblStatusSpeed.Text = "Status: No tweaks selected"
        return
    }

    # Show the progress bar and set its maximum
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
            "Delete Temp Files" { 
                Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue 
            } 
            "Disable Consumer Features" { 
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableWindowsConsumerFeatures" -Value 1 -ErrorAction SilentlyContinue 
            }
            "Disable Telemetry" { 
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0 -ErrorAction SilentlyContinue 
            }
            "Disable Activity History" { 
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -Name "EnableActivityFeed" -Value 0 -ErrorAction SilentlyContinue 
            }
            "Disable GameDVR" { 
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR" -Name "AllowGameDVR" -Value 0 -ErrorAction SilentlyContinue 
            }
            "Disable Hibernation" { 
                powercfg /hibernate off 
            }
            "Disable Homegroup" { 
                Set-Service -Name "HomeGroupListener" -StartupType Disabled -ErrorAction SilentlyContinue 
                Set-Service -Name "HomeGroupProvider" -StartupType Disabled -ErrorAction SilentlyContinue 
            }
            "Disable Location Tracking" { 
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors" -Name "DisableLocation" -Value 1 -ErrorAction SilentlyContinue 
            }
            "Disable Storage Sense" { 
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\StorageSense" -Name "AllowStorageSenseGlobal" -Value 0 -ErrorAction SilentlyContinue 
            }
            "Disable Wifi-Sense" { 
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config" -Name "AutoConnectAllowedOEM" -Value 0 -ErrorAction SilentlyContinue 
            }
            "Disable Recall" { 
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" -Name "DisableRecall" -Value 1 -ErrorAction SilentlyContinue 
            }
            "Debloat Edge" { 
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "HideFirstRunExperience" -Value 1 -ErrorAction SilentlyContinue 
            }
            "Clear Delivery Optimization Cache" {
                Remove-Item -Path "C:\Windows\SoftwareDistribution\DeliveryOptimization" -Recurse -Force -ErrorAction SilentlyContinue
            }
                       "Disable Background Applications" { 
                # Set the policy to disable background apps, but exclude OneDrive
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsRunInBackground" -Value 2 -ErrorAction SilentlyContinue
                # Ensure OneDrive is allowed to run in the background by setting an exception
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications\Microsoft.OneDrive" -Name "Disabled" -Value 0 -ErrorAction SilentlyContinue
            }
            "Set Services" { 
                $services = @{
                    "AssignedAccessManagerSvc"="Disabled";
                    "BDESVC"="Disabled";
                    "BthServ"="Automatic";
                    "DiagTrack"="Disabled";
                    "DmEnrollmentSvc"="Manual";
                    "DPS"="Manual";
                    "Fax"="Disabled";
                    "FontCache"="Manual";
                    "HotspotService"="Disabled";
                    "lfsvc"="Disabled";
                    "MapsBroker"="Manual";
                    "stisvc"="Manual";
                    "SysMain"="Manual";
                    "TermService"="Disabled";
                    "TrkWks"="Manual";
                    "WlanSvc"="Manual";
                    "WManSvc"="Manual";
                    "wuauserv"="Manual";
                    "WSearch"="Manual"; 
                    "XblAuthManager"="Disabled";
                    "XblGameSave"="Disabled";
                    "XboxNetApiSvc"="Disabled"
                    "RetailDemo"="Disabled";
                    "RemoteRegisty"="Disabled";
                    "PhoneSvc"="Disabled";
                    "wisvc"="Disabled";
                    "icssvc"="Disabled";
                    "seclogon"="Disabled";
                }; 
                foreach ($svc in $services.GetEnumerator()) { 
                    Set-Service -Name $svc.Key -StartupType $svc.Value -ErrorAction SilentlyContinue 
                } 
            } 
        }

        # Update progress bar
        $progressBarSpeed.Value = $currentTweak
        $form.Refresh()
    }

    # Hide progress bar and show completion
    $progressBarSpeed.Visible = $false
    $lblStatusSpeed.Text = "Status: Done"
})
$btnPanelSpeed.Controls.Add($btnRunSpeed)
$tabControl.Controls.Add($tabSpeed)

# Preferences tab
$tabPrefs = New-Object System.Windows.Forms.TabPage
$tabPrefs.Text = "Preferences"
$tabPrefs.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)

$panelPrefs = New-Object System.Windows.Forms.Panel
$panelPrefs.Dock = "Top"
$panelPrefs.Height = 400
$panelPrefs.AutoScroll = $true
$panelPrefs.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$tabPrefs.Controls.Add($panelPrefs)

$yPos = 10
$prefsCheckboxes = @("Dark Theme for Windows", "Disable Bing Search in Start Menu", "Disable Recommendations in Start Menu", "Search Button in Taskbar", "Disable Widgets in Taskbar", "Disable Animations", "Turn Off Taskbar Transparency", "Disable Snap Assist", "Add and Activate Ultimate Performance Profile")
foreach ($check in $prefsCheckboxes) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Location = New-Object System.Drawing.Point(10,$yPos)
    $cb.Size = New-Object System.Drawing.Size(380,25)
    $cb.Text = $check
    $cb.ForeColor = [System.Drawing.Color]::White
    $cb.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
    $cb.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $cb.Anchor = "Left,Right"
    $panelPrefs.Controls.Add($cb)
    $yPos += 30
}

$lblStatusPrefs = New-Object System.Windows.Forms.Label
$lblStatusPrefs.Dock = "Top"
$lblStatusPrefs.Height = 30
$lblStatusPrefs.Text = "Status: Ready"
$lblStatusPrefs.ForeColor = [System.Drawing.Color]::White
$lblStatusPrefs.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$lblStatusPrefs.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$lblStatusPrefs.TextAlign = "MiddleLeft"
$tabPrefs.Controls.Add($lblStatusPrefs)

# Add a ProgressBar below the status label
$progressBarPrefs = New-Object System.Windows.Forms.ProgressBar
$progressBarPrefs.Dock = "Top"
$progressBarPrefs.Height = 20
$progressBarPrefs.Minimum = 0
$progressBarPrefs.Maximum = 100
$progressBarPrefs.Value = 0
$progressBarPrefs.Visible = $false
$progressBarPrefs.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$progressBarPrefs.ForeColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$tabPrefs.Controls.Add($progressBarPrefs)

$btnPanelPrefs = New-Object System.Windows.Forms.Panel
$btnPanelPrefs.Dock = "Bottom"
$btnPanelPrefs.Height = 40
$btnPanelPrefs.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$tabPrefs.Controls.Add($btnPanelPrefs)

$btnSelectAllPrefs = New-Object System.Windows.Forms.Button
$btnSelectAllPrefs.Location = New-Object System.Drawing.Point(10,5)
$btnSelectAllPrefs.Size = New-Object System.Drawing.Size(120,30)
$btnSelectAllPrefs.Text = "Select All"
$btnSelectAllPrefs.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$btnSelectAllPrefs.ForeColor = [System.Drawing.Color]::White
$btnSelectAllPrefs.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSelectAllPrefs.FlatStyle = "Flat"
$btnSelectAllPrefs.FlatAppearance.BorderSize = 0
$btnSelectAllPrefs.Add_Click({ foreach ($control in $panelPrefs.Controls) { if ($control -is [System.Windows.Forms.CheckBox]) { $control.Checked = $true } } })
$btnPanelPrefs.Controls.Add($btnSelectAllPrefs)

$btnRunPrefs = New-Object System.Windows.Forms.Button
$btnRunPrefs.Location = New-Object System.Drawing.Point(290,5)
$btnRunPrefs.Size = New-Object System.Drawing.Size(120,30)
$btnRunPrefs.Text = "Apply Tweaks"
$btnRunPrefs.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$btnRunPrefs.ForeColor = [System.Drawing.Color]::White
$btnRunPrefs.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnRunPrefs.FlatStyle = "Flat"
$btnRunPrefs.FlatAppearance.BorderSize = 0
$btnRunPrefs.Anchor = "Right"
$btnRunPrefs.Add_Click({ 
    # Count the number of selected tweaks
    $selectedTweaks = @($panelPrefs.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] -and $_.Checked })
    $totalTweaks = $selectedTweaks.Count

    if ($totalTweaks -eq 0) {
        $lblStatusPrefs.Text = "Status: No tweaks selected"
        return
    }

    # Show the progress bar and set its maximum
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
                Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "UserPreferencesMask" -Value ([byte[]](0x90,0x12,0x03,0x80,0x10,0x00,0x00,0x00)) -ErrorAction SilentlyContinue
                Set-ItemProperty -Path "HKCU:\Control Panel\Desktop\WindowMetrics" -Name "MinAnimate" -Value 0
            }
            "Turn off Taskbar Transparency" { 
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "EnableTransparency" -Value 0
            }
            "Disable Snap Assist" { 
                Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WindowArrangementActive" -Value 0
            }
            "Add and Activate Ultimate Performance Profile" { 
                powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61; 
                $guid = (powercfg -l | Where-Object { $_ -match "Ultimate Performance" } | ForEach-Object { $_.Split()[3] }); 
                if ($guid) { powercfg -s $guid } 
            } 
        }

        # Update progress bar
        $progressBarPrefs.Value = $currentTweak
        $form.Refresh()
    }

    # Hide progress bar and show completion
    $progressBarPrefs.Visible = $false
    $lblStatusPrefs.Text = "Status: Done"
})
$btnPanelPrefs.Controls.Add($btnRunPrefs)
$tabControl.Controls.Add($tabPrefs)

# Uninstall Apps tab
$tabUninstall = New-Object System.Windows.Forms.TabPage
$tabUninstall.Text = "Uninstall Apps"
$tabUninstall.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)

$containerUninstall = New-Object System.Windows.Forms.Panel
$containerUninstall.Dock = "Fill"
$containerUninstall.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$tabUninstall.Controls.Add($containerUninstall)

$clbApps = New-Object System.Windows.Forms.CheckedListBox
$clbApps.Dock = "Fill"
$clbApps.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$clbApps.ForeColor = [System.Drawing.Color]::White
$clbApps.BorderStyle = "None"
$clbApps.CheckOnClick = $true
$clbApps.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$containerUninstall.Controls.Add($clbApps)

$lblStatusUninstall = New-Object System.Windows.Forms.Label
$lblStatusUninstall.Dock = "Bottom"
$lblStatusUninstall.Height = 30
$lblStatusUninstall.Text = "Status: Ready"
$lblStatusUninstall.ForeColor = [System.Drawing.Color]::White
$lblStatusUninstall.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$lblStatusUninstall.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$lblStatusUninstall.TextAlign = "MiddleLeft"
$containerUninstall.Controls.Add($lblStatusUninstall)

# Add a ProgressBar below the status label
$progressBarUninstall = New-Object System.Windows.Forms.ProgressBar
$progressBarUninstall.Dock = "Bottom"
$progressBarUninstall.Height = 20
$progressBarUninstall.Minimum = 0
$progressBarUninstall.Maximum = 100
$progressBarUninstall.Value = 0
$progressBarUninstall.Visible = $false
$progressBarUninstall.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$progressBarUninstall.ForeColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$containerUninstall.Controls.Add($progressBarUninstall)

$btnPanelUninstall = New-Object System.Windows.Forms.Panel
$btnPanelUninstall.Dock = "Bottom"
$btnPanelUninstall.Height = 40
$btnPanelUninstall.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$containerUninstall.Controls.Add($btnPanelUninstall)

# Function to load apps
function Load-Apps {
    $lblStatusUninstall.Text = "Status: Loading apps..."
    [System.Windows.Forms.Application]::DoEvents()
    
    $clbApps.Items.Clear()
    $script:apps = @()

    # Collect apps from registry
    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $registryPaths) {
        if (Test-Path $path) {
            $appsFromRegistry = Get-ItemProperty $path | Where-Object {
                $_.DisplayName -and ($_.UninstallString -or $_.PSChildName -match "^{.+}$")
            } | Select-Object DisplayName, UninstallString, @{Name = "WinGetID"; Expression = { $_.PSChildName }}

            $script:apps += $appsFromRegistry
        }
    }

    # Check for Winget
    $wingetAvailable = $false
    try {
        $null = winget --version 2>$null
        if ($LASTEXITCODE -eq 0) {
            $wingetAvailable = $true
        }
    } catch {}

    if ($wingetAvailable) {
        try {
            $wingetJson = winget list --source winget --accept-source-agreements --output json 2>$null | ConvertFrom-Json
            $wingetApps = @()

            foreach ($entry in $wingetJson) {
                if ($entry.Name -and $entry.Id) {
                    $wingetApps += [PSCustomObject]@{
                        DisplayName     = $entry.Name
                        WinGetID        = $entry.Id
                        UninstallString = $null
                    }
                }
            }

            # Merge Winget apps, de-duplicate by ID first, then DisplayName
            foreach ($wingetApp in $wingetApps) {
                if (-not ($script:apps | Where-Object {
                    $_.WinGetID -eq $wingetApp.WinGetID -or $_.DisplayName -eq $wingetApp.DisplayName
                })) {
                    $script:apps += $wingetApp
                }
            }

            if ($wingetApps.Count -eq 0) {
                $lblStatusUninstall.Text = "Status: Winget found, but no apps returned"
            }

        } catch {
            $lblStatusUninstall.Text = "Status: Winget failed to list apps"
        }
    } else {
        $lblStatusUninstall.Text = "Status: Winget not available"
    }

    # Sort and populate UI
    $script:apps = $script:apps | Sort-Object DisplayName
    foreach ($app in $script:apps) {
        if ($app.DisplayName) {
            $clbApps.Items.Add($app.DisplayName) | Out-Null
        }
    }

    [System.Windows.Forms.Application]::DoEvents()
    $lblStatusUninstall.Text = "Status: $($script:apps.Count) apps loaded"
}

$btnLoadApps = New-Object System.Windows.Forms.Button
$btnLoadApps.Location = New-Object System.Drawing.Point(10,5)
$btnLoadApps.Size = New-Object System.Drawing.Size(120,30)
$btnLoadApps.Text = "Load Apps"
$btnLoadApps.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$btnLoadApps.ForeColor = [System.Drawing.Color]::White
$btnLoadApps.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnLoadApps.FlatStyle = "Flat"
$btnLoadApps.FlatAppearance.BorderSize = 0
$btnLoadApps.Add_Click({ Load-Apps })
$btnPanelUninstall.Controls.Add($btnLoadApps)

$btnSelectAllUninstall = New-Object System.Windows.Forms.Button
$btnSelectAllUninstall.Location = New-Object System.Drawing.Point(140,5)
$btnSelectAllUninstall.Size = New-Object System.Drawing.Size(120,30)
$btnSelectAllUninstall.Text = "Select All"
$btnSelectAllUninstall.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$btnSelectAllUninstall.ForeColor = [System.Drawing.Color]::White
$btnSelectAllUninstall.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSelectAllUninstall.FlatStyle = "Flat"
$btnSelectAllUninstall.FlatAppearance.BorderSize = 0
$btnSelectAllUninstall.Add_Click({ for ($i = 0; $i -lt $clbApps.Items.Count; $i++) { $clbApps.SetItemChecked($i, $true) } })
$btnPanelUninstall.Controls.Add($btnSelectAllUninstall)

$btnUninstall = New-Object System.Windows.Forms.Button
$btnUninstall.Location = New-Object System.Drawing.Point(270,5)
$btnUninstall.Size = New-Object System.Drawing.Size(120,30)
$btnUninstall.Text = "Uninstall Selected"
$btnUninstall.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$btnUninstall.ForeColor = [System.Drawing.Color]::White
$btnUninstall.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnUninstall.FlatStyle = "Flat"
$btnUninstall.FlatAppearance.BorderSize = 0
$btnUninstall.Anchor = "Right"
$btnUninstall.Add_Click({
    $selectedApps = $clbApps.CheckedItems
    $totalApps = $selectedApps.Count

    if ($totalApps -eq 0) {
        $lblStatusUninstall.Text = "Status: No apps selected"
        return
    }

    # Show the progress bar and set its maximum
    $progressBarUninstall.Visible = $true
    $progressBarUninstall.Maximum = $totalApps
    $progressBarUninstall.Value = 0

    $lblStatusUninstall.Text = "Status: Uninstalling apps (0/$totalApps)..."
    $currentApp = 0

    foreach ($appName in $selectedApps) {
        $currentApp++
        $lblStatusUninstall.Text = "Status: Uninstalling app $currentApp of $totalApps ($appName)..."
        $app = $script:apps | Where-Object { $_.DisplayName -eq $appName }
        if ($app) {
            try {
                $wingetId = $app.WinGetID
                if ($wingetId -and (Get-Command winget -ErrorAction SilentlyContinue)) {
                    winget uninstall --id $wingetId --silent --force
                } else {
                    throw "No winget ID or winget not available"
                }
            } catch {
                if ($app.UninstallString) {
                    $uninst = $app.UninstallString -replace '"', ''
                    if ($uninst -match "msiexec") {
                        Start-Process "msiexec.exe" -ArgumentList "/x $($app.WinGetID) /qn" -Wait -NoNewWindow
                    } else {
                        Start-Process -FilePath $uninst -ArgumentList "/S" -Wait -NoNewWindow -ErrorAction SilentlyContinue
                    }
                } else {
                    $lblStatusUninstall.Text = "Status: No uninstall method for $appName"
                }
            }
        }

        # Update progress bar
        $progressBarUninstall.Value = $currentApp
        $form.Refresh()
    }

    # Hide progress bar and show completion
    $progressBarUninstall.Visible = $false
    $lblStatusUninstall.Text = "Status: Done"
    Load-Apps
})
$btnPanelUninstall.Controls.Add($btnUninstall)

Load-Apps
$tabControl.Controls.Add($tabUninstall)

# Install Apps tab
$tabInstall = New-Object System.Windows.Forms.TabPage
$tabInstall.Text = "Install Apps"
$tabInstall.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)

$containerInstall = New-Object System.Windows.Forms.Panel
$containerInstall.Dock = "Fill"
$containerInstall.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$tabInstall.Controls.Add($containerInstall)

$clbInstallApps = New-Object System.Windows.Forms.CheckedListBox
$clbInstallApps.Dock = "Fill"
$clbInstallApps.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$clbInstallApps.ForeColor = [System.Drawing.Color]::White
$clbInstallApps.BorderStyle = "None"
$clbInstallApps.CheckOnClick = $true
$clbInstallApps.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$containerInstall.Controls.Add($clbInstallApps)

$lblStatusInstall = New-Object System.Windows.Forms.Label
$lblStatusInstall.Dock = "Bottom"
$lblStatusInstall.Height = 30
$lblStatusInstall.Text = "Status: Ready"
$lblStatusInstall.ForeColor = [System.Drawing.Color]::White
$lblStatusInstall.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$lblStatusInstall.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$lblStatusInstall.TextAlign = "MiddleLeft"
$containerInstall.Controls.Add($lblStatusInstall)

# Add a ProgressBar below the status label
$progressBarInstall = New-Object System.Windows.Forms.ProgressBar
$progressBarInstall.Dock = "Bottom"
$progressBarInstall.Height = 20
$progressBarInstall.Minimum = 0
$progressBarInstall.Maximum = 100
$progressBarInstall.Value = 0
$progressBarInstall.Visible = $false
$progressBarInstall.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$progressBarInstall.ForeColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$containerInstall.Controls.Add($progressBarInstall)

$btnPanelInstall = New-Object System.Windows.Forms.Panel
$btnPanelInstall.Dock = "Bottom"
$btnPanelInstall.Height = 40
$btnPanelInstall.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$containerInstall.Controls.Add($btnPanelInstall)

$appList = @(
    [PSCustomObject]@{ Name = "7-Zip"; Url = "https://www.7-zip.org/a/7z2301-x64.exe" },
    [PSCustomObject]@{ Name = "RingCentral"; Url = "https://app.ringcentral.com/download/RingCentral.exe?V=20138739841159500" },
    [PSCustomObject]@{ Name = "Google Chrome"; Url = "https://dl.google.com/chrome/install/ChromeSetup.exe" },
    [PSCustomObject]@{ Name = "Microsoft Teams"; Url = "https://go.microsoft.com/fwlink/?linkid=2281613&clcid=0x409&culture=en-us&country=us" }
)

foreach ($app in $appList) {
    $clbInstallApps.Items.Add($app.Name) | Out-Null
}

$btnSelectAllInstall = New-Object System.Windows.Forms.Button
$btnSelectAllInstall.Location = New-Object System.Drawing.Point(10,5)
$btnSelectAllInstall.Size = New-Object System.Drawing.Size(120,30)
$btnSelectAllInstall.Text = "Select All"
$btnSelectAllInstall.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$btnSelectAllInstall.ForeColor = [System.Drawing.Color]::White
$btnSelectAllInstall.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSelectAllInstall.FlatStyle = "Flat"
$btnSelectAllInstall.FlatAppearance.BorderSize = 0
$btnSelectAllInstall.Add_Click({ for ($i = 0; $i -lt $clbInstallApps.Items.Count; $i++) { $clbInstallApps.SetItemChecked($i, $true) } })
$btnPanelInstall.Controls.Add($btnSelectAllInstall)

$btnInstallSelected = New-Object System.Windows.Forms.Button
$btnInstallSelected.Location = New-Object System.Drawing.Point(290,5)
$btnInstallSelected.Size = New-Object System.Drawing.Size(120,30)
$btnInstallSelected.Text = "Install Selected"
$btnInstallSelected.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$btnInstallSelected.ForeColor = [System.Drawing.Color]::White
$btnInstallSelected.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnInstallSelected.FlatStyle = "Flat"
$btnInstallSelected.FlatAppearance.BorderSize = 0
$btnInstallSelected.Anchor = "Right"
$btnInstallSelected.Add_Click({
    $selectedApps = $clbInstallApps.CheckedItems
    $totalApps = $selectedApps.Count

    if ($totalApps -eq 0) {
        $lblStatusInstall.Text = "Status: No apps selected"
        return
    }

    # Show the progress bar and set its maximum
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

        # Update progress bar
        $progressBarInstall.Value = $currentApp
        $form.Refresh()
    }

    # Hide progress bar and show completion
    $progressBarInstall.Visible = $false
    $lblStatusInstall.Text = "Status: Done"
})
$btnPanelInstall.Controls.Add($btnInstallSelected)
$tabControl.Controls.Add($tabInstall)

# Win Updates tab
$tabUpdates = New-Object System.Windows.Forms.TabPage
$tabUpdates.Text = "Win Updates"
$tabUpdates.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)

$containerUpdates = New-Object System.Windows.Forms.Panel
$containerUpdates.Dock = "Fill"
$containerUpdates.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$tabUpdates.Controls.Add($containerUpdates)

$listUpdates = New-Object System.Windows.Forms.ListBox
$listUpdates.Dock = "Fill"
$listUpdates.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$listUpdates.ForeColor = [System.Drawing.Color]::White
$listUpdates.BorderStyle = "None"
$listUpdates.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$containerUpdates.Controls.Add($listUpdates)

$lblStatusUpdates = New-Object System.Windows.Forms.Label
$lblStatusUpdates.Dock = "Bottom"
$lblStatusUpdates.Height = 30
$lblStatusUpdates.Text = "Status: Ready"
$lblStatusUpdates.ForeColor = [System.Drawing.Color]::White
$lblStatusUpdates.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$lblStatusUpdates.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$lblStatusUpdates.TextAlign = "MiddleLeft"
$containerUpdates.Controls.Add($lblStatusUpdates)

# Add a ProgressBar below the status label
$progressBarUpdates = New-Object System.Windows.Forms.ProgressBar
$progressBarUpdates.Dock = "Bottom"
$progressBarUpdates.Height = 20
$progressBarUpdates.Minimum = 0
$progressBarUpdates.Maximum = 100
$progressBarUpdates.Value = 0
$progressBarUpdates.Visible = $false
$progressBarUpdates.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$progressBarUpdates.ForeColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$containerUpdates.Controls.Add($progressBarUpdates)

$btnPanelUpdates = New-Object System.Windows.Forms.Panel
$btnPanelUpdates.Dock = "Bottom"
$btnPanelUpdates.Height = 100
$btnPanelUpdates.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$containerUpdates.Controls.Add($btnPanelUpdates)

$btnLoadSecurity = New-Object System.Windows.Forms.Button
$btnLoadSecurity.Location = New-Object System.Drawing.Point(10,10)
$btnLoadSecurity.Size = New-Object System.Drawing.Size(150,35)
$btnLoadSecurity.Text = "Load Security Updates"
$btnLoadSecurity.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$btnLoadSecurity.ForeColor = [System.Drawing.Color]::White
$btnLoadSecurity.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnLoadSecurity.FlatStyle = "Flat"
$btnLoadSecurity.FlatAppearance.BorderSize = 0
$btnLoadSecurity.Add_Click({
    $lblStatusUpdates.Text = "Status: Loading security updates..."
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        $lblStatusUpdates.Text = "Status: Installing PSWindowsUpdate module..."
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction SilentlyContinue
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            $lblStatusUpdates.Text = "Status: Failed to install module"
            return
        }
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
$btnPanelUpdates.Controls.Add($btnLoadSecurity)

$btnInstallSecurity = New-Object System.Windows.Forms.Button
$btnInstallSecurity.Location = New-Object System.Drawing.Point(170,10)
$btnInstallSecurity.Size = New-Object System.Drawing.Size(150,35)
$btnInstallSecurity.Text = "Install Security Updates"
$btnInstallSecurity.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$btnInstallSecurity.ForeColor = [System.Drawing.Color]::White
$btnInstallSecurity.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnInstallSecurity.FlatStyle = "Flat"
$btnInstallSecurity.FlatAppearance.BorderSize = 0
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
$btnPanelUpdates.Controls.Add($btnInstallSecurity)

$btnLoadAll = New-Object System.Windows.Forms.Button
$btnLoadAll.Location = New-Object System.Drawing.Point(10,55)
$btnLoadAll.Size = New-Object System.Drawing.Size(150,35)
$btnLoadAll.Text = "Load All Updates"
$btnLoadAll.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$btnLoadAll.ForeColor = [System.Drawing.Color]::White
$btnLoadAll.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnLoadAll.FlatStyle = "Flat"
$btnLoadAll.FlatAppearance.BorderSize = 0
$btnLoadAll.Add_Click({
    $lblStatusUpdates.Text = "Status: Loading all updates..."
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        $lblStatusUpdates.Text = "Status: Installing PSWindowsUpdate module..."
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction SilentlyContinue
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            $lblStatusUpdates.Text = "Status: Failed to install module"
            return
        }
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
$btnPanelUpdates.Controls.Add($btnLoadAll)

$btnInstallAll = New-Object System.Windows.Forms.Button
$btnInstallAll.Location = New-Object System.Drawing.Point(170,55)
$btnInstallAll.Size = New-Object System.Drawing.Size(150,35)
$btnInstallAll.Text = "Install All Updates"
$btnInstallAll.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184)
$btnInstallAll.ForeColor = [System.Drawing.Color]::White
$btnInstallAll.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnInstallAll.FlatStyle = "Flat"
$btnInstallAll.FlatAppearance.BorderSize = 0
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
$btnPanelUpdates.Controls.Add($btnInstallAll)
$tabControl.Controls.Add($tabUpdates)

# Adjust button positions on resize
$form.Add_Resize({
    $btnRunSpeed.Location = New-Object System.Drawing.Point(($btnPanelSpeed.Width - $btnRunSpeed.Width - 10), 5)
    $btnRunPrefs.Location = New-Object System.Drawing.Point(($btnPanelPrefs.Width - $btnRunPrefs.Width - 10), 5)
    $btnUninstall.Location = New-Object System.Drawing.Point(($btnPanelUninstall.Width - $btnUninstall.Width - 10), 5)
    $btnInstallSelected.Location = New-Object System.Drawing.Point(($btnPanelInstall.Width - $btnInstallSelected.Width - 10), 5)
    $btnInstallSecurity.Location = New-Object System.Drawing.Point(($btnPanelUpdates.Width - $btnInstallSecurity.Width - 230), 10)
    $btnInstallAll.Location = New-Object System.Drawing.Point(($btnPanelUpdates.Width - $btnInstallAll.Width - 230), 55)
})

# Show form
$form.ShowDialog()