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
$panelSpeed.Height = 450  # Increased height to ensure all checkboxes are visible
$panelSpeed.AutoScroll = $true
$panelSpeed.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$tabSpeed.Controls.Add($panelSpeed)

$yPos = 10
$speedCheckboxes = @("Delete Temp Files", "Disable Consumer Features", "Disable Telemetry", "Disable Activity History", "Disable GameDVR", "Disable Hibernation", "Disable Homegroup", "Disable Location Tracking", "Disable Storage Sense", "Disable Wifi-Sense", "Disable Recall", "Debloat Edge", "Disable Background Applications", "Set Services")
foreach ($check in $speedCheckboxes) {
    Write-Host "Adding checkbox: $check" -ForegroundColor Cyan
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Location = New-Object System.Drawing.Point(10, $yPos)
    $cb.Size = New-Object System.Drawing.Size(380, 25)
    $cb.Text = $check
    $cb.ForeColor = [System.Drawing.Color]::White
    $cb.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
    $cb.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $cb.Anchor = "Left,Right"
    $panelSpeed.Controls.Add($cb)
    $yPos += 28  # Reduced spacing to fit all checkboxes
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
    $lblStatusSpeed.Text = "Status: Running..."; 
    Write-Host "Starting Speed Up tweaks..." -ForegroundColor Cyan; 
    foreach ($control in $panelSpeed.Controls) { 
        if ($control -is [System.Windows.Forms.CheckBox] -and $control.Checked) { 
            $tweak = $control.Text; 
            Write-Host "Executing: $tweak" -ForegroundColor Yellow; 
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
                "Disable Background Applications" { 
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsRunInBackground" -Value 0 -ErrorAction SilentlyContinue 
                }
                "Set Services" { 
                    $services = @{
                        "AssignedAccessManagerSvc"="Disabled";
                        "BDESVC"="Disabled";
                        "BthServ"="Disabled";
                        "DiagTrack"="Disabled";
                        "DmEnrollmentSvc"="Manual";
                        "DPS"="Disabled";
                        "Fax"="Disabled";
                        "FontCache"="Manual";
                        "HotspotService"="Disabled";
                        "lfsvc"="Disabled";
                        "MapsBroker"="Manual";
                        "Spooler"="Disabled";
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
                    }; 
                    foreach ($svc in $services.GetEnumerator()) { 
                        Set-Service -Name $svc.Key -StartupType $svc.Value -ErrorAction SilentlyContinue 
                    } 
                } 
            } 
            Write-Host "$tweak completed" -ForegroundColor Green 
        } 
    } 
    $lblStatusSpeed.Text = "Status: Done"; 
    Write-Host "Speed Up tweaks applied successfully" -ForegroundColor Cyan 
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
$prefsCheckboxes = @("Dark Theme for Windows", "Disable Bing Search in Start Menu", "Disable Recommendations in Start Menu", "Search Button in Taskbar", "Disable Widgets in Taskbar", "Add and Activate Ultimate Performance Profile")
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
    $lblStatusPrefs.Text = "Status: Running..."; 
    Write-Host "Starting Preferences tweaks..." -ForegroundColor Cyan; 
    foreach ($control in $panelPrefs.Controls) { 
        if ($control -is [System.Windows.Forms.CheckBox] -and $control.Checked) { 
            $tweak = $control.Text; 
            Write-Host "Executing: $tweak" -ForegroundColor Yellow; 
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
                "Add and Activate Ultimate Performance Profile" { 
                    powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61; 
                    $guid = (powercfg -l | Where-Object { $_ -match "Ultimate Performance" } | ForEach-Object { $_.Split()[3] }); 
                    if ($guid) { powercfg -s $guid } 
                } 
            } 
            Write-Host "$tweak completed" -ForegroundColor Green 
        } 
    } 
    $lblStatusPrefs.Text = "Status: Done"; 
    Write-Host "Preferences tweaks applied successfully" -ForegroundColor Cyan 
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

$btnPanelUninstall = New-Object System.Windows.Forms.Panel
$btnPanelUninstall.Dock = "Bottom"
$btnPanelUninstall.Height = 40
$btnPanelUninstall.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$containerUninstall.Controls.Add($btnPanelUninstall)

# Function to load apps
function Load-Apps {
    $lblStatusUninstall.Text = "Status: Loading apps..."
    Write-Host "Loading installed applications..." -ForegroundColor Cyan
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
            $appsFromRegistry = Get-ItemProperty $path | Where-Object { $_.DisplayName } | Select-Object DisplayName, UninstallString, @{Name="WinGetID";Expression={$_.PSChildName}}
            $script:apps += $appsFromRegistry
        }
    }

    # Check if winget is available
    $wingetAvailable = $false
    try {
        $wingetVersion = & winget --version 2>$null
        if ($?) {
            $wingetAvailable = $true
        }
    } catch {
        Write-Host "winget not found. Falling back to registry-only app detection." -ForegroundColor Yellow
    }

    if ($wingetAvailable) {
        $wingetAppsRaw = winget list --accept-source-agreements | Where-Object { $_ -match "^[^\-]" } | Select-Object -Skip 2
        $wingetApps = @()
        foreach ($line in $wingetAppsRaw) {
            $columns = $line -split '\s+', 4
            if ($columns.Count -ge 3) {
                $appName = $columns[0].Trim()
                $appId = $columns[2].Trim()
                if ($appName -and $appId) {
                    $wingetApps += [PSCustomObject]@{
                        DisplayName = $appName
                        WinGetID = $appId
                        UninstallString = $null
                    }
                }
            }
        }

        foreach ($wingetApp in $wingetApps) {
            if (-not ($script:apps | Where-Object { $_.DisplayName -eq $wingetApp.DisplayName -or $_.WinGetID -eq $wingetApp.WinGetID })) {
                $script:apps += $wingetApp
            }
        }
    }

    $script:apps = $script:apps | Sort-Object DisplayName
    foreach ($app in $script:apps) {
        if ($app.DisplayName) {
            $clbApps.Items.Add($app.DisplayName) | Out-Null
        }
    }

    $lblStatusUninstall.Text = "Status: $($script:apps.Count) apps loaded"
    Write-Host "$($script:apps.Count) applications loaded." -ForegroundColor Green
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
    $lblStatusUninstall.Text = "Status: Running..."
    Write-Host "Starting uninstallation..." -ForegroundColor Cyan
    $selectedApps = $clbApps.CheckedItems
    foreach ($appName in $selectedApps) {
        Write-Host "Uninstalling: $appName" -ForegroundColor Yellow
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
                    Write-Host "No uninstall method available for $appName" -ForegroundColor Red
                }
            }
            Write-Host "$appName uninstalled" -ForegroundColor Green
        }
    }
    $lblStatusUninstall.Text = "Status: Done"
    Write-Host "Uninstallation completed" -ForegroundColor Cyan
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
    $lblStatusInstall.Text = "Status: Installing..."
    Write-Host "Starting app installation..." -ForegroundColor Cyan
    $selectedApps = $clbInstallApps.CheckedItems

    foreach ($appName in $selectedApps) {
        $app = $appList | Where-Object { $_.Name -eq $appName }
        if ($app) {
            Write-Host "Downloading: $appName" -ForegroundColor Yellow
            $downloadPath = "$env:TEMP\$($app.Name)_installer.exe"
            try {
                Invoke-WebRequest -Uri $app.Url -OutFile $downloadPath -ErrorAction Stop
                Write-Host "$appName downloaded to $downloadPath" -ForegroundColor Green
                Write-Host "Installing: $appName" -ForegroundColor Yellow
                Start-Process -FilePath $downloadPath -ArgumentList "/S" -Wait -NoNewWindow -ErrorAction Stop
                Write-Host "$appName installed successfully" -ForegroundColor Green
                Remove-Item $downloadPath -Force -ErrorAction SilentlyContinue
            } catch {
                Write-Host "Failed to install $appName. Error: $_" -ForegroundColor Red
                $lblStatusInstall.Text = "Status: Error installing $appName"
            }
        }
    }

    $lblStatusInstall.Text = "Status: Done"
    Write-Host "App installation completed" -ForegroundColor Cyan
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
    $lblStatusUpdates.Text = "Status: Loading..."
    Write-Host "Loading security updates..." -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Host "PSWindowsUpdate module not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction SilentlyContinue
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            $lblStatusUpdates.Text = "Status: Failed to install module"
            Write-Host "Failed to install PSWindowsUpdate module." -ForegroundColor Red
            return
        }
    }
    Import-Module PSWindowsUpdate
    $listUpdates.Items.Clear()
    $script:securityUpdates = Get-WindowsUpdate -MicrosoftUpdate -Category "Security Updates" -ErrorAction SilentlyContinue
    if ($script:securityUpdates) {
        foreach ($update in $script:securityUpdates) {
            $listUpdates.Items.Add("[$($update.KB)] $($update.Title)")
        }
        $lblStatusUpdates.Text = "Status: $($script:securityUpdates.Count) security updates loaded"
        Write-Host "$($script:securityUpdates.Count) security updates loaded." -ForegroundColor Green
    } else {
        $lblStatusUpdates.Text = "Status: No security updates found"
        Write-Host "No security updates found." -ForegroundColor Yellow
    }
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
        $lblStatusUpdates.Text = "Status: Installing..."
        Write-Host "Installing security updates..." -ForegroundColor Cyan
        Install-WindowsUpdate -MicrosoftUpdate -KBArticleID ($script:securityUpdates.KB) -AcceptAll -IgnoreReboot -ErrorAction SilentlyContinue
        $lblStatusUpdates.Text = "Status: Installation complete"
        Write-Host "Security updates installed." -ForegroundColor Green
        $listUpdates.Items.Clear()
        $script:securityUpdates = $null
    } else {
        $lblStatusUpdates.Text = "Status: No security updates to install"
        Write-Host "No security updates to install." -ForegroundColor Yellow
    }
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
    $lblStatusUpdates.Text = "Status: Loading..."
    Write-Host "Loading all updates..." -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Host "PSWindowsUpdate module not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction SilentlyContinue
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            $lblStatusUpdates.Text = "Status: Failed to install module"
            Write-Host "Failed to install PSWindowsUpdate module." -ForegroundColor Red
            return
        }
    }
    Import-Module PSWindowsUpdate
    $listUpdates.Items.Clear()
    $script:allUpdates = Get-WindowsUpdate -MicrosoftUpdate -ErrorAction SilentlyContinue
    if ($script:allUpdates) {
        foreach ($update in $script:allUpdates) {
            $listUpdates.Items.Add("[$($update.KB)] $($update.Title)")
        }
        $lblStatusUpdates.Text = "Status: $($script:allUpdates.Count) updates loaded"
        Write-Host "$($script:allUpdates.Count) updates loaded." -ForegroundColor Green
    } else {
        $lblStatusUpdates.Text = "Status: No updates found"
        Write-Host "No updates found." -ForegroundColor Yellow
    }
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
        $lblStatusUpdates.Text = "Status: Installing..."
        Write-Host "Installing all updates..." -ForegroundColor Cyan
        Install-WindowsUpdate -MicrosoftUpdate -KBArticleID ($script:allUpdates.KB) -AcceptAll -IgnoreReboot -ErrorAction SilentlyContinue
        $lblStatusUpdates.Text = "Status: Installation complete"
        Write-Host "All updates installed." -ForegroundColor Green
        $listUpdates.Items.Clear()
        $script:allUpdates = $null
    } else {
        $lblStatusUpdates.Text = "Status: No updates to install"
        Write-Host "No updates to install." -ForegroundColor Yellow
    }
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