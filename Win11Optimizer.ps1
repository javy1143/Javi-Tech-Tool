# Optimize-Win11-DarkGUI.ps1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Win11 Optimizer"
$form.Size = New-Object System.Drawing.Size(450,600)
$form.MinimumSize = New-Object System.Drawing.Size(400,500)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38) # Dark background (#1C2526)
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
$panelSpeed.Height = 400 # Restored original height since no logo panel
$panelSpeed.AutoScroll = $true
$panelSpeed.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
$tabSpeed.Controls.Add($panelSpeed)

$yPos = 10
$speedCheckboxes = @("Delete Temp Files", "Disable Consumer Features", "Disable Telemetry", "Disable Activity History", "Disable GameDVR", "Disable Hibernation", "Disable Homegroup", "Disable Location Tracking", "Disable Storage Sense", "Disable Wifi-Sense", "Disable Recall", "Debloat Edge", "Disable Background Applications", "Windows Services")
foreach ($check in $speedCheckboxes) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Location = New-Object System.Drawing.Point(10,$yPos)
    $cb.Size = New-Object System.Drawing.Size(380,25)
    $cb.Text = $check
    $cb.ForeColor = [System.Drawing.Color]::White
    $cb.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)
    $cb.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $cb.Anchor = "Left,Right"
    $panelSpeed.Controls.Add($cb)
    $yPos += 30
}

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
$btnSelectAllSpeed.BackColor = [System.Drawing.Color]::FromArgb(0, 94, 184) # Blue (#005EB8)
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
$btnRunSpeed.Add_Click({ $lblStatusSpeed.Text = "Status: Running..."; Write-Host "Starting Speed Up tweaks..." -ForegroundColor Cyan; foreach ($control in $panelSpeed.Controls) { if ($control -is [System.Windows.Forms.CheckBox] -and $control.Checked) { $tweak = $control.Text; Write-Host "Executing: $tweak" -ForegroundColor Yellow; switch ($tweak) { "Delete Temp Files" { Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue } "Windows Services" { $services = @{"AssignedAccessManagerSvc"="Disabled";"BDESVC"="Disabled";"BthServ"="Disabled";"DiagTrack"="Disabled";"DmEnrollmentSvc"="Manual";"DPS"="Disabled";"Fax"="Disabled";"FontCache"="Manual";"HotspotService"="Disabled";"lfsvc"="Disabled";"MapsBroker"="Manual";"Spooler"="Disabled";"stisvc"="Manual";"SysMain"="Manual";"TermService"="Disabled";"TrkWks"="Manual";"WlanSvc"="Manual";"WManSvc"="Manual";"wuauserv"="Manual";"WSearch"="Manual";"XblAuthManager"="Disabled";"XblGameSave"="Disabled";"XboxNetApiSvc"="Disabled"}; foreach ($svc in $services.GetEnumerator()) { Set-Service -Name $svc.Key -StartupType $svc.Value -ErrorAction SilentlyContinue } } } Write-Host "$tweak completed" -ForegroundColor Green } } $lblStatusSpeed.Text = "Status: Done"; Write-Host "Speed Up tweaks applied successfully" -ForegroundColor Cyan })
$btnPanelSpeed.Controls.Add($btnRunSpeed)
$tabControl.Controls.Add($tabSpeed)

# Preferences tab
$tabPrefs = New-Object System.Windows.Forms.TabPage
$tabPrefs.Text = "Preferences"
$tabPrefs.BackColor = [System.Drawing.Color]::FromArgb(28, 37, 38)

$panelPrefs = New-Object System.Windows.Forms.Panel
$panelPrefs.Dock = "Top"
$panelPrefs.Height = 400 # Restored original height since no logo panel
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
$btnRunPrefs.Add_Click({ $lblStatusPrefs.Text = "Status: Running..."; Write-Host "Starting Preferences tweaks..." -ForegroundColor Cyan; foreach ($control in $panelPrefs.Controls) { if ($control -is [System.Windows.Forms.CheckBox] -and $control.Checked) { $tweak = $control.Text; Write-Host "Executing: $tweak" -ForegroundColor Yellow; switch ($tweak) { "Dark Theme for Windows" { Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0 -ErrorAction SilentlyContinue } "Add and Activate Ultimate Performance Profile" { powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61; $guid = (powercfg -l | Where-Object { $_ -match "Ultimate Performance" } | ForEach-Object { $_.Split()[3] }); if ($guid) { powercfg -s $guid } } } Write-Host "$tweak completed" -ForegroundColor Green } } $lblStatusPrefs.Text = "Status: Done"; Write-Host "Preferences tweaks applied successfully" -ForegroundColor Cyan })
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

$apps = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName } | Select-Object DisplayName, UninstallString, @{Name="WinGetID";Expression={$_.PSChildName}}
$apps += Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName } | Select-Object DisplayName, UninstallString, @{Name="WinGetID";Expression={$_.PSChildName}}
$apps = $apps | Sort-Object DisplayName
foreach ($app in $apps) { if ($app.DisplayName) { $clbApps.Items.Add($app.DisplayName) | Out-Null } }
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

$btnSelectAllUninstall = New-Object System.Windows.Forms.Button
$btnSelectAllUninstall.Location = New-Object System.Drawing.Point(10,5)
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
$btnUninstall.Location = New-Object System.Drawing.Point(290,5)
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
        $app = $apps | Where-Object { $_.DisplayName -eq $appName }
        if ($app) {
            try {
                $wingetId = (winget list --name "`"$appName`"" --accept-source-agreements | Where-Object { $_ -match $appName }) -split '\s+' | Where-Object { $_ } | Select-Object -Skip 1 -First 1
                if ($wingetId) { winget uninstall --id $wingetId --silent --force } else { throw "No winget ID" }
            } catch {
                if ($app.UninstallString) {
                    $uninst = $app.UninstallString -replace '"', ''
                    if ($uninst -match "msiexec") { Start-Process "msiexec.exe" -ArgumentList "/x $($app.WinGetID) /qn" -Wait -NoNewWindow } else { Start-Process -FilePath $uninst -ArgumentList "/S" -Wait -NoNewWindow -ErrorAction SilentlyContinue }
                }
            }
            Write-Host "$appName uninstalled" -ForegroundColor Green
        }
    }
    $lblStatusUninstall.Text = "Status: Done"
    Write-Host "Uninstallation completed" -ForegroundColor Cyan
    $clbApps.Items.Clear()
    $apps = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName } | Select-Object DisplayName, UninstallString, @{Name="WinGetID";Expression={$_.PSChildName}}
    $apps += Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName } | Select-Object DisplayName, UninstallString, @{Name="WinGetID";Expression={$_.PSChildName}}
    $apps = $apps | Sort-Object DisplayName
    foreach ($app in $apps) { if ($app.DisplayName) { $clbApps.Items.Add($app.DisplayName) | Out-Null } }
})
$btnPanelUninstall.Controls.Add($btnUninstall)
$tabControl.Controls.Add($tabUninstall)

# Adjust button positions on resize
$form.Add_Resize({
    $btnRunSpeed.Location = New-Object System.Drawing.Point(($btnPanelSpeed.Width - $btnRunSpeed.Width - 10), 5)
    $btnRunPrefs.Location = New-Object System.Drawing.Point(($btnPanelPrefs.Width - $btnRunPrefs.Width - 10), 5)
    $btnUninstall.Location = New-Object System.Drawing.Point(($btnPanelUninstall.Width - $btnUninstall.Width - 10), 5)
})

# Show form
$form.ShowDialog()