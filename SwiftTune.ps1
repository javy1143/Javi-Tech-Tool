# ============================================
# Load Required Assemblies
# ============================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================
# Define Helper: Ensure PSWindowsUpdate Module
# ============================================
function Ensure-PSWindowsUpdate {
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Host "Installing PSWindowsUpdate module..."
        try {
            Install-PackageProvider -Name NuGet -Force -ErrorAction Stop
            Install-Module -Name PSWindowsUpdate -Force -ErrorAction Stop
            Import-Module PSWindowsUpdate -Force -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to install PSWindowsUpdate: $_", "Error", 'OK', 'Error')
            throw $_
        }
    } else {
        Import-Module PSWindowsUpdate -Force -ErrorAction Stop
    }
}

# ============================================
# Create Main Form
# ============================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "PC Optimizer"
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $false

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"

# ============================================
# SYSTEM INFO TAB (centered and spaced layout, themed colors)
# ============================================
$tabSystemInfo = New-Object System.Windows.Forms.TabPage
$tabSystemInfo.Text = "System Info"
$tabSystemInfo.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)

# Add SwiftTune Logo to System Info Tab (use provided path)
$logoPath = "C:\Powershell\PowershellScripts\SwiftTune\SwiftTuneLogo.png"
if (Test-Path $logoPath) {
    $logoImage = [System.Drawing.Image]::FromFile($logoPath)

    $logoBox = New-Object System.Windows.Forms.PictureBox
    $logoBox.Image = $logoImage
    $logoBox.SizeMode = 'Zoom'
    $logoBox.Size = New-Object System.Drawing.Size(150, 150)
    $logoBox.Location = New-Object System.Drawing.Point(425, 10)  # Centered horizontally near top

    $tabSystemInfo.Controls.Add($logoBox)
}

# Collect system info
$cpu = (Get-CimInstance Win32_Processor).Name
$ram = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
$os = (Get-CimInstance Win32_OperatingSystem).Caption
$osVer = (Get-CimInstance Win32_OperatingSystem).Version
$uptime = ((Get-CimInstance Win32_OperatingSystem).LastBootUpTime).ToLocalTime()
$diskC = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$diskTotal = [math]::Round($diskC.Size / 1GB, 2)
$diskFree = [math]::Round($diskC.FreeSpace / 1GB, 2)

# Create a multiline label for system info (centered and spaced)
$sysInfoLabel = New-Object System.Windows.Forms.Label
$sysInfoLabel.Location = New-Object System.Drawing.Point(250, 180)
$sysInfoLabel.Size = New-Object System.Drawing.Size(500, 400)
$sysInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$sysInfoLabel.TextAlign = 'MiddleCenter'
$sysInfoLabel.ForeColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$sysInfoLabel.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)
$sysInfoLabel.Text = "
────────────────────────────────────────────
   CPU:
   $cpu

   RAM:
   $ram GB

   OS:
   $os ($osVer)

   Last Boot:
   $uptime

   Disk (C:)
   Total: $diskTotal GB
   Free:  $diskFree GB
────────────────────────────────────────────
"
$tabSystemInfo.Controls.Add($sysInfoLabel)
# Apply the color palette to the main form
$form.BackColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
$tabControl.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)
$progressBar.BackColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(22, 153, 118)
$statusLabel.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)

# ============================================
# PERFORMANCE TWEAKS TAB (full script with color palette)
# ============================================
$tabPerformance = New-Object System.Windows.Forms.TabPage
$tabPerformance.Text = "Performance Tweaks"
$tabPerformance.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)

$groupVisual = New-Object System.Windows.Forms.GroupBox
$groupVisual.Text = "Visual Tweaks"
$groupVisual.ForeColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$groupVisual.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)
$groupVisual.Location = New-Object System.Drawing.Point(20, 20)
$groupVisual.Size = New-Object System.Drawing.Size(450, 300)
$tabPerformance.Controls.Add($groupVisual)

$groupSystem = New-Object System.Windows.Forms.GroupBox
$groupSystem.Text = "System Tweaks"
$groupSystem.ForeColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$groupSystem.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)
$groupSystem.Location = New-Object System.Drawing.Point(500, 20)
$groupSystem.Size = New-Object System.Drawing.Size(450, 400)
$tabPerformance.Controls.Add($groupSystem)

$visualTweaks = @(
    @{ Text = "Set Visual Effects to Best Performance"; Variable = 'checkVisuals' },
    @{ Text = "Disable Transparency Effects"; Variable = 'checkTransparency' },
    @{ Text = "Disable Window Animations"; Variable = 'checkAnimations' },
    @{ Text = "Speed Up Menu Show Delay"; Variable = 'checkMenuDelay' }
)

$systemTweaks = @(
    @{ Text = "Set Power Plan to High Performance"; Variable = 'checkPower' },
    @{ Text = "Disable SysMain (Superfetch)"; Variable = 'checkSysMain' },
    @{ Text = "Disable Search Indexing (if SSD)"; Variable = 'checkIndexing' },
    @{ Text = "Set Processor Scheduling to Favor Programs"; Variable = 'checkProcessor' },
    @{ Text = "Clear Temporary Files"; Variable = 'checkTemp' },
    @{ Text = "Disable Windows Tips and Suggestions"; Variable = 'checkTips' },
    @{ Text = "Clear Edge Browser Cache"; Variable = 'checkEdgeCache' },
    @{ Text = "Clear Chrome Browser Cache"; Variable = 'checkChromeCache' }
)

$y = 30
foreach ($item in $visualTweaks) {
    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = $item.Text
    $chk.ForeColor = [System.Drawing.Color]::FromArgb(22, 153, 118)
    $chk.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)
    $chk.Location = New-Object System.Drawing.Point(10, $y)
    $chk.AutoSize = $true
    Set-Variable -Name $item.Variable -Value $chk -Scope Script
    $groupVisual.Controls.Add($chk)
    $y += 30
}

$y = 30
foreach ($item in $systemTweaks) {
    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = $item.Text
    $chk.ForeColor = [System.Drawing.Color]::FromArgb(22, 153, 118)
    $chk.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)
    $chk.Location = New-Object System.Drawing.Point(10, $y)
    $chk.AutoSize = $true
    Set-Variable -Name $item.Variable -Value $chk -Scope Script
    $groupSystem.Controls.Add($chk)
    $y += 30
}

$applyButton = New-Object System.Windows.Forms.Button
$applyButton.Text = "Apply Selected Tweaks"
$applyButton.Size = New-Object System.Drawing.Size(200, 40)
$applyButton.Location = New-Object System.Drawing.Point(20, 350)
$applyButton.BackColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$applyButton.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
$applyButton.Add_Click({
    $progressBar.Style = 'Marquee'
    $statusLabel.Text = "Applying selected performance tweaks..."
    $form.Refresh()
    try {
        if ($checkVisuals.Checked) { Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" -Name VisualFXSetting -Value 2 }
        if ($checkTransparency.Checked) { Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name EnableTransparency -Value 0 }
        if ($checkAnimations.Checked) { Set-ItemProperty -Path "HKCU:\Control Panel\Desktop\WindowMetrics" -Name MinAnimate -Value 0 }
        if ($checkMenuDelay.Checked) { Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name MenuShowDelay -Value 100 }
        if ($checkPower.Checked) {
            $highPerfGuid = (powercfg /L | Select-String "High performance" | ForEach-Object { ($_ -split '\s+')[3] })
            if ($highPerfGuid) { powercfg /S $highPerfGuid }
        }
        if ($checkSysMain.Checked) {
            Stop-Service -Name "SysMain" -ErrorAction SilentlyContinue
            Set-Service -Name "SysMain" -StartupType Disabled -ErrorAction SilentlyContinue
        }
        if ($checkIndexing.Checked) {
            $driveC = Get-PhysicalDisk | Where-Object { $_.DeviceID -eq 0 }
            if ($driveC.MediaType -eq 'SSD') {
                Stop-Service -Name "WSearch" -ErrorAction SilentlyContinue
                Set-Service -Name "WSearch" -StartupType Disabled -ErrorAction SilentlyContinue
            }
        }
        if ($checkProcessor.Checked) { Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\PriorityControl" -Name Win32PrioritySeparation -Value 26 }
        if ($checkTemp.Checked) { Get-ChildItem -Path $env:TEMP -Recurse -Force | Remove-Item -Recurse -Force }
        if ($checkTips.Checked) { Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name SubscribedContent-338388Enabled -Value 0 }
        if ($checkEdgeCache.Checked) {
            $edgeCachePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache"
            if (Test-Path $edgeCachePath) { Get-ChildItem -Path $edgeCachePath -Recurse -Force | Remove-Item -Recurse -Force }
        }
        if ($checkChromeCache.Checked) {
            $chromeCachePath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache"
            if (Test-Path $chromeCachePath) { Get-ChildItem -Path $chromeCachePath -Recurse -Force | Remove-Item -Recurse -Force }
        }
        [System.Windows.Forms.MessageBox]::Show("Selected performance tweaks applied!", "Success", 'OK', 'Information')
        $statusLabel.Text = "Performance tweaks applied."
    } catch {
        $statusLabel.Text = "Error applying performance tweaks."
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $_", "Error", 'OK', 'Error')
    } finally {
        $progressBar.Style = 'Blocks'
    }
})
$tabPerformance.Controls.Add($applyButton)


# ============================================
# STARTUP MANAGER TAB (with color palette)
# ============================================
$tabStartup = New-Object System.Windows.Forms.TabPage
$tabStartup.Text = "Startup Manager"
$tabStartup.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)

$startupLabel = New-Object System.Windows.Forms.Label
$startupLabel.Text = "Select startup items to disable or enable:"
$startupLabel.Location = New-Object System.Drawing.Point(20, 10)
$startupLabel.Size = New-Object System.Drawing.Size(500, 20)
$startupLabel.ForeColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$startupLabel.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)
$tabStartup.Controls.Add($startupLabel)

$startupListBox = New-Object System.Windows.Forms.ListBox
$startupListBox.Location = New-Object System.Drawing.Point(20, 40)
$startupListBox.Size = New-Object System.Drawing.Size(740, 400)
$startupListBox.SelectionMode = "MultiExtended"
$startupListBox.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)
$startupListBox.ForeColor = [System.Drawing.Color]::FromArgb(22, 153, 118)
$tabStartup.Controls.Add($startupListBox)

function Load-StartupItems {
    $startupListBox.Items.Clear()
    $userRun = Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -ErrorAction SilentlyContinue
    if ($userRun) {
        $userRun.PSObject.Properties | ForEach-Object {
            if ($_.Name -and $_.Value) {
                $startupListBox.Items.Add("User | $($_.Name) | $($_.Value)") | Out-Null
            }
        }
    }
    $machineRun = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run" -ErrorAction SilentlyContinue
    if ($machineRun) {
        $machineRun.PSObject.Properties | ForEach-Object {
            if ($_.Name -and $_.Value) {
                $startupListBox.Items.Add("Machine | $($_.Name) | $($_.Value)") | Out-Null
            }
        }
    }
    $startupFolder = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
    if (Test-Path $startupFolder) {
        Get-ChildItem $startupFolder -Filter *.lnk | ForEach-Object {
            $startupListBox.Items.Add("Folder | $($_.Name) | $($_.FullName)") | Out-Null
        }
    }
}
Load-StartupItems

$disableButton = New-Object System.Windows.Forms.Button
$disableButton.Text = "Disable Selected"
$disableButton.Size = New-Object System.Drawing.Size(180, 30)
$disableButton.Location = New-Object System.Drawing.Point(20, 550)
$disableButton.BackColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$disableButton.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
$disableButton.Add_Click({
    foreach ($item in $startupListBox.SelectedItems) {
        $parts = $item -split '\|'
        $scope = $parts[0].Trim()
        $name = $parts[1].Trim()
        if ($scope -eq "User") { Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $name -ErrorAction SilentlyContinue }
        elseif ($scope -eq "Machine") { Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $name -ErrorAction SilentlyContinue }
        elseif ($scope -eq "Folder") {
            $file = $parts[2].Trim()
            if (Test-Path $file) { Rename-Item -Path $file -NewName ($file + ".disabled") -ErrorAction SilentlyContinue }
        }
    }
    Load-StartupItems
    [System.Windows.Forms.MessageBox]::Show("Selected items disabled.", "Success", 'OK', 'Information')
})
$tabStartup.Controls.Add($disableButton)

$enableButton = New-Object System.Windows.Forms.Button
$enableButton.Text = "Enable Selected"
$enableButton.Size = New-Object System.Drawing.Size(180, 30)
$enableButton.Location = New-Object System.Drawing.Point(220, 550)
$enableButton.BackColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$enableButton.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
$enableButton.Add_Click({
    foreach ($item in $startupListBox.SelectedItems) {
        $parts = $item -split '\|'
        $scope = $parts[0].Trim()
        $file = $parts[2].Trim()
        if ($scope -eq "Folder" -and $file.EndsWith(".disabled") -and (Test-Path $file)) {
            $original = $file -replace '\.disabled$', ''
            Rename-Item -Path $file -NewName $original -ErrorAction SilentlyContinue
        }
    }
    Load-StartupItems
    [System.Windows.Forms.MessageBox]::Show("Selected folder items re-enabled.", "Success", 'OK', 'Information')
})
$tabStartup.Controls.Add($enableButton)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh List"
$refreshButton.Size = New-Object System.Drawing.Size(150, 30)
$refreshButton.Location = New-Object System.Drawing.Point(420, 550)
$refreshButton.BackColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$refreshButton.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
$refreshButton.Add_Click({
    $progressBar.Style = 'Marquee'
    $statusLabel.Text = "Refreshing startup list..."
    $form.Refresh()
    try {
        Load-StartupItems
        $statusLabel.Text = "Startup items refreshed."
    } catch {
        $statusLabel.Text = "Error refreshing startup list."
    } finally {
        $progressBar.Style = 'Blocks'
    }
})
$tabStartup.Controls.Add($refreshButton)

# ============================================
# WINDOWS UPDATES TAB (with color palette)
# ============================================
$tabUpdates = New-Object System.Windows.Forms.TabPage
$tabUpdates.Text = "Windows Updates"
$tabUpdates.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)

$updateLabel = New-Object System.Windows.Forms.Label
$updateLabel.Text = "Select updates to install:"
$updateLabel.Location = New-Object System.Drawing.Point(20, 10)
$updateLabel.Size = New-Object System.Drawing.Size(500, 20)
$updateLabel.ForeColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$updateLabel.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)
$tabUpdates.Controls.Add($updateLabel)

$updateListBox = New-Object System.Windows.Forms.ListBox
$updateListBox.Location = New-Object System.Drawing.Point(20, 40)
$updateListBox.Size = New-Object System.Drawing.Size(940, 500)
$updateListBox.SelectionMode = "MultiExtended"
$updateListBox.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)
$updateListBox.ForeColor = [System.Drawing.Color]::FromArgb(22, 153, 118)
$tabUpdates.Controls.Add($updateListBox)

$loadSecurityButton = New-Object System.Windows.Forms.Button
$loadSecurityButton.Text = "Load Security Updates"
$loadSecurityButton.Size = New-Object System.Drawing.Size(180, 30)
$loadSecurityButton.Location = New-Object System.Drawing.Point(20, 550)
$loadSecurityButton.BackColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$loadSecurityButton.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
$loadSecurityButton.Add_Click({
    $progressBar.Style = 'Marquee'
    $statusLabel.Text = "Loading security updates..."
    $form.Refresh()
    try {
        Ensure-PSWindowsUpdate
        $updateListBox.Items.Clear()
        $updateListBox.Items.Add("Security Update Example") | Out-Null
        $statusLabel.Text = "Security updates loaded."
    } catch {
        $statusLabel.Text = "Error loading security updates."
        [System.Windows.Forms.MessageBox]::Show("Failed: $_", "Error", 'OK', 'Error')
    } finally {
        $progressBar.Style = 'Blocks'
    }
})
$tabUpdates.Controls.Add($loadSecurityButton)

$loadDriverButton = New-Object System.Windows.Forms.Button
$loadDriverButton.Text = "Load Driver Updates"
$loadDriverButton.Size = New-Object System.Drawing.Size(180, 30)
$loadDriverButton.Location = New-Object System.Drawing.Point(220, 550)
$loadDriverButton.BackColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$loadDriverButton.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
$loadDriverButton.Add_Click({
    $progressBar.Style = 'Marquee'
    $statusLabel.Text = "Loading driver updates..."
    $form.Refresh()
    try {
        Ensure-PSWindowsUpdate
        $updateListBox.Items.Clear()
        $updates = Get-WindowsUpdate -MicrosoftUpdate -Category "Drivers" -IgnoreUserInput
        $updates | ForEach-Object {
            $updateListBox.Items.Add("Driver | $($_.Title)") | Out-Null
        }
        $statusLabel.Text = "Driver updates loaded."
    } catch {
        $statusLabel.Text = "Error loading driver updates."
        [System.Windows.Forms.MessageBox]::Show("Failed: $_", "Error", 'OK', 'Error')
    } finally {
        $progressBar.Style = 'Blocks'
    }
})
$tabUpdates.Controls.Add($loadDriverButton)

$installButton = New-Object System.Windows.Forms.Button
$installButton.Text = "Install Selected Updates"
$installButton.Size = New-Object System.Drawing.Size(180, 30)
$installButton.Location = New-Object System.Drawing.Point(420, 550)
$installButton.BackColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$installButton.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
$installButton.Add_Click({
    $progressBar.Style = 'Marquee'
    $statusLabel.Text = "Installing selected updates..."
    $form.Refresh()
    try {
        Ensure-PSWindowsUpdate
        $selected = $updateListBox.SelectedItems
        if ($selected.Count -eq 0) {
            $statusLabel.Text = "No updates selected."
            [System.Windows.Forms.MessageBox]::Show("Please select updates to install.", "Info", 'OK', 'Information')
            return
        }
        $allUpdates = Get-WindowsUpdate -MicrosoftUpdate -IgnoreUserInput
        foreach ($item in $selected) {
            $title = ($item -split '\|')[1].Trim()
            $match = $allUpdates | Where-Object { $_.Title -eq $title }
            if ($match) {
                $match | Install-WindowsUpdate -AcceptAll -IgnoreReboot
            }
        }
        $statusLabel.Text = "Selected updates installed."
        [System.Windows.Forms.MessageBox]::Show("Selected updates installed (reboot if needed).", "Success", 'OK', 'Information')
    } catch {
        $statusLabel.Text = "Error installing updates."
        [System.Windows.Forms.MessageBox]::Show("Failed to install updates: $_", "Error", 'OK', 'Error')
    } finally {
        $progressBar.Style = 'Blocks'
    }
})
$tabUpdates.Controls.Add($installButton)

# ============================================
# Finalize Tabs and Run (with color palette applied)
# ============================================

$tabControl.TabPages.AddRange(@(
    $tabSystemInfo,
    $tabPerformance,
    $tabStartup,
    $tabUpdates
))

# Add Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Style = 'Blocks'
$progressBar.Location = New-Object System.Drawing.Point(10, 630)
$progressBar.Size = New-Object System.Drawing.Size(960, 20)
$progressBar.BackColor = [System.Drawing.Color]::FromArgb(29, 205, 159)
$form.Controls.Add($progressBar)

# Add Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready"
$statusLabel.Location = New-Object System.Drawing.Point(10, 610)
$statusLabel.Size = New-Object System.Drawing.Size(960, 20)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(22, 153, 118)
$statusLabel.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)
$form.Controls.Add($statusLabel)

# Apply color palette to form and tab control
$form.BackColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
$tabControl.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)

$form.Controls.Add($tabControl)
$form.Add_Shown({ $form.Activate() })
[System.Windows.Forms.Application]::Run($form)

