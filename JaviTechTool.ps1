# JaviTechTool.ps1
# Integrated tool with tabs for software install, Windows tweaks, custom preferences,
# Windows updates, and a bulk uninstaller.

# -------------------------
# Ensure script is run as Administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $scriptPath = $MyInvocation.MyCommand.Definition
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    Start-Process powershell.exe -Verb RunAs -ArgumentList $arguments
    exit
}

# -------------------------
# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -------------------------
# Global log text box (will be added to the form)
$global:logBox = $null

# -------------------------
# Helper: Log messages to console and log window
function Log-Message {
    param ([string]$message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $line = "[$timestamp] $message"
    Write-Host $line
    if ($global:logBox -ne $null) {
        $global:logBox.AppendText("$line`r`n")
        # Optionally, scroll to the bottom:
        $global:logBox.SelectionStart = $global:logBox.Text.Length
        $global:logBox.ScrollToCaret()
    }
}

# -------------------------
# Helper: Download file using BITS with fallback to WebClient
function Download-File {
    param(
        [Parameter(Mandatory=$true)][string]$url,
        [Parameter(Mandatory=$true)][string]$destination
    )
    try {
        Log-Message "Attempting download using BITS Transfer..."
        Start-BitsTransfer -Source $url -Destination $destination -ErrorAction Stop
    } catch {
        Log-Message "BITS Transfer failed, falling back to WebClient: $($_.Exception.Message)"
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($url, $destination)
    }
}

# -------------------------
# Software Installation Functions

function Install-GoogleChrome {
    Log-Message "Installing Google Chrome..."
    $url = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi"
    $tempFile = "$env:TEMP\googlechrome_installer.msi"
    try {
        Log-Message "Downloading Google Chrome installer..."
        Download-File -url $url -destination $tempFile
        Log-Message "Download complete. Executing installer..."
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$tempFile`" /quiet /norestart" -NoNewWindow -Wait
        Log-Message "Google Chrome installation completed."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error installing Google Chrome: $($_.Exception.Message)")
    } finally {
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
    }
}

function Install-RingCentral {
    Log-Message "Installing RingCentral..."
    $url = "https://app.ringcentral.com/download/RingCentral.exe?V=20138600535791900"
    $tempFile = "$env:TEMP\RingCentralInstaller.exe"
    try {
        Log-Message "Downloading RingCentral installer..."
        Download-File -url $url -destination $tempFile
        Log-Message "Download complete. Executing installer..."
        # Adjust silent switches as needed; /S is assumed here.
        Start-Process -FilePath $tempFile -ArgumentList "/S" -NoNewWindow -Wait
        Log-Message "RingCentral installation completed."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error installing RingCentral: $($_.Exception.Message)")
    } finally {
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
    }
}

function Install-MicrosoftTeams {
    Log-Message "Installing Microsoft Teams..."
    $url = "https://go.microsoft.com/fwlink/?linkid=2281613&clcid=0x409&culture=en-us&country=us"
    $tempFile = "$env:TEMP\TeamsInstaller.exe"
    try {
        Log-Message "Downloading Microsoft Teams installer..."
        Download-File -url $url -destination $tempFile
        Log-Message "Download complete. Executing installer..."
        Start-Process -FilePath $tempFile -ArgumentList "/S" -NoNewWindow -Wait
        Log-Message "Microsoft Teams installation completed."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error installing Microsoft Teams: $($_.Exception.Message)")
    } finally {
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
    }
}

# -------------------------
# Speed Up Windows Functions

function Clear-TempFiles {
    Log-Message "Clearing Temporary Files..."
    $tempPaths = @("C:\Windows\Temp", $env:TEMP)
    foreach ($path in $tempPaths) {
        try {
            $files = Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue
            $totalFiles = $files.Count
            $currentFile = 0
            foreach ($file in $files) {
                $currentFile++
                $percentComplete = ($currentFile / $totalFiles) * 100
                Write-Progress -Activity "Clearing Temp Files" -Status "Deleting $($file.FullName)" -PercentComplete $percentComplete
                try {
                    if ($file.Attributes -band [System.IO.FileAttributes]::ReadOnly) {
                        $file.Attributes = $file.Attributes -bxor [System.IO.FileAttributes]::ReadOnly
                    }
                    Remove-Item $file.FullName -Force -Recurse -ErrorAction Stop
                } catch {
                    Log-Message "Failed to delete file: $($file.FullName) - $($_.Exception.Message)"
                }
            }
            Log-Message "Cleared files in $path"
        } catch {
            Log-Message "Error processing $($path): $($_.Exception.Message)"
        }
    }
    Write-Progress -Activity "Clearing Temp Files" -Completed
    Log-Message "Temporary files cleanup completed."
}

function Disable-ConsumerFeatures {
    Log-Message "Disabling Consumer Features..."
    try {
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
        if (-not (Test-Path $regPath)) {
            New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows" -Name "CloudContent" -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "DisableWindowsConsumerFeatures" -Value 1 -Type DWord -ErrorAction Stop
        Log-Message "Consumer features disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling consumer features: $($_.Exception.Message)")
    }
}

function Disable-Telemetry {
    Log-Message "Disabling Telemetry..."
    try {
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
        if (-not (Test-Path $regPath)) {
            New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows" -Name "DataCollection" -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "AllowTelemetry" -Value 0 -Type DWord -ErrorAction Stop
        Log-Message "Telemetry disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling telemetry: $($_.Exception.Message)")
    }
}

function Disable-ActiveHistory {
    Log-Message "Disabling Active History..."
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        New-ItemProperty -Path $regPath -Name "EnableActivityFeed" -Value 0 -PropertyType DWord -Force | Out-Null
        New-ItemProperty -Path $regPath -Name "PublishUserActivities" -Value 0 -PropertyType DWord -Force | Out-Null
        New-ItemProperty -Path $regPath -Name "UploadUserActivities" -Value 0 -PropertyType DWord -Force | Out-Null
        Log-Message "Active History disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling active history: $($_.Exception.Message)")
    }
}

function Disable-GameDVR {
    Log-Message "Disabling Game DVR..."
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR"
        if (-not (Test-Path $regPath)) {
            New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion" -Name "GameDVR" -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "GameDVR_FSEBehavior" -Value 2 -Type DWord -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name "GameDVR_Enabled" -Value 0 -Type DWord -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name "GameDVR_HonorUserFSEBehaviorMode" -Value 1 -Type DWord -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name "GameDVR_EFSEFeatureFlags" -Value 0 -Type DWord -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name "AllowGameDVR" -Value 0 -Type DWord -ErrorAction Stop
        Log-Message "Game DVR has been disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling Game DVR: $($_.Exception.Message)")
    }
}

function Disable-Hibernation {
    Log-Message "Disabling Hibernation..."
    try {
        powercfg.exe /hibernate off
        Log-Message "Hibernation disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling hibernation: $($_.Exception.Message)")
    }
}

function Disable-HomeGroup {
    Log-Message "Disabling HomeGroup..."
    try {
        sc.exe config HomeGroupListener start= manual | Out-Null
        sc.exe config HomeGroupProvider start= manual | Out-Null
        Log-Message "HomeGroup disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling HomeGroup: $($_.Exception.Message)")
    }
}

function Disable-LocationTracking {
    Log-Message "Disabling Location Tracking..."
    try {
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors"
        if (-not (Test-Path $regPath)) {
            New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows" -Name "LocationAndSensors" -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "Value" -Value "Deny" -Type String -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name "SensorPermissionState" -Value 0 -Type DWord -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name "Status" -Value 0 -Type DWord -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name "AutoUpdateEnabled" -Value 0 -Type DWord -ErrorAction Stop
        Log-Message "Location Tracking disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling location tracking: $($_.Exception.Message)")
    }
}

function Disable-StorageSense {
    Log-Message "Disabling Storage Sense..."
    try {
        $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy"
        if (-not (Test-Path $regPath)) {
            New-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\StorageSense\Parameters" -Name "StoragePolicy" -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "01" -Value 0 -Type DWord -Force -ErrorAction Stop
        Log-Message "Storage Sense disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling Storage Sense: $($_.Exception.Message)")
    }
}

function Disable-WifiSense {
    Log-Message "Disabling Wifi Sense..."
    try {
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WcmSvc"
        if (-not (Test-Path $regPath)) {
            New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows" -Name "WcmSvc" -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "DisableWifiSense" -Value 1 -Type DWord -ErrorAction Stop
        Log-Message "Wifi Sense disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling Wifi Sense: $($_.Exception.Message)")
    }
}

function Enable-EndTaskRightClick {
    Log-Message "Enabling End Task with Right Click..."
    try {
        $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings"
        if (-not (Test-Path $path)) {
            New-Item -Path $path -Force | Out-Null
        }
        New-ItemProperty -Path $path -Name "TaskbarEndTask" -PropertyType DWord -Value 1 -Force | Out-Null
        Log-Message "End Task with Right Click enabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error enabling End Task with Right Click: $($_.Exception.Message)")
    }
}

function Disable-Powershell7Telemetry {
    Log-Message "Disabling Powershell 7 Telemetry..."
    try {
        [Environment]::SetEnvironmentVariable('POWERSHELL_TELEMETRY_OPTOUT', '1', 'Machine')
        Log-Message "Powershell 7 Telemetry disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling Powershell 7 Telemetry: $($_.Exception.Message)")
    }
}

function Set-Services {
    Log-Message "Setting Services based on embedded table..."
    $services = @(
        @{ ServiceName = "AJRouter";                   StartupType = "Disabled" },
        @{ ServiceName = "ALG";                        StartupType = "Manual" },
        @{ ServiceName = "AppIDSvc";                   StartupType = "Manual" },
        @{ ServiceName = "AppMgmt";                    StartupType = "Manual" },
        @{ ServiceName = "AppReadiness";               StartupType = "Manual" },
        @{ ServiceName = "AppVClient";                 StartupType = "Disabled" },
        @{ ServiceName = "AppXSvc";                    StartupType = "Manual" },
        @{ ServiceName = "Appinfo";                    StartupType = "Manual" },
        @{ ServiceName = "AssignedAccessManagerSvc";   StartupType = "Disabled" },
        @{ ServiceName = "AudioEndpointBuilder";       StartupType = "Automatic" },
        @{ ServiceName = "AudioSrv";                   StartupType = "Automatic" },
        @{ ServiceName = "Audiosrv";                   StartupType = "Automatic" },
        @{ ServiceName = "AxInstSV";                   StartupType = "Manual" },
        @{ ServiceName = "BDESVC";                     StartupType = "Manual" },
        @{ ServiceName = "BFE";                        StartupType = "Automatic" },
        @{ ServiceName = "BITS";                       StartupType = "AutomaticDelayedStart" },
        @{ ServiceName = "BTAGService";                StartupType = "Manual" },
        @{ ServiceName = "BcastDVRUserService_*";      StartupType = "Manual" },
        @{ ServiceName = "BluetoothUserService_*";     StartupType = "Manual" },
        @{ ServiceName = "BrokerInfrastructure";       StartupType = "Automatic" },
        @{ ServiceName = "Browser";                    StartupType = "Manual" },
        @{ ServiceName = "BthAvctpSvc";                StartupType = "Automatic" },
        @{ ServiceName = "BthHFSrv";                   StartupType = "Automatic" },
        @{ ServiceName = "CDPSvc";                     StartupType = "Manual" },
        @{ ServiceName = "CDPUserSvc_*";               StartupType = "Automatic" },
        @{ ServiceName = "COMSysApp";                  StartupType = "Manual" },
        @{ ServiceName = "CaptureService_*";           StartupType = "Manual" },
        @{ ServiceName = "CertPropSvc";                StartupType = "Manual" },
        @{ ServiceName = "ClipSVC";                    StartupType = "Manual" },
        @{ ServiceName = "ConsentUxUserSvc_*";         StartupType = "Manual" },
        @{ ServiceName = "CoreMessagingRegistrar";     StartupType = "Automatic" },
        @{ ServiceName = "CredentialEnrollmentManagerUserSvc_*"; StartupType = "Manual" },
        @{ ServiceName = "CryptSvc";                   StartupType = "Automatic" },
        @{ ServiceName = "CscService";                 StartupType = "Manual" },
        @{ ServiceName = "DPS";                        StartupType = "Automatic" },
        @{ ServiceName = "DcomLaunch";                 StartupType = "Automatic" },
        @{ ServiceName = "DcpSvc";                     StartupType = "Manual" },
        @{ ServiceName = "DevQueryBroker";             StartupType = "Manual" },
        @{ ServiceName = "DeviceAssociationBrokerSvc_*"; StartupType = "Manual" },
        @{ ServiceName = "DeviceAssociationService";   StartupType = "Manual" },
        @{ ServiceName = "DeviceInstall";              StartupType = "Manual" },
        @{ ServiceName = "DevicePickerUserSvc_*";      StartupType = "Manual" },
        @{ ServiceName = "DevicesFlowUserSvc_*";       StartupType = "Manual" },
        @{ ServiceName = "Dhcp";                       StartupType = "Automatic" },
        @{ ServiceName = "DiagTrack";                  StartupType = "Disabled" },
        @{ ServiceName = "DialogBlockingService";      StartupType = "Disabled" },
        @{ ServiceName = "DispBrokerDesktopSvc";       StartupType = "Automatic" },
        @{ ServiceName = "DisplayEnhancementService";  StartupType = "Manual" },
        @{ ServiceName = "DmEnrollmentSvc";            StartupType = "Manual" },
        @{ ServiceName = "Dnscache";                   StartupType = "Automatic" },
        @{ ServiceName = "DoSvc";                      StartupType = "AutomaticDelayedStart" },
        @{ ServiceName = "DsSvc";                      StartupType = "Manual" },
        @{ ServiceName = "DsmSvc";                     StartupType = "Manual" },
        @{ ServiceName = "DusmSvc";                    StartupType = "Automatic" },
        @{ ServiceName = "EFS";                        StartupType = "Manual" },
        @{ ServiceName = "EapHost";                    StartupType = "Manual" },
        @{ ServiceName = "EntAppSvc";                  StartupType = "Manual" },
        @{ ServiceName = "EventLog";                   StartupType = "Automatic" },
        @{ ServiceName = "EventSystem";                StartupType = "Automatic" },
        @{ ServiceName = "FDResPub";                   StartupType = "Manual" },
        @{ ServiceName = "Fax";                        StartupType = "Manual" },
        @{ ServiceName = "FontCache";                  StartupType = "Automatic" },
        @{ ServiceName = "FrameServer";                StartupType = "Manual" },
        @{ ServiceName = "FrameServerMonitor";         StartupType = "Manual" },
        @{ ServiceName = "GraphicsPerfSvc";            StartupType = "Manual" },
        @{ ServiceName = "HomeGroupListener";          StartupType = "Manual" },
        @{ ServiceName = "HomeGroupProvider";          StartupType = "Manual" },
        @{ ServiceName = "HvHost";                     StartupType = "Manual" },
        @{ ServiceName = "IEEtwCollectorService";      StartupType = "Manual" },
        @{ ServiceName = "IKEEXT";                     StartupType = "Manual" },
        @{ ServiceName = "InstallService";             StartupType = "Manual" },
        @{ ServiceName = "InventorySvc";               StartupType = "Manual" },
        @{ ServiceName = "IpxlatCfgSvc";               StartupType = "Manual" },
        @{ ServiceName = "KeyIso";                     StartupType = "Automatic" },
        @{ ServiceName = "KtmRm";                      StartupType = "Manual" },
        @{ ServiceName = "LSM";                        StartupType = "Automatic" },
        @{ ServiceName = "LanmanServer";               StartupType = "Automatic" },
        @{ ServiceName = "LanmanWorkstation";          StartupType = "Automatic" },
        @{ ServiceName = "LicenseManager";             StartupType = "Manual" },
        @{ ServiceName = "LxpSvc";                     StartupType = "Manual" },
        @{ ServiceName = "MSDTC";                      StartupType = "Manual" },
        @{ ServiceName = "MSiSCSI";                    StartupType = "Manual" },
        @{ ServiceName = "MapsBroker";                 StartupType = "AutomaticDelayedStart" },
        @{ ServiceName = "McpManagementService";       StartupType = "Manual" },
        @{ ServiceName = "MessagingService_*";         StartupType = "Manual" },
        @{ ServiceName = "MicrosoftEdgeElevationService"; StartupType = "Manual" },
        @{ ServiceName = "MixedRealityOpenXRSvc";      StartupType = "Manual" },
        @{ ServiceName = "MpsSvc";                     StartupType = "Automatic" },
        @{ ServiceName = "MsKeyboardFilter";           StartupType = "Manual" },
        @{ ServiceName = "NPSMSvc_*";                  StartupType = "Manual" },
        @{ ServiceName = "NaturalAuthentication";      StartupType = "Manual" },
        @{ ServiceName = "NcaSvc";                     StartupType = "Manual" },
        @{ ServiceName = "NcbService";                 StartupType = "Manual" },
        @{ ServiceName = "NcdAutoSetup";               StartupType = "Manual" },
        @{ ServiceName = "NetSetupSvc";                StartupType = "Manual" },
        @{ ServiceName = "NetTcpPortSharing";          StartupType = "Disabled" },
        @{ ServiceName = "Netlogon";                   StartupType = "Automatic" },
        @{ ServiceName = "Netman";                     StartupType = "Manual" },
        @{ ServiceName = "NgcCtnrSvc";                 StartupType = "Manual" },
        @{ ServiceName = "NgcSvc";                     StartupType = "Manual" },
        @{ ServiceName = "NlaSvc";                     StartupType = "Manual" },
        @{ ServiceName = "OneSyncSvc_*";              StartupType = "Automatic" },
        @{ ServiceName = "P9RdrService_*";            StartupType = "Manual" },
        @{ ServiceName = "PNRPAutoReg";                StartupType = "Manual" },
        @{ ServiceName = "PNRPsvc";                    StartupType = "Manual" },
        @{ ServiceName = "PcaSvc";                     StartupType = "Manual" },
        @{ ServiceName = "PeerDistSvc";                StartupType = "Manual" },
        @{ ServiceName = "PenService_*";              StartupType = "Manual" },
        @{ ServiceName = "PerfHost";                   StartupType = "Manual" },
        @{ ServiceName = "PhoneSvc";                   StartupType = "Manual" },
        @{ ServiceName = "PimIndexMaintenanceSvc_*";   StartupType = "Manual" },
        @{ ServiceName = "PlugPlay";                   StartupType = "Manual" },
        @{ ServiceName = "PolicyAgent";                StartupType = "Manual" },
        @{ ServiceName = "Power";                      StartupType = "Automatic" },
        @{ ServiceName = "PrintNotify";                StartupType = "Manual" },
        @{ ServiceName = "PrintWorkflowUserSvc_*";     StartupType = "Manual" },
        @{ ServiceName = "ProfSvc";                    StartupType = "Automatic" },
        @{ ServiceName = "PushToInstall";              StartupType = "Manual" },
        @{ ServiceName = "QWAVE";                      StartupType = "Manual" },
        @{ ServiceName = "RasAuto";                    StartupType = "Manual" },
        @{ ServiceName = "RasMan";                     StartupType = "Manual" },
        @{ ServiceName = "RemoteAccess";               StartupType = "Disabled" },
        @{ ServiceName = "RemoteRegistry";             StartupType = "Disabled" },
        @{ ServiceName = "RetailDemo";                 StartupType = "Manual" },
        @{ ServiceName = "RmSvc";                      StartupType = "Manual" },
        @{ ServiceName = "RpcEptMapper";               StartupType = "Automatic" },
        @{ ServiceName = "RpcLocator";                 StartupType = "Manual" },
        @{ ServiceName = "RpcSs";                      StartupType = "Automatic" },
        @{ ServiceName = "SCPolicySvc";                StartupType = "Manual" },
        @{ ServiceName = "SCardSvr";                   StartupType = "Manual" },
        @{ ServiceName = "SDRSVC";                     StartupType = "Manual" },
        @{ ServiceName = "SEMgrSvc";                   StartupType = "Manual" },
        @{ ServiceName = "SENS";                       StartupType = "Automatic" },
        @{ ServiceName = "SNMPTRAP";                   StartupType = "Manual" },
        @{ ServiceName = "SNMPTrap";                   StartupType = "Manual" },
        @{ ServiceName = "SSDPSRV";                    StartupType = "Manual" },
        @{ ServiceName = "SamSs";                      StartupType = "Automatic" },
        @{ ServiceName = "ScDeviceEnum";               StartupType = "Manual" },
        @{ ServiceName = "Schedule";                   StartupType = "Automatic" },
        @{ ServiceName = "SecurityHealthService";      StartupType = "Manual" },
        @{ ServiceName = "Sense";                      StartupType = "Manual" },
        @{ ServiceName = "SensorDataService";          StartupType = "Manual" },
        @{ ServiceName = "SensorService";              StartupType = "Manual" },
        @{ ServiceName = "SensrSvc";                   StartupType = "Manual" },
        @{ ServiceName = "SessionEnv";                 StartupType = "Manual" },
        @{ ServiceName = "SgrmBroker";                 StartupType = "Automatic" },
        @{ ServiceName = "SharedAccess";               StartupType = "Manual" },
        @{ ServiceName = "SharedRealitySvc";           StartupType = "Manual" },
        @{ ServiceName = "ShellHWDetection";           StartupType = "Automatic" },
        @{ ServiceName = "SmsRouter";                  StartupType = "Manual" },
        @{ ServiceName = "Spooler";                    StartupType = "Automatic" },
        @{ ServiceName = "SstpSvc";                    StartupType = "Manual" },
        @{ ServiceName = "StateRepository";            StartupType = "Manual" },
        @{ ServiceName = "StiSvc";                     StartupType = "Manual" },
        @{ ServiceName = "StorSvc";                    StartupType = "Manual" },
        @{ ServiceName = "SysMain";                    StartupType = "Automatic" },
        @{ ServiceName = "SystemEventsBroker";         StartupType = "Automatic" },
        @{ ServiceName = "TabletInputService";         StartupType = "Manual" },
        @{ ServiceName = "TapiSrv";                    StartupType = "Manual" },
        @{ ServiceName = "TermService";                StartupType = "Automatic" },
        @{ ServiceName = "TextInputManagementService"; StartupType = "Manual" },
        @{ ServiceName = "Themes";                     StartupType = "Automatic" },
        @{ ServiceName = "TieringEngineService";       StartupType = "Manual" },
        @{ ServiceName = "TimeBroker";                 StartupType = "Manual" },
        @{ ServiceName = "TimeBrokerSvc";              StartupType = "Manual" },
        @{ ServiceName = "TokenBroker";                StartupType = "Manual" },
        @{ ServiceName = "TrkWks";                     StartupType = "Automatic" },
        @{ ServiceName = "TroubleshootingSvc";         StartupType = "Manual" },
        @{ ServiceName = "TrustedInstaller";           StartupType = "Manual" },
        @{ ServiceName = "UI0Detect";                  StartupType = "Manual" },
        @{ ServiceName = "UdkUserSvc_*";               StartupType = "Manual" },
        @{ ServiceName = "UevAgentService";            StartupType = "Disabled" },
        @{ ServiceName = "UmRdpService";               StartupType = "Manual" },
        @{ ServiceName = "UnistoreSvc_*";              StartupType = "Manual" },
        @{ ServiceName = "UserDataSvc_*";              StartupType = "Manual" },
        @{ ServiceName = "UserManager";                StartupType = "Automatic" },
        @{ ServiceName = "UsoSvc";                     StartupType = "Manual" },
        @{ ServiceName = "VGAuthService";              StartupType = "Automatic" },
        @{ ServiceName = "VMTools";                    StartupType = "Automatic" },
        @{ ServiceName = "VSS";                        StartupType = "Manual" },
        @{ ServiceName = "VacSvc";                     StartupType = "Manual" },
        @{ ServiceName = "VaultSvc";                   StartupType = "Automatic" },
        @{ ServiceName = "W32Time";                    StartupType = "Manual" },
        @{ ServiceName = "WEPHOSTSVC";                 StartupType = "Manual" },
        @{ ServiceName = "WFDSConMgrSvc";              StartupType = "Manual" },
        @{ ServiceName = "WMPNetworkSvc";              StartupType = "Manual" },
        @{ ServiceName = "WManSvc";                    StartupType = "Manual" },
        @{ ServiceName = "WPDBusEnum";                 StartupType = "Manual" },
        @{ ServiceName = "WSService";                  StartupType = "Manual" },
        @{ ServiceName = "WSearch";                    StartupType = "AutomaticDelayedStart" },
        @{ ServiceName = "WaaSMedicSvc";               StartupType = "Manual" },
        @{ ServiceName = "WalletService";              StartupType = "Manual" },
        @{ ServiceName = "WarpJITSvc";                StartupType = "Manual" },
        @{ ServiceName = "WbioSrvc";                   StartupType = "Manual" },
        @{ ServiceName = "Wcmsvc";                     StartupType = "Automatic" },
        @{ ServiceName = "WcsPlugInService";           StartupType = "Manual" },
        @{ ServiceName = "WdNisSvc";                   StartupType = "Manual" },
        @{ ServiceName = "WdiServiceHost";             StartupType = "Manual" },
        @{ ServiceName = "WdiSystemHost";              StartupType = "Manual" },
        @{ ServiceName = "WebClient";                  StartupType = "Manual" },
        @{ ServiceName = "Wecsvc";                     StartupType = "Manual" },
        @{ ServiceName = "WerSvc";                     StartupType = "Manual" },
        @{ ServiceName = "WiaRpc";                     StartupType = "Manual" },
        @{ ServiceName = "WinDefend";                  StartupType = "Automatic" },
        @{ ServiceName = "WinHttpAutoProxySvc";        StartupType = "Manual" },
        @{ ServiceName = "WinRM";                      StartupType = "Manual" },
        @{ ServiceName = "Winmgmt";                    StartupType = "Automatic" },
        @{ ServiceName = "WlanSvc";                    StartupType = "Automatic" },
        @{ ServiceName = "WpcMonSvc";                  StartupType = "Manual" },
        @{ ServiceName = "WpnService";                 StartupType = "Manual" },
        @{ ServiceName = "WpnUserService_*";           StartupType = "Automatic" },
        @{ ServiceName = "XblAuthManager";             StartupType = "Manual" },
        @{ ServiceName = "XblGameSave";                StartupType = "Manual" },
        @{ ServiceName = "XboxGipSvc";                 StartupType = "Manual" },
        @{ ServiceName = "XboxNetApiSvc";              StartupType = "Manual" },
        @{ ServiceName = "autotimesvc";                StartupType = "Manual" },
        @{ ServiceName = "bthserv";                    StartupType = "Manual" },
        @{ ServiceName = "camsvc";                     StartupType = "Manual" },
        @{ ServiceName = "cbdhsvc_*";                  StartupType = "Manual" },
        @{ ServiceName = "cloudidsvc";                 StartupType = "Manual" },
        @{ ServiceName = "dcsvc";                      StartupType = "Manual" },
        @{ ServiceName = "defragsvc";                  StartupType = "Manual" },
        @{ ServiceName = "diagnosticshub.standardcollector.service"; StartupType = "Manual" },
        @{ ServiceName = "diagsvc";                    StartupType = "Manual" },
        @{ ServiceName = "dmwappushservice";           StartupType = "Manual" },
        @{ ServiceName = "dot3svc";                    StartupType = "Manual" },
        @{ ServiceName = "edgeupdate";                 StartupType = "Manual" },
        @{ ServiceName = "edgeupdatem";                StartupType = "Manual" },
        @{ ServiceName = "embeddedmode";               StartupType = "Manual" },
        @{ ServiceName = "fdPHost";                    StartupType = "Manual" },
        @{ ServiceName = "fhsvc";                      StartupType = "Manual" },
        @{ ServiceName = "gpsvc";                      StartupType = "Automatic" },
        @{ ServiceName = "hidserv";                    StartupType = "Manual" },
        @{ ServiceName = "icssvc";                     StartupType = "Manual" },
        @{ ServiceName = "iphlpsvc";                   StartupType = "Automatic" },
        @{ ServiceName = "lfsvc";                      StartupType = "Manual" },
        @{ ServiceName = "lltdsvc";                    StartupType = "Manual" },
        @{ ServiceName = "lmhosts";                    StartupType = "Manual" },
        @{ ServiceName = "mpssvc";                     StartupType = "Automatic" },
        @{ ServiceName = "msiserver";                  StartupType = "Manual" },
        @{ ServiceName = "netprofm";                   StartupType = "Manual" },
        @{ ServiceName = "nsi";                        StartupType = "Automatic" },
        @{ ServiceName = "p2pimsvc";                   StartupType = "Manual" },
        @{ ServiceName = "p2psvc";                     StartupType = "Manual" },
        @{ ServiceName = "perceptionsimulation";       StartupType = "Manual" },
        @{ ServiceName = "pla";                        StartupType = "Manual" },
        @{ ServiceName = "seclogon";                   StartupType = "Manual" },
        @{ ServiceName = "shpamsvc";                   StartupType = "Disabled" },
        @{ ServiceName = "smphost";                    StartupType = "Manual" },
        @{ ServiceName = "spectrum";                   StartupType = "Manual" },
        @{ ServiceName = "sppsvc";                     StartupType = "AutomaticDelayedStart" },
        @{ ServiceName = "ssh-agent";                  StartupType = "Disabled" },
        @{ ServiceName = "svsvc";                      StartupType = "Manual" },
        @{ ServiceName = "swprv";                      StartupType = "Manual" },
        @{ ServiceName = "tiledatamodelsvc";           StartupType = "Automatic" },
        @{ ServiceName = "tzautoupdate";               StartupType = "Disabled" },
        @{ ServiceName = "uhssvc";                     StartupType = "Disabled" },
        @{ ServiceName = "upnphost";                   StartupType = "Manual" },
        @{ ServiceName = "vds";                        StartupType = "Manual" },
        @{ ServiceName = "vm3dservice";                StartupType = "Manual" },
        @{ ServiceName = "vmicguestinterface";         StartupType = "Manual" },
        @{ ServiceName = "vmicheartbeat";              StartupType = "Manual" },
        @{ ServiceName = "vmickvpexchange";            StartupType = "Manual" },
        @{ ServiceName = "vmicrdv";                    StartupType = "Manual" },
        @{ ServiceName = "vmicshutdown";               StartupType = "Manual" },
        @{ ServiceName = "vmictimesync";               StartupType = "Manual" },
        @{ ServiceName = "vmicvmsession";              StartupType = "Manual" },
        @{ ServiceName = "vmicvss";                    StartupType = "Manual" },
        @{ ServiceName = "vmvss";                      StartupType = "Manual" },
        @{ ServiceName = "wbengine";                   StartupType = "Manual" },
        @{ ServiceName = "wcncsvc";                    StartupType = "Manual" },
        @{ ServiceName = "webthreatdefsvc";            StartupType = "Manual" },
        @{ ServiceName = "webthreatdefusersvc_*";      StartupType = "Automatic" },
        @{ ServiceName = "wercplsupport";              StartupType = "Manual" },
        @{ ServiceName = "wisvc";                      StartupType = "Manual" },
        @{ ServiceName = "wlidsvc";                    StartupType = "Manual" },
        @{ ServiceName = "wlpasvc";                    StartupType = "Manual" },
        @{ ServiceName = "wmiApSrv";                   StartupType = "Manual" },
        @{ ServiceName = "workfolderssvc";             StartupType = "Manual" },
        @{ ServiceName = "wscsvc";                     StartupType = "AutomaticDelayedStart" },
        @{ ServiceName = "wuauserv";                   StartupType = "Manual" },
        @{ ServiceName = "wudfsvc";                    StartupType = "Manual" }
    )
    
    foreach ($service in $services) {
        $name = $service.ServiceName
        $type = $service.StartupType
        Log-Message "Setting service '$name' to startup type '$type'..."
        try {
            Set-Service -Name $name -StartupType $type -ErrorAction Stop
            Log-Message "Service '$name' set to '$type'."
        } catch {
            Log-Message "Failed to set service '$name': $($_.Exception.Message)"
        }
    }
    Log-Message "Service settings update completed."
}

function Disable-BackgroundApps {
    try {
        Log-Message "Disabling Background Apps..."
        # Define the registry path that holds the setting.
        $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications"
        # Set the GlobalUserDisabled value to 1 (disabled).
        Set-ItemProperty -Path $registryPath -Name "GlobalUserDisabled" -Value 1 -Type DWord -Force
        Log-Message "Background Apps have been disabled."
    } catch {
        Log-Message "Error disabling Background Apps: $_"
    }
}

# -------------------------
# Custom Preferences Functions

function Enable-DarkTheme {
    try {
        Log-Message "Enabling Dark Theme for Windows..."
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        
        # Set dark mode for apps and system
        Set-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -Value 0 -Force
        Set-ItemProperty -Path $regPath -Name "SystemUsesLightTheme" -Value 0 -Force
        
        Log-Message "Dark Theme has been enabled. You may need to sign out or restart Explorer for changes to take effect."
    } catch {
        Log-Message "Error enabling Dark Theme: $_"
    }
}

function Disable-BingSearch {
    Log-Message "Disabling Bing Search in Start Menu..."
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
        if (-not (Test-Path $regPath)) {
            New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion" -Name "Search" -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "BingSearchEnabled" -Value 0 -Type DWord -Force
        Log-Message "Bing Search disabled in Start Menu."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling Bing Search: $($_.Exception.Message)")
    }
}

function Disable-VerboseLogon {
    Log-Message "Disabling Verbose Messages During Logon..."
    try {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
        Set-ItemProperty -Path $regPath -Name "VerboseStatus" -Value 0 -Type DWord -Force
        Log-Message "Verbose Logon messages disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling Verbose Logon messages: $($_.Exception.Message)")
    }
}

function Disable-StartRecommendations {
    Log-Message "Disabling Recommendations in Start Menu..."
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
        Set-ItemProperty -Path $regPath -Name "SystemPaneSuggestionsEnabled" -Value 0 -Type DWord -Force
        Set-ItemProperty -Path $regPath -Name "SubscribedContent-338393Enabled" -Value 0 -Type DWord -Force
        Log-Message "Start Menu Recommendations disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling Start Menu Recommendations: $($_.Exception.Message)")
    }
}

function Disable-SnapWindow {
    Log-Message "Disabling Snap Window..."
    try {
        $regPath = "HKCU:\Control Panel\Desktop"
        Set-ItemProperty -Path $regPath -Name "WindowArrangementActive" -Value 0 -Type String -Force
        Log-Message "Snap Window disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling Snap Window: $($_.Exception.Message)")
    }
}

function Disable-SnapAssist {
    Log-Message "Disabling Snap Assist Suggestions..."
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $regPath -Name "DisableSnapAssistFlyout" -Value 1 -Type DWord -Force
        Log-Message "Snap Assist Suggestions disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling Snap Assist Suggestions: $($_.Exception.Message)")
    }
}

function Disable-StickyKeys {
    Log-Message "Disabling Sticky Keys..."
    try {
        $regPath = "HKCU:\Control Panel\Accessibility\StickyKeys"
        Set-ItemProperty -Path $regPath -Name "Flags" -Value "506" -Type String -Force
        Log-Message "Sticky Keys disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling Sticky Keys: $($_.Exception.Message)")
    }
}

function Disable-TaskbarSearchButton {
    Log-Message "Disabling Search Button in Taskbar..."
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
        Set-ItemProperty -Path $regPath -Name "SearchboxTaskbarMode" -Value 0 -Type DWord -Force
        Log-Message "Search Button in Taskbar disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling Search Button in Taskbar: $($_.Exception.Message)")
    }
}

function Disable-TaskViewButton {
    Log-Message "Disabling Task View Button in Taskbar..."
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $regPath -Name "ShowTaskViewButton" -Value 0 -Type DWord -Force
        Log-Message "Task View Button in Taskbar disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling Task View Button in Taskbar: $($_.Exception.Message)")
    }
}

function Disable-TaskbarWidgets {
    Log-Message "Disabling Widgets in Taskbar..."
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $regPath -Name "EnableWidgets" -Value 0 -Type DWord -Force
        Log-Message "Widgets in Taskbar disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling Widgets in Taskbar: $($_.Exception.Message)")
    }
}

function Activate-UltimatePerformanceProfile {
    Log-Message "Activating Ultimate Performance Profile..."
    try {
        $scheme = "e9a42b02-d5df-448d-aa00-03f14749eb61"
        powercfg -duplicatescheme $scheme | Out-Null
        powercfg -setactive $scheme
        Log-Message "Ultimate Performance Profile activated."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error activating Ultimate Performance Profile: $($_.Exception.Message)")
    }
}

# -------------------------
# Windows Updates Functions

function Ensure-PSWindowsUpdate {
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Log-Message "PSWindowsUpdate module not found. Installing..."
        try {
            Install-Module -Name PSWindowsUpdate -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Log-Message "PSWindowsUpdate installed successfully."
        } catch {
            throw "Failed to install PSWindowsUpdate module: $($_.Exception.Message)"
        }
    }
    Import-Module PSWindowsUpdate -ErrorAction Stop
}

function Check-WindowsSecurityUpdates {
    Log-Message "Checking for Windows Security Updates..."
    try {
        Ensure-PSWindowsUpdate
        $updates = Get-WUList -Category "Security Updates" -AcceptAll -IgnoreReboot
        if ($updates) {
            Log-Message "Security Updates available:"
            foreach ($update in $updates) {
                Log-Message " - $($update.Title)"
            }
        } else {
            Log-Message "No Windows Security Updates available."
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error checking Windows Security Updates: $($_.Exception.Message)")
    }
}

function Install-WindowsSecurityUpdates {
    Log-Message "Installing Windows Security Updates..."
    try {
        Ensure-PSWindowsUpdate
        $updates = Get-WUList -Category "Security Updates" -AcceptAll -IgnoreReboot
        if ($updates) {
            # Pipe the updates and force installation without prompting
            $updates | Install-WindowsUpdate -IgnoreReboot -AutoReboot -Confirm:$false -Verbose 2>&1 | ForEach-Object { Log-Message $_ }
            Log-Message "Windows Security Updates installation process completed."
        } else {
            Log-Message "No security updates to install."
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error installing Windows Security Updates: $($_.Exception.Message)")
    }
}

function Check-AllUpdates {
    Log-Message "Checking for all available Windows Updates..."
    try {
        Ensure-PSWindowsUpdate
        $updates = Get-WUList -AcceptAll -IgnoreReboot
        if ($updates) {
            Log-Message "All Updates available:"
            foreach ($update in $updates) {
                Log-Message " - $($update.Title)"
            }
        } else {
            Log-Message "No Windows Updates available."
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error checking Windows Updates: $($_.Exception.Message)")
    }
}

function Install-AllUpdates {
    Log-Message "Installing all available Windows Updates..."
    try {
        Ensure-PSWindowsUpdate
        $updates = Get-WUList -AcceptAll -IgnoreReboot
        if ($updates) {
            $updates | Install-WindowsUpdate -IgnoreReboot -AutoReboot -Confirm:$false -Verbose 2>&1 | ForEach-Object { Log-Message $_ }
            Log-Message "All Windows Updates installation process completed."
        } else {
            Log-Message "No Windows Updates to install."
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error installing Windows Updates: $($_.Exception.Message)")
    }
}

# -------------------------
# Uninstall Functions

function Get-InstalledSoftware {
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    $installedSoftware = foreach ($path in $registryPaths) {
        Get-ChildItem $path -ErrorAction SilentlyContinue | ForEach-Object {
            $app = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($app.DisplayName -and $app.UninstallString) {
                [PSCustomObject]@{
                    DisplayName     = $app.DisplayName
                    UninstallString = $app.UninstallString
                }
            }
        }
    }
    $installedSoftware = $installedSoftware | Sort-Object DisplayName -Unique
    return $installedSoftware
}

# -------------------------
# GUI Theme Colors (Assuming these are defined above)
$themeColorBackground = [System.Drawing.Color]::FromArgb(45, 45, 48)
$themeColorForeground = [System.Drawing.Color]::White
$themeColorButtonBack = [System.Drawing.Color]::FromArgb(63, 63, 70)
$themeColorButtonHover = [System.Drawing.Color]::FromArgb(80, 80, 85)
$themeColorControlBorder = [System.Drawing.Color]::FromArgb(80, 80, 85)
$themeColorLogBackground = [System.Drawing.Color]::FromArgb(30, 30, 30)
$themeColorGridHeader = $themeColorButtonBack
$themeColorGridSelectionBack = $themeColorControlBorder
# -------------------------

# Build the GUI

$form = New-Object System.Windows.Forms.Form
$form.Text = "Javi Tech Tool"
$form.Size = New-Object System.Drawing.Size(900,700)
$form.MinimumSize = New-Object System.Drawing.Size(750, 600) # Keep minimum size
$form.StartPosition = "CenterScreen"
$form.BackColor = $themeColorBackground
$form.ForeColor = $themeColorForeground
$form.Font = New-Object System.Drawing.Font("Segoe UI",10)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabControl.BackColor = $themeColorBackground
# Removed redundant Anchor for $tabControl

# --- ** NEW: Define Helper Functions Here ** ---
Function Style-Button {
    param(
        [System.Windows.Forms.Button]$Button
    )
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = $themeColorControlBorder
    $Button.BackColor = $themeColorButtonBack
    $Button.ForeColor = $themeColorForeground
    $Button.FlatAppearance.MouseOverBackColor = $themeColorButtonHover
}

Function Style-CheckBox {
     param(
        [System.Windows.Forms.CheckBox]$CheckBox
    )
    # Note: Standard checkboxes have limited background styling support in WinForms.
    # Setting BackColor might not always visually change the box itself, but ensures consistency if themes affect it.
    $CheckBox.BackColor = $themeColorBackground
    $CheckBox.ForeColor = $themeColorForeground
}
# --- End Helper Function Definitions ---


# Create Tabs (Apply BackColor/ForeColor as before)
$tabInstall = New-Object System.Windows.Forms.TabPage; $tabInstall.Text = "Install Software"; $tabInstall.BackColor = $themeColorBackground; $tabInstall.ForeColor = $themeColorForeground
$tabSpeedUp = New-Object System.Windows.Forms.TabPage; $tabSpeedUp.Text = "Speed Up Windows"; $tabSpeedUp.BackColor = $themeColorBackground; $tabSpeedUp.ForeColor = $themeColorForeground
$tabCustom = New-Object System.Windows.Forms.TabPage; $tabCustom.Text = "Custom Preferences"; $tabCustom.BackColor = $themeColorBackground; $tabCustom.ForeColor = $themeColorForeground
$tabUpdates = New-Object System.Windows.Forms.TabPage; $tabUpdates.Text = "Windows Updates"; $tabUpdates.BackColor = $themeColorBackground; $tabUpdates.ForeColor = $themeColorForeground
$tabUninstall = New-Object System.Windows.Forms.TabPage; $tabUninstall.Text = "Uninstall"; $tabUninstall.BackColor = $themeColorBackground; $tabUninstall.ForeColor = $themeColorForeground

# Add Tabs to TabControl FIRST
$tabControl.TabPages.Add($tabInstall)
$tabControl.TabPages.Add($tabSpeedUp)
$tabControl.TabPages.Add($tabCustom)
$tabControl.TabPages.Add($tabUpdates)
$tabControl.TabPages.Add($tabUninstall)


# -------------------------
# Tab: Install Software (Now using defined helpers)
$chkChrome = New-Object System.Windows.Forms.CheckBox; $chkChrome.Location = New-Object System.Drawing.Point(20,20); $chkChrome.Size = New-Object System.Drawing.Size(250,30); $chkChrome.Text = "Install Google Chrome"; Style-CheckBox $chkChrome; $tabInstall.Controls.Add($chkChrome)
$chkRingCentral = New-Object System.Windows.Forms.CheckBox; $chkRingCentral.Location = New-Object System.Drawing.Point(20,60); $chkRingCentral.Size = New-Object System.Drawing.Size(250,30); $chkRingCentral.Text = "Install RingCentral"; Style-CheckBox $chkRingCentral; $tabInstall.Controls.Add($chkRingCentral)
$chkTeams = New-Object System.Windows.Forms.CheckBox; $chkTeams.Location = New-Object System.Drawing.Point(20,100); $chkTeams.Size = New-Object System.Drawing.Size(250,30); $chkTeams.Text = "Install Microsoft Teams"; Style-CheckBox $chkTeams; $tabInstall.Controls.Add($chkTeams)
$btnInstallSelected = New-Object System.Windows.Forms.Button; $btnInstallSelected.Location = New-Object System.Drawing.Point(20,150); $btnInstallSelected.Size = New-Object System.Drawing.Size(250,40); $btnInstallSelected.Text = "Install Selected Software"; Style-Button $btnInstallSelected
$btnInstallSelected.Add_Click({ if ($chkChrome.Checked) { Install-GoogleChrome }; if ($chkRingCentral.Checked) { Install-RingCentral }; if ($chkTeams.Checked) { Install-MicrosoftTeams }; [System.Windows.Forms.MessageBox]::Show("Selected installations completed.") })
$tabInstall.Controls.Add($btnInstallSelected)

# -------------------------
# Tab: Speed Up Windows (Using Scrollable Panel and defined helpers)
$panelSpeedUp = New-Object System.Windows.Forms.Panel; $panelSpeedUp.Dock = [System.Windows.Forms.DockStyle]::Fill; $panelSpeedUp.AutoScroll = $true; $panelSpeedUp.BackColor = $themeColorBackground; $tabSpeedUp.Controls.Add($panelSpeedUp)
$btnSelectAllSpeed = New-Object System.Windows.Forms.Button; $btnSelectAllSpeed.Location = New-Object System.Drawing.Point(20,10); $btnSelectAllSpeed.Size = New-Object System.Drawing.Size(100,30); $btnSelectAllSpeed.Text = "Select All"; Style-Button $btnSelectAllSpeed
$btnSelectAllSpeed.Add_Click({ $chkClearTemp.Checked = $true; $chkDisableConsumer.Checked = $true; $chkDisableTelemetry.Checked = $true; $chkDisableActiveHistory.Checked = $true; $chkDisableGameDVR.Checked = $true; $chkDisableHibernation.Checked = $true; $chkDisableHomeGroup.Checked = $true; $chkDisableLocationTracking.Checked = $true; $chkDisableStorageSense.Checked = $true; $chkDisableWifiSense.Checked = $true; $chkEnableEndTaskRightClick.Checked = $true; $chkDisablePowershell7Telemetry.Checked = $true; $chkSetServices.Checked = $true; $chkDisableBackgroundApps.Checked = $true })
$panelSpeedUp.Controls.Add($btnSelectAllSpeed)
$controlsToStyle = @();
# Define Checkboxes... (same location/size/text as before)
$chkClearTemp = New-Object System.Windows.Forms.CheckBox; $chkClearTemp.Location = New-Object System.Drawing.Point(20,60); $chkClearTemp.Size = New-Object System.Drawing.Size(250,30); $chkClearTemp.Text = "Clear Temporary Files"; $controlsToStyle += $chkClearTemp
$chkDisableConsumer = New-Object System.Windows.Forms.CheckBox; $chkDisableConsumer.Location = New-Object System.Drawing.Point(20,100); $chkDisableConsumer.Size = New-Object System.Drawing.Size(250,30); $chkDisableConsumer.Text = "Disable Consumer Features"; $controlsToStyle += $chkDisableConsumer
$chkDisableTelemetry = New-Object System.Windows.Forms.CheckBox; $chkDisableTelemetry.Location = New-Object System.Drawing.Point(20,140); $chkDisableTelemetry.Size = New-Object System.Drawing.Size(250,30); $chkDisableTelemetry.Text = "Disable Telemetry"; $controlsToStyle += $chkDisableTelemetry
$chkDisableActiveHistory = New-Object System.Windows.Forms.CheckBox; $chkDisableActiveHistory.Location = New-Object System.Drawing.Point(20,180); $chkDisableActiveHistory.Size = New-Object System.Drawing.Size(250,30); $chkDisableActiveHistory.Text = "Disable Active History"; $controlsToStyle += $chkDisableActiveHistory
$chkDisableGameDVR = New-Object System.Windows.Forms.CheckBox; $chkDisableGameDVR.Location = New-Object System.Drawing.Point(20,220); $chkDisableGameDVR.Size = New-Object System.Drawing.Size(250,30); $chkDisableGameDVR.Text = "Disable Game DVR"; $controlsToStyle += $chkDisableGameDVR
$chkDisableHibernation = New-Object System.Windows.Forms.CheckBox; $chkDisableHibernation.Location = New-Object System.Drawing.Point(300,60); $chkDisableHibernation.Size = New-Object System.Drawing.Size(250,30); $chkDisableHibernation.Text = "Disable Hibernation"; $controlsToStyle += $chkDisableHibernation
$chkDisableHomeGroup = New-Object System.Windows.Forms.CheckBox; $chkDisableHomeGroup.Location = New-Object System.Drawing.Point(300,100); $chkDisableHomeGroup.Size = New-Object System.Drawing.Size(250,30); $chkDisableHomeGroup.Text = "Disable HomeGroup"; $controlsToStyle += $chkDisableHomeGroup
$chkDisableLocationTracking = New-Object System.Windows.Forms.CheckBox; $chkDisableLocationTracking.Location = New-Object System.Drawing.Point(300,140); $chkDisableLocationTracking.Size = New-Object System.Drawing.Size(250,30); $chkDisableLocationTracking.Text = "Disable Location Tracking"; $controlsToStyle += $chkDisableLocationTracking
$chkDisableStorageSense = New-Object System.Windows.Forms.CheckBox; $chkDisableStorageSense.Location = New-Object System.Drawing.Point(300,180); $chkDisableStorageSense.Size = New-Object System.Drawing.Size(250,30); $chkDisableStorageSense.Text = "Disable Storage Sense"; $controlsToStyle += $chkDisableStorageSense
$chkDisableWifiSense = New-Object System.Windows.Forms.CheckBox; $chkDisableWifiSense.Location = New-Object System.Drawing.Point(300,220); $chkDisableWifiSense.Size = New-Object System.Drawing.Size(250,30); $chkDisableWifiSense.Text = "Disable Wifi Sense"; $controlsToStyle += $chkDisableWifiSense
$chkEnableEndTaskRightClick = New-Object System.Windows.Forms.CheckBox; $chkEnableEndTaskRightClick.Location = New-Object System.Drawing.Point(580,60); $chkEnableEndTaskRightClick.Size = New-Object System.Drawing.Size(250,30); $chkEnableEndTaskRightClick.Text = "Enable End Task (Right Click)"; $controlsToStyle += $chkEnableEndTaskRightClick
$chkDisablePowershell7Telemetry = New-Object System.Windows.Forms.CheckBox; $chkDisablePowershell7Telemetry.Location = New-Object System.Drawing.Point(580,100); $chkDisablePowershell7Telemetry.Size = New-Object System.Drawing.Size(250,30); $chkDisablePowershell7Telemetry.Text = "Disable Powershell 7 Telemetry"; $controlsToStyle += $chkDisablePowershell7Telemetry
$chkSetServices = New-Object System.Windows.Forms.CheckBox; $chkSetServices.Location = New-Object System.Drawing.Point(580,140); $chkSetServices.Size = New-Object System.Drawing.Size(250,30); $chkSetServices.Text = "Set Services"; $controlsToStyle += $chkSetServices
$chkDisableBackgroundApps = New-Object System.Windows.Forms.CheckBox; $chkDisableBackgroundApps.Location = New-Object System.Drawing.Point(580,180); $chkDisableBackgroundApps.Size = New-Object System.Drawing.Size(250,30); $chkDisableBackgroundApps.Text = "Disable Background Apps"; $controlsToStyle += $chkDisableBackgroundApps
# Apply styles and add to panel
foreach ($control in $controlsToStyle) { Style-CheckBox $control; $panelSpeedUp.Controls.Add($control) }
$btnApplySpeedUp = New-Object System.Windows.Forms.Button; $btnApplySpeedUp.Location = New-Object System.Drawing.Point(20,280); $btnApplySpeedUp.Size = New-Object System.Drawing.Size(250,40); $btnApplySpeedUp.Text = "Apply Selected Tasks"; Style-Button $btnApplySpeedUp
$btnApplySpeedUp.Add_Click({ if ($chkClearTemp.Checked) { Clear-TempFiles }; if ($chkDisableConsumer.Checked) { Disable-ConsumerFeatures }; if ($chkDisableTelemetry.Checked) { Disable-Telemetry }; if ($chkDisableActiveHistory.Checked) { Disable-ActiveHistory }; if ($chkDisableGameDVR.Checked) { Disable-GameDVR }; if ($chkDisableHibernation.Checked) { Disable-Hibernation }; if ($chkDisableHomeGroup.Checked) { Disable-HomeGroup }; if ($chkDisableLocationTracking.Checked) { Disable-LocationTracking }; if ($chkDisableStorageSense.Checked) { Disable-StorageSense }; if ($chkDisableWifiSense.Checked) { Disable-WifiSense }; if ($chkEnableEndTaskRightClick.Checked) { Enable-EndTaskRightClick }; if ($chkDisablePowershell7Telemetry.Checked) { Disable-Powershell7Telemetry }; if ($chkSetServices.Checked) { Set-Services }; if ($chkDisableBackgroundApps.Checked) { Disable-BackgroundApps }; [System.Windows.Forms.MessageBox]::Show("Selected speed up tasks completed.") })
$panelSpeedUp.Controls.Add($btnApplySpeedUp)

# -------------------------
# Tab: Custom Preferences Controls (Using Scrollable Panel and defined helpers)
$panelCustom = New-Object System.Windows.Forms.Panel; $panelCustom.Dock = [System.Windows.Forms.DockStyle]::Fill; $panelCustom.AutoScroll = $true; $panelCustom.BackColor = $themeColorBackground; $tabCustom.Controls.Add($panelCustom)
$btnSelectAllCustom = New-Object System.Windows.Forms.Button; $btnSelectAllCustom.Location = New-Object System.Drawing.Point(20,10); $btnSelectAllCustom.Size = New-Object System.Drawing.Size(100,30); $btnSelectAllCustom.Text = "Select All"; Style-Button $btnSelectAllCustom
$btnSelectAllCustom.Add_Click({ $chkDarkTheme.Checked = $true; $chkDisableBing.Checked = $true; $chkDisableVerbose.Checked = $true; $chkDisableRecommendations.Checked = $true; $chkDisableSnapWindow.Checked = $true; $chkDisableSnapAssist.Checked = $true; $chkDisableStickyKeys.Checked = $true; $chkDisableTaskbarSearch.Checked = $true; $chkDisableTaskView.Checked = $true; $chkDisableWidgets.Checked = $true; $chkActivateUltimatePerf.Checked = $true })
$panelCustom.Controls.Add($btnSelectAllCustom)
$customControlsToStyle = @();
# Define Checkboxes... (same location/size/text as before)
$chkDarkTheme = New-Object System.Windows.Forms.CheckBox; $chkDarkTheme.Location = New-Object System.Drawing.Point(20,60); $chkDarkTheme.Size = New-Object System.Drawing.Size(350,30); $chkDarkTheme.Text = "Enable Dark Theme for Windows"; $customControlsToStyle += $chkDarkTheme
$chkDisableBing = New-Object System.Windows.Forms.CheckBox; $chkDisableBing.Location = New-Object System.Drawing.Point(20,100); $chkDisableBing.Size = New-Object System.Drawing.Size(350,30); $chkDisableBing.Text = "Disable Bing Search in Start Menu"; $customControlsToStyle += $chkDisableBing
$chkDisableVerbose = New-Object System.Windows.Forms.CheckBox; $chkDisableVerbose.Location = New-Object System.Drawing.Point(20,140); $chkDisableVerbose.Size = New-Object System.Drawing.Size(350,30); $chkDisableVerbose.Text = "Disable Verbose Messages During Logon"; $customControlsToStyle += $chkDisableVerbose
$chkDisableRecommendations = New-Object System.Windows.Forms.CheckBox; $chkDisableRecommendations.Location = New-Object System.Drawing.Point(20,180); $chkDisableRecommendations.Size = New-Object System.Drawing.Size(350,30); $chkDisableRecommendations.Text = "Disable Recommendations in Start Menu"; $customControlsToStyle += $chkDisableRecommendations
$chkDisableSnapWindow = New-Object System.Windows.Forms.CheckBox; $chkDisableSnapWindow.Location = New-Object System.Drawing.Point(20,220); $chkDisableSnapWindow.Size = New-Object System.Drawing.Size(350,30); $chkDisableSnapWindow.Text = "Disable Snap Window"; $customControlsToStyle += $chkDisableSnapWindow
$chkDisableSnapAssist = New-Object System.Windows.Forms.CheckBox; $chkDisableSnapAssist.Location = New-Object System.Drawing.Point(20,260); $chkDisableSnapAssist.Size = New-Object System.Drawing.Size(350,30); $chkDisableSnapAssist.Text = "Disable Snap Assist Suggestions"; $customControlsToStyle += $chkDisableSnapAssist
$chkDisableStickyKeys = New-Object System.Windows.Forms.CheckBox; $chkDisableStickyKeys.Location = New-Object System.Drawing.Point(20,300); $chkDisableStickyKeys.Size = New-Object System.Drawing.Size(350,30); $chkDisableStickyKeys.Text = "Disable Sticky Keys"; $customControlsToStyle += $chkDisableStickyKeys
$chkDisableTaskbarSearch = New-Object System.Windows.Forms.CheckBox; $chkDisableTaskbarSearch.Location = New-Object System.Drawing.Point(20,340); $chkDisableTaskbarSearch.Size = New-Object System.Drawing.Size(350,30); $chkDisableTaskbarSearch.Text = "Disable Search Button in Taskbar"; $customControlsToStyle += $chkDisableTaskbarSearch
$chkDisableTaskView = New-Object System.Windows.Forms.CheckBox; $chkDisableTaskView.Location = New-Object System.Drawing.Point(20,380); $chkDisableTaskView.Size = New-Object System.Drawing.Size(350,30); $chkDisableTaskView.Text = "Disable Task View Button in Taskbar"; $customControlsToStyle += $chkDisableTaskView
$chkDisableWidgets = New-Object System.Windows.Forms.CheckBox; $chkDisableWidgets.Location = New-Object System.Drawing.Point(20,420); $chkDisableWidgets.Size = New-Object System.Drawing.Size(350,30); $chkDisableWidgets.Text = "Disable Widgets in Taskbar"; $customControlsToStyle += $chkDisableWidgets
$chkActivateUltimatePerf = New-Object System.Windows.Forms.CheckBox; $chkActivateUltimatePerf.Location = New-Object System.Drawing.Point(20,460); $chkActivateUltimatePerf.Size = New-Object System.Drawing.Size(350,30); $chkActivateUltimatePerf.Text = "Add and Activate Ultimate Performance Profile"; $customControlsToStyle += $chkActivateUltimatePerf
# Apply styles and add to panel
foreach ($control in $customControlsToStyle) { Style-CheckBox $control; $panelCustom.Controls.Add($control) }
$btnApplyCustom = New-Object System.Windows.Forms.Button; $btnApplyCustom.Location = New-Object System.Drawing.Point(20,510); $btnApplyCustom.Size = New-Object System.Drawing.Size(250,40); $btnApplyCustom.Text = "Apply Custom Preferences"; Style-Button $btnApplyCustom
$btnApplyCustom.Add_Click({ if ($chkDarkTheme.Checked) { Enable-DarkTheme }; if ($chkDisableBing.Checked) { Disable-BingSearch }; if ($chkDisableVerbose.Checked) { Disable-VerboseLogon }; if ($chkDisableRecommendations.Checked) { Disable-StartRecommendations }; if ($chkDisableSnapWindow.Checked) { Disable-SnapWindow }; if ($chkDisableSnapAssist.Checked) { Disable-SnapAssist }; if ($chkDisableStickyKeys.Checked) { Disable-StickyKeys }; if ($chkDisableTaskbarSearch.Checked) { Disable-TaskbarSearchButton }; if ($chkDisableTaskView.Checked) { Disable-TaskViewButton }; if ($chkDisableWidgets.Checked) { Disable-TaskbarWidgets }; if ($chkActivateUltimatePerf.Checked) { Activate-UltimatePerformanceProfile }; [System.Windows.Forms.MessageBox]::Show("Custom Preferences applied.") })
$panelCustom.Controls.Add($btnApplyCustom)

# -------------------------
# Tab: Windows Updates Controls (Using defined helpers)
$btnCheckSecUpdates = New-Object System.Windows.Forms.Button; $btnCheckSecUpdates.Location = New-Object System.Drawing.Point(20,20); $btnCheckSecUpdates.Size = New-Object System.Drawing.Size(220,40); $btnCheckSecUpdates.Text = "Check Security Updates"; Style-Button $btnCheckSecUpdates; $btnCheckSecUpdates.Add_Click({ Check-WindowsSecurityUpdates }); $tabUpdates.Controls.Add($btnCheckSecUpdates)
$btnInstallSecUpdates = New-Object System.Windows.Forms.Button; $btnInstallSecUpdates.Location = New-Object System.Drawing.Point(260,20); $btnInstallSecUpdates.Size = New-Object System.Drawing.Size(220,40); $btnInstallSecUpdates.Text = "Install Security Updates"; Style-Button $btnInstallSecUpdates; $btnInstallSecUpdates.Add_Click({ Install-WindowsSecurityUpdates }); $tabUpdates.Controls.Add($btnInstallSecUpdates)
$btnCheckAllUpdates = New-Object System.Windows.Forms.Button; $btnCheckAllUpdates.Location = New-Object System.Drawing.Point(20,80); $btnCheckAllUpdates.Size = New-Object System.Drawing.Size(220,40); $btnCheckAllUpdates.Text = "Check All Updates"; Style-Button $btnCheckAllUpdates; $btnCheckAllUpdates.Add_Click({ Check-AllUpdates }); $tabUpdates.Controls.Add($btnCheckAllUpdates)
$btnInstallAllUpdates = New-Object System.Windows.Forms.Button; $btnInstallAllUpdates.Location = New-Object System.Drawing.Point(260,80); $btnInstallAllUpdates.Size = New-Object System.Drawing.Size(220,40); $btnInstallAllUpdates.Text = "Install All Updates"; Style-Button $btnInstallAllUpdates; $btnInstallAllUpdates.Add_Click({ Install-AllUpdates }); $tabUpdates.Controls.Add($btnInstallAllUpdates)

# -------------------------
# Tab: Uninstall Controls (Using defined helpers, DGV anchored)
$btnLoadSoftware = New-Object System.Windows.Forms.Button; $btnLoadSoftware.Location = New-Object System.Drawing.Point(20,20); $btnLoadSoftware.Size = New-Object System.Drawing.Size(200,40); $btnLoadSoftware.Text = "Load Installed Software"; Style-Button $btnLoadSoftware; $tabUninstall.Controls.Add($btnLoadSoftware)
$btnUninstallSelected = New-Object System.Windows.Forms.Button; $btnUninstallSelected.Location = New-Object System.Drawing.Point(240,20); $btnUninstallSelected.Size = New-Object System.Drawing.Size(200,40); $btnUninstallSelected.Text = "Uninstall Selected"; Style-Button $btnUninstallSelected; $tabUninstall.Controls.Add($btnUninstallSelected)
$dgvUninstall = New-Object System.Windows.Forms.DataGridView
$dgvUninstall.Location = New-Object System.Drawing.Point(20,80)
$dgvUninstall.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$dgvUninstall.SelectionMode = "FullRowSelect"; $dgvUninstall.MultiSelect = $false; $dgvUninstall.AutoGenerateColumns = $false; $dgvUninstall.ReadOnly = $false
# Styling (as before)...
$dgvUninstall.BackgroundColor = $themeColorBackground; $dgvUninstall.ForeColor = $themeColorForeground; $dgvUninstall.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle; $dgvUninstall.GridColor = $themeColorControlBorder
$dgvUninstall.DefaultCellStyle.BackColor = $themeColorBackground; $dgvUninstall.DefaultCellStyle.ForeColor = $themeColorForeground; $dgvUninstall.DefaultCellStyle.SelectionBackColor = $themeColorGridSelectionBack; $dgvUninstall.DefaultCellStyle.SelectionForeColor = $themeColorForeground
$dgvUninstall.ColumnHeadersDefaultCellStyle.BackColor = $themeColorGridHeader; $dgvUninstall.ColumnHeadersDefaultCellStyle.ForeColor = $themeColorForeground; $dgvUninstall.ColumnHeadersDefaultCellStyle.SelectionBackColor = $themeColorGridHeader; $dgvUninstall.ColumnHeadersDefaultCellStyle.SelectionForeColor = $themeColorForeground; $dgvUninstall.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single; $dgvUninstall.ColumnHeadersHeight = 30; $dgvUninstall.EnableHeadersVisualStyles = $false
$dgvUninstall.RowHeadersDefaultCellStyle.BackColor = $themeColorBackground; $dgvUninstall.RowHeadersDefaultCellStyle.SelectionBackColor = $themeColorGridSelectionBack; $dgvUninstall.RowHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single; $dgvUninstall.RowHeadersVisible = $false
$tabUninstall.Controls.Add($dgvUninstall)
# Click Handlers (as before)...
$btnLoadSoftware.Add_Click({ try { $installedSoftware = Get-InstalledSoftware; $dgvUninstall.DataSource = $null; $dgvUninstall.Columns.Clear(); if ($installedSoftware -and $installedSoftware.Count -gt 0) { $dataTable = New-Object System.Data.DataTable "InstalledSoftware"; [void]$dataTable.Columns.Add("DisplayName",[string]); [void]$dataTable.Columns.Add("UninstallString",[string]); foreach ($app in $installedSoftware) { if ($app -and $app.DisplayName -and $app.UninstallString) { try { $row = $dataTable.NewRow(); $row["DisplayName"] = [string]$app.DisplayName; $row["UninstallString"] = [string]$app.UninstallString; $dataTable.Rows.Add($row) } catch { Log-Message "Skipping an item due to error: $($_.Exception.Message)" } } else { Log-Message "Skipping item missing DisplayName or UninstallString." } }; $cbCol = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn; $cbCol.Name = "Select"; $cbCol.HeaderText = "Select"; $cbCol.Width = 50; $cbCol.ReadOnly = $false; [void]$dgvUninstall.Columns.Add($cbCol); $dispCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $dispCol.Name = "DisplayName"; $dispCol.HeaderText = "Application Name"; $dispCol.DataPropertyName = "DisplayName"; $dispCol.ReadOnly = $true; $dispCol.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill; [void]$dgvUninstall.Columns.Add($dispCol); $uninstCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $uninstCol.Name = "UninstallString"; $uninstCol.DataPropertyName = "UninstallString"; $uninstCol.Visible = $false; $uninstCol.ReadOnly = $true; [void]$dgvUninstall.Columns.Add($uninstCol); $dgvUninstall.DataSource = $dataTable; $dgvUninstall.AllowUserToAddRows = $false; $dgvUninstall.AllowUserToDeleteRows = $false; $dgvUninstall.Refresh(); Log-Message "Installed software list loaded. ($($dataTable.Rows.Count) items)" } else { Log-Message "No installed software found." } } catch { [System.Windows.Forms.MessageBox]::Show("Error loading software list: $($_.Exception.Message)") } })
$btnUninstallSelected.Add_Click({ $rowsToUninstall = @(); foreach ($row in $dgvUninstall.Rows) { if ($row.Cells["Select"] -ne $null -and $row.Cells["Select"].Value -eq $true) { $rowsToUninstall += $row } }; if ($rowsToUninstall.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No applications selected for uninstallation."); return }; $confirmResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to uninstall $($rowsToUninstall.Count) selected application(s)?", "Confirm Uninstall", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning); if ($confirmResult -eq [System.Windows.Forms.DialogResult]::No) { Log-Message "Uninstallation cancelled by user."; return }; foreach ($row in $rowsToUninstall) { $displayName = $row.Cells["DisplayName"].Value; $uninstallString = $row.Cells["UninstallString"].Value; if (-not $uninstallString) { Log-Message "Could not find uninstall string for '$displayName'. Skipping."; continue }; Log-Message "Attempting to uninstall: $displayName"; try { $silentArgs = $uninstallString -replace '(?i)msiexec\.exe\s*/i', 'msiexec.exe /x'; $silentArgs = $silentArgs -replace '/passive', '/quiet' -replace '/qb', '/quiet'; if ($silentArgs -notmatch '/quiet' -and $silentArgs -notmatch '/qn' -and $silentArgs -notmatch '/s') { $argumentsToTry = $uninstallString; Log-Message "No obvious silent switch found, running: cmd /c $argumentsToTry"; Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$argumentsToTry`"" -Wait -WindowStyle Hidden -ErrorAction Stop } else { Log-Message "Attempting silent uninstall: cmd /c $silentArgs"; Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$silentArgs`"" -Wait -WindowStyle Hidden -ErrorAction Stop }; Log-Message "'$displayName' uninstallation process initiated (Check system for completion/prompts)." } catch { Log-Message "Failed to initiate uninstall for '$displayName': $($_.Exception.Message). Attempting non-silent..."; try { Log-Message "Running original command: cmd /c $uninstallString"; Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$uninstallString`"" -Wait; Log-Message "'$displayName' uninstallation command executed (non-silent)." } catch { Log-Message "Also failed non-silent uninstall for '$displayName': $($_.Exception.Message)" } } }; Log-Message "Uninstall process finished for selected items. Refreshing list..."; $btnLoadSoftware.PerformClick() })

# -------------------------
# Add Log Box and THEN TabControl to Form Controls (Order Matters for Docking)
$global:logBox = New-Object System.Windows.Forms.RichTextBox
$global:logBox.Multiline = $true; $global:logBox.ScrollBars = "Vertical"; $global:logBox.Dock = [System.Windows.Forms.DockStyle]::Bottom; $global:logBox.Height = 80; $global:logBox.ReadOnly = $true; $global:logBox.BackColor = $themeColorLogBackground; $global:logBox.ForeColor = $themeColorForeground; $global:logBox.Font = New-Object System.Drawing.Font("Consolas",10)
$form.Controls.Add($global:logBox) # Add Bottom Docked First
$form.Controls.Add($tabControl)    # Add Fill Docked Last

# Add Form Load event
$form.Add_Load({ Log-Message "JaviTechTool Initialized. Ready."})

# --------------------------------------------------
# Show the form
$form.ShowDialog() | Out-Null