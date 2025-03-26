#--------------------------------------------------
# Ensure the script is running as Administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $scriptPath = $MyInvocation.MyCommand.Definition
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    Start-Process powershell.exe -Verb RunAs -ArgumentList $arguments
    exit
}

#--------------------------------------------------
# Load Windows Forms and Drawing assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#--------------------------------------------------
# Global Log Box (will be assigned later)
$global:logBox = $null

#--------------------------------------------------
# Define a helper function to log messages
function Log-Message {
    param ([string]$message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $line = "[$timestamp] $message"
    Write-Host $line
    if ($global:logBox -ne $null) {
        $global:logBox.AppendText("$line`r`n")
    }
}

#--------------------------------------------------
# Define Functions
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

# --- Software Install Functions ---
function Install-GoogleChrome {
    Log-Message "Installing Google Chrome..."
    $url = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi"
    $tempFile = "$env:TEMP\googlechrome_installer.msi"
    
    # Ensure BITS is running
    $bitsService = Get-Service -Name bits -ErrorAction SilentlyContinue
    if ($bitsService -and $bitsService.Status -ne "Running") {
        Start-Service -Name bits
        Log-Message "BITS service started."
    }
    
    try {
        Log-Message "Downloading Google Chrome installer..."
        Download-File -url $url -destination $tempFile
        Log-Message "Download complete. Executing installer..."
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$tempFile`" /quiet /norestart" -NoNewWindow -Wait
        Log-Message "Google Chrome installation completed."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error installing Google Chrome: $($_.Exception.Message)")
    } finally {
        if (Test-Path $tempFile) {
            Remove-Item $tempFile -Force
        }
    }
}

function Install-RingCentral {
    Log-Message "Installing RingCentral..."
    $url = "https://app.ringcentral.com/download/RingCentral.exe?V=20138600535791900"
    $tempFile = "$env:TEMP\RingCentralInstaller.exe"
    
    # Ensure BITS is running
    $bitsService = Get-Service -Name bits -ErrorAction SilentlyContinue
    if ($bitsService -and $bitsService.Status -ne "Running") {
        Start-Service -Name bits
        Log-Message "BITS service started."
    }
    
    try {
        Log-Message "Downloading RingCentral installer..."
        Download-File -url $url -destination $tempFile
        Log-Message "Download complete. Executing installer..."
        # Execute the installer with a silent argument (adjust if needed)
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
    
    # Ensure BITS is running
    $bitsService = Get-Service -Name bits -ErrorAction SilentlyContinue
    if ($bitsService -and $bitsService.Status -ne "Running") {
        Start-Service -Name bits
        Log-Message "BITS service started."
    }
    
    try {
        Log-Message "Downloading Microsoft Teams installer..."
        Download-File -url $url -destination $tempFile
        Log-Message "Download complete. Executing installer..."
        # Use the appropriate silent argument; adjust if needed (/S is common)
        Start-Process -FilePath $tempFile -ArgumentList "/S" -NoNewWindow -Wait
        Log-Message "Microsoft Teams installation completed."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error installing Microsoft Teams: $($_.Exception.Message)")
    } finally {
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
    }
}

# --- Speed Up Windows Functions ---
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
                    # Remove the read-only attribute if it is set.
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

# --- Set Services Function (Embedded Table) ---
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
    Log-Message "Disabling Background Apps..."
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        if (-not (Test-Path $regPath)) {
            New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "Advanced" -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "GlobalUserDisabled" -Value 1 -Type DWord -Force -ErrorAction Stop
        Log-Message "Background apps disabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disabling background apps: $($_.Exception.Message)")
    }
}

# --- Custom Preferences Functions ---
function Enable-DarkTheme {
    Log-Message "Enabling Dark Theme..."
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        if (-not (Test-Path $regPath)) {
            New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes" -Name "Personalize" -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -Value 0 -Type DWord -Force
        Set-ItemProperty -Path $regPath -Name "SystemUsesLightTheme" -Value 0 -Type DWord -Force
        Log-Message "Dark Theme enabled."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error enabling Dark Theme: $($_.Exception.Message)")
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

# --- Windows Updates Functions ---
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

# --- Windows Updates Functions for Security Updates ---
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
            # Pipe the updates list to Install-WindowsUpdate
            $updates | Install-WindowsUpdate -IgnoreReboot -AutoReboot -Verbose 2>&1 | ForEach-Object { Log-Message $_ }
            Log-Message "Windows Security Updates installation process completed."
        } else {
            Log-Message "No security updates to install."
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error installing Windows Security Updates: $($_.Exception.Message)")
    }
}

# --- New Functions for ALL Updates ---
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
            $updates | Install-WindowsUpdate -IgnoreReboot -AutoReboot -Verbose 2>&1 | ForEach-Object { Log-Message $_ }
            Log-Message "All Windows Updates installation process completed."
        } else {
            Log-Message "No Windows Updates to install."
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error installing Windows Updates: $($_.Exception.Message)")
    }
}

# --- Build the GUI with a modern dark theme ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Javi Tech Tool"
$form.Size = New-Object System.Drawing.Size(900,700)
$form.StartPosition = "CenterScreen"
# Set dark theme for the main form
$form.BackColor = [System.Drawing.Color]::FromArgb(45,45,48)
$form.ForeColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Segoe UI",10)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"
# Set dark theme for the TabControl
$tabControl.BackColor = [System.Drawing.Color]::FromArgb(45,45,48)
$tabControl.ForeColor = [System.Drawing.Color]::White

# Create three tabs: Install Software, Speed Up Windows, Custom Preferences
$tabInstall = New-Object System.Windows.Forms.TabPage
$tabInstall.Text = "Install Software"
$tabInstall.BackColor = [System.Drawing.Color]::FromArgb(45,45,48)
$tabInstall.ForeColor = [System.Drawing.Color]::White

$tabSpeedUp = New-Object System.Windows.Forms.TabPage
$tabSpeedUp.Text = "Speed Up Windows"
$tabSpeedUp.BackColor = [System.Drawing.Color]::FromArgb(45,45,48)
$tabSpeedUp.ForeColor = [System.Drawing.Color]::White

$tabCustom = New-Object System.Windows.Forms.TabPage
$tabCustom.Text = "Custom Preferences"
$tabCustom.BackColor = [System.Drawing.Color]::FromArgb(45,45,48)
$tabCustom.ForeColor = [System.Drawing.Color]::White

$tabControl.TabPages.Add($tabInstall)
$tabControl.TabPages.Add($tabSpeedUp)
$tabControl.TabPages.Add($tabCustom)
$form.Controls.Add($tabControl)

# For buttons, we use a flat style with a slightly lighter background
$buttonBackColor = [System.Drawing.Color]::FromArgb(63,63,70)
$buttonForeColor = [System.Drawing.Color]::White

#-------------------------
# Tab: Install Software
$chkChrome = New-Object System.Windows.Forms.CheckBox
$chkChrome.Location = New-Object System.Drawing.Point(20,20)
$chkChrome.Size = New-Object System.Drawing.Size(250,30)
$chkChrome.Text = "Install Google Chrome"
$chkChrome.BackColor = $form.BackColor
$chkChrome.ForeColor = $form.ForeColor
$tabInstall.Controls.Add($chkChrome)

$chkRingCentral = New-Object System.Windows.Forms.CheckBox
$chkRingCentral.Location = New-Object System.Drawing.Point(20,60)
$chkRingCentral.Size = New-Object System.Drawing.Size(250,30)
$chkRingCentral.Text = "Install RingCentral"
$chkRingCentral.BackColor = $form.BackColor
$chkRingCentral.ForeColor = $form.ForeColor
$tabInstall.Controls.Add($chkRingCentral)

$chkTeams = New-Object System.Windows.Forms.CheckBox
$chkTeams.Location = New-Object System.Drawing.Point(20,100)
$chkTeams.Size = New-Object System.Drawing.Size(250,30)
$chkTeams.Text = "Install Microsoft Teams"
$chkTeams.BackColor = $form.BackColor
$chkTeams.ForeColor = $form.ForeColor
$tabInstall.Controls.Add($chkTeams)

$btnInstallSelected = New-Object System.Windows.Forms.Button
$btnInstallSelected.Location = New-Object System.Drawing.Point(20,150)
$btnInstallSelected.Size = New-Object System.Drawing.Size(250,40)
$btnInstallSelected.Text = "Install Selected Software"
$btnInstallSelected.FlatStyle = "Flat"
$btnInstallSelected.BackColor = $buttonBackColor
$btnInstallSelected.ForeColor = $buttonForeColor
$btnInstallSelected.Add_Click({
    if ($chkChrome.Checked) { Install-GoogleChrome }
    if ($chkRingCentral.Checked) { Install-RingCentral }
    if ($chkTeams.Checked) { Install-MicrosoftTeams }
    [System.Windows.Forms.MessageBox]::Show("Selected installations completed.")
})
$tabInstall.Controls.Add($btnInstallSelected)

#-------------------------
# Tab: Speed Up Windows
# Add a "Select All" button for this tab
$btnSelectAllSpeed = New-Object System.Windows.Forms.Button
$btnSelectAllSpeed.Location = New-Object System.Drawing.Point(20,10)
$btnSelectAllSpeed.Size = New-Object System.Drawing.Size(100,30)
$btnSelectAllSpeed.Text = "Select All"
$btnSelectAllSpeed.FlatStyle = "Flat"
$btnSelectAllSpeed.BackColor = $buttonBackColor
$btnSelectAllSpeed.ForeColor = $buttonForeColor
$btnSelectAllSpeed.Add_Click({
    $chkClearTemp.Checked = $true
    $chkDisableConsumer.Checked = $true
    $chkDisableTelemetry.Checked = $true
    $chkDisableActiveHistory.Checked = $true
    $chkDisableGameDVR.Checked = $true
    $chkDisableHibernation.Checked = $true
    $chkDisableHomeGroup.Checked = $true
    $chkDisableLocationTracking.Checked = $true
    $chkDisableStorageSense.Checked = $true
    $chkDisableWifiSense.Checked = $true
    $chkEnableEndTaskRightClick.Checked = $true
    $chkDisablePowershell7Telemetry.Checked = $true
    $chkSetServices.Checked = $true
    $chkDisableBackgroundApps.Checked = $true
})
$tabSpeedUp.Controls.Add($btnSelectAllSpeed)

# Arrange Speed Up Windows checkboxes in three columns
# Column 1 (x = 20)
$chkClearTemp = New-Object System.Windows.Forms.CheckBox
$chkClearTemp.Location = New-Object System.Drawing.Point(20,60)
$chkClearTemp.Size = New-Object System.Drawing.Size(250,30)
$chkClearTemp.Text = "Clear Temporary Files"
$chkClearTemp.BackColor = $form.BackColor
$chkClearTemp.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkClearTemp)

$chkDisableConsumer = New-Object System.Windows.Forms.CheckBox
$chkDisableConsumer.Location = New-Object System.Drawing.Point(20,100)
$chkDisableConsumer.Size = New-Object System.Drawing.Size(250,30)
$chkDisableConsumer.Text = "Disable Consumer Features"
$chkDisableConsumer.BackColor = $form.BackColor
$chkDisableConsumer.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkDisableConsumer)

$chkDisableTelemetry = New-Object System.Windows.Forms.CheckBox
$chkDisableTelemetry.Location = New-Object System.Drawing.Point(20,140)
$chkDisableTelemetry.Size = New-Object System.Drawing.Size(250,30)
$chkDisableTelemetry.Text = "Disable Telemetry"
$chkDisableTelemetry.BackColor = $form.BackColor
$chkDisableTelemetry.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkDisableTelemetry)

$chkDisableActiveHistory = New-Object System.Windows.Forms.CheckBox
$chkDisableActiveHistory.Location = New-Object System.Drawing.Point(20,180)
$chkDisableActiveHistory.Size = New-Object System.Drawing.Size(250,30)
$chkDisableActiveHistory.Text = "Disable Active History"
$chkDisableActiveHistory.BackColor = $form.BackColor
$chkDisableActiveHistory.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkDisableActiveHistory)

$chkDisableGameDVR = New-Object System.Windows.Forms.CheckBox
$chkDisableGameDVR.Location = New-Object System.Drawing.Point(20,220)
$chkDisableGameDVR.Size = New-Object System.Drawing.Size(250,30)
$chkDisableGameDVR.Text = "Disable Game DVR"
$chkDisableGameDVR.BackColor = $form.BackColor
$chkDisableGameDVR.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkDisableGameDVR)

# Column 2 (x = 300)
$chkDisableHibernation = New-Object System.Windows.Forms.CheckBox
$chkDisableHibernation.Location = New-Object System.Drawing.Point(300,60)
$chkDisableHibernation.Size = New-Object System.Drawing.Size(250,30)
$chkDisableHibernation.Text = "Disable Hibernation"
$chkDisableHibernation.BackColor = $form.BackColor
$chkDisableHibernation.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkDisableHibernation)

$chkDisableHomeGroup = New-Object System.Windows.Forms.CheckBox
$chkDisableHomeGroup.Location = New-Object System.Drawing.Point(300,100)
$chkDisableHomeGroup.Size = New-Object System.Drawing.Size(250,30)
$chkDisableHomeGroup.Text = "Disable HomeGroup"
$chkDisableHomeGroup.BackColor = $form.BackColor
$chkDisableHomeGroup.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkDisableHomeGroup)

$chkDisableLocationTracking = New-Object System.Windows.Forms.CheckBox
$chkDisableLocationTracking.Location = New-Object System.Drawing.Point(300,140)
$chkDisableLocationTracking.Size = New-Object System.Drawing.Size(250,30)
$chkDisableLocationTracking.Text = "Disable Location Tracking"
$chkDisableLocationTracking.BackColor = $form.BackColor
$chkDisableLocationTracking.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkDisableLocationTracking)

$chkDisableStorageSense = New-Object System.Windows.Forms.CheckBox
$chkDisableStorageSense.Location = New-Object System.Drawing.Point(300,180)
$chkDisableStorageSense.Size = New-Object System.Drawing.Size(250,30)
$chkDisableStorageSense.Text = "Disable Storage Sense"
$chkDisableStorageSense.BackColor = $form.BackColor
$chkDisableStorageSense.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkDisableStorageSense)

$chkDisableWifiSense = New-Object System.Windows.Forms.CheckBox
$chkDisableWifiSense.Location = New-Object System.Drawing.Point(300,220)
$chkDisableWifiSense.Size = New-Object System.Drawing.Size(250,30)
$chkDisableWifiSense.Text = "Disable Wifi Sense"
$chkDisableWifiSense.BackColor = $form.BackColor
$chkDisableWifiSense.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkDisableWifiSense)

# Column 3 (x = 580)
$chkEnableEndTaskRightClick = New-Object System.Windows.Forms.CheckBox
$chkEnableEndTaskRightClick.Location = New-Object System.Drawing.Point(580,60)
$chkEnableEndTaskRightClick.Size = New-Object System.Drawing.Size(250,30)
$chkEnableEndTaskRightClick.Text = "Enable End Task (Right Click)"
$chkEnableEndTaskRightClick.BackColor = $form.BackColor
$chkEnableEndTaskRightClick.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkEnableEndTaskRightClick)

$chkDisablePowershell7Telemetry = New-Object System.Windows.Forms.CheckBox
$chkDisablePowershell7Telemetry.Location = New-Object System.Drawing.Point(580,100)
$chkDisablePowershell7Telemetry.Size = New-Object System.Drawing.Size(250,30)
$chkDisablePowershell7Telemetry.Text = "Disable Powershell 7 Telemetry"
$chkDisablePowershell7Telemetry.BackColor = $form.BackColor
$chkDisablePowershell7Telemetry.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkDisablePowershell7Telemetry)

$chkSetServices = New-Object System.Windows.Forms.CheckBox
$chkSetServices.Location = New-Object System.Drawing.Point(580,140)
$chkSetServices.Size = New-Object System.Drawing.Size(250,30)
$chkSetServices.Text = "Set Services"
$chkSetServices.BackColor = $form.BackColor
$chkSetServices.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkSetServices)

$chkDisableBackgroundApps = New-Object System.Windows.Forms.CheckBox
$chkDisableBackgroundApps.Location = New-Object System.Drawing.Point(580,180)
$chkDisableBackgroundApps.Size = New-Object System.Drawing.Size(250,30)
$chkDisableBackgroundApps.Text = "Disable Background Apps"
$chkDisableBackgroundApps.BackColor = $form.BackColor
$chkDisableBackgroundApps.ForeColor = $form.ForeColor
$tabSpeedUp.Controls.Add($chkDisableBackgroundApps)

$btnApplySpeedUp = New-Object System.Windows.Forms.Button
$btnApplySpeedUp.Location = New-Object System.Drawing.Point(20,280)
$btnApplySpeedUp.Size = New-Object System.Drawing.Size(250,40)
$btnApplySpeedUp.Text = "Apply Selected Tasks"
$btnApplySpeedUp.FlatStyle = "Flat"
$btnApplySpeedUp.BackColor = $buttonBackColor
$btnApplySpeedUp.ForeColor = $buttonForeColor
$btnApplySpeedUp.Add_Click({
    if ($chkClearTemp.Checked) { Clear-TempFiles }
    if ($chkDisableConsumer.Checked) { Disable-ConsumerFeatures }
    if ($chkDisableTelemetry.Checked) { Disable-Telemetry }
    if ($chkDisableActiveHistory.Checked) { Disable-ActiveHistory }
    if ($chkDisableGameDVR.Checked) { Disable-GameDVR }
    if ($chkDisableHibernation.Checked) { Disable-Hibernation }
    if ($chkDisableHomeGroup.Checked) { Disable-HomeGroup }
    if ($chkDisableLocationTracking.Checked) { Disable-LocationTracking }
    if ($chkDisableStorageSense.Checked) { Disable-StorageSense }
    if ($chkDisableWifiSense.Checked) { Disable-WifiSense }
    if ($chkEnableEndTaskRightClick.Checked) { Enable-EndTaskRightClick }
    if ($chkDisablePowershell7Telemetry.Checked) { Disable-Powershell7Telemetry }
    if ($chkSetServices.Checked) { Set-Services }
    if ($chkDisableBackgroundApps.Checked) { Disable-BackgroundApps }
    [System.Windows.Forms.MessageBox]::Show("Selected speed up tasks completed.")
})
$tabSpeedUp.Controls.Add($btnApplySpeedUp)

#-------------------------
# Tab: Custom Preferences
$btnSelectAllCustom = New-Object System.Windows.Forms.Button
$btnSelectAllCustom.Location = New-Object System.Drawing.Point(20,10)
$btnSelectAllCustom.Size = New-Object System.Drawing.Size(100,30)
$btnSelectAllCustom.Text = "Select All"
$btnSelectAllCustom.FlatStyle = "Flat"
$btnSelectAllCustom.BackColor = $buttonBackColor
$btnSelectAllCustom.ForeColor = $buttonForeColor
$btnSelectAllCustom.Add_Click({
    $chkDarkTheme.Checked = $true
    $chkDisableBing.Checked = $true
    $chkDisableVerbose.Checked = $true
    $chkDisableRecommendations.Checked = $true
    $chkDisableSnapWindow.Checked = $true
    $chkDisableSnapAssist.Checked = $true
    $chkDisableStickyKeys.Checked = $true
    $chkDisableTaskbarSearch.Checked = $true
    $chkDisableTaskView.Checked = $true
    $chkDisableWidgets.Checked = $true
    $chkActivateUltimatePerf.Checked = $true
})
$tabCustom.Controls.Add($btnSelectAllCustom)

$chkDarkTheme = New-Object System.Windows.Forms.CheckBox
$chkDarkTheme.Location = New-Object System.Drawing.Point(20,60)
$chkDarkTheme.Size = New-Object System.Drawing.Size(350,30)
$chkDarkTheme.Text = "Enable Dark Theme for Windows"
$chkDarkTheme.BackColor = $form.BackColor
$chkDarkTheme.ForeColor = $form.ForeColor
$tabCustom.Controls.Add($chkDarkTheme)

$chkDisableBing = New-Object System.Windows.Forms.CheckBox
$chkDisableBing.Location = New-Object System.Drawing.Point(20,100)
$chkDisableBing.Size = New-Object System.Drawing.Size(350,30)
$chkDisableBing.Text = "Disable Bing Search in Start Menu"
$chkDisableBing.BackColor = $form.BackColor
$chkDisableBing.ForeColor = $form.ForeColor
$tabCustom.Controls.Add($chkDisableBing)

$chkDisableVerbose = New-Object System.Windows.Forms.CheckBox
$chkDisableVerbose.Location = New-Object System.Drawing.Point(20,140)
$chkDisableVerbose.Size = New-Object System.Drawing.Size(350,30)
$chkDisableVerbose.Text = "Disable Verbose Messages During Logon"
$chkDisableVerbose.BackColor = $form.BackColor
$chkDisableVerbose.ForeColor = $form.ForeColor
$tabCustom.Controls.Add($chkDisableVerbose)

$chkDisableRecommendations = New-Object System.Windows.Forms.CheckBox
$chkDisableRecommendations.Location = New-Object System.Drawing.Point(20,180)
$chkDisableRecommendations.Size = New-Object System.Drawing.Size(350,30)
$chkDisableRecommendations.Text = "Disable Recommendations in Start Menu"
$chkDisableRecommendations.BackColor = $form.BackColor
$chkDisableRecommendations.ForeColor = $form.ForeColor
$tabCustom.Controls.Add($chkDisableRecommendations)

$chkDisableSnapWindow = New-Object System.Windows.Forms.CheckBox
$chkDisableSnapWindow.Location = New-Object System.Drawing.Point(20,220)
$chkDisableSnapWindow.Size = New-Object System.Drawing.Size(350,30)
$chkDisableSnapWindow.Text = "Disable Snap Window"
$chkDisableSnapWindow.BackColor = $form.BackColor
$chkDisableSnapWindow.ForeColor = $form.ForeColor
$tabCustom.Controls.Add($chkDisableSnapWindow)

$chkDisableSnapAssist = New-Object System.Windows.Forms.CheckBox
$chkDisableSnapAssist.Location = New-Object System.Drawing.Point(20,260)
$chkDisableSnapAssist.Size = New-Object System.Drawing.Size(350,30)
$chkDisableSnapAssist.Text = "Disable Snap Assist Suggestions"
$chkDisableSnapAssist.BackColor = $form.BackColor
$chkDisableSnapAssist.ForeColor = $form.ForeColor
$tabCustom.Controls.Add($chkDisableSnapAssist)

$chkDisableStickyKeys = New-Object System.Windows.Forms.CheckBox
$chkDisableStickyKeys.Location = New-Object System.Drawing.Point(20,300)
$chkDisableStickyKeys.Size = New-Object System.Drawing.Size(350,30)
$chkDisableStickyKeys.Text = "Disable Sticky Keys"
$chkDisableStickyKeys.BackColor = $form.BackColor
$chkDisableStickyKeys.ForeColor = $form.ForeColor
$tabCustom.Controls.Add($chkDisableStickyKeys)

$chkDisableTaskbarSearch = New-Object System.Windows.Forms.CheckBox
$chkDisableTaskbarSearch.Location = New-Object System.Drawing.Point(20,340)
$chkDisableTaskbarSearch.Size = New-Object System.Drawing.Size(350,30)
$chkDisableTaskbarSearch.Text = "Disable Search Button in Taskbar"
$chkDisableTaskbarSearch.BackColor = $form.BackColor
$chkDisableTaskbarSearch.ForeColor = $form.ForeColor
$tabCustom.Controls.Add($chkDisableTaskbarSearch)

$chkDisableTaskView = New-Object System.Windows.Forms.CheckBox
$chkDisableTaskView.Location = New-Object System.Drawing.Point(20,380)
$chkDisableTaskView.Size = New-Object System.Drawing.Size(350,30)
$chkDisableTaskView.Text = "Disable Task View Button in Taskbar"
$chkDisableTaskView.BackColor = $form.BackColor
$chkDisableTaskView.ForeColor = $form.ForeColor
$tabCustom.Controls.Add($chkDisableTaskView)

$chkDisableWidgets = New-Object System.Windows.Forms.CheckBox
$chkDisableWidgets.Location = New-Object System.Drawing.Point(20,420)
$chkDisableWidgets.Size = New-Object System.Drawing.Size(350,30)
$chkDisableWidgets.Text = "Disable Widgets in Taskbar"
$chkDisableWidgets.BackColor = $form.BackColor
$chkDisableWidgets.ForeColor = $form.ForeColor
$tabCustom.Controls.Add($chkDisableWidgets)

$chkActivateUltimatePerf = New-Object System.Windows.Forms.CheckBox
$chkActivateUltimatePerf.Location = New-Object System.Drawing.Point(20,460)
$chkActivateUltimatePerf.Size = New-Object System.Drawing.Size(350,30)
$chkActivateUltimatePerf.Text = "Add and Activate Ultimate Performance Profile"
$chkActivateUltimatePerf.BackColor = $form.BackColor
$chkActivateUltimatePerf.ForeColor = $form.ForeColor
$tabCustom.Controls.Add($chkActivateUltimatePerf)

$btnApplyCustom = New-Object System.Windows.Forms.Button
$btnApplyCustom.Location = New-Object System.Drawing.Point(20,510)
$btnApplyCustom.Size = New-Object System.Drawing.Size(250,40)
$btnApplyCustom.Text = "Apply Custom Preferences"
$btnApplyCustom.FlatStyle = "Flat"
$btnApplyCustom.BackColor = $buttonBackColor
$btnApplyCustom.ForeColor = $buttonForeColor
$btnApplyCustom.Add_Click({
    if ($chkDarkTheme.Checked) { Enable-DarkTheme }
    if ($chkDisableBing.Checked) { Disable-BingSearch }
    if ($chkDisableVerbose.Checked) { Disable-VerboseLogon }
    if ($chkDisableRecommendations.Checked) { Disable-StartRecommendations }
    if ($chkDisableSnapWindow.Checked) { Disable-SnapWindow }
    if ($chkDisableSnapAssist.Checked) { Disable-SnapAssist }
    if ($chkDisableStickyKeys.Checked) { Disable-StickyKeys }
    if ($chkDisableTaskbarSearch.Checked) { Disable-TaskbarSearchButton }
    if ($chkDisableTaskView.Checked) { Disable-TaskViewButton }
    if ($chkDisableWidgets.Checked) { Disable-TaskbarWidgets }
    if ($chkActivateUltimatePerf.Checked) { Activate-UltimatePerformanceProfile }
    [System.Windows.Forms.MessageBox]::Show("Custom Preferences applied.")
})
$tabCustom.Controls.Add($btnApplyCustom)

#-------------------------
$tabWindowsUpdates = New-Object System.Windows.Forms.TabPage
$tabWindowsUpdates.Text = "Windows Updates"
$tabWindowsUpdates.BackColor = [System.Drawing.Color]::FromArgb(45,45,48)
$tabWindowsUpdates.ForeColor = [System.Drawing.Color]::White

# Define button colors (if not defined already)
$buttonBackColor = [System.Drawing.Color]::FromArgb(63,63,70)
$buttonForeColor = [System.Drawing.Color]::White

# Create buttons for Security Updates (check and install)
$btnCheckSecUpdates = New-Object System.Windows.Forms.Button
$btnCheckSecUpdates.Location = New-Object System.Drawing.Point(20,20)
$btnCheckSecUpdates.Size = New-Object System.Drawing.Size(220,40)
$btnCheckSecUpdates.Text = "Check Security Updates"
$btnCheckSecUpdates.FlatStyle = "Flat"
$btnCheckSecUpdates.BackColor = $buttonBackColor
$btnCheckSecUpdates.ForeColor = $buttonForeColor
$btnCheckSecUpdates.Add_Click({ Check-WindowsSecurityUpdates })
$tabWindowsUpdates.Controls.Add($btnCheckSecUpdates)

$btnInstallSecUpdates = New-Object System.Windows.Forms.Button
$btnInstallSecUpdates.Location = New-Object System.Drawing.Point(260,20)
$btnInstallSecUpdates.Size = New-Object System.Drawing.Size(220,40)
$btnInstallSecUpdates.Text = "Install Security Updates"
$btnInstallSecUpdates.FlatStyle = "Flat"
$btnInstallSecUpdates.BackColor = $buttonBackColor
$btnInstallSecUpdates.ForeColor = $buttonForeColor
$btnInstallSecUpdates.Add_Click({ Install-WindowsSecurityUpdates })
$tabWindowsUpdates.Controls.Add($btnInstallSecUpdates)

# Create buttons for ALL Updates (check and install)
$btnCheckAllUpdates = New-Object System.Windows.Forms.Button
$btnCheckAllUpdates.Location = New-Object System.Drawing.Point(20,80)
$btnCheckAllUpdates.Size = New-Object System.Drawing.Size(220,40)
$btnCheckAllUpdates.Text = "Check All Updates"
$btnCheckAllUpdates.FlatStyle = "Flat"
$btnCheckAllUpdates.BackColor = $buttonBackColor
$btnCheckAllUpdates.ForeColor = $buttonForeColor
$btnCheckAllUpdates.Add_Click({ Check-AllUpdates })
$tabWindowsUpdates.Controls.Add($btnCheckAllUpdates)

$btnInstallAllUpdates = New-Object System.Windows.Forms.Button
$btnInstallAllUpdates.Location = New-Object System.Drawing.Point(260,80)
$btnInstallAllUpdates.Size = New-Object System.Drawing.Size(220,40)
$btnInstallAllUpdates.Text = "Install All Updates"
$btnInstallAllUpdates.FlatStyle = "Flat"
$btnInstallAllUpdates.BackColor = $buttonBackColor
$btnInstallAllUpdates.ForeColor = $buttonForeColor
$btnInstallAllUpdates.Add_Click({ Install-AllUpdates })
$tabWindowsUpdates.Controls.Add($btnInstallAllUpdates)

# Add the Windows Updates tab to the main TabControl (e.g., after your other tabs)
$tabControl.TabPages.Add($tabWindowsUpdates)

#--------------------------------------------------
[void]$form.ShowDialog()
