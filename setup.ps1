function Set-DevSetupStage {
    Write-Host "Advancing to stage $StageValue"
    param(
        [Parameter(Mandatory = $true)][string]$StageValue
    )
    [System.Environment]::SetEnvironmentVariable("DEV_SETUP_STAGE", $StageValue, "User")
}

function Install-VisualStudio2015 {
    $primarySource = "\\newforma.local\data\departments\Development\Installation Kits\Microsoft\Visual Studio 2015 Pro\vs_professional.exe"
    $backupSource = "\\winnas01\Aperus\Installation Kits\Visual Studio 2015 Pro\vs_professional.exe"
    $localDir = "$env:TEMP\VS2015Install"
    $localInstaller = Join-Path $localDir "vs_professional.exe"
    Write-Host "Starting Visual Studio 2015 installation process..."
    Write-Host "Ensuring local directory exists: $localDir"
    Write-Host "Attempting to copy VS2015 installer from primary location: $primarySource"
    $success = $true
    try {
        if (-not (Test-Path $localDir)) {
            New-Item -ItemType Directory -Path $localDir -Force | Out-Null
        }
        $copied = $false
        if (Test-Path $primarySource) {
            try {
                Copy-Item -Path $primarySource -Destination $localInstaller -Force -ErrorAction Stop 2>$null
                $copied = $true
            }
            catch {
                Write-Warning "Failed to copy from primary location: $($_.Exception.Message)"
            }
        }
        if (-not $copied -and (Test-Path $backupSource)) {
            Write-Host "Attempting to copy VS2015 installer from backup location: $backupSource"
            try {
                Copy-Item -Path $backupSource -Destination $localInstaller -Force -ErrorAction Stop 2>$null
                $copied = $true
            }
            catch {
                Write-Warning "Failed to copy from backup location: $($_.Exception.Message)"
            }
        }
        if ($copied) {
            Write-Host "VS2015 installer successfully copied to $localInstaller"
        }
        if (-not $copied) {
            Write-Warning "Neither primary nor backup VS2015 installer could be copied."
            $success = $false
        }
    }
    catch {
        Write-Warning "Unexpected error during installer copy: $($_.Exception.Message)"
        $success = $false
    }

    $installArgs = '/quiet /norestart /log "%TEMP%\VS2015Install.log" ' +
    '/features ' +
    'OfficeTools,VC,VC_MFC,VC_MFC_XP,VC_Common'
    Write-Host "Launching VS2015 installer with arguments: $installArgs"
    try {
        Start-Process -FilePath $localInstaller -ArgumentList $installArgs -Wait -NoNewWindow -ErrorAction SilentlyContinue
        Write-Host "VS2015 installer process completed."
    }
    catch {
        Write-Warning "Failed to start VS2015 installer."
        $success = $false
    }
    return $success
}

function Install-VisualStudio2022 {
    $vsInstallerUrl = "https://aka.ms/vs/17/release/vs_professional.exe"
    $localDir = "$env:TEMP\VS2022Install"
    $localInstaller = Join-Path $localDir "vs_professional.exe"
    Write-Host "Starting Visual Studio 2022 installation process..."
    Write-Host "Ensuring local directory exists: $localDir"
    Write-Host "Downloading VS2022 installer from $vsInstallerUrl to $localInstaller"
    $vsInstallerUrl = "https://aka.ms/vs/17/release/vs_professional.exe"
    $localDir = "$env:TEMP\VS2022Install"
    $localInstaller = Join-Path $localDir "vs_professional.exe"
    $success = $true
    try {
        if (-not (Test-Path $localDir)) {
            New-Item -ItemType Directory -Path $localDir -Force | Out-Null
        }
        Invoke-WebRequest -Uri $vsInstallerUrl -OutFile $localInstaller -UseBasicParsing
    }
    catch {
        Write-Warning "Failed to download Visual Studio 2022 installer."
        $success = $false
    }

    $installArgs = @(
        "--quiet"
        "--wait"
        "--norestart"
        "--nocache"
        "--installPath `"C:\Program Files\Microsoft Visual Studio\2022\Professional`""
        "--add Microsoft.VisualStudio.Workload.ManagedDesktop"
        "--add Microsoft.VisualStudio.Workload.NativeDesktop"
        "--add Microsoft.VisualStudio.Workload.Office"
        # .NET Desktop Development
        "--add Microsoft.Net.Component.4.8.SDK"
        "--add Microsoft.Net.Component.4.8.TargetingPack"
        "--add Microsoft.VisualStudio.Component.EntityFramework"
        "--add Microsoft.VisualStudio.Component.DiagnosticTools"
        "--add Microsoft.VisualStudio.Component.IntelliCode"
        "--add Microsoft.VisualStudio.Component.Debugger.JustInTime"
        "--add Microsoft.VisualStudio.LiveShare"
        "--add Microsoft.VisualStudio.Component.ML.NetModelBuilder"
        "--add Microsoft.VisualStudio.Component.GitHub.Copilot"
        "--add Microsoft.VisualStudio.Component.Blend"
        "--add Microsoft.Net.Component.4.6.2.TargetingPack"
        "--add Microsoft.Net.Component.4.7.1.TargetingPack"
        "--add Microsoft.VisualStudio.Component.Wcf.Tooling"
        "--add Microsoft.VisualStudio.Component.SQL.LocalDB"
        "--add Microsoft.VisualStudio.Component.JavaScript.Diagnostics"
        "--add Microsoft.Net.Component.4.8.1.TargetingPack"
        # Desktop Development with C++
        "--add Microsoft.VisualStudio.Component.VC.Tools.x86.x64"
        "--add Microsoft.VisualStudio.Component.VC.ATL"
        "--add Microsoft.VisualStudio.Component.VC.ATL.ARM"
        "--add Microsoft.VisualStudio.Component.VC.CMake.Project"
        "--add Microsoft.VisualStudio.Component.VC.Llvm.Clang"
        "--add Microsoft.VisualStudio.Component.VC.MFC"
        "--add Microsoft.VisualStudio.Component.VC.TestAdapterForBoostTest"
        "--add Microsoft.VisualStudio.Component.VC.TestAdapterForGoogleTest"
        "--add Microsoft.VisualStudio.Component.VC.ASAN"
        "--add Microsoft.VisualStudio.Component.Windows11SDK.22621"
        "--add Microsoft.VisualStudio.Component.VC.vcpkg"
        "--add Microsoft.VisualStudio.Component.GitHub.Copilot"
        "--add Microsoft.VisualStudio.Component.VC.CLI.Support"
        "--add Microsoft.VisualStudio.Component.Windows10SDK.19041"
        "--add Microsoft.VisualStudio.Component.VC.Tools.ARM64"
        "--add Microsoft.VisualStudio.Component.VC.Tools.x86.x64"
        "--add Microsoft.VisualStudio.Component.VC.v142.x86.x64"
        "--add Microsoft.VisualStudio.Component.VC.v141.x86.x64"
        "--add Microsoft.VisualStudio.Component.VC.v140.x86.x64"
        "--add Microsoft.VisualStudio.Component.Security.StaticAnalysis"
        "--add Microsoft.VisualStudio.Component.VC.Profiler"
        "--add Microsoft.VisualStudio.Component.IntelliCode"
        "--add Microsoft.VisualStudio.LiveShare"
        "--add Microsoft.VisualStudio.Component.JavaScript.Diagnostics"
        # Office/SharePoint Development
        "--add Microsoft.VisualStudio.Component.Office.Tools"
        "--add Microsoft.VisualStudio.Component.WebDeploy"
        "--add Microsoft.VisualStudio.Component.IntelliCode"
        "--add Microsoft.Net.Component.4.8.TargetingPack"
        "--add Microsoft.VisualStudio.Component.GitHub.Copilot"
        "--add Microsoft.Net.Component.4.6.2.TargetingPack"
        "--add Microsoft.Net.Component.4.7.1.TargetingPack"
        "--add Microsoft.Net.Component.4.8.1.TargetingPack"
        # Individual components
        "--add Microsoft.VisualStudio.Component.VC.ATL.141"
        "--add Microsoft.VisualStudio.Component.VC.MFC.141"
        "--add Microsoft.VisualStudio.Component.VC.ATL.142"
        "--add Microsoft.VisualStudio.Component.VC.MFC.142"
        "--add Microsoft.VisualStudio.Component.VC.CLI.Support.142"
    )

    Write-Host "Launching VS2022 installer with arguments: $($installArgs -join ' ')"

    try {
        Start-Process -FilePath $localInstaller -ArgumentList $installArgs -Wait -NoNewWindow -ErrorAction Stop
        Write-Host "VS2022 installer process completed."
    }
    catch {
        Write-Warning "Failed to start Visual Studio 2022 installer."
        $success = $false
    }
    return $success
}

function main {
    Write-Host "Checking for administrator privileges..."
    Write-Host "Ensuring script is in user profile: $profilePath"
    if (-not (Test-Path $profilePath)) {
        New-Item -ItemType File -Path $profilePath -Force | Out-Null
    }
    $profileContent = Get-Content $profilePath -Raw
    if ($profileContent -notmatch [regex]::Escape($profileLine)) {
        Add-Content -Path $profilePath -Value $profileLine
        Write-Host "Added script invocation to profile."
    }
    else {
        Write-Host "Script invocation already present in profile."
    }
    if (-not $stage) {
        Write-Host "DEV_SETUP_STAGE not set. Defaulting to 1."
    }
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "This script must be run as an administrator." -ForegroundColor Red
        exit 1
    }

    $profileLine = "powershell -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $profilePath = $PROFILE
    if (-not (Test-Path $profilePath)) {
        New-Item -ItemType File -Path $profilePath -Force | Out-Null
    }
    $profileContent = Get-Content $profilePath -Raw
    if ($profileContent -notmatch [regex]::Escape($profileLine)) {
        Add-Content -Path $profilePath -Value $profileLine
    }

    $stage = [int]($env:DEV_SETUP_STAGE)
    if (-not $stage) {
        $stage = 1
        Set-DevSetupStage "$stage"
    }

    switch ($stage) {
        1 {
            Write-Host "Stage 1: Installing Visual Studio 2015 and 2022..."
            $vs2015Ok = Install-VisualStudio2015
            if (-not $vs2015Ok) {
                Write-Host "ERROR: Visual Studio 2015 installation failed. Please install Visual Studio 2015 manually and restart Powershell." -ForegroundColor Red
                exit 1
            }
            $vs2022Ok = Install-VisualStudio2022
            if (-not $vs2022Ok) {
                Write-Host "ERROR: Visual Studio 2022 installation failed. Please install Visual Studio 2022 manually and restart Powershell." -ForegroundColor Red
                exit 1
            }
            Set-DevSetupStage "2"
            Write-Host "Stage 1 complete. Restarting shell for next stage..."
            Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-File", "`"$PSCommandPath`""
            exit
        }
        2 {
            Write-Host "Stage 2: Setup already completed or next steps go here."
            # Add additional setup stages as needed
        }
        default {
            Write-Host "Unknown stage: $stage"
        }
    }
}

main