function Set-DevSetupStage {
    param(
        [Parameter(Mandatory=$true)][string]$StageValue
    )
    [System.Environment]::SetEnvironmentVariable("DEV_SETUP_STAGE", $StageValue, "User")
}

function Install-VisualStudio2015 {
    $sourcePath = "\\newforma.local\data\departments\Development\Installation Kits\Microsoft\Visual Studio 2015 Pro\vs_professional.exe"
    $localDir = "$env:TEMP\VS2015Install"
    $localInstaller = Join-Path $localDir "vs_professional.exe"
    $success = $true
    try {
        if (-not (Test-Path $localDir)) {
            New-Item -ItemType Directory -Path $localDir -Force | Out-Null
        }
        Copy-Item -Path $sourcePath -Destination $localInstaller -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Warning "Failed to copy installer."
        $success = $false
    }

    $installArgs = '/quiet /norestart /log "%TEMP%\VS2015Install.log" ' +
        '/features ' +
        'OfficeTools,VC,VC_MFC,VC_MFC_XP,VC_Common'

    try {
        Start-Process -FilePath $localInstaller -ArgumentList $installArgs -Wait -NoNewWindow -ErrorAction SilentlyContinue
    } catch {
        Write-Warning "Failed to start VS2015 installer."
        $success = $false
    }
    return $success
}

function Install-VisualStudio2022 {
    Write-Host "Downloading Visual Studio 2022 Professional installer..."
    $vsInstallerUrl = "https://aka.ms/vs/17/release/vs_professional.exe"
    $localDir = "$env:TEMP\VS2022Install"
    $localInstaller = Join-Path $localDir "vs_professional.exe"
    $success = $true
    try {
        if (-not (Test-Path $localDir)) {
            New-Item -ItemType Directory -Path $localDir -Force | Out-Null
        }
        Invoke-WebRequest -Uri $vsInstallerUrl -OutFile $localInstaller -UseBasicParsing
    } catch {
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

    try {
        Start-Process -FilePath $localInstaller -ArgumentList $installArgs -Wait -NoNewWindow -ErrorAction Stop
    } catch {
        Write-Warning "Failed to start Visual Studio 2022 installer."
        $success = $false
    }
    return $success
}

function main {
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