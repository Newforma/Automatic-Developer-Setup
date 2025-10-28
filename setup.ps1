function Set-DevSetupStage {
    param(
        [Parameter(Mandatory = $true)][string]$StageValue
    )
    Write-Host "Advancing to stage $StageValue"
    $env:DEV_SETUP_STAGE = $StageValue
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
        if (Test-Path $primarySource 2>$null) {
            try {
                Copy-Item -Path $primarySource -Destination $localInstaller -Force -ErrorAction Stop 2>$null
                $copied = $true
            }
            catch {
                Write-Warning "Failed to copy from primary location: $($_.Exception.Message)"
            }
        }
        if (-not $copied -and (Test-Path $backupSource 2>$null)) {
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

    Write-Host "Installing VS2022. This might take a while..."

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

function Install-Git {
    Write-Host "Starting Git installation..."
    $success = $true
    $gitUrl = "https://github.com/git-for-windows/git/releases/latest/download/Git-2.43.0-64-bit.exe"
    $localDir = "$env:TEMP\GitInstall"
    $localInstaller = Join-Path $localDir "Git-Setup.exe"
    try {
        if (-not (Test-Path $localDir)) {
            New-Item -ItemType Directory -Path $localDir -Force | Out-Null
        }
        Write-Host "Downloading Git installer from $gitUrl to $localInstaller"
        Invoke-WebRequest -Uri $gitUrl -OutFile $localInstaller -UseBasicParsing
    }
    catch {
        Write-Host "Failed to download Git installer: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }

    $installArgs = "/VERYSILENT /NORESTART /NOCANCEL /SP- /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS"
    try {
        Write-Host "Running Git installer silently..."
        Start-Process -FilePath $localInstaller -ArgumentList $installArgs -Wait -NoNewWindow -ErrorAction Stop
        Write-Host "Git installed successfully."
    }
    catch {
        Write-Host "Failed to install Git: $($_.Exception.Message)" -ForegroundColor Red
        $success = $false
    }
    return $success
}

function Install-GitHubDesktop {
    Write-Host "Starting GitHub Desktop installation..."
    $success = $true
    $ghdUrl = "https://central.github.com/deployments/desktop/desktop/latest/win32"
    $localDir = "$env:TEMP\GitHubDesktopInstall"
    $localInstaller = Join-Path $localDir "GitHubDesktopSetup.exe"
    try {
        if (-not (Test-Path $localDir)) {
            New-Item -ItemType Directory -Path $localDir -Force | Out-Null
        }
        Write-Host "Downloading GitHub Desktop installer from $ghdUrl to $localInstaller"
        Invoke-WebRequest -Uri $ghdUrl -OutFile $localInstaller -UseBasicParsing
    }
    catch {
        Write-Host "Failed to download GitHub Desktop installer: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }

    $installArgs = "/silent"
    try {
        Write-Host "Running GitHub Desktop installer silently..."
        Start-Process -FilePath $localInstaller -ArgumentList $installArgs -Wait -NoNewWindow -ErrorAction Stop
        Write-Host "GitHub Desktop installed successfully."
    }
    catch {
        Write-Host "Failed to install GitHub Desktop: $($_.Exception.Message)" -ForegroundColor Red
        $success = $false
    }
    return $success
}

function Get-Repositories {
    $defaultRepoPath = "C:/repos"
    $GitRepoPath = Read-Host "Where should git repositories be cloned? (Default: $defaultRepoPath)"
    if ([string]::IsNullOrWhiteSpace($GitRepoPath)) {
        $GitRepoPath = $defaultRepoPath
    }

    Write-Log "Creating git repository directory: $GitRepoPath"
    New-Item -ItemType Directory -Force -Path $GitRepoPath | Out-Null
    
    $repos = @(
        "enterprise-suite",
        "enterprise-technical-documentation",
        "enterprise-tools",
        "enterprise-api"
    )
    
    foreach ($repo in $repos) {
        $repoPath = Join-Path $GitRepoPath $repo
        if (Test-Path $repoPath) {
            Write-Log "Repository $repo already exists, skipping clone"
        }
        else {
            Write-Log "Cloning $repo..."
            Set-Location $GitRepoPath
            git clone "https://github.com/Newforma/$repo.git"
        }
    }
}

function main {
    # Ensure running as administrator
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "This script must be run as an administrator." -ForegroundColor Red
        exit 1
    }

    $profilePath = $PROFILE
    $profileLine = "powershell -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Write-Host "Ensuring script is in user profile: $profilePath"
    if (-not (Test-Path $profilePath)) {
        New-Item -ItemType File -Path $profilePath -Force | Out-Null
    }
    $profileContent = Get-Content $profilePath -Raw
    if ($profileContent -notmatch [regex]::Escape($profileLine)) {
        Add-Content -Path $profilePath -Value $profileLine
    }

    if (-not (Test-Path $profilePath)) {
        New-Item -ItemType File -Path $profilePath -Force | Out-Null
    }
    $profileContent = Get-Content $profilePath -Raw
    if ($profileContent -notmatch [regex]::Escape($profileLine)) {
        Add-Content -Path $profilePath -Value $profileLine
    }

    $stage = [int]($env:DEV_SETUP_STAGE)
    if (-not $stage) {
        Write-Host "Setup stage not set. Defaulting to 1."
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
            Stop-Process -Id $PID
        }
        2 {
            Write-Host "Stage 2: Installing NUnit console runners..."
            try {
                Write-Host "Running: dotnet add package NUnit.Runners --version 3.9.0"
                dotnet add package NUnit.Runners --version 3.9.0
                Write-Host "NUnit console runners installed."
                $gitOk = Install-Git
                if (-not $gitOk) {
                    Write-Host "ERROR: Git installation failed. Please install Git manually and restart Powershell." -ForegroundColor Red
                    exit 1
                }
                $ghdOk = Install-GitHubDesktop
                if (-not $ghdOk) {
                    Write-Host "ERROR: GitHub Desktop installation failed. Please install GitHub Desktop manually and restart Powershell." -ForegroundColor Red
                    exit 1
                }
                Set-DevSetupStage "3"
                Write-Host "Stage 2 complete. Restarting shell for next stage..."
                Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-File", "`"$PSCommandPath`""
                Stop-Process -Id $PID
            }
            catch {
                Write-Host "Failed to install NUnit console runners. Please install manually and restart Powershell." -ForegroundColor Red
                exit 1
            }
        }
        3 {
            Write-Host "Stage 3: Cloning repositories..."
            Get-Repositories
            Set-DevSetupStage "4"
            Write-Host "Stage 3 complete. Restarting shell for next stage..."
            Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-File", "`"$PSCommandPath`""
            Stop-Process -Id $PID
        }
        default {
            Write-Host "Unknown stage: $stage"
        }
    }
}

main