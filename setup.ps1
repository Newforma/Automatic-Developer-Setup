function Install-NUnitConsoleRunners {
    Write-Host "Installing NUnit Console Runners v3.9.0 from GitHub..."
    $nunitUrl = "https://github.com/nunit/nunit-console/releases/download/v3.9/NUnit.Console-3.9.0.zip"
    $programFiles = ${env:ProgramFiles}
    $nunitDir = Join-Path $programFiles "NUnit.Console.3.9.0"
    $zipPath = Join-Path $env:TEMP "NUnit.Console-3.9.0.zip"
    try {
        $nunitExe = Join-Path $nunitDir "nunit3-console.exe"
        if (Test-Path $nunitExe) {
            Write-Host "NUnit Console Runners already installed at $nunitDir. Skipping download."
        }
        else {
            if (-not (Test-Path $nunitDir)) {
                New-Item -ItemType Directory -Path $nunitDir -Force | Out-Null
            }
            Write-Host "Downloading NUnit Console Runners from $nunitUrl to $zipPath"
            Invoke-WebRequest -Uri $nunitUrl -OutFile $zipPath -UseBasicParsing
            Write-Host "Extracting NUnit Console Runners to $nunitDir"
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $nunitDir)
            Write-Host "NUnit Console Runners installed to $nunitDir"
        }
        # Add to system PATH if not already present
        $currentPath = [Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::Machine)
        if ($currentPath -notlike "*${nunitDir}*") {
            Write-Host "Adding $nunitDir to system PATH..."
            [Environment]::SetEnvironmentVariable("Path", "$currentPath;${nunitDir}", [System.EnvironmentVariableTarget]::Machine)
            Write-Host "NUnit Console directory added to system PATH. You may need to restart your shell."
        }
        else {
            Write-Host "NUnit Console directory already in system PATH."
        }
        return $true
    }
    catch {
        Write-Host "Failed to install NUnit Console Runners: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}
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

    $features = @(
        # Workloads
        @{ Name = "Managed Desktop"; Arg = "--add Microsoft.VisualStudio.Workload.ManagedDesktop" }
        @{ Name = "Native Desktop"; Arg = "--add Microsoft.VisualStudio.Workload.NativeDesktop" }
        @{ Name = "Office/SharePoint"; Arg = "--add Microsoft.VisualStudio.Workload.Office" }
        # .NET Desktop Development
        @{ Name = ".NET 4.8 SDK"; Arg = "--add Microsoft.Net.Component.4.8.SDK" }
        @{ Name = ".NET 4.8 Targeting Pack"; Arg = "--add Microsoft.Net.Component.4.8.TargetingPack" }
        @{ Name = "Entity Framework"; Arg = "--add Microsoft.VisualStudio.Component.EntityFramework" }
        @{ Name = "Diagnostic Tools"; Arg = "--add Microsoft.VisualStudio.Component.DiagnosticTools" }
        @{ Name = "IntelliCode"; Arg = "--add Microsoft.VisualStudio.Component.IntelliCode" }
        @{ Name = "Just-In-Time Debugger"; Arg = "--add Microsoft.VisualStudio.Component.Debugger.JustInTime" }
        @{ Name = "Live Share"; Arg = "--add Microsoft.VisualStudio.LiveShare" }
        @{ Name = "ML .NET Model Builder"; Arg = "--add Microsoft.VisualStudio.Component.ML.NetModelBuilder" }
        @{ Name = "GitHub Copilot"; Arg = "--add Microsoft.VisualStudio.Component.GitHub.Copilot" }
        @{ Name = "Blend"; Arg = "--add Microsoft.VisualStudio.Component.Blend" }
        @{ Name = ".NET 4.6.2 Targeting Pack"; Arg = "--add Microsoft.Net.Component.4.6.2.TargetingPack" }
        @{ Name = ".NET 4.7.1 Targeting Pack"; Arg = "--add Microsoft.Net.Component.4.7.1.TargetingPack" }
        @{ Name = "WCF Tooling"; Arg = "--add Microsoft.VisualStudio.Component.Wcf.Tooling" }
        @{ Name = "SQL LocalDB"; Arg = "--add Microsoft.VisualStudio.Component.SQL.LocalDB" }
        @{ Name = "JavaScript Diagnostics"; Arg = "--add Microsoft.VisualStudio.Component.JavaScript.Diagnostics" }
        @{ Name = ".NET 4.8.1 Targeting Pack"; Arg = "--add Microsoft.Net.Component.4.8.1.TargetingPack" }
        # Desktop Development with C++
        @{ Name = "VC Tools x86/x64"; Arg = "--add Microsoft.VisualStudio.Component.VC.Tools.x86.x64" }
        @{ Name = "VC ATL"; Arg = "--add Microsoft.VisualStudio.Component.VC.ATL" }
        @{ Name = "VC ATL ARM"; Arg = "--add Microsoft.VisualStudio.Component.VC.ATL.ARM" }
        @{ Name = "VC CMake Project"; Arg = "--add Microsoft.VisualStudio.Component.VC.CMake.Project" }
        @{ Name = "VC LLVM Clang"; Arg = "--add Microsoft.VisualStudio.Component.VC.Llvm.Clang" }
        @{ Name = "VC MFC"; Arg = "--add Microsoft.VisualStudio.Component.VC.MFC" }
        @{ Name = "VC TestAdapterForBoostTest"; Arg = "--add Microsoft.VisualStudio.Component.VC.TestAdapterForBoostTest" }
        @{ Name = "VC TestAdapterForGoogleTest"; Arg = "--add Microsoft.VisualStudio.Component.VC.TestAdapterForGoogleTest" }
        @{ Name = "VC ASAN"; Arg = "--add Microsoft.VisualStudio.Component.VC.ASAN" }
        @{ Name = "Windows 11 SDK v22621"; Arg = "--add Microsoft.VisualStudio.Component.Windows11SDK.22621" }
        @{ Name = "VC vcpkg"; Arg = "--add Microsoft.VisualStudio.Component.VC.vcpkg" }
        @{ Name = "VC CLI Support"; Arg = "--add Microsoft.VisualStudio.Component.VC.CLI.Support" }
        @{ Name = "Windows 10 SDK v19041"; Arg = "--add Microsoft.VisualStudio.Component.Windows10SDK.19041" }
        @{ Name = "VC Tools ARM64"; Arg = "--add Microsoft.VisualStudio.Component.VC.Tools.ARM64" }
        @{ Name = "VC Tools x86/x64 (repeat)"; Arg = "--add Microsoft.VisualStudio.Component.VC.Tools.x86.x64" }
        @{ Name = "VC v142 x86/x64"; Arg = "--add Microsoft.VisualStudio.Component.VC.v142.x86.x64" }
        @{ Name = "VC v141 x86/x64"; Arg = "--add Microsoft.VisualStudio.Component.VC.v141.x86.x64" }
        @{ Name = "VC v140 x86/x64"; Arg = "--add Microsoft.VisualStudio.Component.VC.v140.x86.x64" }
        @{ Name = "Security Static Analysis"; Arg = "--add Microsoft.VisualStudio.Component.Security.StaticAnalysis" }
        @{ Name = "VC Profiler"; Arg = "--add Microsoft.VisualStudio.Component.VC.Profiler" }
        @{ Name = "IntelliCode (repeat)"; Arg = "--add Microsoft.VisualStudio.Component.IntelliCode" }
        @{ Name = "LiveShare (repeat)"; Arg = "--add Microsoft.VisualStudio.LiveShare" }
        @{ Name = "JavaScript Diagnostics (repeat)"; Arg = "--add Microsoft.VisualStudio.Component.JavaScript.Diagnostics" }
        # Office/SharePoint Development
        @{ Name = "Office Tools"; Arg = "--add Microsoft.VisualStudio.Component.Office.Tools" }
        @{ Name = "Web Deploy"; Arg = "--add Microsoft.VisualStudio.Component.WebDeploy" }
        @{ Name = "IntelliCode (repeat 2)"; Arg = "--add Microsoft.VisualStudio.Component.IntelliCode" }
        @{ Name = ".NET 4.8 Targeting Pack (repeat)"; Arg = "--add Microsoft.Net.Component.4.8.TargetingPack" }
        @{ Name = "GitHub Copilot (repeat)"; Arg = "--add Microsoft.VisualStudio.Component.GitHub.Copilot" }
        @{ Name = ".NET 4.6.2 Targeting Pack (repeat)"; Arg = "--add Microsoft.Net.Component.4.6.2.TargetingPack" }
        @{ Name = ".NET 4.7.1 Targeting Pack (repeat)"; Arg = "--add Microsoft.Net.Component.4.7.1.TargetingPack" }
        @{ Name = ".NET 4.8.1 Targeting Pack (repeat)"; Arg = "--add Microsoft.Net.Component.4.8.1.TargetingPack" }
        # Individual components
        @{ Name = "VC ATL 141"; Arg = "--add Microsoft.VisualStudio.Component.VC.ATL.141" }
        @{ Name = "VC MFC 141"; Arg = "--add Microsoft.VisualStudio.Component.VC.MFC.141" }
        @{ Name = "VC ATL 142"; Arg = "--add Microsoft.VisualStudio.Component.VC.ATL.142" }
        @{ Name = "VC MFC 142"; Arg = "--add Microsoft.VisualStudio.Component.VC.MFC.142" }
        @{ Name = "VC CLI Support 142"; Arg = "--add Microsoft.VisualStudio.Component.VC.CLI.Support.142" }
    )

    Write-Host "Installing VS2022 features one at a time. This will take a while..."

    $baseArgs = @(
        "--quiet"
        "--wait"
        "--norestart"
        "--nocache"
        "--installPath `"C:\Program Files\Microsoft Visual Studio\2022\Professional`""
    )

    $total = $features.Count
    $current = 1
    foreach ($feature in $features) {
        Write-Host ("Installing feature ({0}/{1}): {2}" -f $current, $total, $feature.Name)
        $vsargs = $baseArgs + $feature.Arg
        try {
            Start-Process -FilePath $localInstaller -ArgumentList $vsargs -Wait -NoNewWindow -ErrorAction Stop
            Write-Host ("Successfully installed: {0} ({1}/{2})" -f $feature.Name, $current, $total)
        }
        catch {
            Write-Warning ("Failed to install feature: {0} ({1}/{2})" -f $feature.Name, $current, $total)
            $success = $false
            break
        }
        $current++
    }
    if ($success) {
        Write-Host "VS2022 all features installed successfully."
    }
    else {
        Write-Warning "VS2022 installation incomplete due to errors."
    }
    return $success
}

function Install-Git {
    Write-Host "Starting Git installation..."
    $success = $true
    $gitUrl = "https://github.com/git-for-windows/git/releases/download/v2.51.2.windows.1/Git-2.51.2-64-bit.exe"
    $localDir = "$env:TEMP\GitInstall"
    $localInstaller = Join-Path $localDir "Git-Setup.exe"
    try {
        # Check for existing installation
        $gitDefaultDir = "C:\Program Files\Git\cmd"
        $gitOtherDir = "C:\Program Files (x86)\Git\cmd";
        $gitExe = Join-Path $gitDefaultDir "git.exe"
        $gitExe2 = Join-Path $gitOtherDir "git.exe"
        if ((Test-Path $gitExe) -or (Test-Path $gitExe2)) {
            Write-Host "Git is already installed. Skipping installation."
        }
        else {
            if (-not (Test-Path $localDir)) {
                New-Item -ItemType Directory -Path $localDir -Force | Out-Null
            }
            Write-Host "Downloading Git installer from $gitUrl to $localInstaller"
            Invoke-WebRequest -Uri $gitUrl -OutFile $localInstaller -UseBasicParsing
            $installArgs = "/VERYSILENT /NORESTART /NOCANCEL /SP- /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS"
            Write-Host "Running Git installer silently..."
            Start-Process -FilePath $localInstaller -ArgumentList $installArgs -Wait -NoNewWindow -ErrorAction Stop
            Write-Host "Git installed successfully."
        }

        # Add to PATH if not already present
        if (Test-Path $gitExe) {
            $currentPath = [Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::Machine)
            if ($currentPath -notlike "*${gitDefaultDir}*") {
                Write-Host "Adding $gitDefaultDir to system PATH..."
                [Environment]::SetEnvironmentVariable("Path", "$currentPath;${gitDefaultDir}", [System.EnvironmentVariableTarget]::Machine)
                Write-Host "Git directory added to system PATH. You may need to restart your shell."
            }
            else {
                Write-Host "Git directory already in system PATH."
            }
        }
        else {
            Write-Warning "Could not find git.exe in the default install locations. Please ensure Git is in your PATH."
        }
    }
    catch {
        Write-Host "Failed to install Git: $($_.Exception.Message)" -ForegroundColor Red
        $success = $false
    }
    return $success
}

function Install-GitHubDesktop {
    Write-Host "Starting GitHub Desktop installation..."
    $ghdUrl = "https://central.github.com/deployments/desktop/desktop/latest/win32"
    $localDir = "$env:TEMP\GitHubDesktopInstall"
    $localInstaller = Join-Path $localDir "GitHubDesktopSetup.exe"
    try {
        # Check for existing installation (default location)
        $ghdExe = "$env:LOCALAPPDATA\GitHub Desktop\GitHubDesktop.exe"
        if (Test-Path $ghdExe) {
            Write-Host "GitHub Desktop is already installed at $ghdExe. Skipping installation."
            return $true
        }
        if (-not (Test-Path $localDir)) {
            New-Item -ItemType Directory -Path $localDir -Force | Out-Null
        }
        Write-Host "Downloading GitHub Desktop installer from $ghdUrl to $localInstaller"
        Invoke-WebRequest -Uri $ghdUrl -OutFile $localInstaller -UseBasicParsing
        $installArgs = "/silent"
        Write-Host "Running GitHub Desktop installer silently..."
        Start-Process -FilePath $localInstaller -ArgumentList $installArgs -Wait -NoNewWindow -ErrorAction Stop
        Write-Host "GitHub Desktop installed successfully."
        return $true
    }
    catch {
        Write-Host "Failed to install GitHub Desktop: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Get-Repositories {
    $defaultRepoPath = "C:/repos"
    $GitRepoPath = Read-Host "Where should git repositories be cloned? (Default: $defaultRepoPath)"
    if ([string]::IsNullOrWhiteSpace($GitRepoPath)) {
        $GitRepoPath = $defaultRepoPath
    }

    Write-Host "Creating git repository directory: $GitRepoPath"
    New-Item -ItemType Directory -Force -Path $GitRepoPath | Out-Null
    
    $repos = @(
        "enterprise-suite",
        "enterprise-technical-documentation",
        "enterprise-tools",
        "enterprise-api"
    )
    
    # Set registry key for DeveloperRepositoryPath to enterprise-suite repo path
    $enterpriseSuitePath = Join-Path $GitRepoPath "enterprise-suite"
    try {
        Write-Host "Setting DeveloperRepositoryPath in registry to $enterpriseSuitePath"
        $regPath = "Registry::HKEY_CLASSES_ROOT\Newforma Installation"
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "DeveloperRepositoryPath" -Value $enterpriseSuitePath -Type String
        Write-Host "Registry key set successfully."
    }
    catch {
        Write-Warning "Failed to set registry key: $($_.Exception.Message)"
    }

    foreach ($repo in $repos) {
        $repoPath = Join-Path $GitRepoPath $repo
        if (Test-Path $repoPath) {
            Write-Host "Repository $repo already exists, skipping clone"
        }
        else {
            Write-Host "Cloning $repo..."
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
            $nunitOk = Install-NUnitConsoleRunners
            if (-not $nunitOk) {
                Write-Host "Failed to install NUnit console runners. Please install manually and restart Powershell." -ForegroundColor Red
                exit 1
            }
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
        3 {
            Write-Host "Stage 3: Cloning repositories..."
            Get-Repositories
            Set-DevSetupStage "4"
            Write-Host "Stage 3 complete. Restarting shell for next stage..."
            Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-File", "`"$PSCommandPath`""
            Stop-Process -Id $PID
        }
        4 {
            Write-Host "Setup complete! All stages finished."
            Set-DevSetupStage "complete"
        }
        default {
            Write-Host "Unknown stage: $stage"
        }
    }
}

main