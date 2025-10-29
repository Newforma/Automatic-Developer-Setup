function Set-DeveloperProvisionNixName {
    # Sets HKLM\Software\Wow6432Node\Newforma\20XX\DeveloperProvisionNixName to the local machine name
    $regPath = "HKLM:\Software\Wow6432Node\Newforma\20XX"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    $machineName = $env:COMPUTERNAME
    Set-ItemProperty -Path $regPath -Name "DeveloperProvisionNixName" -Value $machineName
}

function Install-IISCertificate {
    $env:COMPUTERNAME = $env:COMPUTERNAME

    Write-Host "Setting up certificate for: $env:COMPUTERNAME" -ForegroundColor Cyan
    Import-Module WebAdministration -ErrorAction SilentlyContinue

    # Create certificate
    $subject = "CN=$env:COMPUTERNAME.newforma.local, O=Newforma, OU=Dev, L=Manchester, ST=NH, C=US"
    $dnsNames = "$env:COMPUTERNAME", "$DevMachineName.newforma.local"
    
    $cert = New-SelfSignedCertificate `
        -Subject $subject `
        -DnsName $dnsNames `
        -CertStoreLocation "Cert:\LocalMachine\My" `
        -KeyExportPolicy Exportable `
        -KeySpec KeyExchange `
        -KeyUsage DigitalSignature, KeyEncipherment `
        -KeyAlgorithm RSA `
        -KeyLength 2048 `
        -FriendlyName $env:COMPUTERNAME `
        -NotAfter (Get-Date).AddYears(5)
    
    Write-Host "Certificate created: $($cert.Thumbprint)"

    # Configure IIS binding
    $existingBinding = Get-WebBinding -Name "Default Web Site" -Protocol https -Port "443" -ErrorAction SilentlyContinue
    if ($existingBinding) {
        Remove-WebBinding -Name "Default Web Site" -Protocol https -Port "443" -Confirm:$false
    }

    New-WebBinding -Name "Default Web Site" -Protocol https -Port "443" -IPAddress "*" | Out-Null
    $binding = Get-WebBinding -Name "Default Web Site" -Protocol https -Port "443"
    $binding.AddSslCertificate($cert.Thumbprint, "My")
    
    Write-Host "HTTPS binding configured"

    # Configure authentication
    try {
        Set-WebConfigurationProperty -Filter "/system.webServer/security/authentication/anonymousAuthentication" -Name enabled -Value $true -PSPath "IIS:\" -Location "Default Web Site" -ErrorAction Stop
        Set-WebConfigurationProperty -Filter "/system.webServer/security/authentication/windowsAuthentication" -Name enabled -Value $false -PSPath "IIS:\" -Location "Default Web Site" -ErrorAction Stop
        Set-WebConfigurationProperty -Filter "/system.webServer/security/authentication/basicAuthentication" -Name enabled -Value $false -PSPath "IIS:\" -Location "Default Web Site" -ErrorAction Stop
        Set-WebConfigurationProperty -Filter "/system.webServer/security/authentication/digestAuthentication" -Name enabled -Value $false -PSPath "IIS:\" -Location "Default Web Site" -ErrorAction Stop
        Write-Host "Authentication configured"
    }
    catch {
        Write-Warning "Could not configure authentication settings"
    }

    # Clean up old certificates
    Get-ChildItem -Path "Cert:\LocalMachine\My" | 
    Where-Object { $_.FriendlyName -eq $env:COMPUTERNAME -and $_.Thumbprint -ne $cert.Thumbprint -and $_.NotAfter -lt (Get-Date) } |
    ForEach-Object { Remove-Item -Path "Cert:\LocalMachine\My\$($_.Thumbprint)" -Force }

    Write-Host "Certificate setup complete for $env:COMPUTERNAME"
    return $cert
}

function Enable-IISFeatures {
    Write-Host "Enabling IIS features..."
    
    $features = @(
        "IIS-WebServerRole",
        "IIS-WebServer",
        "IIS-CommonHttpFeatures",
        "IIS-HttpErrors",
        "IIS-HttpRedirect",
        "IIS-ApplicationDevelopment",
        "IIS-NetFxExtensibility45",
        "IIS-HealthAndDiagnostics",
        "IIS-HttpLogging",
        "IIS-LoggingLibraries",
        "IIS-RequestMonitor",
        "IIS-HttpTracing",
        "IIS-Security",
        "IIS-RequestFiltering",
        "IIS-Performance",
        "IIS-WebServerManagementTools",
        "IIS-IIS6ManagementCompatibility",
        "IIS-Metabase",
        "IIS-ManagementConsole",
        "IIS-BasicAuthentication",
        "IIS-WindowsAuthentication",
        "IIS-StaticContent",
        "IIS-DefaultDocument",
        "IIS-WebSockets",
        "IIS-ApplicationInit",
        "IIS-ISAPIExtensions",
        "IIS-ISAPIFilter",
        "IIS-HttpCompressionStatic",
        "IIS-ASPNET45",
        "IIS-CGI"
    )
    
    foreach ($feature in $features) {
        Enable-WindowsOptionalFeature -Online -FeatureName $feature -All -NoRestart -ErrorAction SilentlyContinue
    }
    
    Write-Host "IIS features enabled"
}

function Install-MySql {
    param(
        [Parameter(Mandatory = $true)][string]$GitRepoPath
    )
    $installerPath = Find-MySQLInstaller
    if (-not $installerPath) {
        return $false
    }
    Write-Host "Starting silent MySQL installation from $installerPath..."
    $arguments = @(
        "/i"
        "`"$installerPath`""
        "/qn"
        "/norestart"
        'INSTALLDIR="C:\Program Files\MySQL\MySQL Server"'
        "SERVICENAME=MySQL"
        "ROOTPASSWORD=root"
        "ADDLOCAL=ALL"
    )
    try {
        $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -PassThru -NoNewWindow -ErrorAction Stop
        if ($process.ExitCode -eq 0) {
            Write-Host "MySQL installed successfully."
            # Find MySQL installation path
            $mysqlBase = "C:\Program Files\MySQL"
            $mysqlDir = Get-ChildItem -Path $mysqlBase -Directory | Where-Object { $_.Name -like "MySQL Server*" } | Sort-Object Name -Descending | Select-Object -First 1
            if ($mysqlDir) {
                $pluginDir = Join-Path $mysqlDir.FullName "lib\plugin"
                $srcDir = Join-Path $GitRepoPath "enterprise-suite\Solutions\MySqlWordbreaker\release\x64"
                $files = Get-ChildItem -Path $srcDir -File -ErrorAction SilentlyContinue
                if ($files) {
                    Write-Host "Copying MySqlWordbreaker plugin files to $pluginDir..."
                    if (-not (Test-Path $pluginDir)) { New-Item -ItemType Directory -Path $pluginDir -Force | Out-Null }
                    foreach ($file in $files) {
                        Copy-Item -Path $file.FullName -Destination $pluginDir -Force
                    }
                    Write-Host "Plugin files copied."
                }
                else {
                    Write-Warning "No plugin files found in $srcDir"
                }
            }
            else {
                Write-Warning "Could not find MySQL installation directory in $mysqlBase"
            }
            return $true
        }
        else {
            Write-Host "MySQL installer exited with code $($process.ExitCode)" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "Failed to install MySQL: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}
function Find-MySQLInstaller {
    $basePath = "\\winnas01\qadata$\Builds\Project_Center\develop"
    if (-not (Test-Path $basePath)) {
        Write-Warning "Base path not found: $basePath"
        return $null
    }
    $folders = Get-ChildItem -Path $basePath -Directory | Where-Object { $_.Name -match '^20\d{2}\.\d+\.\d+$' }
    if (-not $folders) {
        Write-Warning "No versioned folders found in $basePath"
        return $null
    }
    # Sort by version descending (latest first)
    $latest = $folders | Sort-Object { [version]($_.Name -replace '^([0-9]+\.[0-9]+\.[0-9]+)$', '$1') } -Descending | Select-Object -First 1
    Write-Host "Latest versioned folder: $($latest.Name)"

    # Find latest build subfolder (5-digit number)
    $buildFolders = Get-ChildItem -Path $latest.FullName -Directory | Where-Object { $_.Name -match '^\d{5}$' }
    if (-not $buildFolders) {
        Write-Warning "No build subfolders found in $($latest.FullName)"
        return $null
    }
    $latestBuild = $buildFolders | Sort-Object Name -Descending | Select-Object -First 1
    Write-Host "Latest build subfolder: $($latestBuild.Name)"

    $mysqlDir = Join-Path $latestBuild.FullName "MySQLMigrator\MySQL"
    if (-not (Test-Path $mysqlDir)) {
        Write-Warning "MySQL directory not found: $mysqlDir"
        return $null
    }
    $msi = Get-ChildItem -Path $mysqlDir -Filter *.msi -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($msi) {
        Write-Host "Found MySQL installer (msi): $($msi.FullName)"
        return $msi.FullName
    }
    $exe = Get-ChildItem -Path $mysqlDir -Filter *.exe -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($exe) {
        Write-Host "Found MySQL installer (exe): $($exe.FullName)"
        return $exe.FullName
    }
    Write-Warning "No MySQL installer (.msi or .exe) found in $mysqlDir"
    return $null
}

function Get-NPCVersionFromRepo {
    $versionFile = Join-Path $GitRepoPath 'enterprise-suite\Solutions\enterprise-core\Core\Remote\RemoteConstants.cs'
    if (-not (Test-Path $versionFile)) {
        Write-Warning "Could not find version file: $versionFile"
        return $null
    }
    $content = Get-Content $versionFile | Where-Object { $_ -match 'public const string VERSION' }
    if ($content) {
        if ($content -match '"([0-9]+\.[0-9]+)"') {
            return $matches[1]
        }
        else {
            Write-Warning "VERSION line found but could not extract version."
            return $null
        }
    }
    else {
        Write-Warning "VERSION line not found in $versionFile"
        return $null
    }
}

function Update-SessionPath {
    # Updates the current session's $env:PATH from the system (machine) PATH
    $machinePath = [Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::Machine)
    $env:PATH = $machinePath
    Write-Host "Session PATH updated from system PATH."
}

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
    [Environment]::SetEnvironmentVariable("DEV_SETUP_STAGE", $StageValue, [System.EnvironmentVariableTarget]::Machine)
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
        if ($($(Test-Path $primarySource) 2>$null)) {
            try {
                Copy-Item -Path $primarySource -Destination $localInstaller -Force -ErrorAction Stop 2>$null
                $copied = $true
            }
            catch {
                Write-Warning "Failed to copy from primary location: $($_.Exception.Message)"
            }
        }
        if (-not $copied -and $($(Test-Path $backupSource) 2>$null)) {
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
    
    # User will need to log in to github, so play a sound to get their attention
    Write-Host "Please log in to GitHub if prompted to authenticate."
    [System.Media.SystemSounds]::Exclamation.Play()

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

    return $GitRepoPath
}

function main {
    # Ensure running as administrator
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "This script must be run as an administrator." -ForegroundColor Red
        exit 1
    }

    $profilePath = $PROFILE
    $profileLine = "& `"$PSCommandPath`""
    Write-Host "Ensuring script is in user profile: $profilePath"
    if (-not (Test-Path $profilePath)) {
        New-Item -ItemType File -Path $profilePath -Force | Out-Null
    }
    $profileContent = Get-Content $profilePath -Raw
    $escapedProfileLine = [regex]::Escape($profileLine)
    if ($profileContent -notmatch $escapedProfileLine) {
        Add-Content -Path $profilePath -Value $profileLine
    }

    if (-not (Test-Path $profilePath)) {
        New-Item -ItemType File -Path $profilePath -Force | Out-Null
    }
    $profileContent = Get-Content $profilePath -Raw
    if ([String]::IsNullOrEmpty(($profileContent) -or $profileContent -notmatch [regex]::Escape($profileLine))) {
        Add-Content -Path $profilePath -Value $profileLine
    }

    $stage = [Environment]::GetEnvironmentVariable("DEV_SETUP_STAGE", [System.EnvironmentVariableTarget]::Machine)
    if (-not $stage) {
        $stage = $env:DEV_SETUP_STAGE
    }

    if (-not $stage) {
        Write-Host "Setup stage not set. Defaulting to 1."
        $stage = 1
        Set-DevSetupStage "$stage"
    }

    switch ($stage) {
        1 {
            Write-Host "Stage 1: Installing Visual Studio 2015 and 2022..."
            Update-SessionPath
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
            Start-Process -FilePath "powershell.exe" -ArgumentList "-File", "`"$PSCommandPath`""
            Stop-Process -Id $PID
        }
        2 {
            Write-Host "Stage 2: Installing NUnit console runners..."
            Update-SessionPath
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
            Start-Process -FilePath "powershell.exe" -ArgumentList "-File", "`"$PSCommandPath`""
            Stop-Process -Id $PID
        }
        3 {
            Write-Host "Stage 3: Cloning repositories..."
            Update-SessionPath
            $GitRepoPath = Get-Repositories
            [System.Media.SystemSounds]::Exclamation.Play()
            Write-Host "Launching Redemption installer, please install to $GitRepoPath/enterprise-suite/Third-Party/Redemption..."
            Start-Process $GitRepoPath/enterprise-suite/Solutions/ThirdParty/Redemption/Install.exe -Wait
            [Environment]::SetEnvironmentVariable("OFFICE64", "1", [System.EnvironmentVariableTarget]::Machine)
            Install-MySql $GitRepoPath
            Enable-IISFeatures
            Install-IISCertificate
            Set-DeveloperProvisionNixName
            Set-DevSetupStage "4"
            Write-Host "Stage 3 complete. Restarting shell for next stage..."
            Start-Process -FilePath "powershell.exe" -ArgumentList "-File", "`"$PSCommandPath`""
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