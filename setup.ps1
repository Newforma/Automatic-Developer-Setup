function Install-VisualStudio2015 {
    $sourcePath = "\\newforma.local\data\departments\Development\Installation Kits\Microsoft\Visual Studio 2015 Pro\vs_professional.exe"
    $localDir = "$env:TEMP\VS2015Install"
    $localInstaller = Join-Path $localDir "vs_professional.exe"

    try {
        if (-not (Test-Path $localDir)) {
            New-Item -ItemType Directory -Path $localDir -Force | Out-Null
        }
        Copy-Item -Path $sourcePath -Destination $localInstaller -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Warning "Failed to copy installer. Continuing anyway."
    }

    $installArgs = '/quiet /norestart /log "%TEMP%\VS2015Install.log" ' +
        '/features ' +
        'OfficeTools,VC,VC_MFC,VC_MFC_XP,VC_Common'

    try {
        Start-Process -FilePath $localInstaller -ArgumentList $installArgs -Wait -NoNewWindow -ErrorAction SilentlyContinue
    } catch {
        Write-Warning "Failed to start VS2015 installer. Continuing anyway."
    }
}

function Install-VisualStudio2022 {
    Write-Host "Installing Visual Studio 2022 Professional with selected workloads and components..."

    $vsPackageId = "Microsoft.VisualStudio.2022.Professional"
    $workloads = @(
        "--add Microsoft.VisualStudio.Workload.ManagedDesktop"
        "--add Microsoft.VisualStudio.Workload.NativeDesktop"
        "--add Microsoft.VisualStudio.Workload.Office"
    )
    $components = @(
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

    $allArgs = @(
        "install"
        "--id $vsPackageId"
        "--silent"
        "--accept-package-agreements"
        "--accept-source-agreements"
        $workloads
        $components
    ) -join " "

    winget $allArgs
}

function main {
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
        [System.Environment]::SetEnvironmentVariable("DEV_SETUP_STAGE", "$stage", "User")
    }

    switch ($stage) {
        1 {
            Write-Host "Stage 1: Installing Visual Studio 2015 and 2022..."
            Install-VisualStudio2015
            Install-VisualStudio2022
            [System.Environment]::SetEnvironmentVariable("DEV_SETUP_STAGE", "2", "User")
            Write-Host "Stage set to 2. Restarting shell..."
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