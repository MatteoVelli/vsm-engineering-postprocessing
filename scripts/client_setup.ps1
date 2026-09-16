param(
    [string]$ProjectRoot,
    [switch]$AssumeYes,
    [switch]$NoPause,
    [switch]$ProbePythonOnly,
    [switch]$SimulateMissingPython,
    [switch]$AllowBootstrapInProbe,
    [string]$CandidateProbeExecutable
)

$ErrorActionPreference = "Stop"

$ValidatedPythonMinors = @("3.11")
$ValidatedPythonRange = ">=3.11,<3.12"
$BootstrapPythonVersion = "3.11.9"
$BootstrapPythonInstallerUrl = "https://www.python.org/ftp/python/$BootstrapPythonVersion/python-$BootstrapPythonVersion-amd64.exe"
$BootstrapPythonNugetUrl = "https://www.nuget.org/api/v2/package/python/$BootstrapPythonVersion"
$MinimumInstallerBytes = 20MB
$MinimumNugetBytes = 10MB

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = Split-Path -Parent $PSScriptRoot
}

$ProjectRoot = [System.IO.Path]::GetFullPath($ProjectRoot)
$VenvDir = Join-Path $ProjectRoot ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"
$LocalPythonDir = Join-Path $ProjectRoot ".python\Python311"
$LocalPython = Join-Path $LocalPythonDir "python.exe"
$SetupCache = Join-Path $ProjectRoot "setup_cache"
$LogDir = Join-Path $ProjectRoot "logs"
$LogFile = Join-Path $LogDir "setup.log"

New-Item -ItemType Directory -Path $LogDir -Force | Out-Null

function Write-Log {
    param([string]$Message)
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    [System.IO.File]::AppendAllText($LogFile, "[$stamp] $Message`r`n", [System.Text.Encoding]::UTF8)
}

function Append-LogFile {
    param([string]$Path)
    if (Test-Path $Path) {
        [System.IO.File]::AppendAllText($LogFile, [System.IO.File]::ReadAllText($Path), [System.Text.Encoding]::UTF8)
    }
}

function Write-Step {
    param([string]$Message)
    Write-Host $Message
    Write-Log $Message
}

function Stop-Setup {
    param(
        [string]$Message,
        [int]$Code = 1
    )
    Write-Host ""
    Write-Host "ERROR: $Message" -ForegroundColor Red
    Write-Host "Details were written to: $LogFile"
    Write-Log "ERROR: $Message"
    if (-not $NoPause) {
        Write-Host ""
        Write-Host "Press any key to close this window."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    exit $Code
}

function Remove-PackageLocalVenv {
    $resolvedVenv = [System.IO.Path]::GetFullPath($VenvDir)
    $resolvedRoot = [System.IO.Path]::GetFullPath($ProjectRoot)
    if (-not $resolvedVenv.StartsWith($resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        Stop-Setup "Refusing to recreate an unexpected virtual environment path: $resolvedVenv" 32
    }
    Remove-Item -LiteralPath $resolvedVenv -Recurse -Force
}

function Test-WindowsAppsAlias {
    param([string]$Path)
    return $Path -like "*\WindowsApps\python*.exe"
}

function Test-ValidatedPythonVersion {
    param([version]$Version)
    $majorMinor = "$($Version.Major).$($Version.Minor)"
    return $ValidatedPythonMinors -contains $majorMinor
}

function Test-BootstrapUrl {
    param(
        [string]$Url,
        [string[]]$AllowedPrefixes
    )
    foreach ($prefix in $AllowedPrefixes) {
        if ($Url.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            return
        }
    }
    Stop-Setup "Configured Python bootstrap URL is not from an approved official Python distribution source." 23
}

function Test-DownloadedFile {
    param(
        [string]$Path,
        [int64]$MinimumBytes,
        [byte[]]$ExpectedPrefix,
        [string]$Description
    )
    if (-not (Test-Path $Path)) {
        Stop-Setup "$Description download completed without creating the file." 24
    }
    $downloadedFile = Get-Item $Path
    Write-Log "$Description bytes: $($downloadedFile.Length)"
    if ($downloadedFile.Length -lt $MinimumBytes) {
        Stop-Setup "Downloaded $Description is unexpectedly small. The official URL may not have returned a valid runtime package." 25
    }
    $stream = [System.IO.File]::OpenRead($downloadedFile.FullName)
    try {
        foreach ($expected in $ExpectedPrefix) {
            $actual = $stream.ReadByte()
            if ($actual -ne $expected) {
                Stop-Setup "Downloaded $Description does not have the expected file signature. Setup will not run or extract it." 27
            }
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Test-InstallerFile {
    param([string]$Path)
    Test-DownloadedFile -Path $Path -MinimumBytes $MinimumInstallerBytes -ExpectedPrefix @(77, 90) -Description "Python installer"
}

function Test-NugetRuntimeFile {
    param([string]$Path)
    Test-DownloadedFile -Path $Path -MinimumBytes $MinimumNugetBytes -ExpectedPrefix @(80, 75) -Description "Python runtime package"
}

function Download-CheckedFile {
    param(
        [string]$Url,
        [string]$Destination,
        [string[]]$AllowedPrefixes,
        [scriptblock]$Validator,
        [string]$Description,
        [int]$MissingCode
    )
    Test-BootstrapUrl -Url $Url -AllowedPrefixes $AllowedPrefixes
    $partial = "$Destination.part"
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $Url -OutFile $partial -UseBasicParsing -ErrorAction Stop
        & $Validator $partial
        Move-Item -LiteralPath $partial -Destination $Destination -Force
    }
    catch {
        Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue
        $response = $_.Exception.Response
        $statusCode = $null
        if ($null -ne $response) {
            try { $statusCode = [int]$response.StatusCode } catch { $statusCode = $null }
        }
        Write-Log "$Description download failed: $($_.Exception.GetType().FullName): $($_.Exception.Message)"
        if ($statusCode -eq 404) {
            Stop-Setup "Configured Python bootstrap $Description was not found at the expected official URL. This is a packaging configuration issue, not an internet connection problem." $MissingCode
        }
        if ($statusCode) {
            Stop-Setup "Setup could not download the required Python runtime. The official server returned HTTP $statusCode. See README_SERGIO.txt troubleshooting instructions." 20
        }
        if ($_.Exception -is [System.UnauthorizedAccessException] -or $_.Exception -is [System.IO.IOException]) {
            Stop-Setup "Setup could not write the Python runtime download to the extracted tool folder. Check that the folder is writable and that antivirus or file permissions are not blocking setup." 28
        }
        Stop-Setup "Setup could not download the required Python runtime. Check your internet connection, TLS/proxy settings or company download policy, then run SETUP_VSM_TOOL.bat again. If downloads are blocked, follow README_SERGIO.txt troubleshooting instructions." 20
    }
}

function Install-NugetPythonRuntime {
    $nugetPackage = Join-Path $SetupCache "python-$BootstrapPythonVersion-amd64.nupkg"
    $extractDir = Join-Path $SetupCache "python-$BootstrapPythonVersion-nuget"
    if (-not (Test-Path $nugetPackage)) {
        Write-Step "Downloading official CPython $BootstrapPythonVersion app-local runtime package..."
        Write-Log "Download URL: $BootstrapPythonNugetUrl"
        Download-CheckedFile -Url $BootstrapPythonNugetUrl -Destination $nugetPackage -AllowedPrefixes @("https://www.nuget.org/") -Validator ${function:Test-NugetRuntimeFile} -Description "runtime package" -MissingCode 29
    }
    else {
        Test-NugetRuntimeFile -Path $nugetPackage
    }

    if (Test-Path $extractDir) {
        Remove-Item -LiteralPath $extractDir -Recurse -Force
    }
    if (Test-Path $LocalPythonDir) {
        Remove-Item -LiteralPath $LocalPythonDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($nugetPackage, $extractDir)
    $toolsDir = Join-Path $extractDir "tools"
    if (-not (Test-Path (Join-Path $toolsDir "python.exe"))) {
        Stop-Setup "The official Python runtime package did not contain tools\python.exe." 30
    }
    Move-Item -LiteralPath $toolsDir -Destination $LocalPythonDir -Force
}

function Test-PythonCommand {
    param(
        [string]$Executable,
        [string[]]$PrefixArgs = @(),
        [string]$Source = "candidate"
    )

    if ([string]::IsNullOrWhiteSpace($Executable)) {
        return $null
    }
    if ((Test-Path $Executable) -and (Test-WindowsAppsAlias -Path ([System.IO.Path]::GetFullPath($Executable)))) {
        Write-Log "Skipping Microsoft Store execution alias: $Executable"
        return $null
    }

    $probe = "import platform, sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}|{platform.machine()}|{sys.executable}')"
    try {
        $output = & $Executable @PrefixArgs -c $probe 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $output) {
            Write-Log "Rejected Python candidate from $Source because it did not run: $Executable $($PrefixArgs -join ' ')"
            return $null
        }
    }
    catch {
        Write-Log "Rejected Python candidate from $Source because it failed to start: $Executable $($PrefixArgs -join ' ') :: $($_.Exception.Message)"
        return $null
    }

    $parts = ($output | Select-Object -First 1).Split("|")
    if ($parts.Count -lt 3) {
        Write-Log "Rejected Python candidate from $Source because probe output was unexpected: $output"
        return $null
    }

    $version = $null
    if (-not [version]::TryParse($parts[0], [ref]$version)) {
        Write-Log "Rejected Python candidate from $Source because probe version was unexpected: $($parts[0])"
        return $null
    }
    $machine = $parts[1]
    $resolvedExecutable = $parts[2]
    if (-not (Test-ValidatedPythonVersion -Version $version)) {
        Write-Log "Rejected Python candidate from $Source because version $version is outside the validated VSM runtime range $ValidatedPythonRange."
        return $null
    }
    if ($machine -notin @("AMD64", "x86_64")) {
        Write-Log "Rejected Python candidate from $Source because architecture $machine is unsupported."
        return $null
    }
    if (Test-WindowsAppsAlias -Path $resolvedExecutable) {
        Write-Log "Rejected Microsoft Store execution alias resolved as $resolvedExecutable"
        return $null
    }

    [pscustomobject]@{
        Executable = $Executable
        PrefixArgs = $PrefixArgs
        Version = $version.ToString()
        Architecture = $machine
        ResolvedExecutable = $resolvedExecutable
        Source = $Source
    }
}

function Find-CompatiblePython {
    if ($SimulateMissingPython) {
        Write-Log "Python discovery bypassed for safe missing-Python simulation."
        return $null
    }

    $candidates = @()
    if (Test-Path $VenvPython) {
        $candidates += @{ Exe = $VenvPython; Args = @(); Source = "package-local virtual environment" }
    }
    if (Test-Path $LocalPython) {
        $candidates += @{ Exe = $LocalPython; Args = @(); Source = "package-local CPython runtime" }
    }
    if ($env:VSM_PYTHON_EXE) {
        $candidates += @{ Exe = $env:VSM_PYTHON_EXE; Args = @(); Source = "VSM_PYTHON_EXE" }
    }

    $py = Get-Command "py.exe" -ErrorAction SilentlyContinue
    if ($py -and -not (Test-WindowsAppsAlias -Path $py.Source)) {
        $candidates += @{ Exe = $py.Source; Args = @("-3.11"); Source = "Python launcher py -3.11" }
    }

    foreach ($commandName in @("python.exe", "python3.exe")) {
        $command = Get-Command $commandName -ErrorAction SilentlyContinue
        if ($command -and -not (Test-WindowsAppsAlias -Path $command.Source)) {
            $candidates += @{ Exe = $command.Source; Args = @(); Source = "PATH command $commandName" }
        }
    }

    foreach ($root in @($env:LOCALAPPDATA, $env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        if ($root -eq $env:LOCALAPPDATA) {
            foreach ($folder in @("Python310", "Python311", "Python312", "Python313", "Python314")) {
                $path = Join-Path $root "Programs\Python\$folder\python.exe"
                if (Test-Path $path) {
                    $candidates += @{ Exe = $path; Args = @(); Source = "standard per-user install location" }
                }
            }
        }
        else {
            foreach ($folder in @("Python310", "Python311", "Python312", "Python313", "Python314")) {
                $path = Join-Path $root "$folder\python.exe"
                if (Test-Path $path) {
                    $candidates += @{ Exe = $path; Args = @(); Source = "standard all-users install location" }
                }
            }
        }
    }

    foreach ($candidate in $candidates) {
        $result = Test-PythonCommand -Executable $candidate.Exe -PrefixArgs $candidate.Args -Source $candidate.Source
        if ($null -ne $result) {
            return $result
        }
    }
    return $null
}

if (-not [string]::IsNullOrWhiteSpace($CandidateProbeExecutable)) {
    $candidateProbe = Test-PythonCommand -Executable $CandidateProbeExecutable -Source "candidate probe"
    if ($null -eq $candidateProbe) {
        Write-Host "REJECTED_PYTHON"
        exit 0
    }
    Write-Host "VALIDATED_PYTHON $($candidateProbe.Version) $($candidateProbe.ResolvedExecutable)"
    exit 0
}

function Install-PythonRuntime {
    Write-Host ""
    Write-Host "VSM Engineering Post-Processing Setup" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "A compatible VSM runtime was not found."
    Write-Host ""
    Write-Host "This release has been validated with Python 3.11.x."
    Write-Host ""
    Write-Host "The setup will install the supported Python $BootstrapPythonVersion"
    Write-Host "64-bit runtime for this VSM tool using the official CPython Windows installer."
    Write-Host ""
    Write-Host "Your existing Python installations will not be modified."
    Write-Host "No administrator access should be required in the normal case."
    Write-Host ""

    if ($ProbePythonOnly -and -not $AllowBootstrapInProbe) {
        Write-Host "MISSING_PYTHON_BOOTSTRAP_AVAILABLE"
        Write-Log "Probe-only mode reached missing-Python bootstrap prompt."
        exit 0
    }

    $answer = "N"
    if ($AssumeYes) {
        $answer = "Y"
    }
    else {
        $answer = Read-Host "Press Y to continue or N to cancel"
    }
    if ($answer -notmatch "^[Yy]$") {
        Stop-Setup "Setup cancelled. Install Python 3.11 manually or rerun setup and choose Y." 10
    }

    New-Item -ItemType Directory -Path $SetupCache -Force | Out-Null
    New-Item -ItemType Directory -Path $LocalPythonDir -Force | Out-Null
    $installer = Join-Path $SetupCache "python-$BootstrapPythonVersion-amd64.exe"

    if (-not (Test-Path $installer)) {
        Write-Step "Downloading official CPython $BootstrapPythonVersion runtime..."
        Write-Log "Download URL: $BootstrapPythonInstallerUrl"
        Download-CheckedFile -Url $BootstrapPythonInstallerUrl -Destination $installer -AllowedPrefixes @("https://www.python.org/") -Validator ${function:Test-InstallerFile} -Description "installer" -MissingCode 26
    }
    else {
        Test-InstallerFile -Path $installer
    }

    Write-Step "Installing Python $BootstrapPythonVersion into the package runtime folder..."
    $installLog = Join-Path $LogDir "python-installer.log"
    $installOutLog = Join-Path $LogDir "python-installer.out.log"
    $installErrLog = Join-Path $LogDir "python-installer.err.log"
    $installArgs = "/quiet InstallAllUsers=0 `"TargetDir=$LocalPythonDir`" Include_pip=1 Include_launcher=0 PrependPath=0 Shortcuts=0 Include_test=0"
    $process = Start-Process -FilePath $installer -ArgumentList $installArgs -Wait -PassThru -NoNewWindow -RedirectStandardOutput $installOutLog -RedirectStandardError $installErrLog
    if (Test-Path $installOutLog) { [System.IO.File]::AppendAllText($installLog, [System.IO.File]::ReadAllText($installOutLog), [System.Text.Encoding]::UTF8) }
    if (Test-Path $installErrLog) { [System.IO.File]::AppendAllText($installLog, [System.IO.File]::ReadAllText($installErrLog), [System.Text.Encoding]::UTF8) }
    Write-Log "Python installer exit code: $($process.ExitCode)"
    if ($process.ExitCode -ne 0) {
        Stop-Setup "Python installation failed. See $installLog for installer details." 21
    }

    $python = Test-PythonCommand -Executable $LocalPython -Source "installed package-local CPython runtime"
    if ($null -eq $python) {
        Write-Step "The Windows installer did not create the package-local runtime; extracting the official app-local Python package..."
        Write-Log "Falling back to official CPython NuGet runtime package because python.exe was not available at $LocalPython after installer exit code 0."
        Install-NugetPythonRuntime
        $python = Test-PythonCommand -Executable $LocalPython -Source "installed package-local CPython runtime"
    }
    if ($null -eq $python) {
        Stop-Setup "Python was installed but could not be validated at $LocalPython." 22
    }
    return $python
}

function Invoke-LoggedCommand {
    param(
        [string]$Executable,
        [string[]]$Arguments,
        [string]$FriendlyName,
        [int]$FailureCode
    )

    Write-Step $FriendlyName
    Write-Log "Running: $Executable $($Arguments -join ' ')"
    $logName = ($FriendlyName -replace "[^A-Za-z0-9]+", "_").Trim("_")
    $commandLog = Join-Path $LogDir "$logName.log"
    $commandOutLog = Join-Path $LogDir "$logName.out.log"
    $commandErrLog = Join-Path $LogDir "$logName.err.log"

    $argumentLine = ($Arguments | ForEach-Object {
        $arg = [string]$_
        if ($arg -match '[\s"]') {
            '"' + ($arg -replace '"', '\"') + '"'
        }
        else {
            $arg
        }
    }) -join " "

    try {
        $process = Start-Process -FilePath $Executable -ArgumentList $argumentLine -WorkingDirectory $ProjectRoot -Wait -PassThru -NoNewWindow -RedirectStandardOutput $commandOutLog -RedirectStandardError $commandErrLog
        $exitCode = $process.ExitCode
    }
    catch {
        Write-Log "$FriendlyName could not be started: $($_.Exception.Message)"
        Stop-Setup "$FriendlyName could not be started. See $LogFile for details." $FailureCode
    }
    if (Test-Path $commandOutLog) { [System.IO.File]::WriteAllText($commandLog, [System.IO.File]::ReadAllText($commandOutLog), [System.Text.Encoding]::UTF8) }
    if (Test-Path $commandErrLog) { [System.IO.File]::AppendAllText($commandLog, [System.IO.File]::ReadAllText($commandErrLog), [System.Text.Encoding]::UTF8) }
    Append-LogFile -Path $commandLog
    Write-Log "$FriendlyName exit code: $exitCode"
    if ($exitCode -ne 0) {
        Stop-Setup "$FriendlyName failed. See $commandLog for details." $FailureCode
    }
}

Clear-Host
Write-Host "============================================================"
Write-Host " VSM Engineering Post-Processing"
Write-Host " Setup"
Write-Host "============================================================"
Write-Host ""
Write-Log "Setup started for project root: $ProjectRoot"

if (-not (Test-Path (Join-Path $ProjectRoot "pyproject.toml"))) {
    Stop-Setup "Could not find pyproject.toml in $ProjectRoot. Extract the complete ZIP and run SETUP_VSM_TOOL.bat from the extracted folder." 2
}

$python = Find-CompatiblePython
if ($null -eq $python) {
    $python = Install-PythonRuntime
}
elseif ($ProbePythonOnly) {
    Write-Host "FOUND_PYTHON $($python.Version) $($python.ResolvedExecutable)"
    Write-Log "Probe-only mode found Python: $($python.ResolvedExecutable)"
    exit 0
}

Write-Step "Using Python $($python.Version) [$($python.Architecture)] from $($python.ResolvedExecutable)"

$venvCheck = $null
if (Test-Path $VenvPython) {
    $venvCheck = Test-PythonCommand -Executable $VenvPython -Source "package-local virtual environment"
}
if ((Test-Path $VenvPython) -and $null -eq $venvCheck) {
    Write-Step "Existing package-local virtual environment is outside the validated runtime range; recreating it..."
    Remove-PackageLocalVenv
}

if (-not (Test-Path $VenvPython)) {
    Write-Step "Creating package-local virtual environment..."
    $venvArgs = @($python.PrefixArgs + @("-m", "venv", $VenvDir))
    try {
        & $python.Executable @venvArgs >> $LogFile 2>&1
    }
    catch {
        Write-Log "Virtual environment creation failed: $($_.Exception.Message)"
        Stop-Setup "Virtual environment creation failed. Check that the extracted folder is writable, then run setup again." 30
    }
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $VenvPython)) {
        Stop-Setup "Virtual environment creation failed. Check that the extracted folder is writable, then run setup again." 30
    }
}
else {
    Write-Step "Package-local virtual environment already exists."
}

$venvCheck = Test-PythonCommand -Executable $VenvPython -Source "package-local virtual environment"
if ($null -eq $venvCheck) {
    Stop-Setup "The package-local virtual environment exists but its Python interpreter could not be validated." 31
}

Invoke-LoggedCommand -Executable $VenvPython -Arguments @("-m", "pip", "--version") -FriendlyName "Checking Python packaging tools" -FailureCode 40
Invoke-LoggedCommand -Executable $VenvPython -Arguments @("-m", "pip", "install", "-e", $ProjectRoot) -FriendlyName "Installing VSM application dependencies" -FailureCode 41

Write-Step "Validating installation..."
$doctorLog = Join-Path $LogDir "doctor.log"
& $VenvPython -m vsm_postprocessing.doctor_cli --project-root $ProjectRoot > $doctorLog 2>&1
Append-LogFile -Path $doctorLog
Get-Content $doctorLog
if ($LASTEXITCODE -ne 0) {
    Stop-Setup "Installation validation failed. See $doctorLog for details." 50
}

Write-Host ""
Write-Host "============================================================"
Write-Host " VSM Engineering Post-Processing"
Write-Host " Setup completed successfully."
Write-Host "============================================================"
Write-Host ""
Write-Host "The application is ready to use."
Write-Host ""
Write-Host "Double-click:"
Write-Host ""
Write-Host "START_VSM_TOOL.bat"
Write-Host ""
Write-Host "to launch the application."
Write-Log "Setup completed successfully."

if (-not $NoPause) {
    Write-Host "Press any key to close this window."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
