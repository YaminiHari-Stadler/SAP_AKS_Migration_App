<#
    AKS Migration App launcher.

    Normal users never run this directly: they double-click
    "Start AKS Migration.bat", which calls this script.

    What it does, in order:

      1. Is a working environment already installed?  -> start the app.
      2. Is Python installed on this PC?              -> create .venv and install.
      3. No Python at all?                            -> download a private,
         self-contained Python into runtime\ and install there.

    Nothing is installed system-wide, no administrator rights are needed, the
    user's own Python (if any) is not modified, and PATH is not changed.
    The script is idempotent: a second run just starts the app.
#>

[CmdletBinding()]
param(
    [int]    $Port = 8501,
    [switch] $Reinstall,      # rebuild the environment from scratch
    [switch] $NoBrowser,      # do not open the browser automatically
    [switch] $DebugErrors,    # show technical error details inside the app
    [string] $PythonVersion = "3.12.10"
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$AppRoot      = Split-Path -Parent $MyInvocation.MyCommand.Definition

# The environment lives in the user profile, not in the application folder. The
# application folder is usually on OneDrive, and a 400 MB Python environment
# there would be uploaded to SharePoint and is prone to sync locking. The user
# profile needs no administrator rights either.
function Get-EnvironmentRoot {
    $local = $env:LOCALAPPDATA
    if ($local -and $local -notmatch "OneDrive" -and (Test-Path $local)) {
        return (Join-Path $local "AKS Migration App")
    }
    return (Join-Path $AppRoot ".aks-env")
}

$EnvRoot      = Get-EnvironmentRoot
$VenvDir      = Join-Path $EnvRoot "env"
$RuntimeDir   = Join-Path $EnvRoot "python"
$LegacyVenv   = Join-Path $AppRoot ".venv"
$LogDir       = Join-Path $AppRoot "logs"
$Requirements = Join-Path $AppRoot "requirements.txt"
$AppScript    = Join-Path $AppRoot "app.py"
$StampName    = "aks_env.stamp"

if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$LauncherLog = Join-Path $LogDir ("launcher_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

function Write-Log {
    param([string] $Message, [string] $Level = "INFO")
    $line = "{0}  {1,-7}  {2}" -f (Get-Date -Format "HH:mm:ss"), $Level, $Message
    Add-Content -Path $LauncherLog -Value $line -Encoding utf8
    switch ($Level) {
        "ERROR" { Write-Host $Message -ForegroundColor Red }
        "WARN"  { Write-Host $Message -ForegroundColor Yellow }
        "STEP"  { Write-Host ""; Write-Host $Message -ForegroundColor Cyan }
        default { Write-Host $Message }
    }
}

function Stop-WithMessage {
    param([string] $Message, [string[]] $Hints = @())
    Write-Log $Message "ERROR"
    foreach ($hint in $Hints) { Write-Host "   - $hint" -ForegroundColor Yellow }
    Write-Host ""
    Write-Host "Details for IT: $LauncherLog" -ForegroundColor DarkGray
    exit 1
}

Write-Host ""
Write-Host "  AKS Migration App" -ForegroundColor White
Write-Host "  Convert legacy AKS lists to the AKS V2 format" -ForegroundColor DarkGray
Write-Host "  --------------------------------------------" -ForegroundColor DarkGray
Write-Log "Launcher started in $AppRoot"
Write-Log "Windows $([Environment]::OSVersion.Version) / PowerShell $($PSVersionTable.PSVersion)"

if (-not (Test-Path $AppScript))    { Stop-WithMessage "app.py was not found in $AppRoot." @("Keep the whole AKS Migration App folder together when copying it.") }
if (-not (Test-Path $Requirements)) { Stop-WithMessage "requirements.txt was not found in $AppRoot." @("Keep the whole AKS Migration App folder together when copying it.") }

# ---------------------------------------------------------------------------
# Has the environment already been installed for exactly these requirements?
# ---------------------------------------------------------------------------

function Get-RequirementsStamp {
    $hash = (Get-FileHash -Path $Requirements -Algorithm SHA256).Hash
    return "requirements=$hash"
}

function Test-EnvironmentReady {
    param([string] $PythonExe)
    if (-not (Test-Path $PythonExe)) { return $false }
    $stampFile = Join-Path (Split-Path -Parent $PythonExe) $StampName
    if (-not (Test-Path $stampFile)) { return $false }
    return ((Get-Content $stampFile -Raw).Trim() -eq (Get-RequirementsStamp))
}

function Set-EnvironmentReady {
    param([string] $PythonExe)
    $stampFile = Join-Path (Split-Path -Parent $PythonExe) $StampName
    Set-Content -Path $stampFile -Value (Get-RequirementsStamp) -Encoding utf8
}

$VenvPython       = Join-Path $VenvDir "Scripts\python.exe"
$RuntimePython    = Join-Path $RuntimeDir "python.exe"
$LegacyVenvPython = Join-Path $LegacyVenv "Scripts\python.exe"

if ($Reinstall) {
    Write-Log "Rebuilding the environment because -Reinstall was given." "STEP"
    # Discard the environments themselves, not just their stamps, so a broken
    # installation is actually repaired rather than reused.
    foreach ($stale in @($VenvDir, $RuntimeDir)) {
        if (Test-Path $stale) { Remove-Item $stale -Recurse -Force -ErrorAction SilentlyContinue }
    }
    $legacyStamp = Join-Path $LegacyVenv "Scripts\$StampName"
    if (Test-Path $legacyStamp) { Remove-Item $legacyStamp -Force }
}

$Python = $null
if (Test-EnvironmentReady $VenvPython) {
    $Python = $VenvPython
    Write-Log "Using the existing environment in $VenvDir"
} elseif (Test-EnvironmentReady $RuntimePython) {
    $Python = $RuntimePython
    Write-Log "Using the existing private Python in $RuntimeDir"
} elseif (Test-EnvironmentReady $LegacyVenvPython) {
    # An environment created by an earlier version, inside the application folder.
    $Python = $LegacyVenvPython
    Write-Log "Using the existing environment in $LegacyVenv"
}

# ---------------------------------------------------------------------------
# Install, if needed
# ---------------------------------------------------------------------------

function Find-SystemPython {
    <# Return the path of a usable 64-bit CPython 3.10-3.13 already on this PC. #>
    $candidates = @()
    $launcher = Get-Command "py" -ErrorAction SilentlyContinue
    if ($launcher) {
        foreach ($tag in @("-3.12", "-3.11", "-3.13", "-3.10", "-3")) {
            try {
                $found = & $launcher.Source $tag -c "import sys;print(sys.executable)" 2>$null
                if ($LASTEXITCODE -eq 0 -and $found) { $candidates += $found.Trim() }
            } catch { }
        }
    }
    foreach ($name in @("python", "python3")) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        # The Microsoft Store stub in WindowsApps is not a real interpreter.
        if ($command -and $command.Source -notmatch "WindowsApps") { $candidates += $command.Source }
    }

    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if (-not (Test-Path $candidate)) { continue }
        try {
            $info = & $candidate -c "import sys,struct;print(sys.version_info[0],sys.version_info[1],struct.calcsize(chr(80))*8)" 2>$null
        } catch { continue }
        if ($LASTEXITCODE -ne 0 -or -not $info) { continue }
        $parts = $info.Trim() -split "\s+"
        if ($parts.Count -lt 3) { continue }
        [int] $major = $parts[0]; [int] $minor = $parts[1]; [int] $bits = $parts[2]
        if ($major -eq 3 -and $minor -ge 10 -and $minor -le 13 -and $bits -eq 64) {
            Write-Log "Found Python $major.$minor (64-bit) at $candidate"
            return $candidate
        }
        Write-Log "Ignoring Python $major.$minor ($bits-bit) at $candidate" "WARN"
    }
    return $null
}

function Install-PrivatePython {
    <# Download a self-contained Python into runtime\python. No admin rights. #>
    $zipUrl    = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-embed-amd64.zip"
    $shortVer  = ($PythonVersion -split "\.")[0..1] -join ""     # 3.12.10 -> 312
    $zipPath   = Join-Path $env:TEMP "python-$PythonVersion-embed-amd64.zip"
    $getPipUrl = "https://bootstrap.pypa.io/get-pip.py"
    $getPip    = Join-Path $env:TEMP "get-pip.py"

    Write-Log "No suitable Python was found on this PC." "STEP"
    Write-Host "   Downloading a private Python $PythonVersion for this application only."
    Write-Host "   Nothing is installed system-wide and no administrator rights are needed."

    if (Test-Path $RuntimeDir) { Remove-Item $RuntimeDir -Recurse -Force }
    New-Item -ItemType Directory -Path $RuntimeDir -Force | Out-Null

    try {
        Write-Log "Downloading $zipUrl"
        Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
        Expand-Archive -Path $zipPath -DestinationPath $RuntimeDir -Force
        Remove-Item $zipPath -Force
    } catch {
        Stop-WithMessage "The Python runtime could not be downloaded." @(
            "Check the internet connection.",
            "If this PC uses a proxy, ask IT to allow https://www.python.org and https://pypi.org.",
            "Alternatively install Python 3.12 for Windows from the Microsoft Store or python.org and run this launcher again."
        )
    }

    # The embedded distribution disables site-packages by default; enable it so
    # pip-installed packages (and pywin32's path file) are importable.
    $pthFile = Join-Path $RuntimeDir "python$shortVer._pth"
    if (-not (Test-Path $pthFile)) { Stop-WithMessage "The downloaded Python runtime looks incomplete ($pthFile is missing)." }
    Set-Content -Path $pthFile -Encoding ascii -Value @(
        "python$shortVer.zip",
        ".",
        "Lib\site-packages",
        "import site"
    )

    try {
        Write-Log "Installing pip into the private runtime"
        Invoke-WebRequest -Uri $getPipUrl -OutFile $getPip -UseBasicParsing
        & $RuntimePython $getPip --no-warn-script-location --quiet
        if ($LASTEXITCODE -ne 0) { throw "get-pip.py returned $LASTEXITCODE" }
        Remove-Item $getPip -Force -ErrorAction SilentlyContinue
    } catch {
        Stop-WithMessage "pip could not be installed into the private Python runtime. ($_)" @(
            "Check the internet connection and any proxy settings.",
            "Ask IT to allow https://bootstrap.pypa.io and https://pypi.org."
        )
    }
    return $RuntimePython
}

function Install-Requirements {
    param([string] $PythonExe)
    Write-Log "Installing the application components. The first run takes a few minutes." "STEP"
    Write-Host "   (Streamlit, openpyxl, pyxlsb, pywin32, pandas)"
    & $PythonExe -m pip install --disable-pip-version-check --no-warn-script-location --upgrade pip 2>&1 |
        Tee-Object -FilePath $LauncherLog -Append | Out-Null
    & $PythonExe -m pip install --disable-pip-version-check --no-warn-script-location -r $Requirements 2>&1 |
        Tee-Object -FilePath $LauncherLog -Append | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Stop-WithMessage "The application components could not be installed." @(
            "Check the internet connection.",
            "If this PC uses a proxy, ask IT to allow https://pypi.org and https://files.pythonhosted.org.",
            "Then start the application again."
        )
    }
    Write-Log "All components installed."
}

if (-not $Python) {
    $systemPython = Find-SystemPython
    if ($systemPython) {
        if (-not (Test-Path $VenvPython)) {
            Write-Log "Creating a private environment in $VenvDir" "STEP"
            & $systemPython -m venv $VenvDir 2>&1 | Tee-Object -FilePath $LauncherLog -Append | Out-Null
            if ($LASTEXITCODE -ne 0 -or -not (Test-Path $VenvPython)) {
                Stop-WithMessage "A private Python environment could not be created in $VenvDir." @(
                    "Make sure the AKS Migration App folder is writable.",
                    "OneDrive 'Files On-Demand' can block this: right-click the folder and choose 'Always keep on this device'."
                )
            }
        }
        $Python = $VenvPython
    } else {
        $Python = Install-PrivatePython
    }
    Install-Requirements $Python
    Set-EnvironmentReady $Python
}

Write-Log "Python: $Python"

# Streamlit asks for an e-mail address on its very first start. Answer it once.
$credentials = Join-Path $env:USERPROFILE ".streamlit\credentials.toml"
if (-not (Test-Path $credentials)) {
    New-Item -ItemType Directory -Path (Split-Path -Parent $credentials) -Force | Out-Null
    Set-Content -Path $credentials -Encoding ascii -Value @("[general]", 'email = ""')
}

# ---------------------------------------------------------------------------
# Start the application
# ---------------------------------------------------------------------------

function Get-FreePort {
    param([int] $Start)
    for ($candidate = $Start; $candidate -lt ($Start + 25); $candidate++) {
        $listener = $null
        try {
            $listener = New-Object System.Net.Sockets.TcpListener([Net.IPAddress]::Loopback, $candidate)
            $listener.Start()
            return $candidate
        } catch {
            continue
        } finally {
            if ($listener) { try { $listener.Stop() } catch { } }
        }
    }
    return $Start
}

$Port = Get-FreePort -Start $Port
$Url  = "http://localhost:$Port"

$env:PYTHONUTF8                              = "1"
$env:STREAMLIT_SERVER_PORT                   = "$Port"
$env:STREAMLIT_SERVER_HEADLESS               = "true"
# Reachable from this PC only. Legacy AKS lists are internal engineering data
# and must not be served to the rest of the network.
$env:STREAMLIT_SERVER_ADDRESS               = "localhost"
$env:STREAMLIT_BROWSER_GATHER_USAGE_STATS    = "false"
$env:STREAMLIT_GLOBAL_DEVELOPMENT_MODE       = "false"
if ($DebugErrors) { $env:STREAMLIT_CLIENT_SHOW_ERROR_DETAILS = "true" }

Write-Log "Starting the AKS Migration App on $Url" "STEP"

# The application folder usually contains spaces ("AKS Migration App"), so the
# script path has to reach Streamlit as a single quoted argument.
$streamlit = Start-Process -FilePath $Python `
    -ArgumentList @("-m", "streamlit", "run", "`"$AppScript`"",
                    "--server.port", "$Port", "--server.address", "localhost") `
    -WorkingDirectory $AppRoot -NoNewWindow -PassThru

$ready = $false
for ($attempt = 0; $attempt -lt 90; $attempt++) {
    if ($streamlit.HasExited) { break }
    try {
        $health = Invoke-WebRequest -Uri "$Url/_stcore/health" -UseBasicParsing -TimeoutSec 2
        if ($health.StatusCode -eq 200) { $ready = $true; break }
    } catch { Start-Sleep -Milliseconds 700 }
}

if (-not $ready) {
    if ($streamlit.HasExited) {
        Stop-WithMessage "The application stopped while starting up." @(
            "Run the launcher once more.",
            "If it keeps failing, start it with:  .\install_and_run.ps1 -Reinstall"
        )
    }
    Write-Log "The application is taking longer than usual to start. Open $Url in your browser." "WARN"
} else {
    Write-Log "Ready."
    if (-not $NoBrowser) { Start-Process $Url | Out-Null }
}

Write-Host ""
Write-Host "  The AKS Migration App is running at $Url" -ForegroundColor Green
Write-Host "  Leave this window open while you work."   -ForegroundColor DarkGray
Write-Host "  Close this window, or press Ctrl+C, to stop the application." -ForegroundColor DarkGray
Write-Host ""

try {
    $streamlit.WaitForExit()
} finally {
    if (-not $streamlit.HasExited) { try { $streamlit.Kill() } catch { } }
    Write-Log "The application was closed."
}
