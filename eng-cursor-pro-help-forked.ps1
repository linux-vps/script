# Auto-elevate to admin rights if not already running as admin
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Requesting administrator privileges..."
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -ExecutionFromElevated"
    Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs
    Exit
}

# Set TLS to 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Output colors
$Red = "`e[31m"
$Green = "`e[32m"
$Blue = "`e[36m"
$Yellow = "`e[33m"
$Reset = "`e[0m"

# Messages
$Messages = @{
    "START_INSTALL"           = "Starting installation...";
    "ARCH_DETECTED"           = "Detected system architecture: ";
    "ARCH_UNSUPPORTED"        = "Only 64-bit Windows is supported.";
    "LATEST_VERSION"          = "Latest version: ";
    "CREATE_INSTALL_DIR"      = "Creating installation directory...";
    "DOWNLOAD_FROM"           = "Downloading the latest release from: ";
    "DOWNLOAD_FAILED"         = "Failed to download the binary.";
    "FILE_NOT_FOUND"          = "Downloaded file not found.";
    "INSTALL_BINARY"          = "Installing the binary...";
    "INSTALL_FAILED"          = "Failed to install the binary.";
    "ADD_TO_PATH"             = "Adding the installation directory to PATH...";
    "CLEANUP"                 = "Cleaning up temporary files...";
    "INSTALL_COMPLETE"        = "Installation completed successfully!";
    "USAGE_HINT"              = "You can now use 'cursor-id-modifier' directly.";
    "CHECK_RUNNING_PROCESSES" = "Checking for running Cursor processes...";
    "FOUND_RUNNING_PROCESSES" = "Found running Cursor processes. Attempting to close them...";
    "CLOSE_SUCCESS"           = "Successfully closed all Cursor processes.";
    "CLOSE_FAILED"            = "Failed to close Cursor processes. Please close them manually.";
    "BACKUP_STORAGE"          = "Backing up storage.json...";
    "BACKUP_SUCCESS"          = "Backup created at: "
}

# Functions for colored output
function Write-Status($Message) {
    Write-Host "${Blue}[*]${Reset} $Message"
}

function Write-Success($Message) {
    Write-Host "${Green}[✓]${Reset} $Message"
}

function Write-Warning($Message) {
    Write-Host "${Yellow}[!]${Reset} $Message"
}

function Write-ErrorExit($Message) {
    Write-Host "${Red}[✗]${Reset} $Message"
    Exit 1
}

# Close Cursor instances
function Close-CursorInstances {
    Write-Status $Messages["CHECK_RUNNING_PROCESSES"]
    $cursorProcesses = Get-Process "Cursor" -ErrorAction SilentlyContinue

    if ($cursorProcesses) {
        Write-Status $Messages["FOUND_RUNNING_PROCESSES"]
        try {
            $cursorProcesses | ForEach-Object { $_.CloseMainWindow() | Out-Null }
            Start-Sleep -Seconds 2
            $cursorProcesses | Where-Object { !$_.HasExited } | Stop-Process -Force
            Write-Success $Messages["CLOSE_SUCCESS"]
        } catch {
            Write-ErrorExit $Messages["CLOSE_FAILED"]
        }
    }
}

# Backup storage.json
function Backup-StorageJson {
    Write-Status $Messages["BACKUP_STORAGE"]
    $storageJsonPath = "$env:APPDATA\Cursor\User\globalStorage\storage.json"
    if (Test-Path $storageJsonPath) {
        $backupPath = "$storageJsonPath.backup"
        Copy-Item -Path $storageJsonPath -Destination $backupPath -Force
        Write-Success "$($Messages["BACKUP_SUCCESS"]) $backupPath"
    }
}

# Get latest release version from GitHub
function Get-LatestVersion {
    $repo = "yuaotian/go-cursor-help"
    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/releases/latest"
    return $release.tag_name
}

# Download file with progress
function Download-File {
    param(
        [string]$Url,
        [string]$OutFile
    )
    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell Script")
        $webClient.DownloadFileAsync($Url, $OutFile)

        while ($webClient.IsBusy) {
            Write-Progress -Activity "Downloading binary..." -Status "In progress..."
            Start-Sleep -Milliseconds 100
        }

        Write-Progress -Activity "Downloading binary..." -Completed
        return $true
    } catch {
        Write-ErrorExit "$($Messages["DOWNLOAD_FAILED"]): $_"
    } finally {
        $webClient.Dispose()
    }
}

# Main installation process
Write-Status $Messages["START_INSTALL"]

# Close running Cursor instances
Close-CursorInstances

# Backup storage.json
Backup-StorageJson

# Check system architecture
$arch = if ([Environment]::Is64BitOperatingSystem) { "amd64" } else { "386" }
Write-Status "$($Messages["ARCH_DETECTED"]) $arch"
if ($arch -ne "amd64") {
    Write-ErrorExit $Messages["ARCH_UNSUPPORTED"]
}

# Get the latest version
$version = Get-LatestVersion
Write-Status "$($Messages["LATEST_VERSION"]) $version"

# Setup installation paths
$installDir = "$env:ProgramFiles\cursor-id-modifier"
$binaryName = "cursor_id_modifier_${version.TrimStart('v')}_windows_amd64.exe"
$downloadUrl = "https://github.com/yuaotian/go-cursor-help/releases/download/$version/$binaryName"
$tempFile = "$env:TEMP\$binaryName"

# Create installation directory
Write-Status $Messages["CREATE_INSTALL_DIR"]
if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
}

# Download the binary
Write-Status "$($Messages["DOWNLOAD_FROM"]) $downloadUrl"
Download-File -Url $downloadUrl -OutFile $tempFile

if (-not (Test-Path $tempFile)) {
    Write-ErrorExit $Messages["FILE_NOT_FOUND"]
}

# Install the binary
Write-Status $Messages["INSTALL_BINARY"]
try {
    Move-Item -Force $tempFile "$installDir\cursor-id-modifier.exe"
} catch {
    Write-ErrorExit "$($Messages["INSTALL_FAILED"]): $_"
}

# Add to PATH
$userPath = [Environment]::GetEnvironmentVariable("Path", "User")
if ($userPath -notlike "*$installDir*") {
    Write-Status $Messages["ADD_TO_PATH"]
    [Environment]::SetEnvironmentVariable("Path", "$userPath;$installDir", "User")
}

# Cleanup temporary files
Write-Status $Messages["CLEANUP"]
if (Test-Path $tempFile) {
    Remove-Item -Force $tempFile
}

Write-Success $Messages["INSTALL_COMPLETE"]
Write-Success $Messages["USAGE_HINT"]
