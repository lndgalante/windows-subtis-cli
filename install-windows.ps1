<#
Installs Subtis CLI on Windows by downloading the release zip, extracting subtis.exe,
copying it to %LOCALAPPDATA%\Programs\Subtis (default), and updating PATH for the current user.

Usage (one-liner example):
  powershell -ExecutionPolicy Bypass -c "irm https://raw.githubusercontent.com/lndgalante/windows-subtis-cli/main/scripts/install-windows.ps1 | iex"

Flags:
  -Version <X.Y.Z>     Install a specific release tag (vX.Y.Z). Defaults to latest if omitted.
  -DownloadUrl <url>   Skip GitHub API lookup and use this zip URL directly.
  -RepoOwner <owner>   Override repo owner (default: lndgalante).
  -RepoName <name>     Override repo name (default: windows-subtis-cli).
  -TargetDir <path>    Install destination (default: %LOCALAPPDATA%\Programs\Subtis).
  -Sha256 <hash>       Optional checksum for the zip; install aborts if mismatched.
  -NoPathUpdate        Skip adding TargetDir to user PATH.
#>
[CmdletBinding()]
param(
  [string]$Version,
  [string]$DownloadUrl,
  [string]$RepoOwner = "lndgalante",
  [string]$RepoName = "windows-subtis-cli",
  [string]$TargetDir = "$env:LOCALAPPDATA\Programs\Subtis",
  [string]$Sha256,
  [switch]$NoPathUpdate
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Write-Info([string]$Message) {
  Write-Host $Message -ForegroundColor Cyan
}

function Fail([string]$Message) {
  Write-Host $Message -ForegroundColor Red
  exit 1
}

function Resolve-AssetUrl {
  if ($DownloadUrl) {
    return $DownloadUrl
  }

  $apiUrl = if ($Version) {
    "https://api.github.com/repos/$RepoOwner/$RepoName/releases/tags/v$Version"
  } else {
    "https://api.github.com/repos/$RepoOwner/$RepoName/releases/latest"
  }

  Write-Info "Fetching release metadata from $apiUrl"
  $release = Invoke-RestMethod -Uri $apiUrl -Headers @{ "User-Agent" = "subtis-installer" }
  $asset = $release.assets | Where-Object { $_.name -eq "subtis-windows-x64.zip" } | Select-Object -First 1
  if (-not $asset) {
    Fail "subtis-windows-x64.zip not found in the release assets."
  }

  if (-not $Version -and $release.tag_name) {
    $Version = $release.tag_name.TrimStart("v")
  }

  return $asset.browser_download_url
}

try {
  $assetUrl = Resolve-AssetUrl

  $tempRoot = Join-Path $env:TEMP "subtis-install-$([System.Guid]::NewGuid().ToString("N"))"
  $zipPath = Join-Path $tempRoot "subtis-windows-x64.zip"
  $extractDir = Join-Path $tempRoot "unzipped"

  New-Item -ItemType Directory -Path $tempRoot | Out-Null

  Write-Info "Downloading $assetUrl"
  Invoke-WebRequest -Uri $assetUrl -OutFile $zipPath -UseBasicParsing

  if ($Sha256) {
    $computed = (Get-FileHash -Path $zipPath -Algorithm SHA256).Hash.ToLower()
    if ($computed -ne $Sha256.ToLower()) {
      Fail "Checksum mismatch. Expected $Sha256 but got $computed"
    }
  }

  Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
  $exePath = Join-Path $extractDir "subtis.exe"
  if (-not (Test-Path $exePath)) {
    Fail "subtis.exe not found in the archive."
  }

  New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
  Copy-Item -Path $exePath -Destination (Join-Path $TargetDir "subtis.exe") -Force

  $pathUpdated = $false
  if (-not $NoPathUpdate) {
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $parts = $userPath -split ";" | Where-Object { $_ }
    if ($parts -notcontains $TargetDir) {
      $newPath = ($parts + $TargetDir) -join ";"
      [Environment]::SetEnvironmentVariable("Path", $newPath, "User")
      $pathUpdated = $true
    }
  }

  Write-Info "Installed to $TargetDir"
  if ($pathUpdated) {
    Write-Info "PATH updated for current user. Open a new terminal to pick it up."
  } elseif (-not $NoPathUpdate) {
    Write-Info "PATH already contained $TargetDir"
  } else {
    Write-Info "PATH update skipped (NoPathUpdate set)."
  }

  if ($Version) {
    Write-Info "Installed version: $Version"
  }
}
catch {
  Fail "Install failed: $_"
}
finally {
  if ($tempRoot -and (Test-Path $tempRoot)) {
    Remove-Item $tempRoot -Recurse -Force
  }
}
