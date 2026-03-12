$ErrorActionPreference = "Stop"
$RepoDir = $PSScriptRoot
$PublishDir = "$RepoDir\publish"
$InstallDir = "$env:LOCALAPPDATA\access-cli"

Write-Host "=== access-cli install ===" -ForegroundColor Cyan
Write-Host "Building (framework-dependent)..." -ForegroundColor Yellow

Push-Location "$RepoDir\src\AccessCli"
dotnet publish -c Release -r win-x64 --no-self-contained -o $PublishDir --source "D:\nuget-local"
if ($LASTEXITCODE -ne 0) { Pop-Location; exit 1 }
Pop-Location

if (-not (Test-Path $InstallDir)) { New-Item -ItemType Directory -Path $InstallDir | Out-Null }
Copy-Item "$PublishDir\*" "$InstallDir\" -Force -Recurse

$currentPath = [Environment]::GetEnvironmentVariable("PATH", "User")
if ($currentPath -notlike "*$InstallDir*") {
    [Environment]::SetEnvironmentVariable("PATH", "$InstallDir;$currentPath", "User")
    Write-Host "Added to PATH: $InstallDir" -ForegroundColor Green
    Write-Host "(Restart terminal to take effect)" -ForegroundColor Gray
} else {
    Write-Host "PATH already set." -ForegroundColor Gray
}

Write-Host "Done! Run: access-cli --help" -ForegroundColor Cyan
