$gh = "C:\Program Files\GitHub CLI\gh.exe"

if (-not (Test-Path $gh)) {
  Write-Host "GitHub CLI is not installed. Install it with:"
  Write-Host "winget install --id GitHub.cli -e"
  exit 1
}

& $gh auth login -h github.com -p https -w

Write-Host ""
& $gh auth status
