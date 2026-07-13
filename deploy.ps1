$ErrorActionPreference = "Stop"

$git = "C:\Program Files\Git\bin\git.exe"
$gh = "C:\Program Files\GitHub CLI\gh.exe"
$repoName = "archie-museum"

Set-Location $PSScriptRoot

& $gh auth status *> $null
if ($LASTEXITCODE -ne 0) {
  Write-Host "Please sign in to GitHub first."
  & $gh auth login -h github.com -p https -w
}

$branch = (& $git rev-parse --abbrev-ref HEAD).Trim()
$owner = (& $gh api user --jq .login).Trim()

$remoteExists = & $git remote get-url origin 2>$null
if (-not $remoteExists) {
  Write-Host "Creating repository $repoName ..."
  & $gh repo create $repoName --public --source . --remote origin --push
} else {
  Write-Host "Pushing to existing remote ..."
  & $git push -u origin $branch
}

Write-Host "Enabling GitHub Pages ..."
& $gh api "repos/$owner/$repoName/pages" -X POST -f build_type=legacy -f "source[branch]=$branch" -f "source[path]=/" 2>$null
if ($LASTEXITCODE -ne 0) {
  & $gh api "repos/$owner/$repoName/pages" -X PUT -f build_type=legacy -f "source[branch]=$branch" -f "source[path]=/"
}

$url = "https://$owner.github.io/$repoName/"
Write-Host ""
Write-Host "Deployment complete."
Write-Host "Site URL: $url"
Write-Host "It may take 1-3 minutes before the site is live."
