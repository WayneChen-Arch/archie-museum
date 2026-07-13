$ErrorActionPreference = "Stop"

$git = "C:\Program Files\Git\bin\git.exe"
$gh = "C:\Program Files\GitHub CLI\gh.exe"
$repoName = "archie-museum"

Set-Location $PSScriptRoot

& $gh auth status *> $null
if ($LASTEXITCODE -ne 0) {
  Write-Host "请先登录 GitHub：" -ForegroundColor Yellow
  & $gh auth login -h github.com -p https -w
}

$branch = (& $git rev-parse --abbrev-ref HEAD).Trim()

Write-Host "创建并推送仓库 $repoName ..." -ForegroundColor Cyan
& $gh repo create $repoName --public --source . --remote origin --push

Write-Host "启用 GitHub Pages ..." -ForegroundColor Cyan
& $gh api repos/{owner}/$repoName/pages -X POST -f build_type=legacy -f "source[branch]=$branch" -f "source[path]=/"

$owner = (& $gh api user --jq .login).Trim()
$url = "https://$owner.github.io/$repoName/"

Write-Host ""
Write-Host "部署完成！" -ForegroundColor Green
Write-Host "手机访问地址：$url"
Write-Host "若页面暂时打不开，请等待 1-3 分钟后刷新。"
