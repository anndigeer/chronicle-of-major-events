# 用于本地推送项目文件；不保存任何访问令牌。
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw '请先安装 Git for Windows' }
$expected = 'https://github.com/anndigeer/chronicle-of-major-events.git'
if (-not (Test-Path '.git')) { git init -b main }
$origin = git remote get-url origin 2>$null
if ($LASTEXITCODE -ne 0) { git remote add origin $expected }
elseif ($origin -ne $expected) { throw "远程仓库不匹配: $origin" }
$paths = @('.github/workflows/python-ci.yml','.gitignore','README.md','README_使用说明.md','jrjj_v2.py','jrjj_config.json','requirements_v2.txt','test_jrjj_v2.py','scripts/live_smoke.py','一键采集生成.bat','首次安装依赖.bat','推送至GitHub.ps1')
git add -- $paths
if ($LASTEXITCODE -ne 0) { throw 'git add 失败' }
git diff --cached --quiet
if ($LASTEXITCODE -eq 1) {
 git commit -m 'Update dual-source collector and CI tests'
 if ($LASTEXITCODE -ne 0) { throw 'git commit 失败' }
}
git branch -M main
git push -u origin main
if ($LASTEXITCODE -ne 0) { throw 'git push 失败，请检查权限及远程分支变化' }
Write-Host '推送成功，请检查 GitHub Actions。'
