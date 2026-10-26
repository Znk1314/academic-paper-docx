<#
.SYNOPSIS
    把 academic-paper-docx 安装到 agent 的 skills 目录。

.DESCRIPTION
    skill 以目录形式生效（SKILL.md + scripts/），所以安装 = 把本仓库复制/链接到
    <skills 根目录>/academic-paper-docx。

.PARAMETER Target
    dsh / codex / claude / project / custom

.PARAMETER Path
    Target=custom 时的目标目录（会在此目录下建 academic-paper-docx）。

.PARAMETER Link
    生成目录符号链接而不是复制（改代码后立即生效；Windows 需开发者模式或管理员）。

.EXAMPLE
    pwsh -File .\install.ps1
    pwsh -File .\install.ps1 -Target codex -Force
    pwsh -File .\install.ps1 -Target custom -Path D:\my-agent\skills
#>
[CmdletBinding()]
param(
    [ValidateSet("dsh", "codex", "claude", "project", "custom")]
    [string]$Target = "dsh",
    [string]$Path,
    [switch]$Link,
    [switch]$Force,
    [switch]$SkipDeps
)

$ErrorActionPreference = "Stop"

$Name = "academic-paper-docx"
$Source = $PSScriptRoot

if (-not (Test-Path (Join-Path $Source "SKILL.md"))) {
    throw "install.ps1 必须放在仓库根目录（找不到 SKILL.md）"
}

function Resolve-SkillsRoot {
    param([string]$Kind, [string]$Explicit)

    switch ($Kind) {
        "dsh" { return (Join-Path $env:USERPROFILE ".dsh\skills") }
        "codex" { return (Join-Path $env:USERPROFILE ".codex\skills") }
        "claude" { return (Join-Path $env:USERPROFILE ".claude\skills") }
        "project" { return (Join-Path (Get-Location) ".dsh\skills") }
        "custom" {
            if (-not $Explicit) { throw "Target=custom 时必须提供 -Path" }
            return $Explicit
        }
    }
}

$root = Resolve-SkillsRoot -Kind $Target -Explicit $Path
$dest = Join-Path $root $Name

Write-Host "源目录：$Source"
Write-Host "目标：  $dest"

if (Test-Path $dest) {
    if (-not $Force) {
        throw "目标已存在。加 -Force 覆盖，或手动删除：$dest"
    }
    Remove-Item -Recurse -Force $dest
}

New-Item -ItemType Directory -Force -Path $root | Out-Null

if ($Link) {
    $isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    try {
        New-Item -ItemType SymbolicLink -Path $dest -Target $Source | Out-Null
        Write-Host "已创建符号链接（改仓库代码立即生效）。"
    }
    catch {
        if (-not $isAdmin) {
            Write-Host "创建符号链接失败（需要开发者模式或管理员）。改用复制。" -ForegroundColor Yellow
            Copy-Item -Recurse -Force $Source $dest
        }
        else { throw }
    }
}
else {
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    # /XD 用裸目录名 = 任意层级同名目录都排除（含 examples/demo/build 这类生成物）
    $excludes = @(".git", ".github", "__pycache__", ".pytest_cache", ".ruff_cache",
        ".idea", ".vscode", "build", "dist", "node_modules")
    robocopy $Source $dest /E /NFL /NDL /NJH /NJS /NP /XD $excludes /XF "*.pyc" | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy 失败（exit $LASTEXITCODE）" }
    # robocopy 成功码是 0-7
    $global:LASTEXITCODE = 0
    Write-Host "已复制。"
}

Write-Host ""
if (-not $SkipDeps) {
    Write-Host "安装依赖（python-docx / lxml / matplotlib / numpy / Pillow）…"
    $req = Join-Path $dest "requirements.txt"
    try {
        python -m pip install --quiet -r $req
        Write-Host "依赖已安装。" -ForegroundColor Green
    }
    catch {
        Write-Host "自动安装依赖失败，请手动执行：" -ForegroundColor Yellow
        Write-Host "  python -m pip install -r `"$req`""
    }
}

Write-Host ""
Write-Host "下一步："
Write-Host "  1) 自检：     python `"$dest\scripts\apd.py`" doctor"
Write-Host "  2) 生成示例： python `"$dest\scripts\apd.py`" demo -o build\demo"
Write-Host ""
Write-Host "只排文字不画图？ matplotlib 可以不装，docx 仍能生成。"
Write-Host "重启 agent（或新开会话）后技能即可被发现。"
