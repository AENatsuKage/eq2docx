<#
eq2docx.ps1 —— 剪贴板上的 markdown 行间公式 -> 单公式 docx（走 pandoc）

用法（在 cmd 里，先用 -Init 生成 eq2docx.cmd 与 PATH 入口）：
    eq2docx -o EQ01                         读剪贴板 -> EQ01.md + EQ01.docx
    eq2docx -o EQ01 -OutDir D:\work          指定输出目录（默认脚本所在目录）
    eq2docx -o EQ01 -InputFile a.md          不从剪贴板，从文件读
    eq2docx -Init                            自动探测 pandoc，写配置 + 生成 eq2docx.cmd + 加 PATH
    eq2docx -Init -PandocPath "D:\anaconda3\Library\bin\pandoc.exe"
    eq2docx -Check                           只看当前解析到的 pandoc 与版本

pandoc 解析顺序：-PandocPath > eq2docx.config.json > 环境变量 EQ2DOCX_PANDOC / PANDOC
                > PATH 里的 pandoc > 常见安装位置

退出码：0 成功 / 1 Init 失败 / 2 找不到 pandoc / 3 输入为空 / 4 pandoc 转换失败
#>
[CmdletBinding()]
param(
    [Alias('Out')][string]$o,
    [string]$PandocPath,
    [string]$OutDir,
    [string]$InputFile,
    [switch]$Init,
    [switch]$Check
)

$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$configPath = Join-Path $scriptDir 'eq2docx.config.json'
$shimPath = Join-Path $scriptDir 'eq2docx.cmd'

function Get-ExePath {
    param([string]$Candidate)
    if (-not $Candidate) { return $null }
    $c = $Candidate.Trim().Trim('"')
    if ([IO.Directory]::Exists($c)) { $c = Join-Path $c 'pandoc.exe' }
    if ([IO.File]::Exists($c)) { return (Resolve-Path -LiteralPath $c).Path }
    return $null
}

function Resolve-Pandoc {
    param([string]$Explicit)
    # 显式指定优先，且必须有效（不许静默回退）
    if ($Explicit) {
        $hit = Get-ExePath $Explicit
        if (-not $hit) {
            Write-Host "× -PandocPath 不是有效的 pandoc.exe：$Explicit" -ForegroundColor Red
            exit 2
        }
        return $hit
    }
    $cands = New-Object System.Collections.Generic.List[string]
    # 运行时覆盖用专用变量，优先于配置文件
    if ($env:EQ2DOCX_PANDOC) { $cands.Add($env:EQ2DOCX_PANDOC) }
    if (Test-Path -LiteralPath $configPath) {
        try {
            $cfg = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
            if ($cfg.pandoc) { $cands.Add($cfg.pandoc) }
        } catch { }
    }
    if ($env:PANDOC) { $cands.Add($env:PANDOC) }
    $cmd = Get-Command pandoc -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source) { $cands.Add($cmd.Source) }
    $cands.Add('D:\anaconda3\Library\bin\pandoc.exe')
    if ($env:LOCALAPPDATA) { $cands.Add((Join-Path $env:LOCALAPPDATA 'Pandoc\pandoc.exe')) }
    $cands.Add('C:\Program Files\Pandoc\pandoc.exe')
    if ($env:ProgramFiles) { $cands.Add((Join-Path $env:ProgramFiles 'Pandoc\pandoc.exe')) }
    foreach ($c in $cands) {
        $hit = Get-ExePath $c
        if ($hit) { return $hit }
    }
    return $null
}

function Add-DirToUserPath {
    param([string]$Dir)
    $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Environment', $true)
    if (-not $key) { return $false }
    try {
        $cur = $key.GetValue('Path', '', [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        $kind = 'ExpandString'
        try { $kind = $key.GetValueKind('Path') } catch { }
        if (($cur.TrimEnd(';') -split ';') -contains $Dir) { return $true }
        $new = if ($cur.TrimEnd(';')) { $cur.TrimEnd(';') + ';' + $Dir } else { $Dir }
        $key.SetValue('Path', $new, $kind)
    } finally { $key.Close() }
    # 广播 WM_SETTINGCHANGE，让新开的窗口立即生效
    try {
        Add-Type -Namespace EQ2DOCX -Name Native -MemberDefinition @'
[DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);
'@ -ErrorAction SilentlyContinue
        $r = [UIntPtr]::Zero
        [void][EQ2DOCX.Native]::SendMessageTimeout([IntPtr]0xffff, 0x1A, [UIntPtr]::Zero, 'Environment', 2, 5000, [ref]$r)
    } catch { }
    return $true
}

# ---------------- Check ----------------
if ($Check) {
    $p = Resolve-Pandoc $PandocPath
    if (-not $p) {
        Write-Host '× 没找到 pandoc.exe。跑一次  eq2docx -Init -PandocPath "<pandoc.exe 的完整路径>"' -ForegroundColor Red
        exit 2
    }
    $v = (& $p --version | Select-Object -First 1)
    Write-Host "pandoc : $p"
    Write-Host "版本   : $v"
    Write-Host "配置   : $(if (Test-Path -LiteralPath $configPath) { $configPath } else { '(无 eq2docx.config.json)' })"
    Write-Host "输出目录: $scriptDir"
    exit 0
}

# ---------------- Init ----------------
if ($Init) {
    $p = Resolve-Pandoc $PandocPath
    if (-not $p) {
        Write-Host '× 没探测到 pandoc.exe。' -ForegroundColor Red
        Write-Host '  显式指定：  eq2docx -Init -PandocPath "D:\xxx\pandoc.exe"' -ForegroundColor Yellow
        Write-Host '  或先装官方 MSI，或用 winget install JohnMacFarlane.Pandoc' -ForegroundColor Yellow
        exit 1
    }
    $v = (& $p --version | Select-Object -First 1)
    @{ pandoc = $p; version = $v; saved = (Get-Date).ToString('s') } |
        ConvertTo-Json | Set-Content -LiteralPath $configPath -Encoding UTF8

    $shim = "@echo off`r`nsetlocal`r`nset `"PS=pwsh`"`r`nwhere pwsh >nul 2>nul || set `"PS=powershell`"`r`n`"%PS%`" -NoLogo -NoProfile -ExecutionPolicy Bypass -File `"%~dp0eq2docx.ps1`" %*`r`nexit /b %ERRORLEVEL%`r`n"
    [IO.File]::WriteAllText($shimPath, $shim, (New-Object System.Text.ASCIIEncoding))

    $added = Add-DirToUserPath -Dir $scriptDir
    Write-Host '√ 初始化完成' -ForegroundColor Green
    Write-Host "  pandoc     : $p"
    Write-Host "  版本       : $v"
    Write-Host "  配置文件   : $configPath"
    Write-Host "  cmd 入口   : $shimPath"
    if ($added) {
        Write-Host "  已加入用户 PATH: $scriptDir（新开的 cmd 窗口生效）" -ForegroundColor Green
    } else {
        Write-Host "  × 写 PATH 失败，请手动把 $scriptDir 加入 PATH" -ForegroundColor Yellow
    }
    Write-Host ''
    Write-Host '下一步：复制公式到剪贴板，然后  eq2docx -o EQ01' -ForegroundColor Cyan
    exit 0
}

# ---------------- 主流程 ----------------
$pandoc = Resolve-Pandoc $PandocPath
if (-not $pandoc) {
    Write-Host '× 没找到 pandoc.exe：先跑  eq2docx -Init -PandocPath "<路径>"' -ForegroundColor Red
    exit 2
}

$base = if ($o) { $o } else { 'EQ_' + (Get-Date -Format 'yyyyMMdd_HHmmss') }
$base = $base -replace '\.(md|docx)$', ''
if ($base -match '[\\/:*?"<>|]') {
    Write-Host "× 输出名里不能有这些字符：$base" -ForegroundColor Red
    exit 3
}

$outDir = if ($OutDir) { $OutDir } else { $scriptDir }
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

$mdPath = Join-Path $outDir ($base + '.md')
$docxPath = Join-Path $outDir ($base + '.docx')

if ($InputFile) {
    if (-not (Test-Path -LiteralPath $InputFile)) {
        Write-Host "× 找不到输入文件：$InputFile" -ForegroundColor Red
        exit 3
    }
    $text = Get-Content -LiteralPath $InputFile -Raw -Encoding UTF8
    $source = "文件 $InputFile"
} else {
    try { $text = Get-Clipboard -Raw } catch { $text = $null }
    $source = '剪贴板'
}

if (-not $text -or -not $text.Trim()) {
    Write-Host "× $source 里没有内容。" -ForegroundColor Red
    exit 3
}

$text = $text.Trim() + "`r`n"
if ($text -notmatch '\$') {
    Write-Warning "内容里没看到 `$ 分隔符：pandoc 会把它当普通文字，导出的 docx 里不会有公式。"
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText($mdPath, $text, $utf8NoBom)
Write-Host "· md   : $mdPath"

$pandocOut = & $pandoc $mdPath -o $docxPath 2>&1
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $docxPath)) {
    Write-Host "× pandoc 转换失败（退出码 $LASTEXITCODE）" -ForegroundColor Red
    if ($pandocOut) { $pandocOut | ForEach-Object { Write-Host "  $_" } }
    exit 4
}
if ($pandocOut) { $pandocOut | ForEach-Object { Write-Host "  [pandoc] $_" -ForegroundColor DarkYellow } }
Write-Host "· docx : $docxPath"

# 校验 docx 里真的有 OMML 公式（不是被降级成纯文本）
$hasMath = $null
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $zip = [IO.Compression.ZipFile]::OpenRead($docxPath)
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' } | Select-Object -First 1
        if ($entry) {
            $sr = New-Object IO.StreamReader($entry.Open())
            $xml = $sr.ReadToEnd()
            $sr.Close()
            $hasMath = ($xml -match '<m:oMath')
        }
    } finally { $zip.Dispose() }
} catch { $hasMath = $null }

if ($hasMath -eq $true) {
    Write-Host '√ 已生成 Office 公式（OMML），打开 docx 即可复制那个行间公式' -ForegroundColor Green
} elseif ($hasMath -eq $false) {
    Write-Host '! docx 里没检测到 OMML 公式：检查代码是否被 $ ... $ 包住' -ForegroundColor Yellow
} else {
    Write-Host '· （跳过 OMML 校验）' -ForegroundColor DarkGray
}
exit 0
