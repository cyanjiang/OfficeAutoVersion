# 打包安装程序：把三个加载项的产物做成 payload.zip，再编译成单文件 Setup.exe
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File build-setup.ps1
#   powershell -ExecutionPolicy Bypass -File build-setup.ps1 -Configuration Release
#
# 前提：先在 Visual Studio 里成功生成 OfficeAutoVersion.sln（脚本只打包 bin 里的产物）

param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
$apps = @("Word", "Excel", "PowerPoint")
$shared = @(
    "OfficeAutoVersion.Common.dll",
    "OfficeAutoVersion.Core.dll",
    "Microsoft.Office.Tools.Common.v4.0.Utilities.dll"
)

Write-Host "== 1/4 收集构建产物（$Configuration）==" -ForegroundColor Cyan
$payloadDir = Join-Path $here "payload"
$stage = Join-Path $payloadDir "bin"
if (Test-Path $payloadDir) { Remove-Item $payloadDir -Recurse -Force }
New-Item -ItemType Directory -Path $stage -Force | Out-Null

$missing = @()
foreach ($app in $apps) {
    $bin = Join-Path $root "src\OfficeAutoVersion.$app\bin\$Configuration"
    foreach ($name in @("OfficeAutoVersion.$app.dll", "OfficeAutoVersion.$app.dll.manifest", "OfficeAutoVersion.$app.vsto")) {
        $source = Join-Path $bin $name
        if (-not (Test-Path $source)) { $missing += $source; continue }
        Copy-Item $source $stage -Force
    }
}
foreach ($name in $shared) {
    $source = Join-Path $root "src\OfficeAutoVersion.Word\bin\$Configuration\$name"
    if (-not (Test-Path $source)) { $missing += $source; continue }
    Copy-Item $source $stage -Force
}

if ($missing.Count -gt 0) {
    Write-Host "缺少构建产物，请先在 Visual Studio 里生成 OfficeAutoVersion.sln：" -ForegroundColor Red
    foreach ($item in $missing) { Write-Host "  $item" -ForegroundColor Red }
    exit 1
}

Write-Host "== 2/4 打包 payload.zip ==" -ForegroundColor Cyan
# 使用说明随包安装到 <安装目录>\docs\，也一并放到 dist 方便单独阅读
$docSource = Join-Path $root "docs\Office自动版本-使用说明.docx"
if (Test-Path $docSource) {
    $docStage = Join-Path $payloadDir "docs"
    New-Item -ItemType Directory -Path $docStage -Force | Out-Null
    Copy-Item $docSource $docStage -Force
    Write-Host "   已加入使用说明：docs\Office自动版本-使用说明.docx" -ForegroundColor DarkGray
}
else {
    Write-Host "   [警告] 未找到 docs\Office自动版本-使用说明.docx，安装包将不含使用说明" -ForegroundColor Yellow
}

$payload = Join-Path $here "payload.zip"
if (Test-Path $payload) { Remove-Item $payload -Force }
$zipInputs = @(Join-Path $payloadDir "bin")
if (Test-Path (Join-Path $payloadDir "docs")) { $zipInputs += (Join-Path $payloadDir "docs") }
Compress-Archive -Path $zipInputs -DestinationPath $payload -CompressionLevel Optimal
$payloadKb = [math]::Round((Get-Item $payload).Length / 1KB, 1)
Write-Host "   payload.zip  $payloadKb KB"

if (-not $SkipBuild) {
    Write-Host "== 3/4 编译安装程序 ==" -ForegroundColor Cyan
    & dotnet build (Join-Path $here "Setup.csproj") -c Release -v:q -nologo
    if ($LASTEXITCODE -ne 0) {
        Write-Host "编译安装程序失败。" -ForegroundColor Red
        exit $LASTEXITCODE
    }
}
else {
    Write-Host "== 3/4 跳过编译（-SkipBuild）==" -ForegroundColor Yellow
}

$exe = Join-Path $here "bin\Release\net48\OfficeAutoVersionSetup.exe"
if (-not (Test-Path $exe)) {
    Write-Host "找不到安装程序输出：$exe" -ForegroundColor Red
    exit 1
}

$version = "1.0.0"
[xml]$project = Get-Content (Join-Path $here "Setup.csproj")
foreach ($group in $project.Project.PropertyGroup) {
    if ($group.Version) { $version = $group.Version; break }
}

Write-Host "== 4/4 输出到 dist ==" -ForegroundColor Cyan
$dist = Join-Path $root "dist"
New-Item -ItemType Directory -Path $dist -Force | Out-Null
Copy-Item $exe (Join-Path $dist "OfficeAutoVersionSetup.exe") -Force
Copy-Item $exe (Join-Path $dist "OfficeAutoVersionSetup-$version.exe") -Force
if (Test-Path $docSource) {
    Copy-Item $docSource (Join-Path $dist "Office自动版本-使用说明.docx") -Force
}

$exeKb = [math]::Round((Get-Item $exe).Length / 1KB, 1)
Write-Host ""
Write-Host "完成：$dist\OfficeAutoVersionSetup-$version.exe（$exeKb KB）" -ForegroundColor Green
Write-Host "安装：双击运行，或静默安装 OfficeAutoVersionSetup.exe /S"
Write-Host "卸载：OfficeAutoVersionSetup.exe /uninstall（也可在“设置 → 应用”里卸载）"
