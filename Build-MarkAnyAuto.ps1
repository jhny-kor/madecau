[CmdletBinding()]
param(
    [string]$Python = "python",
    [switch]$SkipInstaller,
    [switch]$SelfTest
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$script = Join-Path $root "markany_auto.py"
$icon = Join-Path $root "MarkAnyAuto.ico"
$venv = Join-Path $root ".build\venv"
$venvPython = Join-Path $venv "Scripts\python.exe"
$dist = Join-Path $root "dist"

if (-not (Get-Command $Python -ErrorAction SilentlyContinue)) {
    throw "Python을 찾지 못했습니다. python.org에서 3.10 이상을 설치하고 PATH에 추가하세요."
}

if ($SelfTest) {
    & $Python --version
    if (-not (Test-Path $script)) { throw "markany_auto.py 가 없습니다: $script" }
    & $Python $script --selftest
    Write-Output "self-test: ok"
    exit 0
}

if (-not (Test-Path $icon)) { throw "아이콘이 없습니다: $icon" }

if (-not (Test-Path $venvPython)) {
    Write-Output "빌드용 가상환경 생성: $venv"
    & $Python -m venv $venv
}

Write-Output "의존성 설치 (pywinauto, tkinterdnd2, pyinstaller)..."
& $venvPython -m pip install --upgrade pip | Out-Null
& $venvPython -m pip install pywinauto tkinterdnd2 pyinstaller
if ($LASTEXITCODE -ne 0) { throw "pip 설치에 실패했습니다." }

& $venvPython $script --selftest
if ($LASTEXITCODE -ne 0) { throw "self-test 실패. 빌드를 중단합니다." }

Write-Output "실행 파일 빌드..."
& $venvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name MarkAnyAuto `
    --icon $icon `
    --add-data "$icon;." `
    --distpath $dist `
    --workpath (Join-Path $root ".build\work") `
    --specpath (Join-Path $root ".build") `
    --collect-all pywinauto `
    --collect-all comtypes `
    --collect-all tkinterdnd2 `
    $script
if ($LASTEXITCODE -ne 0) { throw "PyInstaller 빌드에 실패했습니다." }

$exe = Join-Path $dist "MarkAnyAuto.exe"
if (-not (Test-Path $exe)) { throw "빌드 결과가 없습니다: $exe" }
Write-Output "실행 파일: $exe"

if ($SkipInstaller) { exit 0 }

$iscc = Get-ChildItem -Path @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
) -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $iscc) {
    Write-Output ""
    Write-Output "Inno Setup이 없어 설치 파일은 만들지 않았습니다."
    Write-Output "MarkAnyAuto.exe 자체가 단일 실행 파일이라 옮겨서 바로 실행할 수 있습니다."
    Write-Output "설치 형태가 필요하면 https://jrsoftware.org/isdl.php 에서 Inno Setup 6을 설치한 뒤 다시 실행하세요."
    exit 0
}

# .iss 는 저장소에 그대로 둔다. PowerShell 문자열로 만들어 다시 쓰면
# 인코딩을 한 번 더 거치면서 한글이 깨지고 지시자가 통째로 사라진다.
$issPath = Join-Path $root "MarkAnyAuto.iss"
if (-not (Test-Path $issPath)) { throw "설치 스크립트가 없습니다: $issPath" }

Write-Output "설치 파일 빌드..."
& $iscc.FullName "/DExePath=$exe" "/O$dist" $issPath
if ($LASTEXITCODE -ne 0) { throw "Inno Setup 빌드에 실패했습니다." }

Write-Output ""
Write-Output "설치 파일: $(Join-Path $dist 'MarkAnyAuto-Setup.exe')"
Write-Output "이 파일 하나만 옮겨서 설치하면 됩니다."
