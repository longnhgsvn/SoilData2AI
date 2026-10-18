# Build EXE cho GUI Xlsx2CSV (PySide6 + PyInstaller)
# Mặc định KHÔNG chạy pip (tránh treo mạng/PATH). Muốn cài phụ thuộc:
#   & .\scripts\build_exe.ps1 -InstallDeps
param(
    [switch]$InstallDeps
)
$ErrorActionPreference = "Stop"

$script_dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $script_dir
Set-Location $root

# Dùng đường dẫn tuyệt đối, không phụ thuộc PATH hệ thống.
$py = "C:\Python312\python.exe"
if (-not (Test-Path $py)) {
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd) { $py = $cmd.Source } else { throw "Không tìm thấy Python. Sửa biến `$py trong build_exe.ps1." }
}
$spec = Join-Path $root "Xlsx2CSV_Converter.spec"
$dist = Join-Path $root "dist"
$work = Join-Path $root "build"
$exe  = Join-Path $dist "Xlsx2CSV_Converter.exe"

Write-Host "== Python: $py" -ForegroundColor Cyan
& $py --version

if ($InstallDeps) {
    Write-Host "== Cài phụ thuộc" -ForegroundColor Cyan
    & $py -m pip install -r (Join-Path $root "requirements.txt")
    if ($LASTEXITCODE -ne 0) { throw "pip install thất bại." }
} else {
    Write-Host "== Bỏ qua pip (dùng thư viện đã cài). Thêm -InstallDeps nếu cần." -ForegroundColor DarkGray
}

Write-Host "== Tạo icon" -ForegroundColor Cyan
& $py (Join-Path $root "scripts\make_icon.py")
if ($LASTEXITCODE -ne 0) { throw "Tạo icon thất bại." }

Write-Host "== Chạy test lõi" -ForegroundColor Cyan
& $py -m unittest discover -s (Join-Path $root "tests") -v
if ($LASTEXITCODE -ne 0) { throw "Test lõi thất bại." }

Write-Host "== Xoá build cũ" -ForegroundColor Cyan
Get-Process | Where-Object { $_.Path -like "*Xlsx2CSV_Converter.exe" } | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 400
Remove-Item -Recurse -Force $work, $dist -ErrorAction SilentlyContinue

Write-Host "== PyInstaller (spec tối giản, onefile, windowed)" -ForegroundColor Cyan
& $py -m PyInstaller --noconfirm --clean --distpath $dist --workpath $work $spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller thất bại." }

if (-not (Test-Path $exe)) { throw "Không thấy EXE: $exe" }

Write-Host "== Kiểm tra EXE (selftest)" -ForegroundColor Cyan
$proc = Start-Process -FilePath $exe -ArgumentList "--selftest" -Wait -PassThru
if ($proc.ExitCode -ne 0) { throw "EXE selftest thất bại (exit $($proc.ExitCode))." }

$size = [math]::Round((Get-Item $exe).Length / 1MB, 1)
Write-Host ""
Write-Host "OK: $exe  ($size MB)" -ForegroundColor Green


