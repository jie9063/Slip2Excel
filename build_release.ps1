$ErrorActionPreference = 'Stop'

$version = (Select-String -Path .\version.py -Pattern '^VERSION\s*=\s*"([0-9]+\.[0-9]{2})"').Matches[0].Groups[1].Value
if (-not $version) {
    throw 'A version such as 0.00 is required in version.py.'
}

python -m PyInstaller --noconfirm --clean --windowed --onefile --name "Slip2Excel-v$version" --icon .\assets\slip2excel-icon.ico --add-data ".\assets\slip2excel-icon.ico;assets" main.py
if ($LASTEXITCODE -ne 0) {
    throw 'PyInstaller build failed.'
}

$isccCommand = Get-Command ISCC.exe -ErrorAction SilentlyContinue
if ($isccCommand) {
    $isccPath = $isccCommand.Source
} else {
    $isccPath = (Get-ChildItem 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe', 'C:\Program Files\Inno Setup 6\ISCC.exe', "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe" -ErrorAction SilentlyContinue | Select-Object -First 1).FullName
}
if (-not $isccPath) {
    throw 'Inno Setup 6 was not found. Install it, then run this script again.'
}

& $isccPath "/DMyAppVersion=$version" installer.iss
if ($LASTEXITCODE -ne 0) {
    throw 'Installer build failed.'
}
Write-Host "Created: release\Slip2Excel-v$version-Setup.exe"
