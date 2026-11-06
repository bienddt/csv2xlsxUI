$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    $python = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $python)) {
        python -m venv .venv
        if ($LASTEXITCODE -ne 0) { throw 'Falha ao criar o ambiente Python.' }
    }
    & $python -m pip install -r requirements-build.txt
    if ($LASTEXITCODE -ne 0) { throw 'Falha ao instalar dependências de build.' }
    & $python -m PyInstaller --noconfirm --clean --onefile --windowed --name CSV2XLSX --add-data 'LICENSE;.' main.py
    if ($LASTEXITCODE -ne 0) { throw 'Falha ao gerar o executável.' }
    Write-Host "Executável gerado em $PSScriptRoot\dist\CSV2XLSX.exe"
} finally {
    Pop-Location
}
