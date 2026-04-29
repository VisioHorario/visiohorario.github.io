$ErrorActionPreference = "Stop"

$origem = "g:\Meu Drive\JECKSON RUBENS\PROJETOS APPs\Aplicação Web horarios professor\APLICAÇÃO\VISIO-Gestao de Horários E"
$destino = "C:\VisioGestao"

$arquivos = @(
    "index.html",
    "script.js",
    "backend-ia\server.js"
)

Write-Host "Iniciando sincronização..." -ForegroundColor Cyan

foreach ($arquivo in $arquivos) {
    $src = Join-Path $origem $arquivo
    $dst = Join-Path $destino $arquivo

    if (-not (Test-Path -LiteralPath $src)) {
        Write-Host "Arquivo não encontrado: $src" -ForegroundColor Yellow
        continue
    }

    $pastaDestino = Split-Path -Parent $dst
    if (-not (Test-Path -LiteralPath $pastaDestino)) {
        New-Item -ItemType Directory -Path $pastaDestino -Force | Out-Null
    }

    Copy-Item -LiteralPath $src -Destination $dst -Force
    Write-Host "OK -> $arquivo" -ForegroundColor Green
}

Write-Host "Sincronização concluída." -ForegroundColor Cyan
