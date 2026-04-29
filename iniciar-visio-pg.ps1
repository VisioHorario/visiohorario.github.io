param(
    [switch]$ForceRestart,
    [switch]$TestOnly
)

$ErrorActionPreference = "Stop"
$root = "C:\VisioGestao"
$backend = Join-Path $root "backend-ia"
$server = Join-Path $backend "server.js"

if (-not (Test-Path $server)) {
    throw "Arquivo não encontrado: $server"
}

$env:PGHOST = "localhost"
$env:PGPORT = "5432"
$env:PGDATABASE = "postgres"
$env:PGUSER = "postgres"
$env:PGPASSWORD = "123456"
$env:PERSIST_PG_ENABLED = "true"
$env:PERSIST_TABLE = "visio_snapshots"

if ($TestOnly) {
    Write-Host "Configuração pronta para iniciar backend PostgreSQL."
    Write-Host "Backend: $backend"
    exit 0
}

$listener = Get-NetTCPConnection -LocalPort 3001 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if ($listener) {
    if ($ForceRestart) {
        Stop-Process -Id $listener.OwningProcess -Force
        Start-Sleep -Milliseconds 500
    } else {
        Write-Host "Já existe processo na porta 3001 (PID $($listener.OwningProcess))."
        Write-Host "Use: .\iniciar-visio-pg.ps1 -ForceRestart"
        exit 0
    }
}

Set-Location -LiteralPath $backend
node .\server.js