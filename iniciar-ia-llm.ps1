param(
    [Parameter(Mandatory = $false)]
    [string]$ApiKey = $env:LLM_API_KEY,
    [Parameter(Mandatory = $false)]
    [string]$Model = "gpt-4o-mini",
    [Parameter(Mandatory = $false)]
    [string]$BaseUrl = "https://api.openai.com/v1",
    [switch]$ValidateOnly,
    [switch]$ForceRestart,
    [switch]$NewWindow
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    throw "Informe a chave da LLM via -ApiKey ou variável de ambiente LLM_API_KEY."
}

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$serverPath = Join-Path $root "backend-ia\server.js"

if (-not (Test-Path $serverPath)) {
    throw "Arquivo não encontrado: $serverPath"
}

$env:AJUSTE_IA_MODE = "llm"
$env:LLM_BASE_URL = $BaseUrl
$env:LLM_MODEL = $Model
$env:LLM_API_KEY = $ApiKey

function Get-BackendPidByPort {
    $conn = Get-NetTCPConnection -LocalPort 3001 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($conn) {
        return $conn.OwningProcess
    }
    return $null
}

function Test-BackendEndpoint {
    param(
        [string]$PayloadJson
    )
    $resp = Invoke-RestMethod -Uri "http://localhost:3001/ajuste-ia" -Method Post -ContentType "application/json" -Body $PayloadJson
    "OK backend respondeu:"
    $resp | ConvertTo-Json -Depth 20
}

if ($NewWindow -and -not $ValidateOnly) {
    $apiKeyEsc = $ApiKey.Replace("'", "''")
    $modelEsc = $Model.Replace("'", "''")
    $baseEsc = $BaseUrl.Replace("'", "''")
    $selfPath = $MyInvocation.MyCommand.Path
    $cmd = "Set-Location -LiteralPath '$root'; powershell -NoProfile -ExecutionPolicy Bypass -File '$selfPath' -ApiKey '$apiKeyEsc' -Model '$modelEsc' -BaseUrl '$baseEsc'"
    if ($ForceRestart) {
        $cmd += " -ForceRestart"
    }
    $proc = Start-Process -FilePath "powershell" -WorkingDirectory $root -ArgumentList @("-NoExit", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", $cmd) -PassThru
    Write-Host "Janela iniciada com PID $($proc.Id)."
    Write-Host "Backend IA ficará nessa nova janela."
    exit 0
}

if ($ValidateOnly) {
    $payload = @{
        escolaId = "EETEPA"
        turno = "MANHA"
        turmas = @()
        professores = @()
        tempos = @(
            @{ id = 1; intervalo = $false },
            @{ id = 2; intervalo = $false }
        )
        aulas = @(
            @{ id = "a1"; turno = "MANHA"; dia = "SEGUNDA"; turmaId = "t1"; tempoId = 1; disciplina = "MAT"; professorId = "p1" }
        )
    } | ConvertTo-Json -Depth 20

    $existingPid = Get-BackendPidByPort
    if ($existingPid) {
        "Backend já está ativo na porta 3001 (PID $existingPid). Executando validação sem reiniciar."
        Test-BackendEndpoint -PayloadJson $payload
        exit 0
    }

    $server = Start-Process -FilePath "node" -ArgumentList $serverPath -PassThru
    Start-Sleep -Seconds 2

    try {
        Test-BackendEndpoint -PayloadJson $payload
    } finally {
        if ($server -and -not $server.HasExited) {
            Stop-Process -Id $server.Id -Force
        }
    }
    exit 0
}

Write-Host "Iniciando backend IA em modo LLM..."
Write-Host "Modelo: $Model"
Write-Host "Base URL: $BaseUrl"
Write-Host "Ajuste IA Mode: $env:AJUSTE_IA_MODE"
Write-Host ""
Write-Host "Backend: http://localhost:3001"
Write-Host "Pressione Ctrl+C para encerrar."
Write-Host ""

$existingPid = Get-BackendPidByPort
if ($existingPid) {
    if ($ForceRestart) {
        Write-Host "Porta 3001 já em uso (PID $existingPid). Reiniciando..."
        Stop-Process -Id $existingPid -Force
        Start-Sleep -Milliseconds 500
    } else {
        Write-Host "Backend já está em execução na porta 3001 (PID $existingPid)."
        Write-Host "Use -ForceRestart para encerrar o processo atual e iniciar outro."
        exit 0
    }
}

node $serverPath
