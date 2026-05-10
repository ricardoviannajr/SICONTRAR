<#
.SYNOPSIS
    Sincroniza forçadamente o repositório local com a branch de origem.

.DESCRIPTION
    Destinado ao repositório AIConfig ou clones descartáveis — NÃO use em projetos de aplicação com trabalho local não salvo.

    1. Atualiza as referências remotas (git fetch --all).
    2. Reseta a branch local para o estado exato da origin (git reset --hard).
    3. Remove arquivos e diretórios não rastreados (git clean -fd).
    4. Importa extensões listadas em vscode-extensions.list (via sync-vsc-extensions.ps1).

    Sem -Force, o script exige que você digite SIM para continuar.

.EXAMPLE
    .\sync-hard.ps1 -Force
#>

param (
    [switch]$Force
)

if (-not $Force) {
    Write-Host ""
    Write-Host "AVISO: sync-hard descarta TODAS as alterações locais não commitadas e arquivos não rastreados." -ForegroundColor Red
    Write-Host "Use apenas no AIConfig ou em ambiente descartável. Para cancelar, feche este prompt sem confirmar." -ForegroundColor Yellow
    $confirm = Read-Host "Digite SIM em maiúsculas para continuar"
    if ($confirm -ne "SIM") {
        Write-Host "Operação cancelada." -ForegroundColor Gray
        exit 1
    }
}

$currentBranch = git branch --show-current
Write-Host "Iniciando Sincronização Forçada na branch: $currentBranch" -ForegroundColor Cyan

Write-Host "[1/4] Buscando novidades remotas..." -ForegroundColor Yellow
git fetch --all

Write-Host "[2/4] Executando Hard Reset para origin/$currentBranch..." -ForegroundColor Yellow
git reset --hard "origin/$currentBranch"

Write-Host "[3/4] Removendo arquivos não rastreados..." -ForegroundColor Yellow
git clean -fd

Write-Host "[4/4] Sincronizando extensões (VS Code / lista)..." -ForegroundColor Yellow
& "$PSScriptRoot\sync-vsc-extensions.ps1" -Mode Import

Write-Host "Sincronização concluída com sucesso." -ForegroundColor Green
