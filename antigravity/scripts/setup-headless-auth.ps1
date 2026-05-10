# .SYNOPSIS
#   Configura o Git para usar o GitHub CLI (gh) como helper de credenciais headless.
#
# .DESCRIPTION
#   Este script garante que o Git não abra o navegador para autenticação, utilizando o token do 'gh auth'.
#   Aplica as configurações nos níveis global e local.

Write-Host "Configurando Git para Autenticação Headless (CLI-First)..." -ForegroundColor Cyan

# 1. Verificar se o gh está instalado
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    Write-Error "GitHub CLI (gh) não encontrado. Por favor, instale-o primeiro."
    exit 1
}

# 2. Configurar helper global
Write-Host "[1/3] Configurando helper global..." -ForegroundColor Yellow
git config --global credential.helper "!gh auth git-credential"

# 3. Limpar helpers específicos de host que podem conflitar
Write-Host "[2/3] Removendo helpers redundantes..." -ForegroundColor Yellow
git config --global --unset-all credential.https://github.com.helper
git config --global --add credential.https://github.com.helper "!gh auth git-credential"

# 4. Configurar helper local (se estiver em um repo)
if (git rev-parse --is-inside-work-tree 2>$null) {
    Write-Host "[3/3] Configurando helper local do repositório..." -ForegroundColor Yellow
    git config --local credential.helper "!gh auth git-credential"
}

Write-Host "Configuração concluída. Testando conexão..." -ForegroundColor Green
gh auth status
