<#
.SYNOPSIS
    Sincroniza todos os projetos no diretório de Projetos Antigravity.

.DESCRIPTION
    Itera sobre as subpastas, identifica repositórios Git, 
    garante a autenticação via GitHub CLI e realiza git pull.
#>

$rootPath = Join-Path $env:USERPROFILE "Documents\projetos antigravity"
if (-not (Test-Path $rootPath)) {
    $rootPath = Join-Path $env:USERPROFILE "Documents\# Projetos Antigravity"
}
$directories = Get-ChildItem -Path $rootPath -Directory

Write-Host "--- Iniciando Sincronização Global Antigravity ---" -ForegroundColor Cyan

foreach ($dir in $directories) {
    $path = $dir.FullName
    if (Test-Path "$path\.git") {
        Write-Host "Processando: $($dir.Name)" -ForegroundColor Yellow
        try {
            # Muda para o diretório (apenas para o contexto do git command se necessário, 
            # mas vamos usar o parâmetro -C do git para segurança e elegância)
            
            # 1. Configurar helper CLI
            git -C $path config --local credential.helper "!gh auth git-credential"
            
            # 2. Pull rebase
            $output = git -C $path pull --rebase 2>&1
            Write-Host $output
        } catch {
            Write-Host "Erro ao processar $($dir.Name): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "Ignorando $($dir.Name) (Não é um repositório Git)" -ForegroundColor Gray
    }
}

Write-Host "--- Sincronização Global Concluída ---" -ForegroundColor Green
