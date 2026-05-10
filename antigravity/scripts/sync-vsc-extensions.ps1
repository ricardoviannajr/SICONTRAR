<#
.SYNOPSIS
    Sincroniza extensões do VS Code (base do Antigravity) com base no arquivo vscode-extensions.list.

.DESCRIPTION
    - Import: Lê a lista e instala extensões ausentes.
    - Export: Atualiza a lista com as extensões atualmente instaladas.

.EXAMPLE
    .\sync-vsc-extensions.ps1 -Mode Import
#>

param (
    [Parameter(Mandatory = $false)]
    [ValidateSet("Import", "Export")]
    [string]$Mode = "Import"
)

if (-not (Get-Command code -ErrorAction SilentlyContinue)) {
    Write-Error "Comando 'code' (VS Code) não encontrado na PATH. Certifique-se de que o VS Code está instalado e na PATH."
    exit 1
}

$extensionFile = Join-Path $PSScriptRoot "..\..\vscode-extensions.list"

if ($Mode -eq "Export") {
    Write-Host "Exportando extensões (Antigravity/VS Code) para $extensionFile..." -ForegroundColor Cyan
    & code --list-extensions | Out-File -FilePath $extensionFile -Encoding utf8
    Write-Host "Exportação concluída com sucesso!" -ForegroundColor Green
}
else {
    if (-not (Test-Path $extensionFile)) {
        Write-Error "Arquivo de extensões não encontrado: $extensionFile"
        return
    }

    Write-Host "Iniciando sincronização (Import) - Antigravity/VS Code..." -ForegroundColor Cyan
    $installed = & code --list-extensions
    $toBeInstalled = Get-Content $extensionFile | Where-Object { $_ -and $_ -notin $installed }

    if ($null -eq $toBeInstalled -or @($toBeInstalled).Count -eq 0) {
        Write-Host "Todas as extensões da lista já estão instaladas." -ForegroundColor Green
        return
    }

    Write-Host "Encontradas $(@($toBeInstalled).Count) extensões para instalar." -ForegroundColor Yellow
    foreach ($ext in $toBeInstalled) {
        Write-Host "Instalando: $ext..." -ForegroundColor Gray
        & code --install-extension $ext --force
    }

    Write-Host "Sincronização de extensões concluída!" -ForegroundColor Green
}
