param (
    [Parameter(Mandatory = $true)]
    [string]$SkillId
)

$RepoUrl = "https://github.com/ricardoviannajr/antigravity-awesome-skills.git"
$TempDir = ".temp_skills_fetch"
$TargetDir = ".antigravity/skills/$SkillId"

Write-Host "Invocando skill: $SkillId (Antigravity Exclusive)..." -ForegroundColor Cyan

if (-not (Test-Path ".antigravity/skills")) {
    New-Item -ItemType Directory -Path ".antigravity/skills" -Force | Out-Null
}

if (Test-Path $TempDir) { Remove-Item -Recurse -Force $TempDir }

try {
    Write-Host "Buscando arquivos no repositório remoto..." -ForegroundColor Gray
    git clone --depth 1 --filter=blob:none --sparse $RepoUrl $TempDir

    if ($LASTEXITCODE -ne 0) { throw "Falha ao clonar o repositório." }

    git -C $TempDir sparse-checkout set "skills/$SkillId"

    if (-not (Test-Path "$TempDir/skills/$SkillId")) {
        Write-Host "Localizando caminho exato para $SkillId..." -ForegroundColor Gray
        $ExactPath = git -C $TempDir ls-tree -r --name-only HEAD | Select-String "/$SkillId/SKILL.md" | Select-Object -First 1
        if ($ExactPath) {
            $Folder = Split-Path $ExactPath.ToString()
            git -C $TempDir sparse-checkout set $Folder
            $SkillSource = "$TempDir/$Folder"
        }
        else {
            throw "Skill '$SkillId' não encontrada no repositório remoto."
        }
    }
    else {
        $SkillSource = "$TempDir/skills/$SkillId"
    }

    if (Test-Path $TargetDir) { Remove-Item -Recurse -Force $TargetDir }
    Copy-Item -Path $SkillSource -Destination $TargetDir -Recurse -Force

    Write-Host "Skill '$SkillId' instalada com sucesso em $TargetDir" -ForegroundColor Green
}
catch {
    Write-Error "Erro ao buscar skill: $_"
}
finally {
    if (Test-Path $TempDir) { Remove-Item -Recurse -Force $TempDir }
}
