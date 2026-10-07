# ==============================================================================
#  UNIVERSIDAD CATOLICA BOLIVIANA "SAN PABLO"
#  Materia: Arquitectura de Computadoras (SIS-131)
#  Proyecto: Simulador de CPU von Neumann (8 bits) y Memoria Principal
#  Autor: Javier Torrico Sejas
#
#  Script para inicializar repositorio Git local con commits semanticos atomicos
#  y enlazarlo a: https://github.com/Javiertorricos/Simulador_CPU_Excel.git
# ==============================================================================

param(
    [switch]$Push
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoDir = Split-Path -Parent $ScriptDir
$RemoteUrl = "https://github.com/Javiertorricos/Simulador_CPU_Excel.git"

Write-Host "Inicializando repositorio Git en: $RepoDir"
Set-Location $RepoDir

if (-not (Test-Path (Join-Path $RepoDir ".git"))) {
    git init
    git branch -M main
    git remote add origin $RemoteUrl
    Write-Host "Repositorio inicializado y vinculado a $RemoteUrl"
} else {
    Write-Host "El repositorio Git ya existe en esta carpeta."
}

# Configurar identidad local de Git para el estudiante
git config user.name "Javier Torrico Sejas"
git config user.email "javiertorricos@users.noreply.github.com"

# Crear .gitignore adecuado para Excel y Windows
$gitIgnoreContent = @"
# Archivos temporales de Office / Excel
~$*.xlsm
~$*.xlsx
*.tmp
*.bak
.DS_Store
Thumbs.db
"@
Set-Content -Path (Join-Path $RepoDir ".gitignore") -Value $gitIgnoreContent -Encoding UTF8

Write-Host "Generando commits semanticos atomicos..."

# 1. Commit de Documentación y Arquitectura Inicial
git add GITHUB_ISSUES_KANBAN_EN.md
git commit -m "docs(kanban): define 14 GitHub project issues and acceptance criteria" --allow-empty

git add README.md
git commit -m "docs(readme): add formal technical documentation and Mermaid architecture diagram" --allow-empty

# 2. Commit del Subsistema de Memoria y Primitivas
git add src/ModMemoria.bas
git commit -m "feat(memory): implement 256-byte RAM, logical segmentation, and Read/Write primitives" --allow-empty

# 3. Commit del CPU y ALU
git add src/ModCPU.bas
git commit -m "feat(cpu): implement CPU registers (PC, IR, MAR, MDR, AX, BX) and state reset" --allow-empty

git add src/ModALU.bas
git commit -m "feat(alu): implement 8-bit ALU operations and ZF, CF, SF status flags" --allow-empty

# 4. Commit del Ciclo de Instrucción (4 Fases)
git add src/ModCicloInstruccion.bas
git commit -m "feat(cycle): implement 4-phase instruction cycle (Fetch, Decode, Execute, Store)" --allow-empty

# 5. Commit del Ensamblador y Programas Demo
git add src/ModEnsamblador.bas
git commit -m "feat(asm): implement 2-pass in-sheet assembler and 4 mandatory test programs" --allow-empty

# 6. Commit de Interfaz Gráfica y Dashboard
git add src/ModInterfaz.bas
git commit -m "feat(gui): build interactive Excel dashboard, dynamic highlights, and log panel" --allow-empty

# 7. Commit de Scripts y Generación de Libro
git add scripts/Generar-Simulador-Javier.ps1
git commit -m "infra(build): add automated PowerShell workbook generation script" --allow-empty

git add Simulador_CPU_8bits_Javier_Torrico.xlsm
git commit -m "feat(workbook): generate ready-to-run Excel macro-enabled simulator" --allow-empty

# 8. Commit de Guía de Defensa Oral
git add GUIA_DEFENSA_ORAL.md .gitignore scripts/Iniciar-Git-Commits.ps1
git commit -m "docs(defense): add 15-minute oral defense guide and live exam Q&A" --allow-empty

Write-Host "Historial de commits semanticos generado exitosamente:"
git log --oneline -n 10

if ($Push) {
    Write-Host "Subiendo a GitHub ($RemoteUrl)..."
    git push -u origin main --force
} else {
    Write-Host ""
    Write-Host "Para subir a tu repositorio oficial de GitHub, ejecuta:"
    Write-Host "  git push -u origin main"
}
