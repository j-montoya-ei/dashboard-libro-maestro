# Sync-Excel-To-GitHub.ps1
# Script para sincronizar Excel de OneDrive con GitHub automáticamente

# Configuration
$ExcelPath = "C:\Users\anlcld2\OneDrive - ELECTROINGENIERIA S.A.S\Archivos de Admin - MEJORAMIENTO CONTINUO\CALIDAD\1. SISTEMA DE GESTIÓN\LIBRO MAESTRO\Dashboard_Libro_Maestro\M-FT-1 LIBRO MAESTRO DE DOCUMENTOS v2.xlsx"
$RepoPath = "C:\Users\anlcld2\OneDrive - ELECTROINGENIERIA S.A.S\Archivos de Admin - MEJORAMIENTO CONTINUO\CALIDAD\1. SISTEMA DE GESTIÓN\LIBRO MAESTRO\Dashboard_Libro_Maestro.worktrees\despliegue-github-instrucciones"
$DataJsonPath = Join-Path $RepoPath "data.json"
$LogPath = Join-Path $RepoPath "sync-log.txt"

# Logging function
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Add-Content -Path $LogPath -Value $logEntry
    Write-Host $logEntry
}

Write-Log "=== Iniciando sincronización Excel a JSON ==="

# Verificar que el Excel existe
if (-not (Test-Path $ExcelPath)) {
    Write-Log "ERROR: No se encontró el archivo Excel en: $ExcelPath"
    exit 1
}

# Verificar que el repositorio existe
if (-not (Test-Path $RepoPath)) {
    Write-Log "ERROR: No se encontró el repositorio en: $RepoPath"
    exit 1
}

# Crear backup del data.json actual
$BackupPath = "$DataJsonPath.backup"
if (Test-Path $DataJsonPath) {
    Copy-Item -Path $DataJsonPath -Destination $BackupPath -Force
    Write-Log "Backup creado: $BackupPath"
}

# Convertir Excel a JSON usando Python
$PythonScript = @'
import sys
import pandas as pd
import json
from pathlib import Path

excel_path = sys.argv[1]
output_path = sys.argv[2]

try:
    # Leer Excel (primera hoja)
    df = pd.read_excel(excel_path, sheet_name=0)
    
    # Limpiar datos: remover filas completamente vacías
    df = df.dropna(how='all')
    
    # Reemplazar NaN con None para JSON
    df = df.where(pd.notna(df), None)
    
    # Convertir a list de diccionarios
    records = df.to_dict('records')
    
    # Guardar JSON
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    
    print(f"OK: {len(records)} registros convertidos")
    sys.exit(0)
    
except Exception as e:
    print(f"ERROR: {str(e)}")
    sys.exit(1)
'@

# Guardar script Python temporalmente
$PythonScriptPath = Join-Path $RepoPath "temp_convert.py"
Set-Content -Path $PythonScriptPath -Value $PythonScript

try {
    # Ejecutar conversión
    Write-Log "Convirtiendo Excel a JSON..."
    $output = & python $PythonScriptPath $ExcelPath $DataJsonPath
    
    if ($LASTEXITCODE -ne 0) {
        Write-Log "ERROR en conversión: $output"
        exit 1
    }
    
    Write-Log "Conversión exitosa: $output"
    
    # Cambiar a directorio del repositorio
    Push-Location $RepoPath
    
    # Verificar cambios
    $gitDiff = & git diff --quiet data.json
    $hasChanges = $LASTEXITCODE -ne 0
    
    if ($hasChanges) {
        Write-Log "Cambios detectados en data.json"
        
        # Configurar git
        & git config user.email "script@localhost"
        & git config user.name "Excel Sync Script"
        
        # Agregar y hacer commit
        & git add data.json
        & git commit -m "chore: sync data from Excel - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        
        # Push a remoto
        & git push
        
        Write-Log "✓ Cambios subidos a GitHub exitosamente"
    }
    else {
        Write-Log "✓ No hay cambios - data.json ya está actualizado"
    }
    
    Pop-Location
    Write-Log "=== Sincronización completada exitosamente ==="
    
}
catch {
    Write-Log "ERROR: $_"
    Pop-Location
    exit 1
}
finally {
    # Limpiar archivo temporal
    if (Test-Path $PythonScriptPath) {
        Remove-Item -Path $PythonScriptPath -Force
    }
}
