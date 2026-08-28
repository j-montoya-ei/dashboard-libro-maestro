# 🔄 Automatizar Sincronización de Excel a Dashboard

Este directorio contiene scripts para mantener tu dashboard actualizado automáticamente desde el Excel de OneDrive **sin intervención manual**.

## ¿Cómo funciona?

```
Tu Excel en OneDrive
        ↓
    Cada día a las 1 AM
        ↓
Script ejecuta conversión
        ↓
data.json se actualiza
        ↓
Dashboard se actualiza automáticamente
```

## Requisitos previos

✅ **Ya instalado en tu PC:**
- Windows 10/11
- Python 3.7+
- Git
- pandas (`pip install pandas openpyxl`)

**Verifica que Python está instalado:**
```powershell
python --version
```

Si no está instalado, descárgalo de: https://www.python.org/

## Instalación (Paso a Paso)

### 1️⃣ Instalar librerías Python requeridas

Abre PowerShell y ejecuta:

```powershell
pip install pandas openpyxl
```

### 2️⃣ Permitir ejecución de scripts PowerShell

Abre PowerShell **como administrador** y ejecuta:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Presiona `Y` para confirmar.

### 3️⃣ Crear la tarea automática

Abre PowerShell **como administrador** en esta carpeta y ejecuta:

```powershell
.\Setup-Sync-Schedule.ps1
```

Deberías ver un mensaje como:
```
✓ Tarea creada exitosamente!
```

### 4️⃣ Probar que funciona (OPCIONAL)

Para verificar que todo está configurado correctamente, ejecuta manualmente el script:

```powershell
.\Sync-Excel-To-GitHub.ps1
```

Deberías ver mensajes como:
```
[2026-08-28 17:20:15] === Iniciando sincronización Excel a JSON ===
[2026-08-28 17:20:15] Backup creado: ...
[2026-08-28 17:20:18] Conversión exitosa: OK: 649 registros convertidos
[2026-08-28 17:20:20] ✓ Cambios subidos a GitHub exitosamente
```

## Verificar la configuración

### Abrir Task Scheduler (Windows)

1. Presiona `Win + R`
2. Escribe: `taskschd.msc`
3. Presiona Enter
4. Busca la carpeta `\LibroMaestro\`
5. Deberías ver la tarea `LibroMaestro-Excel-Sync`

### Ver logs de ejecución

Los logs se guardan en `sync-log.txt` en esta carpeta. Abre en cualquier editor de texto para ver:
- Cuándo se ejecutó
- Cuántos registros se sincronizaron
- Si hubo errores

## Cambiar horario de ejecución

Si quieres que se ejecute a otra hora, edita **Setup-Sync-Schedule.ps1**:

Busca esta línea:
```powershell
$trigger = New-ScheduledTaskTrigger -Daily -At "01:00 AM"
```

Cambia `"01:00 AM"` a tu hora preferida. Ejemplos:
- `"03:00 AM"` → 3 AM
- `"06:00 AM"` → 6 AM
- `"12:00 PM"` → 12 PM (mediodía)

Luego ejecuta nuevamente:
```powershell
.\Setup-Sync-Schedule.ps1
```

## Desactivar sincronización automática

Si necesitas pausar la sincronización:

1. Abre Task Scheduler (`Win + R` → `taskschd.msc`)
2. Busca `\LibroMaestro\LibroMaestro-Excel-Sync`
3. Haz clic derecho → **Disable**

Para reactivar:
4. Haz clic derecho → **Enable**

## Troubleshooting

### ❌ "No se encontró Python"

```
ERROR: No se encontró Python en el sistema
```

**Solución:** Instala Python desde https://www.python.org/ y marca la opción "Add Python to PATH"

### ❌ "No se encontró el archivo Excel"

```
ERROR: No se encontró el archivo Excel
```

**Solución:** Verifica que el archivo existe en:
```
C:\Users\anlcld2\OneDrive - ELECTROINGENIERIA S.A.S\Archivos de Admin - MEJORAMIENTO CONTINUO\CALIDAD\1. SISTEMA DE GESTIÓN\LIBRO MAESTRO\Dashboard_Libro_Maestro\M-FT-1 LIBRO MAESTRO DE DOCUMENTOS v2.xlsx
```

### ❌ "No se puede acceder al repositorio de Git"

**Solución:** Asegúrate que tienes credenciales de GitHub guardadas. Ejecuta:
```powershell
git config --global user.email "tu-email@gmail.com"
git config --global user.name "Tu Nombre"
```

### ❌ La tarea no se ejecuta automáticamente

**Solución:** 
1. Verifica que tu PC no está en modo de bajo consumo a la 1 AM
2. Asegúrate que el script PowerShell puede ejecutarse (`Set-ExecutionPolicy RemoteSigned`)
3. Revisa los logs en `sync-log.txt`

## FAQs

**P: ¿Se sincroniza en tiempo real?**
R: No, se sincroniza una vez al día a las 1 AM. Si necesitas tiempo real, contacta al equipo técnico.

**P: ¿Qué pasa si el Excel está abierto cuando se ejecuta el script?**
R: Excel bloquea el archivo y la sincronización falla. El script reintentará el día siguiente.

**P: ¿Puedo cambiar la frecuencia (más de una vez al día)?**
R: Sí, modifica `Setup-Sync-Schedule.ps1` y cambia `$trigger` a:
```powershell
$trigger = New-ScheduledTaskTrigger -Once -At "01:00 AM" -RepetitionInterval (New-TimeSpan -Hours 6) -RepetitionDuration (New-TimeSpan -Days 1)
```
Esto ejecuta cada 6 horas.

**P: ¿Dónde veo cuáles cambios se subieron?**
R: En GitHub, ve a la rama `despliegue-github-instrucciones`, haz clic en "Commits" y verás commits con mensaje "chore: sync data from Excel..."

## Soporte

Si tienes problemas:
1. Revisa los logs en `sync-log.txt`
2. Abre PowerShell como administrador
3. Ejecuta el script manualmente: `.\Sync-Excel-To-GitHub.ps1`
4. Copia el error completo y consulta con el equipo técnico
