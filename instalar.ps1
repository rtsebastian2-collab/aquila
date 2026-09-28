# ============================================================================
#  Aquila — instalador y actualizador
#
#  Instalar (en PowerShell):
#     irm https://raw.githubusercontent.com/rtsebastian2-collab/aquila/main/instalar.ps1 | iex
#
#  Qué hace:
#    1. Descarga la última versión desde GitHub.
#    2. La copia en %LOCALAPPDATA%\Aquila (no toca nada más del sistema).
#    3. Crea el acceso directo «Aquila» en el escritorio y en el menú Inicio (con el manual y el desinstalador).
#  Volver a ejecutarlo actualiza Aquila. Los informes (Documentos\Aquila) se conservan.
# ============================================================================

& {
    $ErrorActionPreference = 'Stop'
    $Repo = 'rtsebastian2-collab/aquila'
    $Rama = 'main'
    $Actualizando = $env:AQUILA_ACTUALIZAR -eq '1'
    $Base = Join-Path $env:LOCALAPPDATA 'Aquila'
    $App = Join-Path $Base 'app'
    $PowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

    function Paso([string]$t) { Write-Host "   > $t" -ForegroundColor Cyan }

    function Crear-Acceso([string]$Ruta, [string]$Destino, [string]$Argumentos, [string]$Icono, [string]$Descripcion, [switch]$Admin, [switch]$Minimizado) {
        [void](New-Item -ItemType Directory -Path (Split-Path $Ruta) -Force)
        $sh = New-Object -ComObject WScript.Shell
        $lnk = $sh.CreateShortcut($Ruta)
        $lnk.TargetPath = $Destino
        if ($Argumentos) { $lnk.Arguments = $Argumentos }
        $lnk.IconLocation = "$Icono,0"
        $lnk.Description = $Descripcion
        $lnk.WorkingDirectory = $App
        if ($Minimizado) { $lnk.WindowStyle = 7 }
        $lnk.Save()
        if ($Admin) {
            # Marca «Ejecutar como administrador» (bit 0x20 del byte 0x15 del .lnk)
            $b = [IO.File]::ReadAllBytes($Ruta); $b[0x15] = $b[0x15] -bor 0x20; [IO.File]::WriteAllBytes($Ruta, $b)
        }
    }

    Write-Host ''
    Write-Host '   ==============================================' -ForegroundColor Yellow
    Write-Host "      AQUILA  ·  $(if ($Actualizando) { 'Actualizando' } else { 'Instalación' })" -ForegroundColor Yellow
    Write-Host '      Diagnóstico y optimización de tu computadora' -ForegroundColor Yellow
    Write-Host '   ==============================================' -ForegroundColor Yellow
    Write-Host ''

    try {
        # --- Requisitos ---
        if ([Environment]::OSVersion.Version.Major -lt 10) { throw 'Aquila funciona en Windows 10 y Windows 11. Esta computadora tiene una versión anterior.' }
        if ($PSVersionTable.PSVersion.Major -lt 5) { throw 'Se necesita PowerShell 5.1 o superior (viene incluido en Windows 10 y 11).' }

        # Si viene desde Aquila, esperar a que la ventana se cierre
        if ($env:AQUILA_ESPERAR_PID) { try { Wait-Process -Id ([int]$env:AQUILA_ESPERAR_PID) -Timeout 30 -ErrorAction SilentlyContinue } catch {} }

        # --- Descarga ---
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $tmp = Join-Path $env:TEMP ('aquila_' + [guid]::NewGuid().ToString('N'))
        [void](New-Item -ItemType Directory -Path $tmp)
        $zip = Join-Path $tmp 'aquila.zip'
        Paso 'Descargando la última versión desde GitHub...'
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri "https://github.com/$Repo/archive/refs/heads/$Rama.zip" -OutFile $zip -UseBasicParsing
        Expand-Archive -Path $zip -DestinationPath $tmp -Force
        $src = Get-ChildItem -Path $tmp -Directory | Select-Object -First 1
        if (-not $src -or -not (Test-Path (Join-Path $src.FullName 'app\Aquila.ps1'))) { throw 'La descarga no contiene Aquila. Intenta de nuevo en unos minutos.' }

        # --- Copia ---
        Paso "Instalando en $Base ..."
        [void](New-Item -ItemType Directory -Path $Base -Force)
        if (Test-Path $App) { Remove-Item -Path $App -Recurse -Force }
        Copy-Item -Path (Join-Path $src.FullName 'app') -Destination $App -Recurse -Force
        Copy-Item -Path (Join-Path $src.FullName 'version.txt') -Destination (Join-Path $Base 'version.txt') -Force
        Get-ChildItem -Path $Base -Recurse -File | Unblock-File
        $version = (Get-Content (Join-Path $Base 'version.txt') -TotalCount 1).Trim()
        Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue

        # --- Accesos directos ---
        Paso 'Creando el acceso directo en el escritorio y en el menú Inicio...'
        $icono = Join-Path $App 'aquila.ico'
        $argsApp = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$App\Aquila.ps1`""
        $menu = Join-Path ([Environment]::GetFolderPath('Programs')) 'Aquila'
        Crear-Acceso (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Aquila.lnk') $PowerShell $argsApp $icono 'Aquila: diagnóstico y optimización de tu PC' -Admin -Minimizado
        Crear-Acceso (Join-Path $menu 'Aquila.lnk') $PowerShell $argsApp $icono 'Aquila: diagnóstico y optimización de tu PC' -Admin -Minimizado
        Crear-Acceso (Join-Path $menu 'Manual de Aquila.lnk') (Join-Path $App 'Manual.html') '' $icono 'Cómo usar Aquila'
        Crear-Acceso (Join-Path $menu 'Desinstalar Aquila.lnk') $PowerShell "-NoProfile -ExecutionPolicy Bypass -File `"$App\Desinstalar.ps1`"" $icono 'Quitar Aquila de esta PC'

        Write-Host ''
        Write-Host "   Aquila $version $(if ($Actualizando) { 'actualizada' } else { 'instalada' }) correctamente." -ForegroundColor Green
        if ($Actualizando) {
            Write-Host '   Abriendo Aquila...' -ForegroundColor Green
            Start-Process -FilePath $PowerShell -ArgumentList $argsApp -WindowStyle Hidden
            Start-Sleep -Seconds 2
        } else {
            Write-Host ''
            Write-Host '   Listo: en tu ESCRITORIO tienes el icono del águila «Aquila».' -ForegroundColor White
            Write-Host '   Haz doble clic, acepta el permiso de administrador y presiona «Diagnosticar».' -ForegroundColor White
            Write-Host '   Se abrirá el manual para que sepas cómo usarla.' -ForegroundColor Gray
            Write-Host ''
            Start-Process (Join-Path $App 'Manual.html')
        }
    } catch {
        Write-Host ''
        Write-Host "   No se pudo completar: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host '   Revisa tu conexión a internet e inténtalo de nuevo.' -ForegroundColor Red
        if ($Actualizando) { Start-Sleep -Seconds 8 }
    }
}
