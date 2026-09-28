#Requires -Version 5.1
<#
    AQUILA — Diagnóstico
    Mide la computadora, explica en palabras simples por qué está lenta y ofrece arreglos.
    Antes de aplicar cualquier arreglo muestra: qué hace, por qué y si afecta algo, y pide confirmación.

    Uso normal: el acceso directo «Aquila» del escritorio (abre la ventana).
    Uso en consola: powershell -ExecutionPolicy Bypass -File Aquila-Diagnostico.ps1

    Parámetros:
      -SegundosMuestreo 60     Tiempo midiendo el uso en vivo (por defecto 40)
      -Comparar ruta.json      Compara contra el diagnóstico de otra PC (o de esta misma, antes/después)
      -SinArreglos             Solo genera el informe
      -SinDuplicados           No busca archivos duplicados
      -NoAbrir                 No abre el informe al terminar
      -RestaurarInicio         Reactiva los programas de inicio que esta herramienta desactivó
      -Aplicar ruta.json       (usado por la ventana) aplica los arreglos listados en ese archivo, sin preguntas
      -Registro / -ArchivoResultado   (usados por la ventana) rutas del registro y del resultado
#>
[CmdletBinding()]
param(
    [int]$SegundosMuestreo = 40,
    [string]$CarpetaSalida,
    [string]$Comparar,
    [switch]$SinArreglos,
    [switch]$SinDuplicados,
    [switch]$NoAbrir,
    [switch]$RestaurarInicio,
    [string]$Aplicar,
    [string]$Registro,
    [string]$ArchivoResultado
)

$ErrorActionPreference = 'SilentlyContinue'
. (Join-Path $PSScriptRoot 'Comun.ps1')
$VersionHerramienta = $AquilaVersion
$HoraInicio = Get-Date
if (-not $CarpetaSalida) { $CarpetaSalida = $AquilaInformes }
[void](New-Item -ItemType Directory -Path $CarpetaSalida -Force)
$Equipo = $env:COMPUTERNAME
$Marca = Get-Date -Format 'yyyy-MM-dd_HHmm'
$ArchivoLog = if ($Registro) { $Registro } else { Join-Path $CarpetaSalida "Registro_${Equipo}_$Marca.txt" }
$ArchivoCambiosInicio = Join-Path $CarpetaSalida "CambiosInicio_$Equipo.json"
$EsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$RxGuid = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
$SisDrive = $env:SystemDrive
$script:Interactivo = -not $Aplicar
$script:NumPaso = 0
$TotalPasos = 11

$Hallazgos = New-Object System.Collections.Generic.List[object]
$Arreglos  = New-Object System.Collections.Generic.List[object]
$Tablas    = New-Object System.Collections.Generic.List[object]
$Datos     = [ordered]@{ Equipo = $Equipo; Fecha = (Get-Date).ToString('yyyy-MM-dd HH:mm'); Version = $VersionHerramienta }

# ============================================================================
#  Utilidades
# ============================================================================
function Write-Log([string]$t) {
    $linea = '[{0:HH:mm:ss}] {1}' -f (Get-Date), $t
    for ($i = 0; $i -lt 5; $i++) {
        try { [IO.File]::AppendAllText($ArchivoLog, $linea + "`r`n", [Text.Encoding]::UTF8); return } catch { Start-Sleep -Milliseconds 100 }
    }
}
function Paso([string]$t) {
    $script:NumPaso++
    Write-Host "`n>> $t" -ForegroundColor Cyan
    Write-Log "($($script:NumPaso)/$TotalPasos) $t"
}
function Sub([string]$t) { Write-Host "   $t" -ForegroundColor DarkGray; Write-Log "   $t" }

function Add-Hallazgo {
    param([string]$Categoria, [ValidateSet('Critico', 'Alto', 'Medio', 'Info', 'OK')][string]$Nivel,
          [string]$Titulo, [string]$Detalle, [string]$Porque, [string]$Solucion, [string[]]$Arreglo = @())
    $Hallazgos.Add([pscustomobject]@{ Categoria = $Categoria; Nivel = $Nivel; Titulo = $Titulo; Detalle = $Detalle
                                      Porque = $Porque; Solucion = $Solucion; Arreglos = @($Arreglo) })
}

# Un arreglo es solo la descripción + los datos; lo que hace está en $Acciones[$Tipo]
function Add-Arreglo {
    param([string]$Id, [string]$Tipo, [string]$Nombre, [string]$QueHace, [string]$Porque, [string]$Afecta,
          [ValidateSet('Bajo', 'Medio', 'Alto')][string]$Riesgo, [double]$Bytes = 0, $Datos = @{}, [switch]$Admin)
    if (-not $Tipo) { $Tipo = $Id }
    if ($Arreglos | Where-Object { $_.Id -eq $Id }) { return }
    $Arreglos.Add([pscustomobject]@{ Id = $Id; Tipo = $Tipo; Numero = 0; Nombre = $Nombre; QueHace = $QueHace; Porque = $Porque; Afecta = $Afecta
                                     Riesgo = $Riesgo; Bytes = $Bytes; Datos = $Datos; Admin = [bool]$Admin })
}

function Add-Tabla([string]$Titulo, [string]$Nota, $Filas, [string[]]$Columnas) {
    $Tablas.Add([pscustomobject]@{ Titulo = $Titulo; Nota = $Nota; Filas = @($Filas); Columnas = $Columnas })
}

# Suma el tamaño de una carpeta sin detenerse por carpetas sin permiso.
# Separa lo que realmente ocupa disco de lo que está "solo en la nube" (OneDrive a petición).
function Get-TamanoCarpeta {
    param([string]$Ruta, [long]$MinGrande = 0, [string]$Filtro, [string[]]$Excluir = @())
    $grandes = New-Object System.Collections.Generic.List[object]
    $bytes = [long]0; $nube = [long]0; $n = [long]0
    if ($Ruta -and (Test-Path -LiteralPath $Ruta -PathType Leaf)) {
        $bytes = (Get-Item -LiteralPath $Ruta -Force).Length; $n = 1
    } elseif ($Ruta -and $Filtro -and (Test-Path -LiteralPath $Ruta)) {
        foreach ($f in Get-ChildItem -LiteralPath $Ruta -Filter $Filtro -File -Force) { $bytes += $f.Length; $n++ }
    } elseif ($Ruta -and (Test-Path -LiteralPath $Ruta)) {
        $excl = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($e in $Excluir) { if ($e) { [void]$excl.Add($e.TrimEnd('\')) } }
        # Pila de (ruta, profundidad): el límite de profundidad evita bucles por accesos directos de carpeta
        $pila = New-Object 'System.Collections.Generic.Stack[object]'
        $pila.Push(@($Ruta, 0))
        while ($pila.Count -gt 0) {
            $item = $pila.Pop(); $prof = $item[1]
            $di = New-Object IO.DirectoryInfo ($item[0])
            try {
                foreach ($f in $di.EnumerateFiles()) {
                    $n++
                    if (([int]$f.Attributes -band 0x441000) -ne 0) { $nube += $f.Length; continue }   # Offline / RecallOnOpen / RecallOnDataAccess
                    $bytes += $f.Length
                    if ($MinGrande -gt 0 -and $f.Length -ge $MinGrande) {
                        $grandes.Add([pscustomobject]@{ Ruta = $f.FullName; Bytes = $f.Length; Fecha = $f.LastWriteTime })
                    }
                }
            } catch {}
            if ($prof -ge 40) { continue }
            try {
                foreach ($s in $di.EnumerateDirectories()) {
                    if (([int]$s.Attributes -band 0x406) -eq 0x406) { continue }   # accesos antiguos de Windows (Mis documentos, etc.)
                    if ($excl.Count -and $excl.Contains($s.FullName)) { continue }
                    $pila.Push(@($s.FullName, ($prof + 1)))
                }
            } catch {}
        }
    }
    [pscustomobject]@{ Bytes = $bytes; BytesNube = $nube; Archivos = $n; Grandes = $grandes }
}

function Resolve-Rutas([string[]]$Patrones) {
    foreach ($p in $Patrones) {
        if (-not $p) { continue }
        if ($p -match '[\*\?]') { Resolve-Path -Path $p -ErrorAction SilentlyContinue | ForEach-Object { $_.ProviderPath } }
        elseif (Test-Path -LiteralPath $p) { $p }
    }
}

function Remove-Contenido {
    param([string]$Ruta, [int]$DiasMin = 0, [string]$Filtro)
    if (-not $Filtro) { $Filtro = '*' }
    if (-not (Test-Path -LiteralPath $Ruta)) { return [long]0 }
    $medir = if ($Filtro -ne '*') { $Filtro } else { $null }
    $antes = (Get-TamanoCarpeta -Ruta $Ruta -Filtro $medir).Bytes
    if (Test-Path -LiteralPath $Ruta -PathType Leaf) {
        Remove-Item -LiteralPath $Ruta -Force -ErrorAction SilentlyContinue
    } else {
        $limite = (Get-Date).AddDays(-$DiasMin)
        Get-ChildItem -LiteralPath $Ruta -Filter $Filtro -Force -ErrorAction SilentlyContinue |
            Where-Object { $DiasMin -le 0 -or $_.LastWriteTime -lt $limite } |
            ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
    }
    $despues = (Get-TamanoCarpeta -Ruta $Ruta -Filtro $medir).Bytes
    [long][math]::Max(0, $antes - $despues)
}

function Confirmar([string]$Pregunta) {
    if (-not $script:Interactivo) { return $true }
    do { $r = Read-Host "$Pregunta (S/N)" } until ($r -match '^[sSnN]')
    $r -match '^[sS]'
}

# Devuelve $true si el programa ya está cerrado (o se cerró). En la ventana, $Forzar indica que el usuario aceptó cerrarlo.
function Esperar-Cierre([string[]]$Procesos, [string]$App, $Forzar = $false) {
    if (-not (Get-Process -Name $Procesos -ErrorAction SilentlyContinue)) { return $true }
    if (-not $script:Interactivo) {
        if ($Forzar) { Get-Process -Name $Procesos -ErrorAction SilentlyContinue | Stop-Process -Force; Start-Sleep -Seconds 3 }
        return -not (Get-Process -Name $Procesos -ErrorAction SilentlyContinue)
    }
    Write-Host "      $App está abierto. Guarda tu trabajo y ciérralo; luego presiona Enter." -ForegroundColor Yellow
    Write-Host '      (F = cerrarlo a la fuerza,  O = omitir este arreglo)' -ForegroundColor Yellow
    $r = Read-Host '      '
    if ($r -match '^[oO]') { return $false }
    if ($r -match '^[fF]') { Get-Process -Name $Procesos -ErrorAction SilentlyContinue | Stop-Process -Force; Start-Sleep -Seconds 3 }
    -not (Get-Process -Name $Procesos -ErrorAction SilentlyContinue)
}

# Activa/desactiva un programa de inicio igual que el Administrador de tareas (reversible, no borra nada)
function Set-EstadoInicio([string]$ClaveAprobado, [string]$Nombre, [bool]$Activar) {
    if (-not (Test-Path -LiteralPath $ClaveAprobado)) { [void](New-Item -Path $ClaveAprobado -Force) }
    $v = if ($Activar) { [byte[]](2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) } else { [byte[]](3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) }
    Set-ItemProperty -LiteralPath $ClaveAprobado -Name $Nombre -Value $v -Type Binary -ErrorAction Stop
}

function Libre-Sistema { [double](Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$SisDrive'").FreeSpace }

function Nuevo-PuntoRestauracion {
    if (-not $EsAdmin) { return $false }
    Write-Host '   Creando punto de restauración (para poder deshacer)...' -ForegroundColor DarkGray
    Write-Log 'Creando punto de restauración...'
    Enable-ComputerRestore -Drive "$SisDrive\" -ErrorAction SilentlyContinue
    New-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore' -Name SystemRestorePointCreationFrequency -Value 0 -PropertyType DWord -Force | Out-Null
    try {
        Checkpoint-Computer -Description 'Aquila - antes de optimizar' -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
        Write-Log 'Punto de restauración creado.'; return $true
    } catch {
        Write-Host "   No se pudo crear el punto de restauración: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Log "No se pudo crear el punto de restauración: $($_.Exception.Message)"; return $false
    }
}

# Enviar a la Papelera (recuperable), sin cuadros de diálogo
if (-not ('PapeleraPC' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class PapeleraPC {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct SHFILEOPSTRUCT {
        public IntPtr hwnd; public uint wFunc; public string pFrom; public string pTo;
        public ushort fFlags; public bool fAnyOperationsAborted; public IntPtr hNameMappings; public string lpszProgressTitle;
    }
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    static extern int SHFileOperation(ref SHFILEOPSTRUCT op);
    // FO_DELETE + FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI
    public static int Enviar(string ruta) {
        SHFILEOPSTRUCT op = new SHFILEOPSTRUCT();
        op.wFunc = 3; op.pFrom = ruta + "\0"; op.fFlags = 0x40 | 0x10 | 0x4 | 0x400;
        return SHFileOperation(ref op);
    }
}
'@
}

# Huella de un archivo: completa hasta 1 GB; para archivos más grandes, 16 muestras de 64 KB repartidas
function Get-Huella([string]$Ruta, [long]$Largo, [switch]$Muestra) {
    $md5 = [Security.Cryptography.MD5]::Create()
    try {
        $fs = [IO.File]::Open($Ruta, 'Open', 'Read', 'ReadWrite')
        try {
            if (-not $Muestra -and $Largo -le 1GB) { return [BitConverter]::ToString($md5.ComputeHash($fs)) }
            $buf = New-Object byte[] 65536
            $ms = New-Object IO.MemoryStream
            $partes = 16
            for ($i = 0; $i -lt $partes; $i++) {
                $fs.Position = [long][math]::Max(0, [math]::Floor(($Largo - 65536) * $i / ($partes - 1)))
                $n = $fs.Read($buf, 0, $buf.Length); $ms.Write($buf, 0, $n)
            }
            return [BitConverter]::ToString($md5.ComputeHash($ms.ToArray()))
        } finally { $fs.Dispose() }
    } catch { return $null } finally { $md5.Dispose() }
}

function Buscar-Duplicados([string[]]$Carpetas, [long]$MinBytes = 1MB, [int]$SegundosMax = 240, [string[]]$Nube = @()) {
    $reloj = [Diagnostics.Stopwatch]::StartNew()
    $vistos = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $porTam = @{}
    # Carpetas y archivos de programación: sus «copias» son normales y borrarlas rompería proyectos
    $rxOmitirDir = '^(node_modules|\.git|\.svn|__pycache__|\.venv|venv|env|AppData|\$RECYCLE\.BIN|\.vs|\.idea|\.gradle|bin|obj|dist|build|build_exe|build_tmp|win-unpacked|site-packages|Lib|target|out|release|debug)$'
    $rxOmitirExt = '\.(ost|pst|lnk|ini|tmp|sys|dll|exe|pyd|pyc|so|jar|node|lib|pdb|msi)$'
    foreach ($c in $Carpetas | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique) {
        $pila = New-Object 'System.Collections.Generic.Stack[object]'
        $pila.Push(@($c, 0))
        while ($pila.Count -gt 0) {
            $item = $pila.Pop(); $di = New-Object IO.DirectoryInfo ($item[0])
            try {
                foreach ($f in $di.EnumerateFiles()) {
                    if ($f.Length -lt $MinBytes) { continue }
                    $a = [int]$f.Attributes
                    if (($a -band 0x441000) -ne 0 -or ($a -band 0x4) -ne 0) { continue }   # solo en la nube / archivos de sistema
                    if ($f.Name -match $rxOmitirExt) { continue }
                    if (-not $vistos.Add($f.FullName)) { continue }
                    $k = [string]$f.Length
                    if (-not $porTam.ContainsKey($k)) { $porTam[$k] = New-Object System.Collections.Generic.List[string] }
                    $porTam[$k].Add($f.FullName)
                }
            } catch {}
            if ($item[1] -ge 30) { continue }
            try {
                foreach ($s in $di.EnumerateDirectories()) {
                    if (([int]$s.Attributes -band 0x406) -eq 0x406 -or $s.Name -match $rxOmitirDir) { continue }
                    $pila.Push(@($s.FullName, ($item[1] + 1)))
                }
            } catch {}
        }
    }
    $desc = Join-Path $env:USERPROFILE 'Downloads'
    $grupos = New-Object System.Collections.Generic.List[object]
    $incompleto = $false
    # Primero los tamaños más grandes: son los que más espacio desperdician
    foreach ($k in ($porTam.Keys | Where-Object { $porTam[$_].Count -gt 1 } | Sort-Object { [long]$_ } -Descending)) {
        if ($reloj.Elapsed.TotalSeconds -gt $SegundosMax) { $incompleto = $true; break }
        $largo = [long]$k
        $porMuestra = $porTam[$k] | Group-Object { Get-Huella $_ $largo -Muestra } | Where-Object { $_.Name -and $_.Count -gt 1 }
        foreach ($gm in $porMuestra) {
            $finales = if ($largo -le 1GB) { $gm.Group | Group-Object { Get-Huella $_ $largo } | Where-Object { $_.Name -and $_.Count -gt 1 } } else { @($gm) }
            foreach ($g in $finales) {
                # Se conserva: la que está en la nube (no se toca lo compartido), la que no está en Descargas,
                # la de nombre original (sin «(1)» ni «copia»), la más antigua y, al final, la de ruta más corta
                $orden = @($g.Group | Sort-Object @{ E = { $r = $_; if ($Nube | Where-Object { $_ -and $r -like "$_\*" }) { 0 } else { 1 } } },
                                                  @{ E = { if ($_ -like "$desc\*") { 1 } else { 0 } } },
                                                  @{ E = { if ((Split-Path $_ -Leaf) -match '\(\d+\)|\bcopia\b|\bcopy\b|- Copy') { 1 } else { 0 } } },
                                                  @{ E = { (Get-Item -LiteralPath $_ -Force).CreationTime } }, @{ E = { $_.Length } })
                $grupos.Add([pscustomobject]@{ Bytes = $largo; Conservar = $orden[0]; Copias = @($orden | Select-Object -Skip 1); Parcial = ($largo -gt 1GB) })
            }
        }
    }
    [pscustomobject]@{ Grupos = $grupos; Incompleto = $incompleto; Segundos = [math]::Round($reloj.Elapsed.TotalSeconds) }
}

# ============================================================================
#  Catálogo de acciones (lo que hace cada arreglo). Reciben $d = los datos del arreglo.
# ============================================================================
$AccionBorrar = {
    param($d)
    if ($d.Procesos -and -not (Esperar-Cierre $d.Procesos $d.App $d.ForzarCierre)) { return 'Omitido: el programa seguía abierto.' }
    $total = [long]0
    foreach ($r in @($d.Rutas)) { $total += Remove-Contenido -Ruta $r -DiasMin ([int]$d.DiasMin) -Filtro $d.Filtro }
    "Se liberaron $(Fmt $total). (Los archivos en uso se saltan.)"
}
$Acciones = @{
    borrar = $AccionBorrar
    papelera = { param($d) $antes = Libre-Sistema; Clear-RecycleBin -Force -ErrorAction SilentlyContinue; "Papelera vaciada (liberado en ${SisDrive}: $(Fmt ([math]::Max(0, (Libre-Sistema) - $antes))))." }
    wu = {
        param($d)
        Stop-Service wuauserv, bits -Force
        $t = [long]0; foreach ($r in @($d.Rutas)) { $t += Remove-Contenido $r }
        Start-Service bits, wuauserv
        "Se liberaron $(Fmt $t)."
    }
    do = {
        param($d)
        $antes = Libre-Sistema
        Delete-DeliveryOptimizationCache -Force -ErrorAction SilentlyContinue
        foreach ($r in @($d.Rutas)) { [void](Remove-Contenido $r) }
        "Liberado: $(Fmt ([math]::Max(0, (Libre-Sistema) - $antes)))."
    }
    windowsold = {
        param($d)
        $antes = Libre-Sistema
        $vc = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches'
        $ok = $false
        if ((Get-Command cleanmgr.exe -ErrorAction SilentlyContinue) -and (Test-Path $vc)) {
            Get-ChildItem $vc | ForEach-Object { Remove-ItemProperty -LiteralPath $_.PSPath -Name StateFlags0077 -ErrorAction SilentlyContinue }
            foreach ($c in 'Previous Installations', 'Temporary Setup Files', 'Windows Upgrade Log Files') {
                if (Test-Path "$vc\$c") { Set-ItemProperty -LiteralPath "$vc\$c" -Name StateFlags0077 -Value 2 -Type DWord; $ok = $true }
            }
            if ($ok) { Write-Log 'Ejecutando el Liberador de espacio de Windows (puede tardar varios minutos)...'; Start-Process cleanmgr.exe -ArgumentList '/sagerun:77' -Wait }
        }
        $si = if ((Get-Culture).TwoLetterISOLanguageName -eq 'es') { 'S' } else { 'Y' }
        foreach ($r in @($d.Rutas)) {
            if (Test-Path -LiteralPath $r) {
                & takeown.exe /F $r /R /A /D $si 2>&1 | Out-Null
                & icacls.exe $r /grant '*S-1-5-32-544:F' /T /C /Q 2>&1 | Out-Null
                Remove-Item -LiteralPath $r -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        "Liberado: $(Fmt ([math]::Max(0, (Libre-Sistema) - $antes)))."
    }
    puntos = { param($d) $antes = Libre-Sistema; & vssadmin.exe resize shadowstorage /for=$SisDrive /on=$SisDrive /maxsize=5% | Out-Null; "Liberado: $(Fmt ([math]::Max(0, (Libre-Sistema) - $antes)))." }
    hibernar = { param($d) & powercfg.exe /h off | Out-Null; 'Hibernación desactivada.' }
    winsxs = { param($d) $antes = Libre-Sistema; & Dism.exe /Online /Cleanup-Image /StartComponentCleanup | Out-Host; "Liberado: $(Fmt ([math]::Max(0, (Libre-Sistema) - $antes)))." }
    escaneo = {
        param($d)
        Update-MpSignature -ErrorAction SilentlyContinue
        Start-MpScan -ScanType QuickScan
        $t = @(Get-MpThreatDetection -ErrorAction SilentlyContinue | Where-Object { $_.InitialDetectionTime -gt (Get-Date).AddHours(-1) })
        "Análisis terminado. Amenazas detectadas ahora: $($t.Count)."
    }
    energia = {
        param($d)
        & powercfg.exe /setactive 8c5e7fda-e8bf-4a96-9a85-cf7e27ae8d8c 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            $g = [regex]::Match(((& powercfg.exe -duplicatescheme 8c5e7fda-e8bf-4a96-9a85-cf7e27ae8d8c) -join ' '), $RxGuid).Value
            if ($g) { & powercfg.exe /setactive $g | Out-Null }
        }
        "Plan activo: $([regex]::Match(((& powercfg.exe /getactivescheme) -join ' '), '\((.+)\)').Groups[1].Value)"
    }
    efectos = {
        param($d)
        Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' -Name VisualFXSetting -Value 3 -Type DWord
        Set-ItemProperty 'HKCU:\Control Panel\Desktop' -Name UserPreferencesMask -Value ([byte[]](0x90, 0x12, 0x03, 0x80, 0x12, 0x00, 0x00, 0x00)) -Type Binary
        Set-ItemProperty 'HKCU:\Control Panel\Desktop\WindowMetrics' -Name MinAnimate -Value '0'
        Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name EnableTransparency -Value 0 -Type DWord
        'Efectos desactivados. Cierra sesión o reinicia para verlo.'
    }
    sfc = { param($d) & Dism.exe /Online /Cleanup-Image /RestoreHealth | Out-Host; & sfc.exe /scannow | Out-Host; 'Revisión terminada (el detalle queda en C:\Windows\Logs\CBS\CBS.log).' }
    optimizar = {
        param($d)
        $letra = $SisDrive.TrimEnd(':')
        if ($d.EsHDD) { Optimize-Volume -DriveLetter $letra -Defrag } else { Optimize-Volume -DriveLetter $letra -ReTrim }
        'Unidad optimizada.'
    }
    onedrive = {
        param($d)
        $antes = Libre-Sistema
        & attrib.exe +U -P "$($d.Ruta)\*" /S /D | Out-Null
        Start-Sleep -Seconds 5
        "OneDrive liberará el espacio en los próximos minutos (liberado hasta ahora: $(Fmt ([math]::Max(0, (Libre-Sistema) - $antes))))."
    }
    outlook = {
        param($d)
        if (-not (Esperar-Cierre @('OUTLOOK') 'Outlook' $d.ForzarCierre)) { return 'Omitido: Outlook seguía abierto.' }
        $k = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode'
        if (-not (Test-Path $k)) { [void](New-Item -Path $k -Force) }
        Set-ItemProperty -Path $k -Name SyncWindowSetting -Value 6 -Type DWord
        $n = 0
        foreach ($f in @($d.Archivos)) {
            try { Rename-Item -LiteralPath $f -NewName ((Split-Path $f -Leaf) + '.anterior') -ErrorAction Stop; $n++ }
            catch { Write-Log "No se pudo renombrar $f : $($_.Exception.Message)" }
        }
        "Listo. Abre Outlook y deja que descargue ($n archivo(s) renombrado(s) como respaldo)."
    }
    inicio = {
        param($d)
        $cambios = @()
        if (Test-Path -LiteralPath $ArchivoCambiosInicio) { $cambios = @(Leer-Json $ArchivoCambiosInicio) }
        $n = 0
        foreach ($p in @($d.Lista)) {
            if (($p.Aprobado -like 'HKLM:*') -and -not $EsAdmin) { Write-Log "Se omite $($p.Nombre): requiere administrador"; continue }
            if ($script:Interactivo) {
                Write-Host ''
                Write-Host "      Programa: $($p.Nombre)" -ForegroundColor White
                Write-Host "      Ubicación: $($p.Comando)" -ForegroundColor DarkGray
                Write-Host "      Sugerencia: $($p.Recomendacion)"
                if (-not (Confirmar '      ¿Desactivar al encender?')) { continue }
            }
            try {
                Set-EstadoInicio $p.Aprobado $p.Nombre $false; $n++
                $cambios += [pscustomobject]@{ Nombre = $p.Nombre; Aprobado = $p.Aprobado; Fecha = (Get-Date).ToString('s') }
                Write-Log "Inicio desactivado: $($p.Nombre)"
            } catch { Write-Log "No se pudo desactivar $($p.Nombre): $($_.Exception.Message)" }
        }
        if ($cambios) { Guardar-Json @($cambios) $ArchivoCambiosInicio }
        "Se desactivaron $n programas. Se aplicará en el próximo encendido."
    }
    duplicados = {
        param($d)
        $revisar = $false
        if ($script:Interactivo) {
            Write-Host '      R = revisar grupo por grupo   |   T = enviar todas las copias a la Papelera' -ForegroundColor Yellow
            $revisar = (Read-Host '      ') -match '^[rR]'
        }
        $n = 0; $b = [double]0; $grandes = 0
        foreach ($g in @($d.Grupos)) {
            if (-not (Test-Path -LiteralPath $g.Conservar)) { Write-Log "Se omite un grupo: la copia a conservar ya no existe ($($g.Conservar))"; continue }
            if ($revisar) {
                Write-Host ''
                Write-Host "      Se conserva: $($g.Conservar)" -ForegroundColor Green
                foreach ($c in @($g.Copias)) { Write-Host "      Copia:       $c" }
                if (-not (Confirmar '      ¿Enviar estas copias a la Papelera?')) { continue }
            }
            foreach ($c in @($g.Copias)) {
                if (-not (Test-Path -LiteralPath $c)) { continue }
                if ([double]$g.Bytes -gt 2GB) { $grandes++; Write-Log "No se mueve (muy grande para la Papelera, bórralo a mano si quieres): $c"; continue }
                $r = [PapeleraPC]::Enviar($c)
                if ($r -eq 0 -and -not (Test-Path -LiteralPath $c)) { $n++; $b += [double]$g.Bytes; Write-Log "A la Papelera: $c" }
                else { Write-Log "No se pudo mover a la Papelera ($r): $c" }
            }
        }
        $extra = if ($grandes) { " $grandes copias de más de 2 GB no se movieron (ver registro)." } else { '' }
        "Se enviaron $n copias a la Papelera ($(Fmt $b)). Puedes recuperarlas desde la Papelera.$extra"
    }
}

function Invoke-Arreglo($a) {
    $acc = $Acciones[[string]$a.Tipo]
    if (-not $acc) { throw "Acción desconocida: $($a.Tipo)" }
    $salida = @(& $acc $a.Datos)
    [string]($salida | Select-Object -Last 1)
}

# ============================================================================
#  Modo: restaurar programas de inicio
# ============================================================================
if ($RestaurarInicio) {
    if (-not (Test-Path -LiteralPath $ArchivoCambiosInicio)) { Write-Host 'No hay cambios de inicio registrados para esta PC.'; Write-Log 'No hay cambios de inicio registrados.'; return }
    $n = 0
    foreach ($c in @(Leer-Json $ArchivoCambiosInicio)) {
        try { Set-EstadoInicio $c.Aprobado $c.Nombre $true; $n++; Write-Host "Reactivado: $($c.Nombre)" -ForegroundColor Green; Write-Log "Reactivado: $($c.Nombre)" }
        catch { Write-Host "No se pudo reactivar $($c.Nombre): $_" -ForegroundColor Red; Write-Log "No se pudo reactivar $($c.Nombre): $_" }
    }
    Rename-Item -LiteralPath $ArchivoCambiosInicio -NewName ("CambiosInicio_${Equipo}_restaurado_$Marca.json")
    Write-Log "Se reactivaron $n programas de inicio."
    return
}

# ============================================================================
#  Modo: aplicar arreglos elegidos en la ventana (sin preguntas)
# ============================================================================
if ($Aplicar) {
    $sel = @(Leer-Json $Aplicar | Sort-Object { [int]$_.Numero })
    Write-Log "Aplicando $($sel.Count) arreglo(s)..."
    $libre0 = Libre-Sistema
    $punto = Nuevo-PuntoRestauracion
    $res = @(); $k = 0
    foreach ($a in $sel) {
        $k++
        Write-Log "($k/$($sel.Count)) $($a.Nombre)..."
        if ($a.Admin -and -not $EsAdmin) { $res += [pscustomobject]@{ Id = $a.Id; Nombre = $a.Nombre; Ok = $false; Mensaje = 'Requiere administrador.' }; continue }
        try { $m = Invoke-Arreglo $a; $ok = $true } catch { $m = $_.Exception.Message; $ok = $false }
        Write-Log ('{0}: {1} -> {2}' -f $(if ($ok) { 'APLICADO' } else { 'ERROR' }), $a.Nombre, $m)
        $res += [pscustomobject]@{ Id = $a.Id; Nombre = $a.Nombre; Ok = $ok; Mensaje = $m }
    }
    $ganado = [math]::Max(0, (Libre-Sistema) - $libre0)
    $salida = if ($ArchivoResultado) { $ArchivoResultado } else { [IO.Path]::ChangeExtension($Aplicar, '.resultado.json') }
    Guardar-Json ([pscustomobject]@{ Resultados = $res; Liberado = $ganado; LiberadoTexto = (Fmt $ganado); PuntoRestauracion = $punto }) $salida
    Write-Log "Fin. Espacio liberado: $(Fmt $ganado)"
    return
}

Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host "   AQUILA v$VersionHerramienta  -  $Equipo" -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan
if (-not $EsAdmin) { Write-Host '  Aviso: sin permisos de administrador algunas mediciones y arreglos no estarán disponibles.' -ForegroundColor Yellow }
Write-Log "Inicio del diagnóstico. Admin=$EsAdmin"

# ============================================================================
#  1. Equipo
# ============================================================================
Paso 'Revisando el equipo (procesador, memoria, Windows)...'
$cs  = Get-CimInstance Win32_ComputerSystem
$os  = Get-CimInstance Win32_OperatingSystem
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
$EsLaptop = [bool](Get-CimInstance Win32_Battery)
$ramGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
$uptime = (Get-Date) - $os.LastBootUpTime
$nLog = [int]$cpu.NumberOfLogicalProcessors; if ($nLog -lt 1) { $nLog = 1 }
$build = [int]$os.BuildNumber

$Datos.Modelo       = ("$($cs.Manufacturer) $($cs.Model)").Trim()
$Datos.Tipo         = if ($EsLaptop) { 'Laptop' } else { 'PC de escritorio' }
$Datos.Procesador   = ([string]$cpu.Name).Trim()
$Datos.Nucleos      = "$($cpu.NumberOfCores) núcleos / $nLog hilos"
$Datos.RAM_GB       = $ramGB
$Datos.Windows      = "$($os.Caption) (compilación $build)"
$Datos.WindowsDesde = $os.InstallDate.ToString('yyyy-MM-dd')
$Datos.DiasEncendida = [math]::Round($uptime.TotalDays, 1)

if ($ramGB -lt 7.5) {
    $nivel = if ($ramGB -lt 4) { 'Critico' } else { 'Alto' }
    Add-Hallazgo 'Memoria' $nivel "Poca memoria RAM ($ramGB GB)" `
        "Esta PC tiene $ramGB GB de RAM. Hoy Windows con navegador, Outlook y Teams necesita al menos 8 GB, idealmente 16 GB." `
        'La RAM es la «mesa de trabajo» de la PC. Cuando se llena, Windows usa el disco como memoria de emergencia, que es decenas de veces más lento. Por eso todo se congela al cambiar de ventana.' `
        'Ampliar la RAM (es una mejora de hardware, generalmente económica). Mientras tanto: menos pestañas abiertas y menos programas al encender.'
}
if ($uptime.TotalDays -ge 7) {
    Add-Hallazgo 'Windows' 'Medio' ('Hace {0:N0} días que no se reinicia' -f $uptime.TotalDays) `
        "Último reinicio: $($os.LastBootUpTime.ToString('dd/MM/yyyy HH:mm'))." `
        'Con los días se acumulan programas colgados y memoria que no se libera. Ojo: «Apagar» en Windows con inicio rápido NO limpia la memoria como «Reiniciar».' `
        'Usar Inicio > Energía > Reiniciar al menos una vez por semana.'
}
if ($build -lt 22000) {
    Add-Hallazgo 'Windows' 'Medio' 'Windows 10 ya no recibe actualizaciones de seguridad' `
        'Microsoft terminó el soporte de Windows 10 en octubre de 2025.' `
        'Sin actualizaciones la PC queda expuesta a virus y fallas que ya no se corrigen.' `
        'Actualizar a Windows 11 si el equipo es compatible, o contratar las actualizaciones extendidas (ESU).'
}

# ============================================================================
#  2. Discos
# ============================================================================
Paso 'Revisando discos (tipo, salud y espacio)...'
$discoSisNum = (Get-Partition -DriveLetter ($SisDrive.TrimEnd(':')) | Get-Disk).Number
$filasDisco = @()
foreach ($pd in Get-PhysicalDisk) {
    $m = [string]$pd.MediaType; $b = [string]$pd.BusType
    $tipo = if ($b -eq '17' -or $b -eq 'NVMe') { 'SSD NVMe' }
            elseif ($m -eq '4' -or $m -eq 'SSD') { 'SSD' }
            elseif ($m -eq '3' -or $m -eq 'HDD') { 'HDD (mecánico)' }
            elseif ([string]$pd.FriendlyName -match 'SSD|NVMe|Solid') { 'SSD' } else { 'Desconocido' }
    $hs = [string]$pd.HealthStatus
    $salud = if ($hs -in '0', 'Healthy') { 'Buena' } elseif ($hs -in '1', 'Warning') { 'Advertencia' } elseif ($hs -in '2', 'Unhealthy') { 'Mala' } else { 'Desconocida' }
    $rel = $pd | Get-StorageReliabilityCounter
    $esSis = ([string]$pd.DeviceId -eq [string]$discoSisNum)
    $filasDisco += [pscustomobject]@{
        'Disco' = $pd.FriendlyName; 'Tipo' = $tipo; 'Tamaño' = (Fmt $pd.Size); 'Salud' = $salud
        'Desgaste' = $(if ($null -ne $rel.Wear) { "$($rel.Wear)%" } else { '-' })
        'Temperatura' = $(if ($rel.Temperature) { "$($rel.Temperature) °C" } else { '-' })
        'Windows aquí' = $(if ($esSis) { 'Sí' } else { '' })
    }
    if ($esSis) { $Datos.DiscoSistema = $tipo }
    if ($salud -in 'Advertencia', 'Mala') {
        Add-Hallazgo 'Disco' 'Critico' "El disco «$($pd.FriendlyName)» reporta salud $($salud.ToLower())" `
            'El propio disco avisa que está fallando.' `
            'Un disco que falla reintenta leer una y otra vez los mismos datos: la PC se congela y existe riesgo real de perder archivos.' `
            'HACER UNA COPIA DE SEGURIDAD YA y reemplazar el disco. Ningún programa lo repara.'
    }
    if ($null -ne $rel.Wear -and [int]$rel.Wear -ge 80) {
        Add-Hallazgo 'Disco' 'Alto' "El disco «$($pd.FriendlyName)» está gastado al $($rel.Wear)%" `
            'Los SSD tienen una vida útil de escritura; este está cerca del final.' `
            'Cerca del final de su vida el disco se vuelve más lento y puede fallar.' 'Planificar el reemplazo y mantener copias de seguridad.'
    }
    if ($esSis -and $tipo -like 'HDD*') {
        Add-Hallazgo 'Disco' 'Critico' 'Windows está instalado en un disco mecánico (HDD)' `
            "El disco del sistema es «$($pd.FriendlyName)», un disco de platos giratorios." `
            'Un disco mecánico es entre 10 y 50 veces más lento que un SSD. Windows 10/11, antivirus, OneDrive y Outlook leen el disco sin parar; con HDD eso se convierte en la causa n.º 1 de lentitud. Es muy probable que por esto la PC «ya no sea como antes».' `
            'Cambiar a un disco SSD (y clonar Windows). Es la mejora de mayor impacto posible, incluso más que la RAM.'
    }
}
Add-Tabla 'Discos' '' $filasDisco @('Disco', 'Tipo', 'Tamaño', 'Salud', 'Desgaste', 'Temperatura', 'Windows aquí')

foreach ($ld in Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3') {
    if (-not $ld.Size) { continue }
    $pct = [math]::Round($ld.FreeSpace / $ld.Size * 100, 1)
    if ($ld.DeviceID -eq $SisDrive) { $Datos.C_LibreGB = [math]::Round($ld.FreeSpace / 1GB, 1); $Datos.C_LibrePct = $pct }
    $nivel = $null
    if ($pct -lt 10 -or $ld.FreeSpace -lt 10GB) { $nivel = 'Critico' } elseif ($pct -lt 20) { $nivel = 'Alto' }
    if ($nivel -and ($ld.DeviceID -eq $SisDrive -or $nivel -eq 'Critico')) {
        Add-Hallazgo 'Disco' $nivel "La unidad $($ld.DeviceID) está casi llena ($pct% libre)" `
            "Quedan $(Fmt $ld.FreeSpace) libres de $(Fmt $ld.Size)." `
            'Windows necesita espacio libre para la memoria de emergencia, las actualizaciones y los archivos temporales. Con el disco lleno todo se ralentiza y las actualizaciones fallan.' `
            'Revisar la sección «Espacio que se puede recuperar»: la mayoría son archivos que ya se «borraron» o que no sirven.'
    }
}

$errDisco = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Level = 1, 2, 3; StartTime = (Get-Date).AddDays(-30) } -ErrorAction SilentlyContinue |
    Where-Object { ($_.ProviderName -eq 'disk' -and $_.Id -in 7, 11, 15, 51, 52, 153, 154) -or
                   ($_.ProviderName -eq 'Ntfs' -and $_.Id -in 55, 57, 137, 140) -or
                   ($_.ProviderName -in 'stornvme', 'storahci', 'iaStorA', 'iaStorAC', 'iaStorAVC' -and $_.Id -eq 129) })
$Datos.ErroresDisco30d = $errDisco.Count
if ($errDisco.Count -gt 0) {
    Add-Hallazgo 'Disco' 'Alto' "Windows registró $($errDisco.Count) errores de disco en los últimos 30 días" `
        "Último: $($errDisco[0].TimeCreated.ToString('dd/MM/yyyy HH:mm')) ($($errDisco[0].ProviderName) #$($errDisco[0].Id))." `
        'Estos errores aparecen cuando el disco tarda en responder o no puede leer una parte. Durante esos segundos la PC se congela.' `
        'Hacer copia de seguridad y ejecutar el arreglo «Revisar y reparar archivos de Windows». Si siguen apareciendo, revisar el disco o su cable.' @('sfc')
}

# ============================================================================
#  3. Medición en vivo
# ============================================================================
Paso "Midiendo el uso en vivo durante $SegundosMuestreo segundos (usa la PC como siempre)..."
$acc = @{}
$cpuTot = @(); $discoOcup = @(); $memDisp = @(); $memComp = @(); $rend = @()
$excluir = @('_Total', 'Idle', 'WmiPrvSE', 'powershell', 'pwsh', 'conhost')
$finMuestreo = (Get-Date).AddSeconds($SegundosMuestreo)
$muestras = 0
while ($muestras -lt 3 -or (Get-Date) -lt $finMuestreo) {
    $muestras++
    $restan = [math]::Max(0, [int]($finMuestreo - (Get-Date)).TotalSeconds)
    Write-Progress -Activity 'Midiendo uso en vivo' -Status "Quedan $restan s" -PercentComplete ([math]::Min(100, (1 - $restan / [math]::Max(1, $SegundosMuestreo)) * 100))
    $p = Get-CimInstance Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'"
    $cpuTot += [double]$p.PercentProcessorTime
    $d = Get-CimInstance Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name='_Total'"
    if ($d) { $discoOcup += [double](100 - [math]::Min(100, $d.PercentIdleTime)) }
    $m = Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
    $memDisp += [double]$m.AvailableMBytes; $memComp += [double]$m.PercentCommittedBytesInUse
    $pi = Get-CimInstance Win32_PerfFormattedData_Counters_ProcessorInformation -Filter "Name='_Total'"
    if ($pi -and $pi.PercentProcessorPerformance) { $rend += [double]$pi.PercentProcessorPerformance }
    $memMuestra = @{}
    foreach ($pr in Get-CimInstance Win32_PerfFormattedData_PerfProc_Process) {
        $n = $pr.Name -replace '#\d+$', ''
        if ($n -in $excluir) { continue }
        if (-not $acc.ContainsKey($n)) { $acc[$n] = [pscustomobject]@{ Nombre = $n; Cpu = 0.0; IO = 0.0; Mem = 0.0 } }
        $acc[$n].Cpu += [double]$pr.PercentProcessorTime
        $acc[$n].IO  += [double]$pr.IODataBytesPersec
        $memMuestra[$n] = [double]$memMuestra[$n] + [double]$pr.WorkingSetPrivate
    }
    foreach ($k in $memMuestra.Keys) { if ($memMuestra[$k] -gt $acc[$k].Mem) { $acc[$k].Mem = $memMuestra[$k] } }
    Start-Sleep -Seconds 1
}
Write-Progress -Activity 'Midiendo uso en vivo' -Completed

$procs = @($acc.Values | ForEach-Object {
    [pscustomobject]@{ Nombre = $_.Nombre; CpuPct = [math]::Round($_.Cpu / $muestras / $nLog, 1); IOs = $_.IO / $muestras; Mem = $_.Mem }
})
$topCpu = @($procs | Sort-Object CpuPct -Descending | Select-Object -First 10)
$topIO  = @($procs | Sort-Object IOs -Descending | Select-Object -First 10)
$topMem = @($procs | Sort-Object Mem -Descending | Select-Object -First 10)
Add-Tabla 'Programas que más usan el procesador' 'Promedio durante la medición (100% = todo el procesador).' `
    ($topCpu | ForEach-Object { [pscustomobject]@{ 'Programa' = (Nombre-Amigable $_.Nombre); 'Procesador' = "$($_.CpuPct)%" } }) @('Programa', 'Procesador')
Add-Tabla 'Programas que más usan el disco' 'Datos leídos/escritos por segundo en promedio.' `
    ($topIO | ForEach-Object { [pscustomobject]@{ 'Programa' = (Nombre-Amigable $_.Nombre); 'Disco' = "$(Fmt $_.IOs)/s" } }) @('Programa', 'Disco')
Add-Tabla 'Programas que más usan memoria RAM' 'Memoria propia (sumando todas sus ventanas/pestañas).' `
    ($topMem | ForEach-Object { [pscustomobject]@{ 'Programa' = (Nombre-Amigable $_.Nombre); 'RAM' = (Fmt $_.Mem) } }) @('Programa', 'RAM')

$cpuProm = Prom $cpuTot; $discoProm = Prom $discoOcup; $memCompProm = Prom $memComp
$memDispMin = if ($memDisp) { ($memDisp | Measure-Object -Minimum).Minimum } else { 0 }
$ramUsoPct = if ($ramGB) { [math]::Round(100 - ($memDispMin / 1024 / $ramGB * 100), 1) } else { 0 }
$Datos.CPU_Prom = $cpuProm; $Datos.Disco_Ocupado_Prom = $discoProm; $Datos.RAM_UsoPct = $ramUsoPct
$rendProm = Prom $rend; $Datos.CPU_VelocidadPct = $rendProm

function Lista-Top($lista, [scriptblock]$valor) { ($lista | Select-Object -First 3 | ForEach-Object { "$(Nombre-Amigable $_.Nombre) ($(& $valor $_))" }) -join '; ' }

if ($cpuProm -ge 60) {
    $t = $topCpu[0]
    Add-Hallazgo 'Rendimiento' 'Alto' "El procesador pasa ocupado ($cpuProm% en promedio)" `
        "Los que más consumen: $(Lista-Top $topCpu { param($x) "$($x.CpuPct)%" })." `
        $(if ($Consejos.ContainsKey($t.Nombre)) { $Consejos[$t.Nombre] } else { 'Cuando el procesador está casi lleno, cada clic tiene que esperar su turno.' }) `
        'Revisar los programas de la lista: cerrar o desactivar al encender los que no se usen. Para atrapar al culpable en el momento exacto, usa «Vigilar».' @('inicio')
}
foreach ($t in $topCpu | Where-Object { $_.CpuPct -ge 20 }) {
    if ($cpuProm -ge 60 -and $t -eq $topCpu[0]) { continue }
    Add-Hallazgo 'Rendimiento' 'Medio' "$(Nombre-Amigable $t.Nombre) consume mucho procesador ($($t.CpuPct)%)" '' `
        $(if ($Consejos.ContainsKey($t.Nombre)) { $Consejos[$t.Nombre] } else { 'Este programa usó una parte importante del procesador durante la medición.' }) `
        'Si no lo usas, ciérralo o desactívalo al encender. Si es del sistema, deja que termine y reinicia.'
}
if ($discoProm -ge 60) {
    $t = $topIO[0]
    Add-Hallazgo 'Rendimiento' 'Alto' "El disco está saturado ($discoProm% del tiempo ocupado)" `
        "Los que más usan el disco: $(Lista-Top $topIO { param($x) "$(Fmt $x.IOs)/s" })." `
        $(if ($Consejos.ContainsKey($t.Nombre)) { $Consejos[$t.Nombre] } else { 'Cuando el disco está al máximo, abrir cualquier programa o archivo tiene que hacer fila. Es la causa más típica de «se queda pensando».' }) `
        $(if ($Datos.DiscoSistema -like 'HDD*') { 'Con disco mecánico esto es casi inevitable: cambiar a SSD. Además, reducir lo que sincroniza (nube/correo).' } else { 'Reducir lo que sincroniza en segundo plano (nube, correo) y revisar los programas de la lista.' })
}
if ($memCompProm -ge 85 -or $ramUsoPct -ge 90) {
    Add-Hallazgo 'Memoria' 'Alto' "La memoria RAM se llena ($ramUsoPct% en uso)" `
        "Los que más memoria usan: $(Lista-Top $topMem { param($x) Fmt $x.Mem })." `
        'Cuando la RAM se llena, Windows mueve cosas al disco para hacer espacio. Cambiar de ventana o pestaña se vuelve lento.' `
        'Cerrar pestañas y programas que no se usan; desactivar programas al encender; si pasa siempre, ampliar la RAM.' @('inicio')
}
if ($rendProm -gt 0 -and $rendProm -lt 55) {
    Add-Hallazgo 'Rendimiento' 'Medio' "El procesador trabaja frenado (a $rendProm% de su velocidad)" `
        'Durante la medición el procesador no llegó a su velocidad normal.' `
        'Windows lo frena para ahorrar energía, o el propio procesador se frena porque está caliente (polvo, ventilador sucio o pasta térmica seca).' `
        'Poner el plan de energía en alto rendimiento (arreglo disponible). Si es una laptop que calienta mucho: limpieza interna y cambio de pasta térmica.' @('energia')
}
if ($cpuProm -lt 60 -and $discoProm -lt 60 -and $ramUsoPct -lt 90) {
    Add-Hallazgo 'Rendimiento' 'OK' 'Durante la medición el uso fue normal' `
        "Procesador $cpuProm%, disco $discoProm%, RAM $ramUsoPct%." `
        'Si la PC se pone lenta en otros momentos (al encender, al abrir Outlook), usa «Vigilar» para atrapar al culpable justo en ese momento.' ''
}

# ============================================================================
#  4. Programas al encender
# ============================================================================
Paso 'Revisando los programas que arrancan solos al encender...'
$fuentes = @(
    @{ Clave = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'; Aprob = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run'; Origen = 'Solo este usuario' }
    @{ Clave = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run'; Aprob = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run'; Origen = 'Todos los usuarios' }
    @{ Clave = 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'; Aprob = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32'; Origen = 'Todos los usuarios' }
)
$ProgramasInicio = @()
foreach ($f in $fuentes) {
    $k = Get-Item -LiteralPath $f.Clave -ErrorAction SilentlyContinue
    if (-not $k) { continue }
    foreach ($n in $k.GetValueNames()) {
        if (-not $n) { continue }
        $ProgramasInicio += [pscustomobject]@{ Nombre = $n; Comando = [string]$k.GetValue($n); Aprobado = $f.Aprob; Origen = $f.Origen }
    }
}
foreach ($c in @(@{ Dir = [Environment]::GetFolderPath('Startup'); Aprob = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder'; Origen = 'Carpeta Inicio (usuario)' },
                 @{ Dir = [Environment]::GetFolderPath('CommonStartup'); Aprob = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder'; Origen = 'Carpeta Inicio (todos)' })) {
    foreach ($a in Get-ChildItem -LiteralPath $c.Dir -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'desktop.ini' }) {
        $ProgramasInicio += [pscustomobject]@{ Nombre = $a.Name; Comando = $a.FullName; Aprobado = $c.Aprob; Origen = $c.Origen }
    }
}
$rxEsencial = 'SecurityHealth|Defender|Realtek|RtkAud|RAVCpl|igfx|Intel|NVIDIA|AMD|Radeon|Synaptics|ETD|Elan|Bluetooth|ctfmon|Dolby|Waves|MaxxAudio|Nahimic|Touchpad|Precision|Lenovo|HPHotkey|Dell|Asus|Acer|Logitech|Antivirus|ESET|Kaspersky|Avast|AVG|Bitdefender|McAfee|Norton|Sophos|Trend|Malwarebytes|VPN|Forti|GlobalProtect|Cisco|Crowd|Sentinel'
$rxNube = 'OneDrive|GoogleDrive|Google Drive|Dropbox|iCloud|Box'
foreach ($p in $ProgramasInicio) {
    $v = (Get-ItemProperty -LiteralPath $p.Aprobado -Name $p.Nombre -ErrorAction SilentlyContinue).($p.Nombre)
    $activo = -not ($v -is [byte[]] -and ($v[0] -band 1))
    $texto = "$($p.Nombre) $($p.Comando)"
    $rec = if ($texto -match $rxEsencial) { 'Mantener (hardware, seguridad o red de la empresa)' }
           elseif ($texto -match $rxNube) { 'Mantener si usas la sincronización' }
           else { 'Se puede desactivar: se abre cuando lo necesites' }
    $p | Add-Member -NotePropertyName Activo -NotePropertyValue $activo
    $p | Add-Member -NotePropertyName Recomendacion -NotePropertyValue $rec
}
$activos = @($ProgramasInicio | Where-Object Activo)
$Datos.InicioActivos = $activos.Count
Add-Tabla 'Programas que arrancan al encender' 'Desactivar uno NO lo desinstala: solo deja de abrirse solo. Es reversible.' `
    ($ProgramasInicio | Sort-Object { -not $_.Activo }, Nombre | ForEach-Object {
        [pscustomobject]@{ 'Programa' = $_.Nombre; 'Estado' = $(if ($_.Activo) { 'Activo' } else { 'Desactivado' }); 'Recomendación' = $_.Recomendacion; 'Origen' = $_.Origen } }) `
    @('Programa', 'Estado', 'Recomendación', 'Origen')

$candidatosInicio = @($activos | Where-Object { $_.Recomendacion -like 'Se puede*' -or $_.Recomendacion -like 'Mantener si*' } |
    Select-Object Nombre, Comando, Aprobado, Origen, Recomendacion)
if ($candidatosInicio.Count) {
    Add-Arreglo -Id 'inicio' -Nombre 'Elegir qué programas NO arrancan al encender' -Riesgo 'Bajo' `
        -QueHace 'Te muestra los programas que se abren solos al encender y desactivas los que elijas.' `
        -Porque 'Cada programa que arranca solo retrasa el encendido y queda consumiendo memoria todo el día aunque no lo uses.' `
        -Afecta 'Esos programas no se abrirán solos; los abres cuando los necesites. No se desinstala nada y se puede revertir (botón «Restaurar programas de inicio» de Aquila, o Administrador de tareas > Aplicaciones de arranque).' `
        -Datos @{ Lista = $candidatosInicio }
}
if ($activos.Count -gt 12) {
    Add-Hallazgo 'Inicio' $(if ($activos.Count -gt 20) { 'Alto' } else { 'Medio' }) "$($activos.Count) programas arrancan solos al encender" `
        'Lo razonable son menos de 10.' `
        'Todos compiten por el disco y el procesador al mismo tiempo durante el encendido, y luego se quedan ocupando memoria aunque no los uses.' `
        'Desactivar los que no necesitas al encender (arreglo disponible, reversible).' @('inicio')
}

# Tiempo de encendido y programas que lo retrasan (registro de Windows)
$evBoot = @(Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-Diagnostics-Performance/Operational'; Id = 100 } -MaxEvents 5 -ErrorAction SilentlyContinue)
$bootMs = @(foreach ($e in $evBoot) { $x = [xml]$e.ToXml(); [double](($x.Event.EventData.Data | Where-Object { $_.Name -eq 'BootTime' }).'#text') })
if ($bootMs.Count) {
    $bootSeg = [math]::Round((Prom $bootMs) / 1000, 0)
    $Datos.Arranque_Seg = $bootSeg
    $evLentos = @(Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-Diagnostics-Performance/Operational'; Id = 101, 102, 103, 106, 109; StartTime = (Get-Date).AddDays(-90) } -ErrorAction SilentlyContinue)
    $lentos = @(foreach ($e in $evLentos) {
        $x = [xml]$e.ToXml(); $h = @{}
        foreach ($dd in $x.Event.EventData.Data) { $h[$dd.Name] = $dd.'#text' }
        $nom = if ($h['FriendlyName']) { $h['FriendlyName'] } elseif ($h['Name']) { $h['Name'] } else { $null }
        if ($nom) { [pscustomobject]@{ Nombre = $nom; Retraso = [double]$h['DegradationTime'] } }
    }) | Group-Object Nombre | ForEach-Object { [pscustomobject]@{ Nombre = $_.Name; Veces = $_.Count; Retraso = [math]::Round((Prom $_.Group.Retraso) / 1000, 1) } } |
        Sort-Object Retraso -Descending | Select-Object -First 10
    if ($lentos) {
        Add-Tabla 'Programas que retrasaron el encendido (últimos 90 días)' 'Según el propio registro de rendimiento de Windows.' `
            ($lentos | ForEach-Object { [pscustomobject]@{ 'Programa / servicio' = $_.Nombre; 'Veces' = $_.Veces; 'Retraso promedio' = "$($_.Retraso) s" } }) @('Programa / servicio', 'Veces', 'Retraso promedio')
    }
    if ($bootSeg -ge 60) {
        Add-Hallazgo 'Inicio' $(if ($bootSeg -ge 120) { 'Alto' } else { 'Medio' }) "Encender la PC tarda unos $bootSeg segundos" `
            $(if ($lentos) { "Los que más lo retrasan: $((($lentos | Select-Object -First 3).Nombre) -join ', ')." } else { '' }) `
            'Un equipo sano con SSD enciende en 15-40 segundos. Lo que más retrasa son los programas al encender, un disco mecánico y los antivirus.' `
            'Desactivar programas al encender y revisar la tabla «Programas que retrasaron el encendido».' @('inicio')
    }
}

# ============================================================================
#  5. Nube (OneDrive, Google Drive, Dropbox)
# ============================================================================
Paso 'Revisando carpetas sincronizadas con la nube (puede tardar)...'
$carpetasNube = @()
foreach ($acc1 in Get-ChildItem 'HKCU:\Software\Microsoft\OneDrive\Accounts' -ErrorAction SilentlyContinue) {
    $uf = (Get-ItemProperty -LiteralPath $acc1.PSPath).UserFolder
    if ($uf -and (Test-Path -LiteralPath $uf)) { $carpetasNube += [pscustomobject]@{ Servicio = 'OneDrive'; Ruta = $uf } }
}
foreach ($ev in @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer)) {
    if ($ev -and (Test-Path -LiteralPath $ev) -and -not ($carpetasNube.Ruta -contains $ev)) { $carpetasNube += [pscustomobject]@{ Servicio = 'OneDrive'; Ruta = $ev } }
}
$dbInfo = Join-Path $env:LOCALAPPDATA 'Dropbox\info.json'
if (Test-Path -LiteralPath $dbInfo) {
    $j = Get-Content -LiteralPath $dbInfo -Raw | ConvertFrom-Json
    foreach ($t in 'personal', 'business') { if ($j.$t.path) { $carpetasNube += [pscustomobject]@{ Servicio = 'Dropbox'; Ruta = $j.$t.path } } }
}
$carpetasNube = @($carpetasNube | Sort-Object Ruta -Unique)
$filasNube = @(); $totalArchNube = 0
foreach ($c in $carpetasNube) {
    Sub "$($c.Servicio): $($c.Ruta)"
    $t = Get-TamanoCarpeta $c.Ruta
    $totalArchNube += $t.Archivos
    $filasNube += [pscustomobject]@{ 'Servicio' = $c.Servicio; 'Carpeta' = $c.Ruta; 'Archivos' = ('{0:N0}' -f $t.Archivos)
                                     'Ocupa en disco' = (Fmt $t.Bytes); 'Solo en la nube' = (Fmt $t.BytesNube) }
    if ($t.Archivos -ge 100000) {
        Add-Hallazgo 'Nube' $(if ($t.Archivos -ge 300000) { 'Alto' } else { 'Medio' }) ("{0} sincroniza {1:N0} archivos" -f $c.Servicio, $t.Archivos) `
            "Carpeta: $($c.Ruta)" `
            'La app de la nube vigila cada archivo para saber si cambió; además el antivirus y el indexador de Windows también los revisan. Con cientos de miles de archivos este trabajo nunca se detiene y consume disco y procesador todo el día.' `
            'Sincronizar solo las carpetas que se usan (en OneDrive: Configuración > Cuenta > Elegir carpetas) y no sincronizar bibliotecas enteras de SharePoint/Teams que no se necesiten.'
    }
    if ($c.Servicio -eq 'OneDrive' -and $t.Bytes -ge 5GB) {
        $id = 'onedrive_' + [math]::Abs($c.Ruta.GetHashCode())
        Add-Arreglo -Id $id -Tipo 'onedrive' -Nombre "Liberar espacio de OneDrive ($($c.Ruta | Split-Path -Leaf))" -Riesgo 'Medio' -Bytes $t.Bytes `
            -QueHace 'Marca los archivos de OneDrive como «solo en línea»: siguen apareciendo en tus carpetas, pero ya no ocupan espacio en el disco.' `
            -Porque 'OneDrive está guardando una copia completa de todo en esta PC; eso llena el disco y hace más pesado el trabajo del antivirus y del indexador.' `
            -Afecta 'Para abrir un archivo necesitarás internet la primera vez (se descarga solo al abrirlo). NO se borra nada de la nube. Si quieres alguno siempre disponible: clic derecho > «Mantener siempre en este dispositivo».' `
            -Datos @{ Ruta = $c.Ruta }
        Add-Hallazgo 'Nube' 'Medio' "OneDrive guarda $(Fmt $t.Bytes) dentro de esta PC" "Carpeta: $($c.Ruta)" `
            'Todo lo que está en la nube también está copiado en el disco. Ocupa espacio y el antivirus e indexador lo revisan.' `
            'Activar «archivos a petición» para que los archivos queden solo en la nube hasta que los abras.' @($id)
    }
}
$Datos.Nube_Archivos = $totalArchNube
if (Get-Process GoogleDriveFS -ErrorAction SilentlyContinue) {
    $gdCache = Get-TamanoCarpeta (Join-Path $env:LOCALAPPDATA 'Google\DriveFS')
    $filasNube += [pscustomobject]@{ 'Servicio' = 'Google Drive'; 'Carpeta' = 'Unidad virtual (caché local)'; 'Archivos' = '-'; 'Ocupa en disco' = (Fmt $gdCache.Bytes); 'Solo en la nube' = '-' }
    if ($gdCache.Bytes -ge 10GB) {
        Add-Hallazgo 'Nube' 'Medio' "La caché de Google Drive ocupa $(Fmt $gdCache.Bytes)" '' `
            'Google Drive guarda copias locales de lo que abriste o marcaste «disponible sin conexión».' `
            'En Google Drive > Configuración, reducir los archivos disponibles sin conexión o mover la caché a otro disco.'
    }
}
if ($filasNube) { Add-Tabla 'Nube: carpetas sincronizadas' '«Solo en la nube» no ocupa espacio en el disco.' $filasNube @('Servicio', 'Carpeta', 'Archivos', 'Ocupa en disco', 'Solo en la nube') }

# ============================================================================
#  6. Correo (Outlook)
# ============================================================================
Paso 'Revisando el correo (Outlook)...'
$docs = [Environment]::GetFolderPath('MyDocuments')
$ost = @(Get-ChildItem -LiteralPath (Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook') -Filter *.ost -File -Force -ErrorAction SilentlyContinue)
$pst = @(foreach ($r in @((Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'), (Join-Path $docs 'Outlook Files'), (Join-Path $docs 'Archivos de Outlook'))) {
    Get-ChildItem -LiteralPath $r -Filter *.pst -File -Force -ErrorAction SilentlyContinue })
$olk = Get-TamanoCarpeta (Join-Path $env:LOCALAPPDATA 'Microsoft\Olk')
$filasCorreo = @()
foreach ($f in $ost) { $filasCorreo += [pscustomobject]@{ 'Archivo' = $f.Name; 'Tipo' = 'OST (copia del servidor)'; 'Tamaño' = (Fmt $f.Length) } }
foreach ($f in $pst) { $filasCorreo += [pscustomobject]@{ 'Archivo' = $f.FullName; 'Tipo' = 'PST (archivo local, ÚNICA copia)'; 'Tamaño' = (Fmt $f.Length) } }
if ($olk.Bytes -gt 0) { $filasCorreo += [pscustomobject]@{ 'Archivo' = 'Caché del nuevo Outlook'; 'Tipo' = 'Caché'; 'Tamaño' = (Fmt $olk.Bytes) } }
if ($filasCorreo) { Add-Tabla 'Correo: archivos de Outlook' '' $filasCorreo @('Archivo', 'Tipo', 'Tamaño') }
$Datos.Outlook_GB = [math]::Round((($ost + $pst) | Measure-Object Length -Sum).Sum / 1GB, 1)

$syncActual = (Get-ItemProperty 'HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode' -ErrorAction SilentlyContinue).SyncWindowSetting
$ostGrandes = @($ost | Where-Object { $_.Length -ge 10GB })
if ($ostGrandes) {
    $mayor = ($ostGrandes | Sort-Object Length -Descending)[0]
    Add-Arreglo -Id 'outlook' -Nombre 'Reducir el correo guardado en esta PC a los últimos 6 meses' -Riesgo 'Medio' `
        -Bytes ([double](($ostGrandes | Measure-Object Length -Sum).Sum) * 0.6) `
        -QueHace 'Configura Outlook para guardar en la PC solo los últimos 6 meses de correo y renombra el archivo de correo actual para que Outlook arme uno nuevo, más liviano.' `
        -Porque 'Outlook guarda en el disco una copia de TODOS tus correos. Con decenas de GB, Outlook, el antivirus y el buscador de Windows trabajan sin parar sobre ese archivo gigante.' `
        -Afecta 'Al abrir Outlook descargará otra vez los últimos 6 meses (puede tardar 30-60 min la primera vez). Los correos antiguos NO se borran: siguen en el servidor y se ven con «Haz clic aquí para ver más en Microsoft Exchange». Si tenías borradores sin sincronizar podrían perderse. El archivo viejo queda renombrado como respaldo (.ost.anterior); bórralo cuando confirmes que todo está bien. Funciona con cuentas Microsoft 365/Exchange. Outlook debe estar cerrado.' `
        -Datos @{ Archivos = @($ostGrandes.FullName); Procesos = @('OUTLOOK'); App = 'Outlook' }
    Add-Hallazgo 'Correo' $(if ($mayor.Length -ge 25GB) { 'Alto' } else { 'Medio' }) "El archivo de correo de Outlook pesa $(Fmt $mayor.Length)" `
        "$($mayor.FullName)$(if ($syncActual) { " · Configurado para guardar $syncActual meses." } else { '' })" `
        'Outlook guarda una copia de todos los correos en la PC. Por encima de ~10 GB, cada vez que llegan correos Outlook, el antivirus y el buscador de Windows reprocesan ese archivo enorme: disco al 100% y Outlook «No responde».' `
        'Guardar en la PC solo los últimos meses (arreglo disponible). Los correos antiguos siguen en el servidor.' @('outlook')
}
foreach ($f in $pst | Where-Object { $_.Length -ge 10GB }) {
    Add-Hallazgo 'Correo' 'Alto' "Archivo de correo local (PST) de $(Fmt $f.Length)" $f.FullName `
        'Un PST es la ÚNICA copia de esos correos (no está en el servidor). Si es muy grande, Outlook se vuelve lento y aumenta el riesgo de que se dañe.' `
        'NO borrarlo. Hacer copia de seguridad y dividirlo por años (Archivo > Herramientas > Limpiar elementos antiguos / Archivar). La herramienta nunca toca los PST.'
}

# ============================================================================
#  7. Espacio recuperable: archivos «eliminados» o que no sirven
# ============================================================================
Paso 'Buscando archivos «borrados» que siguen ocupando espacio y basura acumulada...'
$discosFijos = @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | ForEach-Object { $_.DeviceID })
$navCache = @()
foreach ($base in @('Google\Chrome\User Data', 'Microsoft\Edge\User Data', 'BraveSoftware\Brave-Browser\User Data')) {
    foreach ($sub in 'Cache', 'Code Cache', 'GPUCache', 'Service Worker\CacheStorage') { $navCache += "$env:LOCALAPPDATA\$base\*\$sub" }
}
$navCache += "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles\*\cache2"
$teamsC = @('Cache', 'blob_storage', 'databases', 'GPUCache', 'IndexedDB', 'Local Storage', 'tmp') | ForEach-Object { "$env:APPDATA\Microsoft\Teams\$_" }
$teamsC += "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams"

$Limpiezas = @(
    @{ Id = 'papelera'; Tipo = 'papelera'; Nombre = 'Vaciar la Papelera de reciclaje'; Riesgo = 'Bajo'; Admin = $false; Minimo = 100MB
       Rutas = ($discosFijos | ForEach-Object { "$_\`$Recycle.Bin" })
       QueHace = 'Vacía la Papelera de todas las unidades.'
       Porque = 'Cuando «eliminas» un archivo, Windows solo lo mueve a la Papelera: sigue ocupando exactamente el mismo espacio hasta que la vacías.'
       Afecta = 'Ya no podrás recuperar lo que está en la Papelera. Si tienes dudas, ábrela y revisa antes.' }
    @{ Id = 'temp'; Tipo = 'borrar'; Nombre = 'Borrar archivos temporales'; Riesgo = 'Bajo'; Admin = $false; Minimo = 200MB; DiasMin = 1
       Rutas = @($env:TEMP, "$env:SystemRoot\Temp")
       QueHace = 'Borra archivos temporales de más de 1 día.'
       Porque = 'Los programas crean archivos temporales al instalar, actualizar o trabajar, y casi nunca los borran.'
       Afecta = 'Nada. Los que están en uso ahora mismo se saltan.' }
    @{ Id = 'wu'; Tipo = 'wu'; Nombre = 'Borrar instaladores de actualizaciones ya aplicadas'; Riesgo = 'Bajo'; Admin = $true; Minimo = 300MB
       Rutas = @("$env:SystemRoot\SoftwareDistribution\Download")
       QueHace = 'Detiene Windows Update un momento, borra los instaladores descargados y lo vuelve a iniciar.'
       Porque = 'Son los paquetes de actualizaciones que Windows ya instaló; se quedan guardados sin utilidad.'
       Afecta = 'Nada. Si faltara alguna actualización, Windows la vuelve a descargar.' }
    @{ Id = 'do'; Tipo = 'do'; Nombre = 'Borrar caché de «Optimización de entrega»'; Riesgo = 'Bajo'; Admin = $true; Minimo = 300MB
       Rutas = @("$env:SystemRoot\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache")
       QueHace = 'Borra las copias de actualizaciones que Windows guarda para compartir con otras PCs.'
       Porque = 'Windows guarda actualizaciones para «prestarlas» a otras computadoras de la red. Ocupa espacio sin beneficio para ti.'
       Afecta = 'Nada.' }
    @{ Id = 'windowsold'; Tipo = 'windowsold'; Nombre = 'Borrar la copia de la versión anterior de Windows'; Riesgo = 'Medio'; Admin = $true; Minimo = 500MB
       Rutas = @("$SisDrive\Windows.old", "$SisDrive\`$Windows.~BT", "$SisDrive\`$Windows.~WS")
       QueHace = 'Elimina la carpeta Windows.old y los restos de instalación de la última actualización grande.'
       Porque = 'Tras una actualización grande Windows guarda una copia completa del Windows anterior «por si acaso». Suele pesar entre 10 y 30 GB y nadie la borra.'
       Afecta = 'Ya no podrás usar «Volver a la versión anterior de Windows». Tus archivos actuales no se tocan. Si tras la actualización te faltó algún archivo, búscalo antes en Windows.old\Users.' }
    @{ Id = 'dumps'; Tipo = 'borrar'; Nombre = 'Borrar reportes de errores y volcados de memoria'; Riesgo = 'Bajo'; Admin = $true; Minimo = 200MB
       Rutas = @("$env:SystemRoot\MEMORY.DMP", "$env:SystemRoot\Minidump", "$env:SystemRoot\LiveKernelReports", "$env:ProgramData\Microsoft\Windows\WER\ReportArchive", "$env:ProgramData\Microsoft\Windows\WER\ReportQueue", "$env:LOCALAPPDATA\CrashDumps")
       QueHace = 'Borra los archivos que Windows genera cuando un programa o el sistema se cae.'
       Porque = 'Cada vez que algo falla, Windows guarda una «foto» de la memoria para diagnóstico. Pueden pesar varios GB y quedan para siempre.'
       Afecta = 'Nada en el uso diario. Solo sirven a un técnico para investigar fallas pasadas.' }
    @{ Id = 'navegadores'; Tipo = 'borrar'; Nombre = 'Limpiar la caché de los navegadores'; Riesgo = 'Bajo'; Admin = $false; Minimo = 500MB
       Rutas = $navCache; Procesos = @('chrome', 'msedge', 'brave', 'firefox'); App = 'El navegador (Chrome/Edge/Firefox)'
       QueHace = 'Borra las copias de páginas web guardadas por Chrome, Edge, Brave y Firefox.'
       Porque = 'Los navegadores guardan copias de las páginas visitadas y esa caché crece hasta varios GB, con mucha basura de sitios que ya no visitas.'
       Afecta = 'NO se borran contraseñas, historial, favoritos ni sesiones iniciadas. Las páginas pueden tardar un poco más solo la primera vez. Hay que cerrar el navegador.' }
    @{ Id = 'teams'; Tipo = 'borrar'; Nombre = 'Limpiar la caché de Microsoft Teams'; Riesgo = 'Bajo'; Admin = $false; Minimo = 300MB
       Rutas = $teamsC; Procesos = @('ms-teams', 'Teams'); App = 'Teams'
       QueHace = 'Borra la caché local de Teams (clásico y nuevo).'
       Porque = 'Teams acumula imágenes, videos y datos temporales que lo vuelven lento y a veces causan fallos.'
       Afecta = 'Teams tardará un poco más la primera vez y puede pedirte iniciar sesión. Chats, archivos y reuniones están en la nube: no se pierden. Hay que cerrar Teams.' }
    @{ Id = 'miniaturas'; Tipo = 'borrar'; Nombre = 'Reiniciar la caché de miniaturas'; Riesgo = 'Bajo'; Admin = $false; Minimo = 300MB
       Rutas = @("$env:LOCALAPPDATA\Microsoft\Windows\Explorer"); Filtro = 'thumbcache_*.db'
       QueHace = 'Borra las miniaturas guardadas de fotos, videos y documentos.'
       Porque = 'Windows guarda miniaturas incluso de archivos que ya no existen y ese archivo crece sin límite.'
       Afecta = 'Las carpetas con fotos tardarán un segundo más en mostrar miniaturas la primera vez.' }
    @{ Id = 'parches'; Tipo = 'borrar'; Nombre = 'Borrar la caché de parches de instaladores'; Riesgo = 'Medio'; Admin = $true; Minimo = 500MB
       Rutas = @("$env:SystemRoot\Installer\`$PatchCache`$")
       QueHace = 'Borra la carpeta de respaldo de parches de Office y otros programas.'
       Porque = 'Son copias de respaldo de parches ya instalados. Microsoft indica que se puede borrar.'
       Afecta = 'Si algún día reparas Office u otro programa, podría pedirte el instalador original.' }
    @{ Id = 'logs'; Tipo = 'borrar'; Nombre = 'Borrar registros antiguos de Windows'; Riesgo = 'Bajo'; Admin = $true; Minimo = 300MB; DiasMin = 7
       Rutas = @("$env:SystemRoot\Logs\CBS", "$env:SystemRoot\Logs\DISM", "$env:SystemRoot\Logs\WindowsUpdate")
       QueHace = 'Borra archivos de registro de Windows de más de 7 días.'
       Porque = 'Algunos registros (como CBS) pueden crecer hasta decenas de GB por un error de Windows.'
       Afecta = 'Nada.' }
)
$filasEspacio = @(); $totalRecuperable = [double]0
foreach ($L in $Limpiezas) {
    $rutas = @(Resolve-Rutas $L.Rutas)
    $bytes = [double]0
    foreach ($r in $rutas) { $bytes += (Get-TamanoCarpeta -Ruta $r -Filtro $L.Filtro).Bytes }
    if ($bytes -le 0) { continue }
    $filasEspacio += [pscustomobject]@{ 'Qué es' = $L.Nombre; 'Tamaño' = (Fmt $bytes); 'Por qué sobra' = $L.Porque }
    if ($bytes -ge $L.Minimo) {
        $totalRecuperable += $bytes
        Add-Arreglo -Id $L.Id -Tipo $L.Tipo -Nombre $L.Nombre -Riesgo $L.Riesgo -Bytes $bytes -QueHace $L.QueHace -Porque $L.Porque -Afecta $L.Afecta -Admin:$L.Admin `
            -Datos @{ Rutas = $rutas; DiasMin = [int]$L.DiasMin; Filtro = $L.Filtro; Procesos = $L.Procesos; App = $L.App }
    }
}
Sub 'Carpetas de basura medidas.'

# Puntos de restauración (instantáneas) ocupando espacio
$ssUsado = [double](Get-CimInstance Win32_ShadowStorage | Measure-Object UsedSpace -Sum).Sum
$nSombras = @(Get-CimInstance Win32_ShadowCopy).Count
if ($ssUsado -gt 0) {
    $filasEspacio += [pscustomobject]@{ 'Qué es' = "Puntos de restauración ($nSombras)"; 'Tamaño' = (Fmt $ssUsado); 'Por qué sobra' = 'Copias del sistema para deshacer cambios. Solo se necesitan los más recientes.' }
    $cSize = [double](Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$SisDrive'").Size
    $exceso = $ssUsado - ($cSize * 0.05)
    if ($exceso -ge 5GB) {
        $totalRecuperable += $exceso
        Add-Arreglo -Id 'puntos' -Nombre 'Limitar el espacio de los puntos de restauración al 5% del disco' -Riesgo 'Medio' -Bytes $exceso -Admin `
            -QueHace 'Pone un tope de 5% del disco a los puntos de restauración; Windows borra los más antiguos que no entren.' `
            -Porque "Los puntos de restauración ocupan $(Fmt $ssUsado). Son copias de seguridad del sistema, pero las viejas casi nunca se usan." `
            -Afecta 'Se pierden los puntos de restauración MÁS ANTIGUOS. Se conservan los recientes (incluido el que crea esta herramienta antes de hacer cambios).'
    }
}

# Archivo de hibernación (solo si falta espacio)
$hiber = Get-Item -LiteralPath "$SisDrive\hiberfil.sys" -Force -ErrorAction SilentlyContinue
if ($hiber -and $hiber.Length -gt 0) {
    $filasEspacio += [pscustomobject]@{ 'Qué es' = 'Archivo de hibernación (hiberfil.sys)'; 'Tamaño' = (Fmt $hiber.Length); 'Por qué sobra' = 'Guarda la memoria al hibernar. Solo sobra si no usas la hibernación.' }
    if ($Datos.C_LibrePct -lt 20) {
        Add-Arreglo -Id 'hibernar' -Nombre 'Desactivar la hibernación' -Riesgo 'Medio' -Bytes $hiber.Length -Admin `
            -QueHace 'Desactiva la hibernación y borra el archivo hiberfil.sys.' `
            -Porque "El disco está casi lleno y este archivo reserva $(Fmt $hiber.Length) de forma permanente." `
            -Afecta $(if ($EsLaptop) { 'La laptop ya no podrá «Hibernar» (Suspender sigue funcionando) y se desactiva el inicio rápido: encender puede tardar unos segundos más. Se revierte con: powercfg /h on' } else { 'Se desactiva la hibernación y el inicio rápido (el encendido puede tardar unos segundos más). Se revierte con: powercfg /h on' })
    }
}

# Componentes antiguos de Windows (WinSxS)
if ($EsAdmin) {
    Add-Arreglo -Id 'winsxs' -Nombre 'Limpiar componentes antiguos de Windows (WinSxS)' -Riesgo 'Bajo' -Admin `
        -QueHace 'Usa la herramienta oficial de Microsoft (DISM) para borrar versiones viejas de componentes que las actualizaciones reemplazaron.' `
        -Porque 'Cada actualización deja la versión anterior de los archivos del sistema. Con los años pueden ser varios GB.' `
        -Afecta 'Ya no se podrán desinstalar las actualizaciones que ya están instaladas. Tarda 5-20 minutos.'
}

$Datos.Recuperable_GB = [math]::Round($totalRecuperable / 1GB, 1)
Sub 'Limpiezas medidas.'
if ($filasEspacio) { Add-Tabla 'Espacio que se puede recuperar' 'Incluye archivos que «borraste» pero siguen en el disco, y copias que Windows guarda y nunca limpia.' $filasEspacio @('Qué es', 'Tamaño', 'Por qué sobra') }
if ($totalRecuperable -ge 5GB) {
    Add-Hallazgo 'Espacio' $(if ($totalRecuperable -ge 20GB) { 'Alto' } else { 'Medio' }) "Hay $(Fmt $totalRecuperable) de archivos que no sirven" `
        'Papelera, temporales, copias de Windows anteriores, cachés y volcados de errores.' `
        'Son archivos que ya «eliminaste» o que Windows y los programas dejan acumulados. No los ves, pero ocupan disco, el antivirus los revisa y el indexador los cataloga.' `
        'Aplicar los arreglos de limpieza: cada uno explica qué borra y si afecta algo.' @($Arreglos | Where-Object { $_.Bytes -gt 0 } | ForEach-Object Id)
}

# Tu carpeta de usuario: archivos grandes y Descargas antiguas
Sub 'Buscando archivos grandes en tu carpeta de usuario...'
$perfil = Get-TamanoCarpeta -Ruta $env:USERPROFILE -MinGrande 500MB -Excluir @($carpetasNube.Ruta)
$Datos.Perfil_GB = [math]::Round($perfil.Bytes / 1GB, 1)
$grandes = @($perfil.Grandes | Sort-Object Bytes -Descending | Select-Object -First 15)
if ($grandes) {
    Add-Tabla 'Archivos grandes en tu usuario (más de 500 MB, sin contar la nube)' 'Solo informativo: la herramienta NUNCA borra tus documentos. Revisa si alguno ya no lo necesitas.' `
        ($grandes | ForEach-Object { [pscustomobject]@{ 'Archivo' = $_.Ruta; 'Tamaño' = (Fmt $_.Bytes); 'Última modificación' = $_.Fecha.ToString('dd/MM/yyyy') } }) @('Archivo', 'Tamaño', 'Última modificación')
}
$desc = Join-Path $env:USERPROFILE 'Downloads'
$viejos = @(Get-ChildItem -LiteralPath $desc -File -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) })
$bViejos = [double]($viejos | Measure-Object Length -Sum).Sum
if ($bViejos -ge 2GB) {
    Add-Hallazgo 'Espacio' 'Info' "La carpeta Descargas tiene $(Fmt $bViejos) en archivos de más de 3 meses" "$($viejos.Count) archivos en $desc." `
        'Instaladores, adjuntos y archivos descargados una vez que quedan olvidados.' 'Revisar y borrar manualmente lo que no sirva (la herramienta no borra tus archivos).'
}

# Restos de programas desinstalados
Sub 'Buscando restos de programas desinstalados...'
$regUn = @('HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*')
$inst = @(Get-ItemProperty $regUn -ErrorAction SilentlyContinue | Where-Object DisplayName)
$Datos.ProgramasInstalados = $inst.Count
$textoInst = ((@($inst | ForEach-Object { "$($_.DisplayName) $($_.Publisher) $($_.InstallLocation)" }) + @(Get-AppxPackage -ErrorAction SilentlyContinue | ForEach-Object Name)) -join '|').ToLower() -replace '[^a-z0-9|]', ''
$rutasProc = ((Get-Process | ForEach-Object Path | Where-Object { $_ }) -join '|').ToLower()
$rxSistema = '^(microsoft|windows|packages|temp|comms|connecteddevicesplatform|d3dscache|crashdumps|programs|virtualstore|publishers|common files|windowsapps|modifiablewindowsapps|reference assemblies|msbuild|dotnet|uninstall information|package cache|packagemanagement|internet explorer|ssh|regid|usoshared|usoprivate|softwaredistribution|identities|history|fontcache|peernet|placeholdertilelogofolder|desktop\.ini|default|application data|ssl|nuget|pip|npm|npm-cache|mozilla|google|adobe|intel|nvidia|amd|realtek)'
$restos = @()
foreach ($base in @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramData, $env:APPDATA, $env:LOCALAPPDATA) | Select-Object -Unique) {
    foreach ($dir in Get-ChildItem -LiteralPath $base -Directory -Force -ErrorAction SilentlyContinue) {
        $n = $dir.Name.ToLower()
        if ($n -match $rxSistema) { continue }
        $norm = $n -replace '[^a-z0-9]', ''
        if ($norm.Length -lt 4) { continue }
        if ($textoInst.Contains($norm)) { continue }
        if ($rutasProc.Contains($dir.FullName.ToLower())) { continue }
        if ($dir.LastWriteTime -gt (Get-Date).AddDays(-60)) { continue }
        $t = Get-TamanoCarpeta $dir.FullName
        if ($t.Bytes -ge 100MB) { $restos += [pscustomobject]@{ Ruta = $dir.FullName; Bytes = $t.Bytes; Fecha = $dir.LastWriteTime } }
    }
}
if ($restos) {
    $restos = @($restos | Sort-Object Bytes -Descending)
    $bRestos = [double]($restos | Measure-Object Bytes -Sum).Sum
    Add-Tabla 'Posibles restos de programas desinstalados' 'Carpetas de más de 100 MB que no corresponden a ningún programa instalado y no se usan hace más de 2 meses. REVISAR antes de borrar: puede haber falsos positivos.' `
        ($restos | ForEach-Object { [pscustomobject]@{ 'Carpeta' = $_.Ruta; 'Tamaño' = (Fmt $_.Bytes); 'Sin uso desde' = $_.Fecha.ToString('dd/MM/yyyy') } }) @('Carpeta', 'Tamaño', 'Sin uso desde')
    Add-Hallazgo 'Espacio' 'Info' "Posibles restos de programas desinstalados: $(Fmt $bRestos)" "$($restos.Count) carpetas (ver tabla)." `
        'Al desinstalar, muchos programas dejan sus carpetas de datos. Siguen ocupando espacio aunque el programa ya no exista.' `
        'Revisar la tabla; si reconoces que el programa ya no está, borra la carpeta manualmente. La herramienta no las borra sola porque podría equivocarse.'
}

# ============================================================================
#  8. Archivos duplicados
# ============================================================================
Paso 'Buscando archivos duplicados en tus carpetas (puede tardar unos minutos)...'
if ($SinDuplicados) {
    Sub 'Omitido (-SinDuplicados).'
} else {
    $carpetasDup = @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('MyDocuments'), $desc,
                     [Environment]::GetFolderPath('MyPictures'), [Environment]::GetFolderPath('MyVideos'), [Environment]::GetFolderPath('MyMusic')) + @($carpetasNube.Ruta)
    $dup = Buscar-Duplicados -Carpetas $carpetasDup -MinBytes 1MB -Nube @($carpetasNube.Ruta)
    $gDup = @($dup.Grupos | Sort-Object { [double]$_.Bytes * @($_.Copias).Count } -Descending)
    $bDup = [double]0; foreach ($g in $gDup) { $bDup += [double]$g.Bytes * @($g.Copias).Count }
    $Datos.Duplicados_GB = [math]::Round($bDup / 1GB, 1)
    Sub "Duplicados: $($gDup.Count) grupos, $(Fmt $bDup) repetidos ($($dup.Segundos) s)."
    if ($gDup) {
        $nota = 'Se compara el contenido real de los archivos (no solo el nombre). Solo archivos de 1 MB o más, guardados en esta PC (no los que están solo en la nube).'
        if ($dup.Incompleto) { $nota += ' La búsqueda se detuvo por tiempo: se revisaron primero los archivos más grandes.' }
        Add-Tabla 'Archivos duplicados (copias idénticas)' $nota `
            ($gDup | Select-Object -First 40 | ForEach-Object {
                [pscustomobject]@{ 'Archivo' = (Split-Path $_.Conservar -Leaf); 'Copias extra' = @($_.Copias).Count; 'Tamaño c/u' = (Fmt $_.Bytes)
                                   'Se conserva' = $_.Conservar; 'Copias' = (@($_.Copias) -join "  |  ") } }) `
            @('Archivo', 'Copias extra', 'Tamaño c/u', 'Se conserva', 'Copias')
        if ($bDup -ge 100MB) {
            Add-Arreglo -Id 'duplicados' -Nombre 'Enviar a la Papelera las copias duplicadas' -Riesgo 'Medio' -Bytes $bDup `
                -QueHace 'Envía a la Papelera las copias repetidas y conserva una de cada archivo (prioriza la que está en la nube, fuera de Descargas y con el nombre original).' `
                -Porque "Hay $(Fmt $bDup) en archivos que están repetidos con exactamente el mismo contenido." `
                -Afecta 'Las copias van a la PAPELERA: si te equivocas, las recuperas desde ahí (no vacíes la Papelera hasta revisar). Si una copia está en OneDrive, también se quita de la nube (queda en la papelera de OneDrive). A veces una copia en otra carpeta es intencional (por ejemplo, un archivo enviado a un cliente): revisa la lista antes. Las copias de más de 2 GB no se mueven solas.' `
                -Datos @{ Grupos = @($gDup | Select-Object -First 500 | ForEach-Object { [pscustomobject]@{ Bytes = $_.Bytes; Conservar = $_.Conservar; Copias = @($_.Copias) } }) }
            Add-Hallazgo 'Espacio' $(if ($bDup -ge 5GB) { 'Medio' } else { 'Info' }) "Hay $(Fmt $bDup) en archivos duplicados" "$($gDup.Count) archivos tienen al menos una copia idéntica." `
                'Descargar el mismo adjunto varias veces, copiar carpetas «por si acaso» o tener la misma carpeta dentro y fuera de OneDrive deja copias exactas que ocupan espacio y que la nube y el antivirus procesan dos veces.' `
                'Revisar la lista y enviar las copias a la Papelera (arreglo disponible, recuperable).' @('duplicados')
        }
    }
}

# ============================================================================
#  9. Seguridad
# ============================================================================
Paso 'Revisando antivirus y seguridad...'
$av = @(Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction SilentlyContinue)
$avActivos = @($av | Where-Object { ([int]$_.productState -band 0x1000) -ne 0 })
$Datos.Antivirus = (($avActivos | ForEach-Object displayName) -join ', ')
$mp = Get-MpComputerStatus -ErrorAction SilentlyContinue
if ($mp -and $mp.AntivirusEnabled) {
    Add-Arreglo -Id 'escaneo' -Nombre 'Buscar virus (análisis rápido de Windows Defender)' -Riesgo 'Bajo' -Admin `
        -QueHace 'Ejecuta un análisis rápido con el antivirus de Windows.' `
        -Porque 'Descarta que un virus o programa malicioso esté consumiendo recursos.' `
        -Afecta 'Nada. La PC puede ir algo más lenta durante el análisis (5-15 min). Si encuentra algo, Defender lo pone en cuarentena.'
}
if ($av.Count -and -not $avActivos.Count) {
    Add-Hallazgo 'Seguridad' 'Critico' 'No hay ningún antivirus activo' '' 'Sin antivirus, un virus o minero de criptomonedas puede estar usando la PC sin que lo notes.' 'Activar Windows Defender o el antivirus de la empresa.' @('escaneo')
}
if ($avActivos.Count -ge 2) {
    Add-Hallazgo 'Seguridad' 'Alto' "Hay $($avActivos.Count) antivirus activos a la vez" ($Datos.Antivirus) `
        'Dos antivirus revisan cada archivo dos veces y a veces se bloquean entre sí: el disco y el procesador se duplican.' `
        'Dejar solo uno. Si la empresa instaló el suyo, desinstalar los demás (Windows Defender se desactiva solo).'
}
if ($mp) {
    if ($mp.QuickScanAge -gt 14 -and $mp.QuickScanAge -lt 10000) {
        Add-Hallazgo 'Seguridad' 'Medio' "El antivirus no hace un análisis hace $($mp.QuickScanAge) días" '' `
            'Un programa malicioso (por ejemplo, un minero de criptomonedas) puede ser la causa oculta de la lentitud.' 'Ejecutar un análisis rápido (arreglo disponible).' @('escaneo')
    }
    if ($mp.AntivirusSignatureAge -gt 7) {
        Add-Hallazgo 'Seguridad' 'Medio' "Las firmas del antivirus tienen $($mp.AntivirusSignatureAge) días" '' 'Con firmas viejas el antivirus no reconoce amenazas nuevas.' 'Conectar a internet y ejecutar Windows Update.'
    }
}

# ============================================================================
#  10. Configuración de Windows
# ============================================================================
Paso 'Revisando configuración de Windows (energía, efectos, reinicios pendientes)...'
$pend = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
        (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
if ($pend) {
    Add-Hallazgo 'Windows' 'Medio' 'Hay actualizaciones esperando un reinicio' '' `
        'Mientras no reinicies, Windows sigue trabajando en segundo plano para terminar de instalarlas.' 'Reiniciar la PC.'
}
$planTxt = (& powercfg.exe /getactivescheme) -join ' '
$planGuid = [regex]::Match($planTxt, $RxGuid).Value
$planNom = [regex]::Match($planTxt, '\((.+)\)').Groups[1].Value
$Datos.PlanEnergia = $planNom
if ($planGuid -ne '8c5e7fda-e8bf-4a96-9a85-cf7e27ae8d8c' -and $planGuid -ne 'e9a42b02-d5df-448d-aa00-03f14749eb61') {
    Add-Arreglo -Id 'energia' -Nombre 'Plan de energía «Alto rendimiento»' -Riesgo 'Bajo' -Admin `
        -QueHace 'Cambia el plan de energía a «Alto rendimiento» para que el procesador no se frene.' `
        -Porque "El plan actual («$planNom») frena el procesador para ahorrar energía." `
        -Afecta $(if ($EsLaptop) { 'Con batería, la laptop durará menos y el ventilador puede sonar más. Recomendado si trabaja casi siempre enchufada. Se revierte en Panel de control > Opciones de energía.' } else { 'Consume un poco más de electricidad. Se revierte en Panel de control > Opciones de energía.' })
    if ($planGuid -eq 'a1841308-3541-4fab-bc81-f71556f20b4a') {
        Add-Hallazgo 'Windows' 'Alto' 'La PC está en modo «Economizador de energía»' '' `
            'Este modo limita a propósito la velocidad del procesador. La PC se siente mucho más lenta.' 'Cambiar a «Alto rendimiento» o «Equilibrado» (arreglo disponible).' @('energia')
    }
}
$vfx = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' -ErrorAction SilentlyContinue).VisualFXSetting
if ($vfx -ne 2 -and $vfx -ne 3) {
    Add-Arreglo -Id 'efectos' -Nombre 'Quitar animaciones y efectos visuales' -Riesgo 'Bajo' `
        -QueHace 'Configura Windows para priorizar el rendimiento (sin animaciones, sombras ni transparencias), manteniendo las fuentes suavizadas.' `
        -Porque 'Las animaciones consumen procesador, gráficos y memoria. En equipos modestos se nota al abrir y minimizar ventanas.' `
        -Afecta 'Windows se ve más simple (sin animaciones). Se aplica al cerrar sesión. Se revierte en Sistema > Configuración avanzada > Rendimiento.'
}
if ($EsAdmin) {
    Add-Arreglo -Id 'sfc' -Nombre 'Revisar y reparar archivos de Windows' -Riesgo 'Bajo' -Admin `
        -QueHace 'Ejecuta las herramientas oficiales de Microsoft (DISM y SFC) que comparan los archivos de Windows con los originales y reparan los dañados.' `
        -Porque 'Apagones, discos con errores o actualizaciones interrumpidas pueden dañar archivos del sistema y causar lentitud o errores.' `
        -Afecta 'Nada. Tarda entre 15 y 40 minutos; se puede seguir usando la PC.'
    $esHDD = ($Datos.DiscoSistema -like 'HDD*')
    Add-Arreglo -Id 'optimizar' -Nombre "Optimizar la unidad $SisDrive" -Riesgo 'Bajo' -Admin `
        -QueHace $(if ($esHDD) { 'Desfragmenta el disco (junta las partes dispersas de los archivos).' } else { 'Ejecuta TRIM en el SSD (le avisa qué espacio está libre para que siga siendo rápido).' }) `
        -Porque 'Mantiene el disco trabajando a su velocidad normal, sobre todo después de borrar muchos archivos.' `
        -Afecta 'Nada. En un disco mecánico puede tardar bastante.' -Datos @{ EsHDD = $esHDD }
}

# ============================================================================
#  11. Puntaje, orden, guardado e informe
# ============================================================================
Paso 'Generando el informe...'
$peso = @{ Critico = 20; Alto = 10; Medio = 4; Info = 0; OK = 0 }
$orden = @{ Critico = 0; Alto = 1; Medio = 2; Info = 3; OK = 4 }
$puntaje = [math]::Max(0, 100 - [int](($Hallazgos | ForEach-Object { $peso[$_.Nivel] } | Measure-Object -Sum).Sum))
$estado = if ($puntaje -ge 80) { 'Buena' } elseif ($puntaje -ge 60) { 'Regular' } elseif ($puntaje -ge 40) { 'Mala' } else { 'Crítica' }
$Datos.Puntaje = $puntaje
$HallOrd = @($Hallazgos | Sort-Object { $orden[$_.Nivel] })
# Orden: riesgo bajo primero y luego por espacio. Así la Papelera se vacía ANTES de mandar duplicados a ella.
$i = 0
foreach ($a in @($Arreglos | Sort-Object { @{ Bajo = 0; Medio = 1; Alto = 2 }[$_.Riesgo] }, { -$_.Bytes })) { $i++; $a.Numero = $i }
$ArrOrd = @($Arreglos | Sort-Object Numero)

$comp = $null
if ($Comparar -and (Test-Path -LiteralPath $Comparar)) { $comp = (Leer-Json $Comparar).Datos }

function Html-Tabla($t) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<section class='bloque'><h3>$(Esc $t.Titulo)</h3>")
    if ($t.Nota) { [void]$sb.Append("<p class='nota'>$(Esc $t.Nota)</p>") }
    if (-not $t.Filas.Count) { [void]$sb.Append("<p class='nota'>Sin datos.</p></section>"); return $sb.ToString() }
    [void]$sb.Append("<div class='tabla'><table><thead><tr>")
    foreach ($c in $t.Columnas) { [void]$sb.Append("<th>$(Esc $c)</th>") }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($f in $t.Filas) { [void]$sb.Append('<tr>'); foreach ($c in $t.Columnas) { [void]$sb.Append("<td>$(Esc ([string]$f.$c))</td>") }; [void]$sb.Append('</tr>') }
    [void]$sb.Append('</tbody></table></div></section>')
    $sb.ToString()
}
$etq = @{ Critico = 'Crítico'; Alto = 'Alto'; Medio = 'Medio'; Info = 'Info'; OK = 'Bien' }
$colorPuntaje = if ($puntaje -ge 80) { 'var(--ok)' } elseif ($puntaje -ge 60) { 'var(--medio)' } elseif ($puntaje -ge 40) { 'var(--alto)' } else { 'var(--crit)' }

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append("<!DOCTYPE html><html lang='es'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Aquila · Diagnóstico $(Esc $Equipo)</title><link rel='icon' href='data:image/svg+xml,$([uri]::EscapeDataString($AquilaLogoSvg))'><style>$CssInforme</style></head><body><div class='wrap'>")
[void]$sb.Append("<header><div class='gauge' style='--p:$puntaje;--c:$colorPuntaje'><div><div><b>$puntaje</b><br><span>de 100</span></div></div></div>")
[void]$sb.Append("<div><div class='marca'>$AquilaLogoSvg<span>Aquila</span></div><h1>Diagnóstico de $(Esc $Equipo)</h1><div class='sub'>Salud: <b style='color:$colorPuntaje'>$estado</b> · $(Esc $Datos.Fecha) · $(Esc $Datos.Modelo)</div></div></header>")

$datosVista = [ordered]@{ 'Tipo' = $Datos.Tipo; 'Procesador' = $Datos.Procesador; 'Núcleos' = $Datos.Nucleos; 'Memoria RAM' = "$($Datos.RAM_GB) GB"
    'Disco de Windows' = $Datos.DiscoSistema; "Libre en $SisDrive" = "$($Datos.C_LibreGB) GB ($($Datos.C_LibrePct)%)"; 'Windows' = $Datos.Windows
    'Encendido' = $(if ($Datos.Arranque_Seg) { "$($Datos.Arranque_Seg) s" } else { 'Sin datos' }); 'Días sin reiniciar' = $Datos.DiasEncendida
    'Programas al encender' = $Datos.InicioActivos; 'Espacio recuperable' = "$($Datos.Recuperable_GB) GB"
    'Duplicados' = $(if ($null -ne $Datos.Duplicados_GB) { "$($Datos.Duplicados_GB) GB" } else { 'No revisado' }); 'Antivirus' = $Datos.Antivirus }
[void]$sb.Append("<div class='grid'>")
foreach ($k in $datosVista.Keys) { [void]$sb.Append("<div class='dato'><small>$(Esc $k)</small>$(Esc ([string]$datosVista[$k]))</div>") }
[void]$sb.Append('</div>')

$principales = @($HallOrd | Where-Object { $_.Nivel -in 'Critico', 'Alto' } | Select-Object -First 6)
[void]$sb.Append('<h2>Por qué está lenta: causas principales</h2><div class=''causas''>')
if ($principales) { [void]$sb.Append('<ol>'); foreach ($h in $principales) { [void]$sb.Append("<li><b>$(Esc $h.Titulo)</b> — $(Esc $h.Porque)</li>") }; [void]$sb.Append('</ol>') }
else { [void]$sb.Append('<p>No se encontraron problemas graves. Revisa los puntos de nivel medio más abajo.</p>') }
[void]$sb.Append('</div>')

[void]$sb.Append('<h2>Todo lo que se encontró</h2>')
foreach ($h in $HallOrd) {
    [void]$sb.Append("<article class='hallazgo n-$($h.Nivel)'><div class='cab'><span class='badge'>$($etq[$h.Nivel])</span><span class='cat'>$(Esc $h.Categoria)</span></div><h3>$(Esc $h.Titulo)</h3>")
    if ($h.Detalle) { [void]$sb.Append("<p class='detalle'>$(Esc $h.Detalle)</p>") }
    if ($h.Porque) { [void]$sb.Append("<p class='porque'><b>¿Por qué pasa?</b> $(Esc $h.Porque)</p>") }
    if ($h.Solucion) { [void]$sb.Append("<p class='solucion'><b>Qué hacer:</b> $(Esc $h.Solucion)</p>") }
    $nums = @($ArrOrd | Where-Object { $h.Arreglos -contains $_.Id } | ForEach-Object { "#$($_.Numero) $($_.Nombre)" })
    if ($nums) { [void]$sb.Append("<p class='fix'>🔧 La herramienta puede hacerlo por ti: $(Esc ($nums -join ' · '))</p>") }
    [void]$sb.Append('</article>')
}

[void]$sb.Append('<h2>Arreglos disponibles</h2><div class=''pasos''>Los arreglos se aplican desde la ventana de Aquila (pestaña «Arreglos») o desde la consola. Antes de cada uno se muestra qué hace, por qué y si afecta algo, y se pide confirmación. Antes del primer cambio se crea un <b>punto de restauración</b> de Windows para poder deshacer.</div>')
[void]$sb.Append("<section class='bloque' style='margin-top:12px'><div class='tabla'><table><thead><tr><th>#</th><th>Arreglo</th><th>Libera</th><th>Riesgo</th><th>Qué hace / por qué</th><th>¿Afecta algo?</th></tr></thead><tbody>")
foreach ($a in $ArrOrd) {
    $lib = if ($a.Bytes -gt 0) { Fmt $a.Bytes } else { '—' }
    [void]$sb.Append("<tr><td>$($a.Numero)</td><td><b>$(Esc $a.Nombre)</b></td><td>$lib</td><td class='riesgo-$($a.Riesgo)'>$($a.Riesgo)</td><td>$(Esc $a.QueHace)<br><span class='nota'>$(Esc $a.Porque)</span></td><td>$(Esc $a.Afecta)</td></tr>")
}
[void]$sb.Append('</tbody></table></div></section>')

if ($comp) {
    $claves = [ordered]@{ Puntaje = 'Puntaje de salud'; RAM_GB = 'Memoria RAM (GB)'; DiscoSistema = 'Disco de Windows'; C_LibreGB = 'Libre en disco de Windows (GB)'
        CPU_Prom = 'Procesador ocupado (%)'; Disco_Ocupado_Prom = 'Disco ocupado (%)'; RAM_UsoPct = 'RAM en uso (%)'; InicioActivos = 'Programas al encender'
        Arranque_Seg = 'Tiempo de encendido (s)'; Outlook_GB = 'Correo Outlook (GB)'; Nube_Archivos = 'Archivos sincronizados en la nube'
        Recuperable_GB = 'Espacio recuperable (GB)'; Duplicados_GB = 'Duplicados (GB)'; Perfil_GB = 'Tamaño de tu usuario sin nube (GB)'; DiasEncendida = 'Días sin reiniciar'
        ProgramasInstalados = 'Programas instalados'; ErroresDisco30d = 'Errores de disco (30 días)'; Antivirus = 'Antivirus'; Windows = 'Windows'; Procesador = 'Procesador' }
    [void]$sb.Append("<h2>Comparación</h2><section class='bloque'><p class='nota'>Esta PC ($(Esc $Equipo), $(Esc $Datos.Fecha)) contra $(Esc $comp.Equipo) ($(Esc $comp.Fecha)).</p><div class='tabla'><table><thead><tr><th>Medida</th><th>$(Esc $Equipo) · ahora</th><th>$(Esc $comp.Equipo) · $(Esc $comp.Fecha)</th></tr></thead><tbody>")
    foreach ($k in $claves.Keys) { [void]$sb.Append("<tr><td>$(Esc $claves[$k])</td><td><b>$(Esc ([string]$Datos[$k]))</b></td><td>$(Esc ([string]$comp.$k))</td></tr>") }
    [void]$sb.Append('</tbody></table></div></section>')
}

[void]$sb.Append('<h2>Detalle de las mediciones</h2>')
foreach ($t in $Tablas) { [void]$sb.Append((Html-Tabla $t)) }
[void]$sb.Append("<footer>Aquila v$VersionHerramienta · Duración: $([math]::Round(((Get-Date) - $HoraInicio).TotalMinutes, 1)) min$(if (-not $EsAdmin) { ' · Ejecutado SIN permisos de administrador: algunas mediciones faltan.' })</footer>")
[void]$sb.Append('</div></body></html>')

$archivoHtml = Join-Path $CarpetaSalida "Diagnostico_${Equipo}_$Marca.html"
[IO.File]::WriteAllText($archivoHtml, $sb.ToString(), (New-Object Text.UTF8Encoding $true))

$archivoJson = Join-Path $CarpetaSalida "Diagnostico_${Equipo}_$Marca.json"
$resultado = [pscustomobject]@{
    Datos = $Datos; Puntaje = $puntaje; Estado = $estado; EsAdmin = $EsAdmin; InformeHtml = $archivoHtml; ArchivoJson = $archivoJson
    Hallazgos = $HallOrd
    Arreglos = @($ArrOrd | Select-Object Id, Tipo, Numero, Nombre, QueHace, Porque, Afecta, Riesgo, Bytes, Admin, Datos)
}
Guardar-Json $resultado $archivoJson
if ($ArchivoResultado) { Copy-Item -LiteralPath $archivoJson -Destination $ArchivoResultado -Force }
Write-Log "Informe: $archivoHtml"

# ============================================================================
#  12. Resumen en pantalla
# ============================================================================
$colorCon = if ($puntaje -ge 80) { 'Green' } elseif ($puntaje -ge 60) { 'Yellow' } else { 'Red' }
Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host "   PUNTAJE DE SALUD: $puntaje / 100  ($estado)" -ForegroundColor $colorCon
Write-Host '  ================================================' -ForegroundColor Cyan
if ($principales) {
    Write-Host '   Causas principales:' -ForegroundColor White
    foreach ($h in $principales) { Write-Host "    - [$($etq[$h.Nivel])] $($h.Titulo)" -ForegroundColor $(if ($h.Nivel -eq 'Critico') { 'Red' } else { 'Yellow' }) }
}
Write-Host ''
Write-Host "   Informe completo: $archivoHtml"
Write-Host "   Datos para comparar con otra PC: $archivoJson"
if (-not $NoAbrir) { Start-Process $archivoHtml }

# ============================================================================
#  13. Aplicar arreglos (modo consola)
# ============================================================================
if ($SinArreglos -or -not $ArrOrd.Count) { Write-Log 'Fin del diagnóstico.'; return }

Write-Host ''
Write-Host '  ARREGLOS DISPONIBLES' -ForegroundColor Cyan
foreach ($a in $ArrOrd) {
    $lib = if ($a.Bytes -gt 0) { " — libera ~$(Fmt $a.Bytes)" } else { '' }
    $bloq = if ($a.Admin -and -not $EsAdmin) { '  (requiere administrador)' } else { '' }
    $col = @{ Bajo = 'Green'; Medio = 'Yellow'; Alto = 'Red' }[$a.Riesgo]
    Write-Host ("   [{0,2}] " -f $a.Numero) -NoNewline
    Write-Host "$($a.Nombre)$lib" -NoNewline
    Write-Host "  · riesgo $($a.Riesgo)$bloq" -ForegroundColor $col
}
Write-Host ''
Write-Host '   Escribe los números separados por coma (ej: 1,3,5)'
Write-Host '   B = todos los de riesgo Bajo   |   T = todos   |   Enter = salir sin cambios'
$sel = Read-Host '   Tu elección'
$elegidos = @()
if ($sel -match '^\s*[bB]\s*$') { $elegidos = @($ArrOrd | Where-Object Riesgo -eq 'Bajo') }
elseif ($sel -match '^\s*[tT]\s*$') { $elegidos = $ArrOrd }
elseif ($sel.Trim()) { $nums = $sel -split '[,; ]+' | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }; $elegidos = @($ArrOrd | Where-Object { $nums -contains $_.Numero }) }
if (-not $elegidos) { Write-Host '   No se aplicó ningún cambio.'; Write-Log 'Fin sin arreglos (el usuario no eligió).'; return }

$libreInicial = Libre-Sistema
$puntoCreado = $false
$aplicados = 0
foreach ($a in $elegidos) {
    Write-Host ''
    Write-Host '  ------------------------------------------------' -ForegroundColor DarkGray
    Write-Host "   [$($a.Numero)] $($a.Nombre)" -ForegroundColor White
    if ($a.Bytes -gt 0) { Write-Host "   Libera aprox.:  $(Fmt $a.Bytes)" -ForegroundColor Green }
    Write-Host "   Qué hace:       $($a.QueHace)"
    Write-Host "   Por qué:        $($a.Porque)"
    Write-Host "   ¿Afecta algo?:  $($a.Afecta)" -ForegroundColor Yellow
    Write-Host "   Riesgo:         $($a.Riesgo)"
    if ($a.Admin -and -not $EsAdmin) { Write-Host '   Se omite: requiere ejecutar como administrador.' -ForegroundColor DarkYellow; continue }
    if (-not (Confirmar '   ¿Lo aplico?')) { Write-Log "OMITIDO por el usuario: $($a.Nombre)"; continue }
    if (-not $puntoCreado -and $EsAdmin) {
        if (-not (Nuevo-PuntoRestauracion) -and -not (Confirmar '   ¿Continuar sin punto de restauración?')) { break }
        $puntoCreado = $true
    }
    try {
        Write-Host '   Aplicando...' -ForegroundColor DarkGray
        $res = Invoke-Arreglo $a
        Write-Host "   OK. $res" -ForegroundColor Green
        Write-Log "APLICADO: $($a.Nombre) -> $res"
        $aplicados++
    } catch {
        Write-Host "   ERROR: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "ERROR en $($a.Nombre): $($_.Exception.Message)"
    }
}

$ganado = [math]::Max(0, (Libre-Sistema) - $libreInicial)
Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host "   Arreglos aplicados: $aplicados   ·   Espacio liberado en ${SisDrive}: $(Fmt $ganado)" -ForegroundColor Green
Write-Host '   Recomendado: REINICIAR la PC para que todo tome efecto.'
Write-Host '   Para medir la mejora, ejecuta el diagnóstico otra vez y compáralo con:'
Write-Host "     $archivoJson"
Write-Host "   Registro de todo lo hecho: $ArchivoLog"
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Log "Fin. Aplicados=$aplicados Liberado=$(Fmt $ganado)"
