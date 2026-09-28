#Requires -Version 5.1
<#
    AQUILA — Vigilar
    Se queda midiendo la PC cada pocos segundos. Cada vez que se pone lenta (procesador, disco o memoria al límite)
    anota el momento y qué programas lo causaron. Al final genera un informe con gráfico y culpables.

    Parámetros:
      -Minutos 30          Cuánto tiempo vigilar (por defecto 30)
      -Intervalo 3         Segundos entre mediciones
      -UmbralCpu 85 -UmbralDisco 90 -UmbralRam 92   Cuándo se considera «lenta»
      -ArchivoParar / -ArchivoEstado   (usados por la ventana)
    En consola: presiona Q para terminar antes.
#>
[CmdletBinding()]
param(
    [int]$Minutos = 30,
    [int]$Intervalo = 3,
    [int]$UmbralCpu = 85,
    [int]$UmbralDisco = 90,
    [int]$UmbralRam = 92,
    [string]$CarpetaSalida,
    [string]$ArchivoParar,
    [string]$ArchivoEstado,
    [switch]$NoAbrir
)

$ErrorActionPreference = 'SilentlyContinue'
. (Join-Path $PSScriptRoot 'Comun.ps1')
if (-not $CarpetaSalida) { $CarpetaSalida = $AquilaInformes }
[void](New-Item -ItemType Directory -Path $CarpetaSalida -Force)
$Equipo = $env:COMPUTERNAME
$Marca = Get-Date -Format 'yyyy-MM-dd_HHmm'
$nLog = [int](Get-CimInstance Win32_Processor | Select-Object -First 1).NumberOfLogicalProcessors; if ($nLog -lt 1) { $nLog = 1 }
$ramTotalMB = [double](Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1MB
$excluir = @('_Total', 'Idle', 'WmiPrvSE', 'powershell', 'pwsh', 'conhost')

$muestras = New-Object System.Collections.Generic.List[object]
$episodios = New-Object System.Collections.Generic.List[object]
$culpables = @{}
$ultimos = New-Object System.Collections.Generic.List[string]
$enEp = $false; $altos = 0; $bajos = 0; $ep = $null; $prevProcs = $null
$inicio = Get-Date
$fin = $inicio.AddMinutes($Minutos)

function Tomar-Procesos {
    $h = @{}
    foreach ($pr in Get-CimInstance Win32_PerfFormattedData_PerfProc_Process) {
        $n = $pr.Name -replace '#\d+$', ''
        if ($n -in $excluir) { continue }
        if (-not $h.ContainsKey($n)) { $h[$n] = [pscustomobject]@{ Cpu = 0.0; IO = 0.0; Mem = 0.0 } }
        $h[$n].Cpu += [double]$pr.PercentProcessorTime / $nLog
        $h[$n].IO  += [double]$pr.IODataBytesPersec
        $h[$n].Mem += [double]$pr.WorkingSetPrivate
    }
    $h
}
function Sumar-Procesos($ep, $procs) {
    foreach ($k in $procs.Keys) {
        if (-not $ep.Procs.ContainsKey($k)) { $ep.Procs[$k] = [pscustomobject]@{ Cpu = 0.0; IO = 0.0; Mem = 0.0 } }
        $ep.Procs[$k].Cpu += $procs[$k].Cpu
        $ep.Procs[$k].IO  += $procs[$k].IO
        if ($procs[$k].Mem -gt $ep.Procs[$k].Mem) { $ep.Procs[$k].Mem = $procs[$k].Mem }
    }
    $ep.N++
}
function Cerrar-Episodio($ep, [datetime]$hasta) {
    $ep.Fin = $hasta
    $dur = [math]::Max($Intervalo, [int]($ep.Fin - $ep.Inicio).TotalSeconds)
    # El recurso que más veces pasó el límite define el tipo de lentitud y a quién se culpa
    $motivo = ($ep.Motivos.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
    $lista = @($ep.Procs.GetEnumerator() | ForEach-Object { [pscustomobject]@{ Nombre = $_.Key; Cpu = $_.Value.Cpu / $ep.N; IO = $_.Value.IO / $ep.N; Mem = $_.Value.Mem } })
    $top = switch ($motivo) {
        'Procesador' { @($lista | Sort-Object Cpu -Descending | Select-Object -First 3 | ForEach-Object { [pscustomobject]@{ Nombre = $_.Nombre; Valor = ('{0:N0}%' -f $_.Cpu) } }) }
        'Disco'      { @($lista | Sort-Object IO -Descending | Select-Object -First 3 | ForEach-Object { [pscustomobject]@{ Nombre = $_.Nombre; Valor = "$(Fmt $_.IO)/s" } }) }
        default      { @($lista | Sort-Object Mem -Descending | Select-Object -First 3 | ForEach-Object { [pscustomobject]@{ Nombre = $_.Nombre; Valor = (Fmt $_.Mem) } }) }
    }
    $e = [pscustomobject]@{ Inicio = $ep.Inicio; Fin = $ep.Fin; Segundos = $dur; Motivo = $motivo
                            MaxCpu = $ep.MaxCpu; MaxDisco = $ep.MaxDisco; MaxRam = $ep.MaxRam; Culpables = $top }
    $episodios.Add($e)
    $pos = 0
    foreach ($c in $top) {
        $pos++
        if (-not $culpables.ContainsKey($c.Nombre)) { $culpables[$c.Nombre] = [pscustomobject]@{ Nombre = $c.Nombre; Veces = 0; Principal = 0; Segundos = 0 } }
        $culpables[$c.Nombre].Veces++; $culpables[$c.Nombre].Segundos += $dur
        if ($pos -eq 1) { $culpables[$c.Nombre].Principal++ }
    }
    $txt = '{0:HH:mm:ss}-{1:HH:mm:ss} ({2} s) · {3} al límite · {4}' -f $e.Inicio, $e.Fin, $dur, $motivo, (($top | ForEach-Object { "$(Nombre-Amigable $_.Nombre) ($($_.Valor))" }) -join ', ')
    $ultimos.Insert(0, $txt)
    if ($ultimos.Count -gt 8) { $ultimos.RemoveAt(8) }
    Write-Host "  !! LENTITUD $txt" -ForegroundColor Yellow
}
function Escribir-Estado($c, $d, $r, [bool]$terminado, [string]$informe) {
    if (-not $ArchivoEstado) { return }
    try {
        Guardar-Json ([pscustomobject]@{ Cpu = $c; Disco = $d; Ram = $r; Episodios = $episodios.Count; Ultimos = $ultimos.ToArray()
                                         Restante = [math]::Max(0, [int]($fin - (Get-Date)).TotalSeconds); Terminado = $terminado; Informe = $informe }) $ArchivoEstado
    } catch {}
}

Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host "   AQUILA · VIGILANDO $Equipo durante $Minutos minutos" -ForegroundColor Cyan
Write-Host '   Usa la PC normalmente. Presiona Q para terminar antes.' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan

while ((Get-Date) -lt $fin) {
    if ($ArchivoParar -and (Test-Path -LiteralPath $ArchivoParar)) { break }
    try { if ([Console]::KeyAvailable -and [Console]::ReadKey($true).Key -eq 'Q') { break } } catch {}
    $t0 = Get-Date
    $cpu = [double](Get-CimInstance Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'").PercentProcessorTime
    $dk = Get-CimInstance Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name='_Total'"
    $disco = if ($dk) { [double](100 - [math]::Min(100, $dk.PercentIdleTime)) } else { 0 }
    $ram = [math]::Round(100 - [double](Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory).AvailableMBytes / $ramTotalMB * 100, 1)
    $procs = Tomar-Procesos

    $motivos = @()
    if ($cpu -ge $UmbralCpu) { $motivos += 'Procesador' }
    if ($disco -ge $UmbralDisco) { $motivos += 'Disco' }
    if ($ram -ge $UmbralRam) { $motivos += 'Memoria' }
    $lento = $motivos.Count -gt 0
    $muestras.Add([pscustomobject]@{ T = $t0; Cpu = $cpu; Disco = $disco; Ram = $ram; Lento = $lento })
    if ($lento) { $altos++; $bajos = 0 } else { $bajos++; $altos = 0 }

    # Un episodio empieza con 2 mediciones seguidas al límite y termina con 2 normales (evita picos de un segundo)
    if (-not $enEp -and $altos -ge 2) {
        $enEp = $true
        $ep = [pscustomobject]@{ Inicio = $muestras[$muestras.Count - 2].T; Fin = $null; Procs = @{}; Motivos = @{}; N = 0; MaxCpu = 0; MaxDisco = 0; MaxRam = 0 }
        if ($prevProcs) { Sumar-Procesos $ep $prevProcs }
    }
    if ($enEp) {
        Sumar-Procesos $ep $procs
        foreach ($m in $motivos) { $ep.Motivos[$m] = [int]$ep.Motivos[$m] + 1 }
        if ($cpu -gt $ep.MaxCpu) { $ep.MaxCpu = $cpu }; if ($disco -gt $ep.MaxDisco) { $ep.MaxDisco = $disco }; if ($ram -gt $ep.MaxRam) { $ep.MaxRam = $ram }
        if ($bajos -ge 2) {
            if (-not $ep.Motivos.Count) { $ep.Motivos['Procesador'] = 1 }
            Cerrar-Episodio $ep $t0; $enEp = $false
        }
    }
    $prevProcs = $procs

    $restan = [math]::Max(0, [int]($fin - (Get-Date)).TotalSeconds)
    Write-Host ("`r  {0:HH:mm:ss}  CPU {1,3:N0}%  Disco {2,3:N0}%  RAM {3,3:N0}%  · episodios: {4} · quedan {5:N0} min   " -f $t0, $cpu, $disco, $ram, $episodios.Count, ($restan / 60)) -NoNewline
    Escribir-Estado $cpu $disco $ram $false $null
    $espera = $Intervalo - ((Get-Date) - $t0).TotalSeconds
    if ($espera -gt 0) { Start-Sleep -Milliseconds ([int]($espera * 1000)) }
}
if ($enEp) { if (-not $ep.Motivos.Count) { $ep.Motivos['Procesador'] = 1 }; Cerrar-Episodio $ep (Get-Date) }
Write-Host ''

# ============================================================================
#  Informe
# ============================================================================
$finReal = Get-Date
$durTotal = [math]::Max(1, ($finReal - $inicio).TotalSeconds)
$segLento = [double](($episodios | Measure-Object Segundos -Sum).Sum)
$pctLento = [math]::Round($segLento / $durTotal * 100, 1)
$ranking = @($culpables.Values | Sort-Object @{ E = { $_.Segundos }; Descending = $true }, @{ E = { $_.Principal }; Descending = $true })

# Gráfico SVG (procesador, disco, RAM) con las zonas lentas sombreadas
$W = 1000; $Hg = 260; $mx = 40; $my = 16
$anch = $W - $mx - 10; $alto = $Hg - $my - 30
function X([datetime]$t) { [math]::Round($mx + (($t - $inicio).TotalSeconds / $durTotal) * $anch, 1) }
function Y([double]$v) { [math]::Round($my + $alto - ([math]::Min(100, [math]::Max(0, $v)) / 100) * $alto, 1) }
$paso = [math]::Max(1, [int][math]::Ceiling($muestras.Count / 600))
$pts = @{ Cpu = @(); Disco = @(); Ram = @() }
for ($i = 0; $i -lt $muestras.Count; $i += $paso) {
    $s = $muestras[$i]
    foreach ($k in 'Cpu', 'Disco', 'Ram') { $pts[$k] += "$(X $s.T),$(Y $s.$k)" }
}
$svg = New-Object System.Text.StringBuilder
[void]$svg.Append("<svg viewBox='0 0 $W $Hg' role='img' aria-label='Uso de procesador, disco y memoria durante la vigilancia'>")
foreach ($e in $episodios) { $x1 = X $e.Inicio; $x2 = [math]::Max($x1 + 2, (X $e.Fin)); [void]$svg.Append("<rect class='lento' x='$x1' y='$my' width='$([math]::Round($x2 - $x1, 1))' height='$alto'/>") }
foreach ($v in 0, 50, 100) { $y = Y $v; [void]$svg.Append("<line class='eje' x1='$mx' x2='$($W - 10)' y1='$y' y2='$y'/><text x='4' y='$($y + 4)'>$v%</text>") }
for ($i = 0; $i -le 5; $i++) {
    $t = $inicio.AddSeconds($durTotal * $i / 5); $x = X $t
    [void]$svg.Append("<text x='$x' y='$($Hg - 8)' text-anchor='middle'>$($t.ToString('HH:mm'))</text>")
}
$colores = @{ Cpu = 'var(--s1)'; Disco = 'var(--s2)'; Ram = 'var(--s3)' }
foreach ($k in 'Ram', 'Disco', 'Cpu') { [void]$svg.Append("<polyline fill='none' stroke='$($colores[$k])' stroke-width='1.6' stroke-linejoin='round' points='$($pts[$k] -join ' ')'/>") }
[void]$svg.Append('</svg>')

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append("<!DOCTYPE html><html lang='es'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Aquila · Vigilancia $(Esc $Equipo)</title><link rel='icon' href='data:image/svg+xml,$([uri]::EscapeDataString($AquilaLogoSvg))'><style>$CssInforme</style></head><body><div class='wrap'>")
[void]$sb.Append("<header><div><div class='marca'>$AquilaLogoSvg<span>Aquila</span></div><h1>Vigilancia de $(Esc $Equipo)</h1><div class='sub'>$($inicio.ToString('dd/MM/yyyy HH:mm')) a $($finReal.ToString('HH:mm')) · $([math]::Round($durTotal / 60)) min · $($episodios.Count) momentos de lentitud · lenta el $pctLento% del tiempo</div></div></header>")

[void]$sb.Append('<h2>Conclusión</h2><div class=''causas''>')
if (-not $episodios.Count) {
    [void]$sb.Append("<p>No hubo momentos de lentitud durante la vigilancia (procesador bajo $UmbralCpu%, disco bajo $UmbralDisco% y RAM bajo $UmbralRam%). Si la lentitud aparece en otro momento (al encender, al abrir Outlook o al sincronizar), vigila justo en ese momento.</p>")
} else {
    [void]$sb.Append('<p>Los programas que más veces estaban detrás de la lentitud:</p><ol>')
    foreach ($c in $ranking | Select-Object -First 5) {
        $exp = if ($Consejos.ContainsKey($c.Nombre)) { $Consejos[$c.Nombre] } else { 'Estaba entre los que más recursos consumían cuando la PC se puso lenta.' }
        [void]$sb.Append("<li><b>$(Esc (Nombre-Amigable $c.Nombre))</b> — presente en $($c.Veces) de $($episodios.Count) momentos lentos ($([math]::Round($c.Segundos / 60, 1)) min en total; fue el principal $($c.Principal) veces). $(Esc $exp)</li>")
    }
    [void]$sb.Append('</ol><p class=''nota''>Siguiente paso: ejecuta el diagnóstico y revisa los arreglos relacionados (programas al encender, Outlook, OneDrive, limpieza).</p>')
}
[void]$sb.Append('</div>')

[void]$sb.Append("<h2>Uso durante la vigilancia</h2><section class='bloque grafico'>$($svg.ToString())<div class='leyenda'><span><i style='background:var(--s1)'></i>Procesador</span><span><i style='background:var(--s2)'></i>Disco</span><span><i style='background:var(--s3)'></i>Memoria RAM</span><span><i style='background:var(--lento);height:10px'></i>Momentos lentos</span></div></section>")

if ($episodios.Count) {
    [void]$sb.Append("<h2>Momentos de lentitud</h2><section class='bloque'><div class='tabla'><table><thead><tr><th>Desde</th><th>Hasta</th><th>Duración</th><th>Qué se saturó</th><th>Culpables</th></tr></thead><tbody>")
    foreach ($e in $episodios) {
        $cul = ($e.Culpables | ForEach-Object { "$(Nombre-Amigable $_.Nombre) ($($_.Valor))" }) -join '; '
        [void]$sb.Append("<tr><td>$($e.Inicio.ToString('HH:mm:ss'))</td><td>$($e.Fin.ToString('HH:mm:ss'))</td><td>$($e.Segundos) s</td><td>$(Esc $e.Motivo) (CPU $([math]::Round($e.MaxCpu))% · disco $([math]::Round($e.MaxDisco))% · RAM $([math]::Round($e.MaxRam))%)</td><td>$(Esc $cul)</td></tr>")
    }
    [void]$sb.Append('</tbody></table></div></section>')
    [void]$sb.Append("<h2>Ranking de culpables</h2><section class='bloque'><div class='tabla'><table><thead><tr><th>Programa</th><th>Momentos lentos</th><th>Como principal</th><th>Tiempo total</th></tr></thead><tbody>")
    foreach ($c in $ranking) { [void]$sb.Append("<tr><td>$(Esc (Nombre-Amigable $c.Nombre))</td><td>$($c.Veces)</td><td>$($c.Principal)</td><td>$([math]::Round($c.Segundos / 60, 1)) min</td></tr>") }
    [void]$sb.Append('</tbody></table></div></section>')
}
[void]$sb.Append("<footer>Aquila v$AquilaVersion · Vigilancia cada $Intervalo s · Umbrales: procesador $UmbralCpu%, disco $UmbralDisco%, RAM $UmbralRam% · $($muestras.Count) mediciones</footer></div></body></html>")

$archivoHtml = Join-Path $CarpetaSalida "Vigilancia_${Equipo}_$Marca.html"
[IO.File]::WriteAllText($archivoHtml, $sb.ToString(), (New-Object Text.UTF8Encoding $true))
Guardar-Json ([pscustomobject]@{ Equipo = $Equipo; Inicio = $inicio.ToString('s'); Fin = $finReal.ToString('s'); PctLento = $pctLento
                                 Episodios = $episodios.ToArray(); Ranking = $ranking }) (Join-Path $CarpetaSalida "Vigilancia_${Equipo}_$Marca.json")
$ult = $muestras | Select-Object -Last 1
Escribir-Estado $ult.Cpu $ult.Disco $ult.Ram $true $archivoHtml

Write-Host "  Vigilancia terminada: $($episodios.Count) momentos de lentitud ($pctLento% del tiempo)." -ForegroundColor Green
if ($ranking) { Write-Host "  Principal culpable: $(Nombre-Amigable $ranking[0].Nombre)" -ForegroundColor Yellow }
Write-Host "  Informe: $archivoHtml"
if (-not $NoAbrir) { Start-Process $archivoHtml }
