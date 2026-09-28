# Comun.ps1 — funciones y textos compartidos por Aquila.ps1 (ventana), Aquila-Diagnostico.ps1 y Aquila-Vigilar.ps1.

# --- Identidad de Aquila ---
$AquilaRepo = 'rtsebastian2-collab/aquila'
$AquilaVersion = '0.0.0'
$archivoVersion = Join-Path (Split-Path $PSScriptRoot -Parent) 'version.txt'
if (Test-Path -LiteralPath $archivoVersion) { $AquilaVersion = (Get-Content -LiteralPath $archivoVersion -TotalCount 1).Trim() }
# Los informes van a Documentos\Aquila\Informes: así sobreviven a las actualizaciones
$AquilaInformes = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Aquila\Informes'
$AquilaLogoSvg = ''
$archivoLogo = Join-Path $PSScriptRoot 'aquila.svg'
if (Test-Path -LiteralPath $archivoLogo) { $AquilaLogoSvg = [IO.File]::ReadAllText($archivoLogo) }

function Fmt([double]$b) {
    if ($b -ge 1GB) { '{0:N1} GB' -f ($b / 1GB) } elseif ($b -ge 1MB) { '{0:N0} MB' -f ($b / 1MB) } else { '{0:N0} KB' -f ($b / 1KB) }
}
function Prom($arr) { $a = @($arr); if ($a.Count) { [math]::Round(($a | Measure-Object -Average).Average, 1) } else { 0 } }
function Esc([string]$s) { [System.Net.WebUtility]::HtmlEncode($s) }
function Leer-Json([string]$Ruta) {
    # En PowerShell 5.1 ConvertFrom-Json entrega una lista como UN solo objeto; al asignarla y devolverla se separan sus elementos
    $o = Get-Content -LiteralPath $Ruta -Raw -Encoding UTF8 | ConvertFrom-Json
    $o
}
function Guardar-Json($Objeto, [string]$Ruta) {
    # Escribe a un temporal y lo reemplaza, para que la ventana nunca lea un archivo a medio escribir
    $tmp = "$Ruta.tmp"
    ConvertTo-Json -InputObject $Objeto -Depth 8 | Set-Content -LiteralPath $tmp -Encoding UTF8
    Move-Item -LiteralPath $tmp -Destination $Ruta -Force
}

# Nombres entendibles para los procesos más comunes
$Conocidos = @{
    'MsMpEng' = 'Antivirus de Windows (Defender)'; 'MsSense' = 'Defender para empresas'; 'NisSrv' = 'Antivirus de Windows (red)'
    'SearchIndexer' = 'Indexador de búsqueda de Windows'; 'SearchProtocolHost' = 'Indexador de búsqueda'; 'SearchFilterHost' = 'Indexador de búsqueda'
    'TiWorker' = 'Windows Update instalando'; 'TrustedInstaller' = 'Windows Update instalando'; 'MoUsoCoreWorker' = 'Windows Update'
    'OneDrive' = 'OneDrive (nube)'; 'GoogleDriveFS' = 'Google Drive (nube)'; 'Dropbox' = 'Dropbox (nube)'; 'iCloudDrive' = 'iCloud (nube)'
    'OUTLOOK' = 'Outlook (correo)'; 'olk' = 'Nuevo Outlook (correo)'; 'ms-teams' = 'Microsoft Teams'; 'Teams' = 'Teams clásico'
    'chrome' = 'Google Chrome'; 'msedge' = 'Microsoft Edge'; 'firefox' = 'Firefox'; 'brave' = 'Brave'; 'opera' = 'Opera'
    'Memory Compression' = 'Compresión de memoria (señal de RAM llena)'; 'System' = 'Sistema (drivers y disco)'
    'svchost' = 'Servicios de Windows'; 'explorer' = 'Explorador de Windows'; 'dwm' = 'Dibujo de ventanas'
    'CompatTelRunner' = 'Telemetría de Windows'; 'WINWORD' = 'Word'; 'EXCEL' = 'Excel'; 'POWERPNT' = 'PowerPoint'
    'Code' = 'Visual Studio Code'; 'Zoom' = 'Zoom'; 'AnyDesk' = 'AnyDesk'; 'TeamViewer' = 'TeamViewer'; 'WhatsApp' = 'WhatsApp'
}

# Explicación en palabras simples de por qué ese proceso vuelve lenta la PC
$Consejos = @{
    'MsMpEng'            = 'El antivirus revisa cada archivo que se abre, se descarga o se sincroniza. Si OneDrive o Outlook están moviendo miles de archivos o correos, el antivirus los revisa todos.'
    'SearchIndexer'      = 'Windows está catalogando archivos y correos para que la búsqueda sea rápida. Con una nube enorme o un Outlook muy grande, este catálogo casi nunca termina.'
    'SearchProtocolHost' = 'Es parte del indexador de búsqueda: está leyendo archivos o correos para catalogarlos.'
    'OneDrive'           = 'OneDrive está sincronizando. Con decenas de miles de archivos, revisar qué cambió consume procesador y disco todo el día.'
    'GoogleDriveFS'      = 'Google Drive está sincronizando o armando su caché local.'
    'Dropbox'            = 'Dropbox está sincronizando archivos.'
    'OUTLOOK'            = 'Outlook está descargando o indexando correos. Un archivo de correo muy grande hace que Outlook trabaje constantemente.'
    'olk'                = 'El nuevo Outlook está sincronizando correo.'
    'TiWorker'           = 'Windows está instalando actualizaciones. Es temporal; conviene dejar que termine y reiniciar.'
    'TrustedInstaller'   = 'Windows está instalando actualizaciones. Es temporal; conviene dejar que termine y reiniciar.'
    'MoUsoCoreWorker'    = 'Windows Update está buscando o preparando actualizaciones.'
    'Memory Compression' = 'La RAM se está llenando y Windows comprime memoria para ganar espacio. Es síntoma de falta de RAM.'
    'System'             = 'Actividad del sistema y drivers; con uso alto suele indicar disco lento, drivers o antivirus.'
    'chrome'             = 'El navegador con muchas pestañas abiertas consume mucha memoria y procesador (cada pestaña es un programa aparte).'
    'msedge'             = 'El navegador con muchas pestañas abiertas consume mucha memoria y procesador (cada pestaña es un programa aparte).'
    'ms-teams'           = 'Teams consume bastante memoria y procesador, sobre todo en reuniones con video.'
    'Teams'              = 'Teams clásico es pesado; Microsoft recomienda pasar al nuevo Teams.'
    'CompatTelRunner'    = 'Windows revisa la compatibilidad y envía telemetría. Suele durar unos minutos y termina solo.'
    'svchost'            = 'Son servicios internos de Windows. Con uso alto suele tratarse de Windows Update, la búsqueda o la red.'
}
function Nombre-Amigable([string]$n) { if ($Conocidos.ContainsKey($n)) { "$n — $($Conocidos[$n])" } else { $n } }

$CssInforme = @'
:root{--bg:#f5f6f8;--card:#fff;--tx:#1b1f24;--mut:#5d6773;--bd:#e2e5e9;--crit:#c62828;--alto:#d9480f;--medio:#a67c00;--info:#1c64b8;--ok:#2b7a3a;--acc:#0b6bcb;--soft:#eef2f6;--s1:#2563c9;--s2:#d9480f;--s3:#7c3aed;--lento:rgba(198,40,40,.12)}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#101317;--card:#181d23;--tx:#e5e8ec;--mut:#9aa4af;--bd:#2a3139;--crit:#f06262;--alto:#ff8a4c;--medio:#e3c04a;--info:#6aa9f0;--ok:#5fbf6e;--acc:#5aa9f0;--soft:#1f252c;--s1:#6aa9f0;--s2:#ff8a4c;--s3:#b28cf5;--lento:rgba(240,98,98,.16)}}
:root[data-theme="dark"]{--bg:#101317;--card:#181d23;--tx:#e5e8ec;--mut:#9aa4af;--bd:#2a3139;--crit:#f06262;--alto:#ff8a4c;--medio:#e3c04a;--info:#6aa9f0;--ok:#5fbf6e;--acc:#5aa9f0;--soft:#1f252c;--s1:#6aa9f0;--s2:#ff8a4c;--s3:#b28cf5;--lento:rgba(240,98,98,.16)}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--tx);font:15px/1.55 "Segoe UI",system-ui,sans-serif}
.wrap{max-width:1080px;margin:0 auto;padding:24px 16px 60px}
header{display:flex;gap:24px;align-items:center;flex-wrap:wrap;margin-bottom:24px}
header h1{margin:0;font-size:26px}header .sub{color:var(--mut)}
.gauge{width:120px;height:120px;border-radius:50%;display:grid;place-items:center;background:conic-gradient(var(--c) calc(var(--p)*1%),var(--bd) 0)}
.gauge>div{width:92px;height:92px;border-radius:50%;background:var(--bg);display:grid;place-items:center;text-align:center}
.gauge b{font-size:30px;line-height:1}.gauge span{font-size:12px;color:var(--mut)}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:10px;margin-bottom:24px}
.dato{background:var(--card);border:1px solid var(--bd);border-radius:10px;padding:10px 14px}.dato small{color:var(--mut);display:block;font-size:12px}
h2{font-size:20px;margin:32px 0 12px;border-bottom:1px solid var(--bd);padding-bottom:6px}h3{font-size:16px;margin:0 0 6px}
.causas{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:8px 18px}.causas li{margin:8px 0}
.hallazgo{background:var(--card);border:1px solid var(--bd);border-left:5px solid var(--nc);border-radius:10px;padding:14px 18px;margin-bottom:12px}
.hallazgo .cab{display:flex;gap:8px;align-items:center;margin-bottom:6px;font-size:12px}
.badge{background:var(--nc);color:#fff;border-radius:20px;padding:1px 10px;font-weight:600}.cat{color:var(--mut)}
.n-Critico{--nc:var(--crit)}.n-Alto{--nc:var(--alto)}.n-Medio{--nc:var(--medio)}.n-Info{--nc:var(--info)}.n-OK{--nc:var(--ok)}
.detalle{color:var(--mut);margin:4px 0}.porque,.solucion,.fix{margin:6px 0}.fix{background:var(--soft);border-radius:8px;padding:6px 10px;font-size:14px}
.bloque{background:var(--card);border:1px solid var(--bd);border-radius:10px;padding:14px 18px;margin-bottom:14px}
.nota{color:var(--mut);font-size:13px;margin:0 0 8px}.tabla{overflow-x:auto}
table{border-collapse:collapse;width:100%;font-size:14px}th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--bd);vertical-align:top}
th{color:var(--mut);font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:.03em}td{word-break:break-word}
.riesgo-Bajo{color:var(--ok);font-weight:600}.riesgo-Medio{color:var(--medio);font-weight:600}.riesgo-Alto{color:var(--crit);font-weight:600}
.pasos{background:var(--soft);border-radius:10px;padding:12px 18px}footer{color:var(--mut);font-size:12px;margin-top:40px}
.marca{display:flex;align-items:center;gap:8px;font-weight:700;letter-spacing:.04em;color:var(--mut);margin-bottom:4px}.marca svg{width:28px;height:28px;flex:none}
.grafico svg{width:100%;height:auto;display:block}.grafico .eje{stroke:var(--bd);stroke-width:1}.grafico text{fill:var(--mut);font-size:11px}
.grafico .lento{fill:var(--lento)}.leyenda{display:flex;gap:18px;flex-wrap:wrap;font-size:13px;color:var(--mut);margin-top:6px}
.leyenda i{display:inline-block;width:14px;height:3px;border-radius:2px;vertical-align:middle;margin-right:6px}
'@
