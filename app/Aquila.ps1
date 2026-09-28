#Requires -Version 5.1
<#
    AQUILA — ventana principal
    Ejecuta Aquila-Diagnostico.ps1 y Aquila-Vigilar.ps1 en segundo plano y muestra los resultados.
    Se abre con el acceso directo «Aquila» del escritorio.
    -Probar ruta.json [-Captura ruta.png] : carga un resultado y cierra (solo para verificar la ventana).
#>
param([string]$Probar, [string]$Captura)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
. (Join-Path $PSScriptRoot 'Comun.ps1')

$Raiz = $PSScriptRoot
$ScriptDiag = Join-Path $Raiz 'Aquila-Diagnostico.ps1'
$ScriptVig = Join-Path $Raiz 'Aquila-Vigilar.ps1'
$Manual = Join-Path $Raiz 'Manual.html'
$Informes = $AquilaInformes
[void](New-Item -ItemType Directory -Path $Informes -Force)
$EsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

$script:Resultado = $null; $script:Puntaje = -1
$script:Proc = $null; $script:LogActual = $null; $script:LineasVistas = 0; $script:AlTerminar = $null
$script:ProcVig = $null; $script:EstadoVig = $null; $script:PararVig = $null; $script:InformeVig = $null
$script:ArchivoComparar = $null

# ============================================================================
#  Estilo
# ============================================================================
function Col([string]$hex) { [System.Drawing.ColorTranslator]::FromHtml($hex) }
$Paleta = @{ Marino = (Col '#0B1B3F'); Marino2 = (Col '#1E3A8A'); Oro = (Col '#F2B705'); OroOscuro = (Col '#C99700'); Fondo = (Col '#F3F5F9')
        Blanco = [System.Drawing.Color]::White; Texto = (Col '#1B2330'); Suave = (Col '#64748B'); Borde = (Col '#D5DCE6'); Claro = (Col '#CBD5E1'); Azul = (Col '#0B6BCB') }
$ColorNivel = @{ Critico = (Col '#C62828'); Alto = (Col '#D9480F'); Medio = (Col '#A67C00'); Info = (Col '#1C64B8'); OK = (Col '#2B7A3A') }
$EtqNivel = @{ Critico = 'Crítico'; Alto = 'Alto'; Medio = 'Medio'; Info = 'Info'; OK = 'Bien' }
function Fuente([single]$t, [switch]$Negrita) { New-Object System.Drawing.Font('Segoe UI', $t, $(if ($Negrita) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular })) }
$Fuente = Fuente 10; $FuenteNegrita = Fuente 10 -Negrita; $FuenteTitulo = Fuente 14 -Negrita

function Nuevo([string]$Tipo, [hashtable]$Props = @{}) {
    $c = New-Object "System.Windows.Forms.$Tipo"
    foreach ($k in $Props.Keys) { $c.$k = $Props[$k] }
    $c
}
function Pt([int]$x, [int]$y) { New-Object System.Drawing.Point($x, $y) }
function Sz([int]$w, [int]$h) { New-Object System.Drawing.Size($w, $h) }
function Boton([string]$Texto, [int]$Ancho, [switch]$Primario) {
    $b = Nuevo Button @{ Text = $Texto; Size = (Sz $Ancho 38); FlatStyle = 'Flat'; Cursor = 'Hand'; Margin = (New-Object System.Windows.Forms.Padding(0, 0, 8, 0)) }
    if ($Primario) {
        $b.BackColor = $Paleta.Oro; $b.ForeColor = $Paleta.Marino; $b.Font = $FuenteNegrita
        $b.FlatAppearance.BorderColor = $Paleta.OroOscuro; $b.FlatAppearance.MouseOverBackColor = (Col '#FFC928')
    } else {
        $b.BackColor = $Paleta.Blanco; $b.ForeColor = $Paleta.Texto; $b.Font = $Fuente
        $b.FlatAppearance.BorderColor = $Paleta.Borde; $b.FlatAppearance.MouseOverBackColor = (Col '#EEF2F7')
    }
    $b
}
function Doble-Buffer($ctl) { $ctl.GetType().GetProperty('DoubleBuffered', [Reflection.BindingFlags]'NonPublic,Instance').SetValue($ctl, $true, $null) }
function Rtb-Agregar($rtb, [string]$texto, [switch]$Negrita, $Color = $null, [switch]$Titulo) {
    $rtb.SelectionStart = $rtb.TextLength
    $rtb.SelectionFont = if ($Titulo) { $FuenteTitulo } elseif ($Negrita) { $FuenteNegrita } else { $Fuente }
    $rtb.SelectionColor = if ($Color) { $Color } else { $Paleta.Texto }
    $rtb.AppendText($texto)
}
function Mensaje([string]$texto, [string]$titulo = 'Aquila', $icono = 'Information', $botones = 'OK') {
    [System.Windows.Forms.MessageBox]::Show($form, $texto, $titulo, $botones, $icono)
}
function Leer-Lineas([string]$ruta) {
    if (-not $ruta -or -not (Test-Path -LiteralPath $ruta)) { return @() }
    try {
        $fs = [IO.File]::Open($ruta, 'Open', 'Read', 'ReadWrite, Delete')
        $sr = New-Object IO.StreamReader($fs, [Text.Encoding]::UTF8)
        $txt = $sr.ReadToEnd(); $sr.Dispose()
        @($txt -split "`r?`n" | Where-Object { $_ })
    } catch { @() }
}

# ============================================================================
#  Ventana principal
# ============================================================================
$form = Nuevo Form @{ Text = 'Aquila'; Size = (Sz 1180 800); StartPosition = 'CenterScreen'; Font = $Fuente; BackColor = $Paleta.Fondo
                      MinimumSize = (Sz 960 640); AutoScaleMode = 'Dpi' }
$icono = Join-Path $Raiz 'aquila.ico'
if (Test-Path -LiteralPath $icono) { $form.Icon = New-Object System.Drawing.Icon($icono) }

# --- Cabecera azul marino ---
$cab = Nuevo Panel @{ Dock = 'Top'; Height = 118; BackColor = $Paleta.Marino }
$logo = Nuevo PictureBox @{ Size = (Sz 84 84); Location = (Pt 22 17); SizeMode = 'Zoom'; BackColor = $Paleta.Marino }
$archivoLogo = Join-Path $Raiz 'logo.png'
if (Test-Path -LiteralPath $archivoLogo) { $logo.Image = [System.Drawing.Image]::FromFile($archivoLogo) }
$lblApp = Nuevo Label @{ Text = 'Aquila'; Font = (Fuente 26 -Negrita); ForeColor = $Paleta.Blanco; AutoSize = $true; Location = (Pt 116 12); BackColor = $Paleta.Marino }
$lblLema = Nuevo Label @{ Text = 'Diagnóstico y optimización de tu computadora'; Font = (Fuente 11); ForeColor = $Paleta.Claro; AutoSize = $true; Location = (Pt 120 60); BackColor = $Paleta.Marino }
$lblEquipo = Nuevo Label @{ Text = "$env:COMPUTERNAME  ·  $(if ($EsAdmin) { 'con permisos de administrador' } else { 'SIN permisos de administrador: abre Aquila desde su acceso directo' })"
                            Font = (Fuente 9); ForeColor = (Col '#94A3B8'); AutoSize = $true; Location = (Pt 121 86); BackColor = $Paleta.Marino }
$anillo = Nuevo Panel @{ Size = (Sz 104 104); Anchor = 'Top, Right'; BackColor = $Paleta.Marino }
$lblSaludTit = Nuevo Label @{ Text = 'SALUD DE LA PC'; Font = (Fuente 8 -Negrita); ForeColor = (Col '#94A3B8'); AutoSize = $false; Size = (Sz 190 18); TextAlign = 'MiddleRight'; Anchor = 'Top, Right'; BackColor = $Paleta.Marino }
$lblSalud = Nuevo Label @{ Text = 'Sin diagnosticar'; Font = (Fuente 15 -Negrita); ForeColor = $Paleta.Blanco; AutoSize = $false; Size = (Sz 190 30); TextAlign = 'MiddleRight'; Anchor = 'Top, Right'; BackColor = $Paleta.Marino }
Doble-Buffer $anillo
$anillo.Add_Paint({
    param($s, $e)
    $g = $e.Graphics; $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAliasGridFit'
    $r = New-Object System.Drawing.Rectangle(9, 9, 86, 86)
    $g.DrawArc((New-Object System.Drawing.Pen($Paleta.Marino2, 9)), $r, 0, 360)
    $p = $script:Puntaje
    if ($p -ge 0) {
        $col = if ($p -ge 80) { Col '#4ADE80' } elseif ($p -ge 60) { $Paleta.Oro } elseif ($p -ge 40) { Col '#FB923C' } else { Col '#F87171' }
        $pen = New-Object System.Drawing.Pen($col, 9); $pen.StartCap = 'Round'; $pen.EndCap = 'Round'
        $g.DrawArc($pen, $r, -90, [single](3.6 * [math]::Max(1, $p)))
    }
    $sf = New-Object System.Drawing.StringFormat; $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
    $g.DrawString($(if ($p -ge 0) { "$p" } else { '—' }), (Fuente 22 -Negrita), (New-Object System.Drawing.SolidBrush($Paleta.Blanco)), (New-Object System.Drawing.RectangleF(0, 22, 104, 44)), $sf)
    $g.DrawString('de 100', (Fuente 8), (New-Object System.Drawing.SolidBrush((Col '#94A3B8'))), (New-Object System.Drawing.RectangleF(0, 62, 104, 18)), $sf)
})
$cab.Controls.AddRange(@($logo, $lblApp, $lblLema, $lblEquipo, $anillo, $lblSaludTit, $lblSalud))
function Ubicar-Cabecera {
    $anillo.Location = Pt ($cab.ClientSize.Width - 126) 7
    $lblSaludTit.Location = Pt ($cab.ClientSize.Width - 330) 34
    $lblSalud.Location = Pt ($cab.ClientSize.Width - 330) 52
}

# --- Barra de acciones ---
$barraAcc = Nuevo Panel @{ Dock = 'Top'; Height = 62; BackColor = $Paleta.Blanco; Padding = (New-Object System.Windows.Forms.Padding(20, 12, 20, 12)) }
$izq = Nuevo FlowLayoutPanel @{ Dock = 'Fill'; WrapContents = $false; BackColor = $Paleta.Blanco }
$der = Nuevo FlowLayoutPanel @{ Dock = 'Right'; Width = 290; WrapContents = $false; FlowDirection = 'RightToLeft'; BackColor = $Paleta.Blanco }
$btnDiag = Boton '▶  Diagnosticar' 170 -Primario
$chkDup = Nuevo CheckBox @{ Text = 'Buscar duplicados'; Checked = $true; AutoSize = $true; Margin = (New-Object System.Windows.Forms.Padding(4, 9, 16, 0)); ForeColor = $Paleta.Texto }
$btnComparar = Boton 'Comparar con otra PC…' 190
$btnInforme = Boton 'Ver informe' 120; $btnInforme.Enabled = $false
$btnCarpeta = Boton 'Mis informes' 120
$btnActualizar = Boton '⟳  Actualizar' 130
$btnManual = Boton '?  Manual' 120
$izq.Controls.AddRange(@($btnDiag, $chkDup, $btnComparar, $btnInforme, $btnCarpeta))
$der.Controls.AddRange(@($btnManual, $btnActualizar))
$linea = Nuevo Panel @{ Dock = 'Bottom'; Height = 1; BackColor = $Paleta.Borde }
$barraAcc.Controls.AddRange(@($izq, $der, $linea))

# --- Barra de progreso (visible mientras Aquila trabaja) ---
$panelProg = Nuevo Panel @{ Dock = 'Top'; Height = 56; BackColor = (Col '#FFF7DB'); Visible = $false; Padding = (New-Object System.Windows.Forms.Padding(20, 6, 20, 10)) }
$lblProg = Nuevo Label @{ Dock = 'Fill'; Font = (Fuente 10.5 -Negrita); ForeColor = $Paleta.Marino; TextAlign = 'MiddleLeft'; AutoEllipsis = $true }
$barraProg = Nuevo ProgressBar @{ Dock = 'Bottom'; Height = 12; Maximum = 100; Style = 'Marquee'; MarqueeAnimationSpeed = 25 }
$panelProg.Controls.AddRange(@($lblProg, $barraProg))
$lineaProg = Nuevo Panel @{ Dock = 'Bottom'; Height = 1; BackColor = $Paleta.Oro }
$panelProg.Controls.Add($lineaProg)

# --- Pie ---
$pie = Nuevo Panel @{ Dock = 'Bottom'; Height = 34; BackColor = $Paleta.Blanco; Padding = (New-Object System.Windows.Forms.Padding(14, 6, 14, 6)) }
$barra = Nuevo ProgressBar @{ Dock = 'Right'; Width = 260; Visible = $false; Maximum = 100 }
$lblVersion = Nuevo Label @{ Dock = 'Right'; Width = 130; Text = "Aquila v$AquilaVersion"; TextAlign = 'MiddleRight'; ForeColor = $Paleta.Suave }
$lblEstado = Nuevo Label @{ Dock = 'Fill'; Text = 'Listo. Presiona «Diagnosticar» para empezar.'; TextAlign = 'MiddleLeft'; AutoEllipsis = $true; ForeColor = $Paleta.Texto }
$pie.Controls.AddRange(@($lblEstado, $barra, $lblVersion))

# --- Pestañas ---
$cuerpo = Nuevo Panel @{ Dock = 'Fill'; Padding = (New-Object System.Windows.Forms.Padding(16, 12, 16, 8)); BackColor = $Paleta.Fondo }
$tabs = Nuevo TabControl @{ Dock = 'Fill'; Padding = (Pt 18 6); Font = (Fuente 10) }
$tabDiag = Nuevo TabPage @{ Text = 'Diagnóstico'; BackColor = $Paleta.Blanco }
$tabArr = Nuevo TabPage @{ Text = 'Arreglos'; BackColor = $Paleta.Blanco }
$tabVig = Nuevo TabPage @{ Text = 'Vigilar'; BackColor = $Paleta.Blanco }
$tabLog = Nuevo TabPage @{ Text = 'Registro'; BackColor = $Paleta.Blanco }
$tabs.TabPages.AddRange(@($tabDiag, $tabArr, $tabVig, $tabLog))
$cuerpo.Controls.Add($tabs)

# Diagnóstico: lista + detalle
$splitD = Nuevo SplitContainer @{ Dock = 'Fill'; Orientation = 'Vertical'; BackColor = $Paleta.Borde }
$lvHall = Nuevo ListView @{ Dock = 'Fill'; View = 'Details'; FullRowSelect = $true; HideSelection = $false; MultiSelect = $false; BorderStyle = 'None'; Font = (Fuente 10) }
[void]$lvHall.Columns.Add('Nivel', 80); [void]$lvHall.Columns.Add('Tema', 110); [void]$lvHall.Columns.Add('Qué se encontró', 360)
$rtbHall = Nuevo RichTextBox @{ Dock = 'Fill'; ReadOnly = $true; BorderStyle = 'None'; BackColor = $Paleta.Blanco }
$splitD.Panel1.BackColor = $Paleta.Blanco; $splitD.Panel2.BackColor = $Paleta.Blanco
$splitD.Panel2.Padding = New-Object System.Windows.Forms.Padding(16, 12, 12, 12)
$splitD.Panel1.Controls.Add($lvHall); $splitD.Panel2.Controls.Add($rtbHall)
$tabDiag.Controls.Add($splitD)

# Arreglos
$panelArrBtn = Nuevo FlowLayoutPanel @{ Dock = 'Bottom'; Height = 58; Padding = (New-Object System.Windows.Forms.Padding(10, 10, 10, 8)); BackColor = $Paleta.Fondo }
$btnAplicar = Boton 'Aplicar los marcados…' 200 -Primario; $btnAplicar.Enabled = $false
$btnMarcarBajo = Boton 'Marcar los rápidos y seguros' 230
$btnDesmarcar = Boton 'Desmarcar todo' 140
$btnRestaurar = Boton 'Restaurar programas de inicio' 240
$panelArrBtn.Controls.AddRange(@($btnAplicar, $btnMarcarBajo, $btnDesmarcar, $btnRestaurar))
$splitA = Nuevo SplitContainer @{ Dock = 'Fill'; Orientation = 'Horizontal'; BackColor = $Paleta.Borde }
$lvArr = Nuevo ListView @{ Dock = 'Fill'; View = 'Details'; CheckBoxes = $true; FullRowSelect = $true; HideSelection = $false; MultiSelect = $false; BorderStyle = 'None' }
[void]$lvArr.Columns.Add('#', 40); [void]$lvArr.Columns.Add('Arreglo', 470); [void]$lvArr.Columns.Add('Libera', 90); [void]$lvArr.Columns.Add('Riesgo', 80); [void]$lvArr.Columns.Add('Tiempo', 150); [void]$lvArr.Columns.Add('Estado', 170)
$rtbArr = Nuevo RichTextBox @{ Dock = 'Fill'; ReadOnly = $true; BorderStyle = 'None'; BackColor = $Paleta.Blanco }
$splitA.Panel1.BackColor = $Paleta.Blanco; $splitA.Panel2.BackColor = $Paleta.Blanco
$splitA.Panel2.Padding = New-Object System.Windows.Forms.Padding(16, 10, 12, 10)
$splitA.Panel1.Controls.Add($lvArr); $splitA.Panel2.Controls.Add($rtbArr)
$tabArr.Controls.Add($splitA); $tabArr.Controls.Add($panelArrBtn)

# Vigilar
$lblVigInfo = Nuevo Label @{ Text = "Deja la PC trabajando como siempre. Aquila mide cada 3 segundos y, cada vez que la PC se pone lenta, anota el momento y qué programa lo causó. Al final verás un gráfico y el ranking de culpables.`r`nIdeal: iniciarlo justo antes de lo que suele ponerse lento (abrir Outlook, sincronizar, primera hora de la mañana)."
                             Location = (Pt 20 16); Size = (Sz 1080 64); ForeColor = $Paleta.Texto }
$lblMin = Nuevo Label @{ Text = 'Minutos:'; AutoSize = $true; Location = (Pt 20 100) }
$numMin = Nuevo NumericUpDown @{ Minimum = 1; Maximum = 480; Value = 30; Location = (Pt 92 97); Width = 70 }
$btnVigIni = Boton '▶  Empezar a vigilar' 190 -Primario; $btnVigIni.Location = Pt 180 90
$btnVigFin = Boton '■  Terminar ahora' 160; $btnVigFin.Location = Pt 380 90; $btnVigFin.Enabled = $false
$btnVigInf = Boton 'Ver informe de vigilancia' 210; $btnVigInf.Location = Pt 550 90; $btnVigInf.Enabled = $false
$tabVig.Controls.AddRange(@($lblVigInfo, $lblMin, $numMin, $btnVigIni, $btnVigFin, $btnVigInf))
$medidores = @{}
$yMed = 150
foreach ($k in 'Procesador', 'Disco', 'Memoria RAM') {
    $l = Nuevo Label @{ Text = $k; AutoSize = $false; Size = (Sz 110 24); Location = (Pt 20 $yMed) }
    $pb = Nuevo ProgressBar @{ Maximum = 100; Size = (Sz 520 22); Location = (Pt 134 $yMed) }
    $lv = Nuevo Label @{ Text = '—'; AutoSize = $true; Location = (Pt 666 $yMed); Font = $FuenteNegrita }
    $tabVig.Controls.AddRange(@($l, $pb, $lv))
    $medidores[$k] = @{ Barra = $pb; Valor = $lv }
    $yMed += 34
}
$lblVigEstado = Nuevo Label @{ Text = 'Sin vigilancia en curso.'; AutoSize = $true; Location = (Pt 20 258); Font = $FuenteNegrita }
$lbEpisodios = Nuevo ListBox @{ Location = (Pt 20 286); Size = (Sz 1080 260); HorizontalScrollbar = $true; BorderStyle = 'FixedSingle' }
$tabVig.Controls.AddRange(@($lblVigEstado, $lbEpisodios))

# Registro
$txtLog = Nuevo TextBox @{ Dock = 'Fill'; Multiline = $true; ScrollBars = 'Vertical'; ReadOnly = $true; Font = (New-Object System.Drawing.Font('Consolas', 9.5)); BackColor = $Paleta.Blanco; BorderStyle = 'None' }
$tabLog.Padding = New-Object System.Windows.Forms.Padding(10)
$tabLog.Controls.Add($txtLog)

$form.Controls.AddRange(@($cuerpo, $pie, $panelProg, $barraAcc, $cab))

function Ajustar-Paneles {
    try {
        Ubicar-Cabecera
        $splitD.SplitterDistance = [int]($splitD.Width * 0.5)
        $splitA.SplitterDistance = [int]($splitA.Height * 0.55)
        $lvHall.Columns[2].Width = [math]::Max(200, $lvHall.ClientSize.Width - $lvHall.Columns[0].Width - $lvHall.Columns[1].Width - 4)
        $lvArr.Columns[1].Width = [math]::Max(250, $lvArr.ClientSize.Width - 40 - 90 - 80 - 150 - 170 - 4)
        $lblVigInfo.Width = $tabVig.ClientSize.Width - 40
        $lbEpisodios.Width = $tabVig.ClientSize.Width - 40; $lbEpisodios.Height = [math]::Max(120, $tabVig.ClientSize.Height - 300)
    } catch {}
}
$form.Add_Load({ Ajustar-Paneles })
$form.Add_SizeChanged({ if ($form.WindowState -ne 'Minimized') { Ajustar-Paneles } })

# ============================================================================
#  Contenido
# ============================================================================
function Mostrar-Bienvenida {
    $rtbHall.Clear()
    Rtb-Agregar $rtbHall "Bienvenido a Aquila`n`n" -Titulo -Color $Paleta.Marino
    Rtb-Agregar $rtbHall "Aquila revisa esta computadora, te explica en palabras simples por qué está lenta y te ofrece arreglos. Antes de cambiar cualquier cosa te dice qué hace, por qué y si afecta algo.`n`n"
    Rtb-Agregar $rtbHall "1.  Diagnosticar`n" -Negrita -Color $Paleta.Marino
    Rtb-Agregar $rtbHall "     Presiona el botón dorado «Diagnosticar». Tarda unos 3 minutos; puedes seguir usando la PC.`n`n"
    Rtb-Agregar $rtbHall "2.  Leer el resultado`n" -Negrita -Color $Paleta.Marino
    Rtb-Agregar $rtbHall "     Aquí verás cada problema encontrado. Haz clic en uno para ver por qué pasa y qué hacer.`n`n"
    Rtb-Agregar $rtbHall "3.  Arreglar`n" -Negrita -Color $Paleta.Marino
    Rtb-Agregar $rtbHall "     En la pestaña «Arreglos» marca lo que quieras y presiona «Aplicar los marcados». Antes se crea un punto de restauración para poder deshacer.`n`n"
    Rtb-Agregar $rtbHall "¿Se pone lenta solo a ratos?`n" -Negrita -Color $Paleta.Marino
    Rtb-Agregar $rtbHall "     Usa la pestaña «Vigilar»: atrapa al programa culpable justo en el momento en que ocurre.`n`n"
    Rtb-Agregar $rtbHall "Dudas: botón «Manual» arriba a la derecha." -Color $Paleta.Suave
}
function Mostrar-Hallazgo($h) {
    $rtbHall.Clear()
    if (-not $h) { return }
    Rtb-Agregar $rtbHall "$($EtqNivel[$h.Nivel]) · $($h.Categoria)`n" -Negrita -Color $ColorNivel[$h.Nivel]
    Rtb-Agregar $rtbHall "$($h.Titulo)`n`n" -Titulo -Color $Paleta.Marino
    if ($h.Detalle) { Rtb-Agregar $rtbHall "$($h.Detalle)`n`n" -Color $Paleta.Suave }
    if ($h.Porque) { Rtb-Agregar $rtbHall "¿Por qué pasa?`n" -Negrita -Color $Paleta.Marino; Rtb-Agregar $rtbHall "$($h.Porque)`n`n" }
    if ($h.Solucion) { Rtb-Agregar $rtbHall "Qué hacer`n" -Negrita -Color $Paleta.Marino; Rtb-Agregar $rtbHall "$($h.Solucion)`n`n" }
    $rel = @($script:Resultado.Arreglos | Where-Object { @($h.Arreglos) -contains $_.Id })
    if ($rel) {
        Rtb-Agregar $rtbHall "🔧 Aquila puede hacerlo por ti (pestaña «Arreglos»):`n" -Negrita -Color $Paleta.Azul
        foreach ($a in $rel) { Rtb-Agregar $rtbHall "     #$($a.Numero)  $($a.Nombre)`n" }
    }
}
function Mostrar-Arreglo($a) {
    $rtbArr.Clear()
    if (-not $a) { return }
    Rtb-Agregar $rtbArr "#$($a.Numero)  $($a.Nombre)`n" -Titulo -Color $Paleta.Marino
    $lib = if ([double]$a.Bytes -gt 0) { "Libera aprox. $(Fmt $a.Bytes)   ·   " } else { '' }
    Rtb-Agregar $rtbArr "$lib Riesgo: $($a.Riesgo)   ·   Tarda: $(if ($a.Duracion) { $a.Duracion } else { 'menos de 1 min' })$(if ($a.Admin -and -not $EsAdmin) { '   ·   REQUIERE ADMINISTRADOR' })`n`n" -Color $Paleta.Suave
    Rtb-Agregar $rtbArr "Qué hace:  " -Negrita -Color $Paleta.Marino; Rtb-Agregar $rtbArr "$($a.QueHace)`n`n"
    Rtb-Agregar $rtbArr "Por qué:  " -Negrita -Color $Paleta.Marino; Rtb-Agregar $rtbArr "$($a.Porque)`n`n"
    Rtb-Agregar $rtbArr "¿Afecta algo?  " -Negrita -Color $ColorNivel.Alto; Rtb-Agregar $rtbArr "$($a.Afecta)`n"
}

function Cargar-Resultado([string]$ruta) {
    $r = Leer-Json $ruta
    $script:Resultado = $r
    $script:Puntaje = [int]$r.Puntaje; $anillo.Invalidate()
    $lblSalud.Text = [string]$r.Estado
    $lblEquipo.Text = "$($r.Datos.Modelo)  ·  $($r.Datos.Procesador)  ·  RAM $($r.Datos.RAM_GB) GB  ·  Disco $($r.Datos.DiscoSistema)  ·  $($r.Datos.Fecha)"
    $lvHall.BeginUpdate(); $lvHall.Items.Clear()
    foreach ($h in @($r.Hallazgos)) {
        $it = New-Object System.Windows.Forms.ListViewItem($EtqNivel[$h.Nivel])
        [void]$it.SubItems.Add([string]$h.Categoria); [void]$it.SubItems.Add([string]$h.Titulo)
        $it.ForeColor = $ColorNivel[$h.Nivel]; $it.Tag = $h
        if ($h.Nivel -in 'Critico', 'Alto') { $it.Font = $FuenteNegrita }
        [void]$lvHall.Items.Add($it)
    }
    $lvHall.EndUpdate()
    $lvArr.BeginUpdate(); $lvArr.Items.Clear()
    foreach ($a in @($r.Arreglos)) {
        $it = New-Object System.Windows.Forms.ListViewItem([string]$a.Numero)
        [void]$it.SubItems.Add([string]$a.Nombre)
        [void]$it.SubItems.Add($(if ([double]$a.Bytes -gt 0) { Fmt $a.Bytes } else { '—' }))
        [void]$it.SubItems.Add([string]$a.Riesgo)
        [void]$it.SubItems.Add($(if ($a.Duracion) { [string]$a.Duracion } else { 'Menos de 1 min' }))
        [void]$it.SubItems.Add($(if ($a.Admin -and -not $EsAdmin) { 'Requiere administrador' } else { '' }))
        $it.Tag = $a
        if ($a.Riesgo -eq 'Medio') { $it.ForeColor = $ColorNivel.Medio }
        [void]$lvArr.Items.Add($it)
    }
    $lvArr.EndUpdate()
    $btnInforme.Enabled = $true; $btnAplicar.Enabled = ($lvArr.Items.Count -gt 0)
    if ($lvHall.Items.Count) { $lvHall.Items[0].Selected = $true }
    if ($lvArr.Items.Count) { $lvArr.Items[0].Selected = $true }
    $tabs.SelectedTab = $tabDiag
}

# ============================================================================
#  Diálogos
# ============================================================================
function Dialogo-Lista([string]$Titulo, [string]$Explicacion, $Elementos) {
    $d = Nuevo Form @{ Text = $Titulo; Size = (Sz 920 620); StartPosition = 'CenterParent'; Font = $Fuente; MinimizeBox = $false; BackColor = $Paleta.Blanco; Icon = $form.Icon }
    $lbl = Nuevo Label @{ Text = $Explicacion; Dock = 'Top'; Height = 70; Padding = (New-Object System.Windows.Forms.Padding(14, 12, 14, 0)); ForeColor = $Paleta.Texto }
    $cl = Nuevo CheckedListBox @{ Dock = 'Fill'; CheckOnClick = $true; HorizontalScrollbar = $true; IntegralHeight = $false; BorderStyle = 'None' }
    $i = 0
    foreach ($e in $Elementos) { [void]$cl.Items.Add($e.Texto); $cl.SetItemChecked($i, [bool]$e.Marcado); $i++ }
    $pb = Nuevo FlowLayoutPanel @{ Dock = 'Bottom'; Height = 58; Padding = (New-Object System.Windows.Forms.Padding(12, 10, 12, 8)); BackColor = $Paleta.Fondo }
    $ok = Boton 'Aceptar' 130 -Primario; $ok.DialogResult = 'OK'
    $ca = Boton 'Cancelar' 120; $ca.DialogResult = 'Cancel'
    $todo = Boton 'Marcar todo' 130; $nada = Boton 'Desmarcar todo' 140
    $todo.Add_Click({ for ($j = 0; $j -lt $cl.Items.Count; $j++) { $cl.SetItemChecked($j, $true) } }.GetNewClosure())
    $nada.Add_Click({ for ($j = 0; $j -lt $cl.Items.Count; $j++) { $cl.SetItemChecked($j, $false) } }.GetNewClosure())
    $pb.Controls.AddRange(@($ok, $ca, $todo, $nada))
    $d.Controls.AddRange(@($cl, $pb, $lbl))
    $d.AcceptButton = $ok; $d.CancelButton = $ca
    if ($d.ShowDialog($form) -ne 'OK') { return $null }
    $lista = @($Elementos)
    ,@($cl.CheckedIndices | ForEach-Object { $lista[$_].Valor })
}

function Dialogo-Confirmar($Lista) {
    $d = Nuevo Form @{ Text = 'Antes de aplicar: revisa lo que se va a hacer'; Size = (Sz 920 660); StartPosition = 'CenterParent'; Font = $Fuente; MinimizeBox = $false; BackColor = $Paleta.Blanco; Icon = $form.Icon }
    $rtb = Nuevo RichTextBox @{ Dock = 'Fill'; ReadOnly = $true; BorderStyle = 'None'; BackColor = $Paleta.Blanco }
    $marco = Nuevo Panel @{ Dock = 'Fill'; Padding = (New-Object System.Windows.Forms.Padding(18, 14, 12, 8)) }
    $marco.Controls.Add($rtb)
    Rtb-Agregar $rtb "Se aplicarán $(@($Lista).Count) arreglo(s)`n" -Titulo -Color $Paleta.Marino
    Rtb-Agregar $rtb "$(if ($EsAdmin) { 'Antes de empezar se creará un punto de restauración de Windows para poder deshacer los cambios del sistema.' } else { 'Sin permisos de administrador no se puede crear punto de restauración.' })`n`n" -Color $Paleta.Suave
    foreach ($a in $Lista) {
        Rtb-Agregar $rtb "■ #$($a.Numero)  $($a.Nombre)$(if ([double]$a.Bytes -gt 0) { "   (libera ~$(Fmt $a.Bytes))" })   ·   tarda: $(if ($a.Duracion) { $a.Duracion } else { 'menos de 1 min' })`n" -Negrita -Color $Paleta.Marino
        Rtb-Agregar $rtb "   Qué hace: $($a.QueHace)`n"
        Rtb-Agregar $rtb "   Por qué: $($a.Porque)`n"
        Rtb-Agregar $rtb "   ¿Afecta algo? $($a.Afecta)`n`n" -Color $ColorNivel.Alto
    }
    $pb = Nuevo FlowLayoutPanel @{ Dock = 'Bottom'; Height = 58; Padding = (New-Object System.Windows.Forms.Padding(12, 10, 12, 8)); BackColor = $Paleta.Fondo }
    $ok = Boton 'Sí, aplicar' 150 -Primario; $ok.DialogResult = 'OK'
    $ca = Boton 'Cancelar' 120; $ca.DialogResult = 'Cancel'
    $pb.Controls.AddRange(@($ok, $ca))
    $d.Controls.AddRange(@($marco, $pb))
    $d.CancelButton = $ca
    $d.ShowDialog($form) -eq 'OK'
}

# ============================================================================
#  Procesos en segundo plano
# ============================================================================
function Iniciar-Proceso([string]$Script, [string[]]$Argumentos, [string]$Log, [scriptblock]$AlTerminar, [string]$Texto) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$Script`" " + ($Argumentos -join ' ')
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true
    $script:LogActual = $Log; $script:LineasVistas = 0; $script:AlTerminar = $AlTerminar
    $script:Proc = [System.Diagnostics.Process]::Start($psi)
    $script:InicioProc = Get-Date; $script:PasoTexto = $Texto; $script:PasoN = 0; $script:PasoTotal = 0; $script:Giro = 0
    $barraProg.Style = 'Marquee'; $barraProg.Value = 0
    $lblProg.Text = "$Texto"
    $panelProg.Visible = $true
    $lblEstado.Text = $Texto
    foreach ($b in $btnDiag, $btnAplicar, $btnRestaurar, $btnActualizar) { $b.Enabled = $false }
    $timer.Start()
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 600
$timer.Add_Tick({
    $lineas = Leer-Lineas $script:LogActual
    for ($i = $script:LineasVistas; $i -lt $lineas.Count; $i++) {
        $txtLog.AppendText($lineas[$i] + "`r`n")
        $texto = ($lineas[$i] -replace '^\[[^\]]+\]\s*', '').Trim()
        $m = [regex]::Match($texto, '^\((\d+)/(\d+)\)\s*(.*)$')
        if ($m.Success) {
            $script:PasoN = [int]$m.Groups[1].Value; $script:PasoTotal = [int]$m.Groups[2].Value; $script:PasoTexto = $m.Groups[3].Value
            $barraProg.Style = 'Continuous'
            $barraProg.Value = [math]::Min(100, [math]::Max(3, [int](100 * ($script:PasoN - 1) / $script:PasoTotal)))
        } elseif ($texto -and -not $script:PasoTotal) { $script:PasoTexto = $texto }
        $lblEstado.Text = $texto
    }
    $script:LineasVistas = $lineas.Count
    if ($script:Proc -and -not $script:Proc.HasExited) {
        $script:Giro = ($script:Giro + 1) % 4
        $puntos = '.' * ($script:Giro + 1)
        $trans = (Get-Date) - $script:InicioProc
        $pasoTxt = if ($script:PasoTotal) { "Paso $($script:PasoN) de $($script:PasoTotal)  ·  $($barraProg.Value)%  ·  " } else { '' }
        $lblProg.Text = "Trabajando$puntos   $pasoTxt$($script:PasoTexto)   ·   $([int][math]::Floor($trans.TotalMinutes)):$('{0:00}' -f $trans.Seconds)"
    }
    if ($script:Proc -and $script:Proc.HasExited) {
        $timer.Stop()
        $script:Proc = $null
        $barraProg.Style = 'Continuous'; $barraProg.Value = 100; $panelProg.Visible = $false
        foreach ($b in $btnDiag, $btnRestaurar, $btnActualizar) { $b.Enabled = $true }
        $btnAplicar.Enabled = ($lvArr.Items.Count -gt 0)
        $cb = $script:AlTerminar
        try { & $cb } catch { Mensaje "Ocurrió un error al leer el resultado:`n$($_.Exception.Message)" 'Error' 'Error' }
    }
})

# ============================================================================
#  Actualizaciones
# ============================================================================
function Descargar-Texto([string]$archivo) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $u = "https://raw.githubusercontent.com/$AquilaRepo/main/$($archivo)?t=$([DateTime]::UtcNow.Ticks)"
    ([string](Invoke-WebRequest -Uri $u -UseBasicParsing -TimeoutSec 8).Content).TrimStart([char]0xFEFF)
}
# Pide el archivo a la API de GitHub (sin caché); si falla, usa la copia pública (tarda ~5 min en renovarse)
function Descargar-Fresco([string]$archivo) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    try {
        $w = Invoke-WebRequest -Uri "https://api.github.com/repos/$AquilaRepo/contents/$($archivo)?ref=main" -Headers @{ Accept = 'application/vnd.github.raw'; 'User-Agent' = 'Aquila' } -UseBasicParsing -TimeoutSec 8
        $txt = if ($w.Content -is [byte[]]) { [Text.Encoding]::UTF8.GetString($w.Content) } else { [string]$w.Content }
        $txt.TrimStart([char]0xFEFF)
    } catch { Descargar-Texto $archivo }
}
function Buscar-Actualizacion([switch]$Manual) {
    try { $remota = ([string](Descargar-Fresco 'version.txt')).Trim() }
    catch { if ($Manual) { Mensaje 'No se pudo revisar si hay actualizaciones. Revisa tu conexión a internet.' 'Actualizar' 'Warning' }; return }
    $hay = $false
    try { $hay = [version]$remota -gt [version]$AquilaVersion } catch {}
    if (-not $hay) { if ($Manual) { Mensaje "Ya tienes la última versión de Aquila ($AquilaVersion).`n`nVersión publicada en GitHub: $remota" 'Actualizar' }; return }
    $ocupado = ($script:Proc -and -not $script:Proc.HasExited) -or ($script:ProcVig -and -not $script:ProcVig.HasExited)
    if ($ocupado) { $lblEstado.Text = "Hay una nueva versión de Aquila ($remota). Actualiza cuando termine lo que está en curso."; return }
    $r = Mensaje "Hay una nueva versión de Aquila: $remota (tienes $AquilaVersion).`n`n¿Actualizar ahora? Aquila se cerrará y se abrirá sola otra vez en unos segundos. Tus informes se conservan." 'Actualización disponible' 'Information' 'YesNo'
    if ($r -ne 'Yes') { $lblEstado.Text = "Actualización $remota disponible: botón «Actualizar»."; return }
    try {
        $tmp = Join-Path $env:TEMP "aquila-instalar-$([DateTime]::Now.Ticks).ps1"
        [IO.File]::WriteAllText($tmp, (Descargar-Fresco 'instalar.ps1'), (New-Object Text.UTF8Encoding $true))
    } catch { Mensaje "No se pudo descargar la actualización:`n$($_.Exception.Message)" 'Actualizar' 'Warning'; return }
    # Ventana visible «AQUILA · Actualizando»; hereda estas variables y los permisos de administrador
    $env:AQUILA_ACTUALIZAR = '1'; $env:AQUILA_ESPERAR_PID = "$PID"
    Start-Process -FilePath 'powershell.exe' -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$tmp`""
    $script:SaliendoPorActualizacion = $true
    $form.Close()
}

# ============================================================================
#  Botones
# ============================================================================
$btnDiag.Add_Click({
    $m = Get-Date -Format 'yyyyMMdd_HHmmss'
    $log = Join-Path $env:TEMP "aquila_diag_$m.log"
    $res = Join-Path $env:TEMP "aquila_diag_$m.json"
    $argumentos = @('-NoAbrir', '-SinArreglos', "-Registro `"$log`"", "-ArchivoResultado `"$res`"")
    if (-not $chkDup.Checked) { $argumentos += '-SinDuplicados' }
    if ($script:ArchivoComparar) { $argumentos += "-Comparar `"$($script:ArchivoComparar)`"" }
    $txtLog.AppendText("`r`n===== Diagnóstico $(Get-Date -Format 'dd/MM/yyyy HH:mm') =====`r`n")
    $script:ResDiag = $res
    $rtbHall.Clear(); Rtb-Agregar $rtbHall "Diagnosticando...`n`n" -Titulo -Color $Paleta.Marino
    Rtb-Agregar $rtbHall 'Aquila está revisando el equipo, los discos, lo que corre en este momento, los programas al encender, la nube, el correo, el espacio y los duplicados. Tarda unos 3 minutos; puedes seguir usando la PC normalmente.' -Color $Paleta.Suave
    Iniciar-Proceso $ScriptDiag $argumentos $log {
        if (Test-Path -LiteralPath $script:ResDiag) {
            Cargar-Resultado $script:ResDiag
            Remove-Item -LiteralPath $script:ResDiag -Force
            $lblEstado.Text = "Diagnóstico terminado: $($script:Resultado.Puntaje)/100. Revisa «Diagnóstico» y luego «Arreglos»."
            $script:ArchivoComparar = $null; $btnComparar.Text = 'Comparar con otra PC…'
        } else {
            $lblEstado.Text = 'El diagnóstico no terminó correctamente. Revisa la pestaña «Registro».'
            $tabs.SelectedTab = $tabLog
        }
    } 'Diagnosticando... (unos 3 minutos)'
})

$btnComparar.Add_Click({
    $ofd = Nuevo OpenFileDialog @{ Title = 'Elige el diagnóstico (.json) de la otra PC'; Filter = 'Diagnóstico de Aquila (*.json)|Diagnostico_*.json|Todos los .json|*.json'; InitialDirectory = $Informes }
    if ($ofd.ShowDialog($form) -eq 'OK') {
        $script:ArchivoComparar = $ofd.FileName
        $btnComparar.Text = 'Comparar: ' + [IO.Path]::GetFileNameWithoutExtension($ofd.FileName).Replace('Diagnostico_', '')
        $lblEstado.Text = 'El próximo diagnóstico incluirá la comparación. Presiona «Diagnosticar».'
    }
})
$btnInforme.Add_Click({ if ($script:Resultado -and (Test-Path -LiteralPath $script:Resultado.InformeHtml)) { Start-Process $script:Resultado.InformeHtml } })
$btnCarpeta.Add_Click({ Start-Process explorer.exe $Informes })
$btnManual.Add_Click({ if (Test-Path -LiteralPath $Manual) { Start-Process $Manual } else { Start-Process "https://github.com/$AquilaRepo#readme" } })
$btnActualizar.Add_Click({ Buscar-Actualizacion -Manual })

$lvHall.Add_SelectedIndexChanged({ if ($lvHall.SelectedItems.Count) { Mostrar-Hallazgo $lvHall.SelectedItems[0].Tag } })
$lvArr.Add_SelectedIndexChanged({ if ($lvArr.SelectedItems.Count) { Mostrar-Arreglo $lvArr.SelectedItems[0].Tag } })
$lvArr.Add_ItemCheck({
    param($s, $e)
    $a = $lvArr.Items[$e.Index].Tag
    if ($e.NewValue -eq 'Checked' -and (($a.Admin -and -not $EsAdmin) -or $lvArr.Items[$e.Index].SubItems[5].Text -like 'Aplicado*')) { $e.NewValue = 'Unchecked' }
})
$btnMarcarBajo.Add_Click({ foreach ($it in $lvArr.Items) { $it.Checked = ($it.Tag.Riesgo -eq 'Bajo' -and -not $it.Tag.Lento) } })
$btnDesmarcar.Add_Click({ foreach ($it in $lvArr.Items) { $it.Checked = $false } })

$btnAplicar.Add_Click({
    $marcados = @($lvArr.CheckedItems | ForEach-Object { $_.Tag } | Sort-Object { [int]$_.Numero })
    if (-not $marcados) { Mensaje 'Marca con la casilla los arreglos que quieres aplicar.'; return }
    $final = @()
    foreach ($a in $marcados) {
        $c = $a | ConvertTo-Json -Depth 8 | ConvertFrom-Json    # copia independiente
        if ($c.Tipo -eq 'inicio') {
            $el = @($c.Datos.Lista | ForEach-Object { [pscustomobject]@{ Texto = "$($_.Nombre)   —   $($_.Recomendacion)"; Marcado = ($_.Recomendacion -like 'Se puede*'); Valor = $_ } })
            $elegidos = Dialogo-Lista 'Programas que arrancan al encender' 'Marca los programas que NO quieres que se abran solos al encender. No se desinstala nada; se puede revertir con «Restaurar programas de inicio».' $el
            if ($null -eq $elegidos -or -not @($elegidos).Count) { continue }
            $c.Datos.Lista = @($elegidos)
        }
        if ($c.Tipo -eq 'duplicados') {
            $el = @(); $gi = 0
            foreach ($g in @($c.Datos.Grupos)) {
                foreach ($cp in @($g.Copias)) { $el += [pscustomobject]@{ Texto = "$cp   ($(Fmt $g.Bytes))   —   se conserva: $($g.Conservar)"; Marcado = $true; Valor = [pscustomobject]@{ G = $gi; Ruta = $cp } } }
                $gi++
            }
            $elegidos = Dialogo-Lista 'Copias duplicadas que irán a la Papelera' 'Las copias marcadas se envían a la PAPELERA (recuperables). De cada archivo siempre se conserva una copia (indicada a la derecha). Desmarca las que quieras mantener.' $el
            if ($null -eq $elegidos -or -not @($elegidos).Count) { continue }
            $grupos = @($c.Datos.Grupos)
            $nuevos = @()
            for ($j = 0; $j -lt $grupos.Count; $j++) {
                $cps = @($elegidos | Where-Object { $_.G -eq $j } | ForEach-Object { $_.Ruta })
                if ($cps) { $nuevos += [pscustomobject]@{ Bytes = $grupos[$j].Bytes; Conservar = $grupos[$j].Conservar; Copias = $cps } }
            }
            $c.Datos.Grupos = $nuevos
            $c.Bytes = [double](($nuevos | ForEach-Object { [double]$_.Bytes * @($_.Copias).Count } | Measure-Object -Sum).Sum)
        }
        if ($c.Datos.Procesos) {
            $abiertos = @(Get-Process -Name @($c.Datos.Procesos) -ErrorAction SilentlyContinue)
            if ($abiertos) {
                $r = Mensaje "$($c.Datos.App) está abierto y hay que cerrarlo para «$($c.Nombre)».`n`nGuarda tu trabajo. Al presionar «Sí» se cerrará automáticamente.`n«No» = omitir este arreglo." 'Hay que cerrar un programa' 'Warning' 'YesNo'
                if ($r -ne 'Yes') { continue }
                $c.Datos | Add-Member -NotePropertyName ForzarCierre -NotePropertyValue $true -Force
            }
        }
        $final += $c
    }
    if (-not $final) { return }
    if (-not (Dialogo-Confirmar $final)) { return }
    $m = Get-Date -Format 'yyyyMMdd_HHmmss'
    $archivo = Join-Path $env:TEMP "aquila_aplicar_$m.json"
    $script:ResAplicar = Join-Path $env:TEMP "aquila_aplicar_$m.resultado.json"
    $log = Join-Path $Informes "Arreglos_$($env:COMPUTERNAME)_$m.log"
    Guardar-Json @($final) $archivo
    $txtLog.AppendText("`r`n===== Aplicando arreglos $(Get-Date -Format 'dd/MM/yyyy HH:mm') =====`r`n")
    $tabs.SelectedTab = $tabLog
    Iniciar-Proceso $ScriptDiag @("-Aplicar `"$archivo`"", "-Registro `"$log`"", "-ArchivoResultado `"$($script:ResAplicar)`"") $log {
        if (-not (Test-Path -LiteralPath $script:ResAplicar)) { Mensaje 'No se pudo leer el resultado. Revisa la pestaña «Registro».' 'Aviso' 'Warning'; return }
        $r = Leer-Json $script:ResAplicar
        $lineas = foreach ($x in @($r.Resultados)) {
            foreach ($it in $lvArr.Items) { if ($it.Tag.Id -eq $x.Id) { $it.Checked = $false; $it.SubItems[5].Text = $(if ($x.Ok) { 'Aplicado' } else { 'Error' }); $it.ForeColor = [System.Drawing.Color]::Gray } }
            "$(if ($x.Ok) { '✔' } else { '✖' })  $($x.Nombre)`n      $($x.Mensaje)"
        }
        $lblEstado.Text = "Arreglos aplicados. Espacio liberado: $($r.LiberadoTexto)."
        $resp = Mensaje ("Resultado:`n`n" + ($lineas -join "`n`n") + "`n`nEspacio liberado: $($r.LiberadoTexto)`n`nRecomendado: reiniciar la PC.`n`n¿Diagnosticar de nuevo ahora para ver la mejora (comparando con el diagnóstico anterior)?") 'Arreglos aplicados' 'Information' 'YesNo'
        if ($resp -eq 'Yes') {
            $script:ArchivoComparar = $script:Resultado.ArchivoJson
            $btnComparar.Text = 'Comparar: diagnóstico anterior'
            $btnDiag.PerformClick()
        }
    } 'Aplicando arreglos... (no apagues la PC)'
})

$btnRestaurar.Add_Click({
    $r = Mensaje 'Esto vuelve a activar todos los programas de inicio que Aquila desactivó en esta PC. ¿Continuar?' 'Restaurar programas de inicio' 'Question' 'YesNo'
    if ($r -ne 'Yes') { return }
    $log = Join-Path $env:TEMP "aquila_restaurar_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Iniciar-Proceso $ScriptDiag @('-RestaurarInicio', "-Registro `"$log`"") $log {
        $txt = (Leer-Lineas $script:LogActual | ForEach-Object { $_ -replace '^\[[^\]]+\]\s*', '' }) -join "`n"
        Mensaje $(if ($txt) { $txt } else { 'Listo.' }) 'Restaurar programas de inicio'
    } 'Restaurando programas de inicio...'
})

# --- Vigilar ---
$timerVig = New-Object System.Windows.Forms.Timer
$timerVig.Interval = 1000
$timerVig.Add_Tick({
    if ($script:EstadoVig -and (Test-Path -LiteralPath $script:EstadoVig)) {
        try {
            $e = Leer-Json $script:EstadoVig
            foreach ($p in @(@('Procesador', $e.Cpu), @('Disco', $e.Disco), @('Memoria RAM', $e.Ram))) {
                $v = [int][math]::Min(100, [math]::Max(0, [double]$p[1]))
                $medidores[$p[0]].Barra.Value = $v; $medidores[$p[0]].Valor.Text = "$v%"
            }
            $lblVigEstado.Text = if ($e.Terminado) { "Vigilancia terminada: $($e.Episodios) momentos de lentitud." } else { "Vigilando... quedan $([math]::Ceiling($e.Restante / 60)) min · momentos de lentitud detectados: $($e.Episodios)" }
            $lbEpisodios.BeginUpdate(); $lbEpisodios.Items.Clear(); foreach ($u in @($e.Ultimos)) { [void]$lbEpisodios.Items.Add($u) }; $lbEpisodios.EndUpdate()
            if ($e.Informe) { $script:InformeVig = $e.Informe }
        } catch {}
    }
    if ($script:ProcVig -and $script:ProcVig.HasExited) {
        $timerVig.Stop(); $script:ProcVig = $null
        $btnVigIni.Enabled = $true; $btnVigFin.Enabled = $false; $numMin.Enabled = $true
        if ($script:InformeVig) {
            $btnVigInf.Enabled = $true
            if ((Mensaje 'La vigilancia terminó. ¿Abrir el informe con el gráfico y los culpables?' 'Vigilancia' 'Question' 'YesNo') -eq 'Yes') { Start-Process $script:InformeVig }
        }
    }
})
$btnVigIni.Add_Click({
    $m = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:EstadoVig = Join-Path $env:TEMP "aquila_vig_$m.estado.json"
    $script:PararVig = Join-Path $env:TEMP "aquila_vig_$m.parar"
    $script:InformeVig = $null
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptVig`" -Minutos $([int]$numMin.Value) -NoAbrir -ArchivoEstado `"$($script:EstadoVig)`" -ArchivoParar `"$($script:PararVig)`""
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true
    $script:ProcVig = [System.Diagnostics.Process]::Start($psi)
    $btnVigIni.Enabled = $false; $btnVigFin.Enabled = $true; $btnVigInf.Enabled = $false; $numMin.Enabled = $false
    $lbEpisodios.Items.Clear(); $lblVigEstado.Text = 'Iniciando vigilancia...'
    $timerVig.Start()
})
$btnVigFin.Add_Click({ if ($script:PararVig) { Set-Content -LiteralPath $script:PararVig -Value 'parar'; $lblVigEstado.Text = 'Terminando y armando el informe...'; $btnVigFin.Enabled = $false } })
$btnVigInf.Add_Click({ if ($script:InformeVig -and (Test-Path -LiteralPath $script:InformeVig)) { Start-Process $script:InformeVig } })

$form.Add_FormClosing({
    param($s, $e)
    if ($script:SaliendoPorActualizacion) { return }
    $ocupado = ($script:Proc -and -not $script:Proc.HasExited) -or ($script:ProcVig -and -not $script:ProcVig.HasExited)
    if ($ocupado) {
        $r = Mensaje 'Hay un proceso en curso (diagnóstico, arreglos o vigilancia). Si cierras ahora se detendrá. ¿Cerrar de todos modos?' 'Cerrar Aquila' 'Warning' 'YesNo'
        if ($r -ne 'Yes') { $e.Cancel = $true; return }
        foreach ($p in $script:Proc, $script:ProcVig) { if ($p -and -not $p.HasExited) { try { $p.Kill() } catch {} } }
    }
})

# Revisar actualizaciones poco después de abrir (sin bloquear el arranque)
$timerInicio = New-Object System.Windows.Forms.Timer
$timerInicio.Interval = 1500
$timerInicio.Add_Tick({ $timerInicio.Stop(); Buscar-Actualizacion })
$form.Add_Shown({ Ajustar-Paneles; if (-not $Probar) { $timerInicio.Start() } })

Mostrar-Bienvenida

if ($Probar) {
    Cargar-Resultado $Probar
    foreach ($it in $lvArr.Items) { Mostrar-Arreglo $it.Tag }
    foreach ($it in $lvHall.Items) { Mostrar-Hallazgo $it.Tag }
    "OK: $($lvHall.Items.Count) hallazgos, $($lvArr.Items.Count) arreglos cargados en la ventana."
    if ($Captura) {
        $form.StartPosition = 'Manual'; $form.Location = New-Object System.Drawing.Point(-3000, 0)
        $form.Show(); Ajustar-Paneles
        foreach ($tp in $tabs.TabPages) {
            $tabs.SelectedTab = $tp; [System.Windows.Forms.Application]::DoEvents()
            $bmp = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
            $form.DrawToBitmap($bmp, (New-Object System.Drawing.Rectangle(0, 0, $form.Width, $form.Height)))
            $bmp.Save(($Captura -replace '\.png$', "_$($tp.Text).png")); $bmp.Dispose()
        }
        $form.Hide()
    }
    $form.Dispose()
    return
}
[void]$form.ShowDialog()
