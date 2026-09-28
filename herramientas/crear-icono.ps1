# Genera el símbolo de Aquila (un águila) en app\aquila.ico, app\logo.png y docs\aquila.svg.
# Las formas están en un lienzo de 100x100; el SVG y las imágenes usan exactamente los mismos puntos.
# Uso: powershell -ExecutionPolicy Bypass -File herramientas\crear-icono.ps1

Add-Type -AssemblyName System.Drawing
$raiz = Split-Path $PSScriptRoot -Parent

$Formas = [ordered]@{
    # Plumas oscuras del cuerpo (abajo)
    plumas = @{ Color = '#5B3A1E'; Puntos = @(4,80, 14,72, 22,80, 30,71, 38,81, 46,72, 52,82, 60,74, 66,84, 74,76, 82,86, 96,78, 100,100, 0,100) }
    # Cabeza blanca
    cabeza = @{ Color = '#F8FAFC'; Puntos = @(20,90, 16,64, 20,46, 29,31, 42,22, 56,19, 67,22, 75,29, 79,36, 75,48, 66,55, 60,64, 58,78, 60,90) }
    # Pico dorado en gancho
    pico   = @{ Color = '#F2B705'; Puntos = @(75,31, 85,32, 93,37, 97,45, 95,53, 91,58, 90,52, 86,48, 79,50, 73,48, 78,40) }
    # Ceja (le da la mirada de águila)
    ceja   = @{ Color = '#1E293B'; Puntos = @(50,29, 60,25, 71,27, 70,31, 60,30, 52,33) }
}
$Ojo = @{ X = 62; Y = 35; R = 3.6; Color = '#0B1220' }
$Fondo = @{ Desde = '#1E3A8A'; Hasta = '#0B1B3F' }

function Color([string]$hex) { [System.Drawing.ColorTranslator]::FromHtml($hex) }

function Dibujar([int]$lado) {
    $bmp = New-Object System.Drawing.Bitmap($lado, $lado, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'; $g.InterpolationMode = 'HighQualityBicubic'; $g.PixelOffsetMode = 'HighQuality'
    $g.Clear([System.Drawing.Color]::Transparent)
    $k = $lado / 100.0
    $circulo = New-Object System.Drawing.Drawing2D.GraphicsPath
    $circulo.AddEllipse(0, 0, $lado - 1, $lado - 1)
    $grad = New-Object System.Drawing.Drawing2D.LinearGradientBrush((New-Object System.Drawing.PointF(0, 0)), (New-Object System.Drawing.PointF($lado, $lado)), (Color $Fondo.Desde), (Color $Fondo.Hasta))
    $g.FillPath($grad, $circulo)
    $g.SetClip($circulo)
    foreach ($f in $Formas.Values) {
        $pts = for ($i = 0; $i -lt $f.Puntos.Count; $i += 2) { New-Object System.Drawing.PointF(($f.Puntos[$i] * $k), ($f.Puntos[$i + 1] * $k)) }
        $g.FillPolygon((New-Object System.Drawing.SolidBrush(Color $f.Color)), [System.Drawing.PointF[]]$pts)
    }
    $g.FillEllipse((New-Object System.Drawing.SolidBrush(Color $Ojo.Color)), [single](($Ojo.X - $Ojo.R) * $k), [single](($Ojo.Y - $Ojo.R) * $k), [single](2 * $Ojo.R * $k), [single](2 * $Ojo.R * $k))
    $g.ResetClip()
    # Aro dorado
    $grosor = [math]::Max(1, $lado * 0.035)
    $g.DrawEllipse((New-Object System.Drawing.Pen((Color '#F2B705'), $grosor)), [single]($grosor / 2), [single]($grosor / 2), [single]($lado - 1 - $grosor), [single]($lado - 1 - $grosor))
    $g.Dispose()
    $bmp
}

# PNG del logo
$logo = Dibujar 256
$logo.Save((Join-Path $raiz 'app\logo.png'), [System.Drawing.Imaging.ImageFormat]::Png)
$logo.Dispose()

# ICO con varios tamaños (cada uno guardado como PNG dentro del ICO)
$tamanos = 16, 24, 32, 48, 64, 128, 256
$datos = foreach ($t in $tamanos) {
    $b = Dibujar $t; $ms = New-Object IO.MemoryStream
    $b.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png); $b.Dispose()
    , $ms.ToArray()
}
$ico = New-Object IO.MemoryStream
$w = New-Object IO.BinaryWriter($ico)
$w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]$tamanos.Count)
$offset = 6 + 16 * $tamanos.Count
for ($i = 0; $i -lt $tamanos.Count; $i++) {
    $t = $tamanos[$i]; $lado = if ($t -ge 256) { 0 } else { $t }
    $w.Write([byte]$lado); $w.Write([byte]$lado); $w.Write([byte]0); $w.Write([byte]0)
    $w.Write([uint16]1); $w.Write([uint16]32); $w.Write([uint32]$datos[$i].Length); $w.Write([uint32]$offset)
    $offset += $datos[$i].Length
}
foreach ($d in $datos) { $w.Write($d) }
[IO.File]::WriteAllBytes((Join-Path $raiz 'app\aquila.ico'), $ico.ToArray())

# SVG (mismas formas)
$svg = New-Object System.Text.StringBuilder
[void]$svg.Append("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100' role='img' aria-label='Aquila'><defs><linearGradient id='f' x1='0' y1='0' x2='1' y2='1'><stop offset='0' stop-color='$($Fondo.Desde)'/><stop offset='1' stop-color='$($Fondo.Hasta)'/></linearGradient><clipPath id='c'><circle cx='50' cy='50' r='50'/></clipPath></defs>")
[void]$svg.Append("<circle cx='50' cy='50' r='50' fill='url(#f)'/><g clip-path='url(#c)'>")
foreach ($f in $Formas.Values) {
    $p = for ($i = 0; $i -lt $f.Puntos.Count; $i += 2) { "$($f.Puntos[$i]),$($f.Puntos[$i + 1])" }
    [void]$svg.Append("<polygon fill='$($f.Color)' points='$($p -join ' ')'/>")
}
[void]$svg.Append("<circle cx='$($Ojo.X)' cy='$($Ojo.Y)' r='$($Ojo.R)' fill='$($Ojo.Color)'/></g><circle cx='50' cy='50' r='48.25' fill='none' stroke='#F2B705' stroke-width='3.5'/></svg>")
[void](New-Item -ItemType Directory (Join-Path $raiz 'docs') -Force)
[IO.File]::WriteAllText((Join-Path $raiz 'docs\aquila.svg'), $svg.ToString(), (New-Object Text.UTF8Encoding $false))
'Símbolo generado: app\aquila.ico, app\logo.png, docs\aquila.svg'
