# Aquila — desinstalador. Quita el programa y sus accesos directos. Los informes (Documentos\Aquila) se conservan.
Add-Type -AssemblyName System.Windows.Forms
$r = [System.Windows.Forms.MessageBox]::Show("¿Quitar Aquila de esta computadora?`n`nTus informes en Documentos\Aquila se conservan.", 'Desinstalar Aquila', 'YesNo', 'Question')
if ($r -ne 'Yes') { return }
$ErrorActionPreference = 'SilentlyContinue'
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like '*\Aquila\app\Aquila*' -and $_.ProcessId -ne $PID } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
Remove-Item (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Aquila.lnk') -Force
Remove-Item (Join-Path ([Environment]::GetFolderPath('Programs')) 'Aquila') -Recurse -Force
Remove-Item (Join-Path $env:LOCALAPPDATA 'Aquila') -Recurse -Force
[void][System.Windows.Forms.MessageBox]::Show('Aquila se quitó de esta computadora.', 'Desinstalar Aquila', 'OK', 'Information')
