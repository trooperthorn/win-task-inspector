<#
.SYNOPSIS
  Builds dist\TaskInspector.exe from src\Program.cs with TaskInspector.ps1 embedded.

.DESCRIPTION
  Uses the C# compiler that ships with .NET Framework 4.x, so nothing needs installing.
  The script is embedded as a resource; edit TaskInspector.ps1 and rebuild.
  Generates src\TaskInspector.ico on first run if it is missing.
#>
[CmdletBinding()]
param([switch]$Run)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$src  = Join-Path $root 'src'
$dist = Join-Path $root 'dist'
New-Item -ItemType Directory -Force $dist | Out-Null

$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) { $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
if (-not (Test-Path $csc)) { throw 'csc.exe not found; .NET Framework 4.x is required.' }

$sma = [PSObject].Assembly.Location   # System.Management.Automation from the GAC
$ico = Join-Path $src 'TaskInspector.ico'

if (-not (Test-Path $ico)) {
    # Simple generated icon: dark tile with a clock face and a check mark.
    Add-Type -AssemblyName System.Drawing
    $bmp = New-Object System.Drawing.Bitmap 256, 256
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $bg = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 30, 60, 100))
    $g.FillEllipse($bg, 8, 8, 240, 240)
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), 14
    $g.DrawEllipse($pen, 40, 40, 176, 176)
    $g.DrawLine($pen, 128, 128, 128, 68)
    $g.DrawLine($pen, 128, 128, 172, 150)
    $green = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 80, 200, 120)), 22
    $green.StartCap = 'Round'; $green.EndCap = 'Round'
    $g.DrawLine($green, 150, 200, 178, 228)
    $g.DrawLine($green, 178, 228, 236, 170)
    $g.Dispose()
    $hIcon = $bmp.GetHicon()
    $icon = [System.Drawing.Icon]::FromHandle($hIcon)
    $fs = [System.IO.File]::Create($ico)
    $icon.Save($fs); $fs.Close()
    $bmp.Dispose()
    Write-Host "Generated $ico"
}

$out = Join-Path $dist 'TaskInspector.exe'
$args = @(
    '/nologo', '/target:winexe', '/optimize+', '/platform:anycpu',
    "/out:$out",
    "/win32icon:$ico",
    "/win32manifest:$(Join-Path $src 'app.manifest')",
    "/resource:$(Join-Path $root 'TaskInspector.ps1'),TaskInspector.ps1",
    "/r:$sma",
    '/r:System.Windows.Forms.dll',
    (Join-Path $src 'Program.cs')
)
Write-Host "Compiling $out"
& $csc @args
if ($LASTEXITCODE -ne 0) { throw "csc failed with exit code $LASTEXITCODE" }
$fi = Get-Item $out
Write-Host ("Built {0} ({1:N0} bytes) version {2}" -f $fi.Name, $fi.Length, $fi.VersionInfo.FileVersion)
if ($Run) { Start-Process $out }
