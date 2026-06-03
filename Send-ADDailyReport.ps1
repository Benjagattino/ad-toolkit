<#
.SYNOPSIS
    Ejecuta los reportes de AD y envia un email HTML consolidado con graficos SVG.
    Genera Excel con multiples hojas (sin modulos externos - nativo .NET 4.5).
    Disenado para correr como tarea programada (headless).
.PARAMETER ConfigPath
    Ruta al config.json. Default: mismo directorio que el script.
.PARAMETER LogPath
    Archivo de log. Default: .\Reports\daily_log.txt
.PARAMETER NoMail
    Si se especifica, genera los reportes pero no envia email.
.EXAMPLE
    .\Send-ADDailyReport.ps1
.EXAMPLE
    .\Send-ADDailyReport.ps1 -NoMail
#>
#Requires -Module ActiveDirectory, GroupPolicy
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json'),
    [string]$LogPath    = (Join-Path $PSScriptRoot 'Reports\daily_log.txt'),
    [switch]$NoMail
)

$ErrorActionPreference = 'Continue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Cargar generador de Excel nativo (sin modulos externos)
$xlsxHelper = Join-Path $PSScriptRoot 'Export-MultiSheetXlsx.ps1'
if (Test-Path $xlsxHelper) { . $xlsxHelper } else { Write-Warning "No se encontro Export-MultiSheetXlsx.ps1" }
$logBuf = New-Object 'System.Collections.Generic.List[string]'

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Msg"
    $logBuf.Add($entry)
    Write-Host $entry
}

function Flush-Log {
    $dir = Split-Path $LogPath
    if ($dir) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $logBuf | Out-File $LogPath -Append -Encoding UTF8
}

#region -- CONFIG ---------------------------------------------------------------
if (-not (Test-Path $ConfigPath)) {
    Write-Log "config.json no encontrado en: $ConfigPath. Configure la GUI primero." 'ERROR'
    Flush-Log; exit 1
}
$cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
Write-Log "Configuracion cargada desde: $ConfigPath"

# Defaults para campos opcionales
$daysPassExp = if ($cfg.PSObject.Properties['DaysPasswordExpiry'] -and $cfg.DaysPasswordExpiry) { [int]$cfg.DaysPasswordExpiry } else { 15 }

$reportDir = Join-Path $PSScriptRoot 'Reports'
New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
$dateStr = Get-Date -Format 'yyyyMMdd'
#endregion

#region -- MODULO 1: USUARIOS INACTIVOS -----------------------------------------
Write-Log "--- Modulo 1: Usuarios Inactivos (umbral: $($cfg.DaysInactiveUsers) dias) ---"
$cutU = (Get-Date).AddDays(-[int]$cfg.DaysInactiveUsers)
$inactiveUsers    = @()
$totalEnabledUsers = 0

try {
    $p = @{ Filter = '*'; Properties = 'LastLogonDate','EmailAddress','Department','Enabled','PasswordLastSet' }
    if ($cfg.SearchBase) { $p['SearchBase'] = $cfg.SearchBase }
    $allEnabled = @(Get-ADUser @p | Where-Object { $_.Enabled })
    $totalEnabledUsers = $allEnabled.Count

    $inactiveUsers = $allEnabled | Where-Object {
        $null -eq $_.LastLogonDate -or $_.LastLogonDate -lt $cutU
    } | ForEach-Object {
        $d = if ($_.LastLogonDate) { [math]::Round(((Get-Date) - $_.LastLogonDate).TotalDays) } else { 9999 }
        [PSCustomObject]@{
            Usuario          = $_.SamAccountName
            Nombre           = $_.Name
            Email            = $_.EmailAddress
            Departamento     = $_.Department
            UltimoLogin      = if ($_.LastLogonDate) { $_.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Nunca' }
            DiasInactivo     = $d
            UltimoCambioPass = if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString('yyyy-MM-dd') } else { 'Nunca' }
        }
    } | Sort-Object DiasInactivo -Descending

    Write-Log "Usuarios: total habilitados=$totalEnabledUsers, inactivos=$(@($inactiveUsers).Count)"
}
catch {
    Write-Log "Error en usuarios: $($_.Exception.Message)" 'ERROR'
}
#endregion

#region -- MODULO 2: BACKUP DE GPOs --------------------------------------------
Write-Log "--- Modulo 2: Backup de GPOs ---"
$gpoOK = 0; $gpoErr = 0; $gpoRows = @(); $gpoDestino = ''
try {
    # Preferir NASBackupPath si esta configurado, sino usar BackupRoot local
    $hasNAS = ($cfg.PSObject.Properties.Name -contains 'NASBackupPath') -and (-not [string]::IsNullOrWhiteSpace($cfg.NASBackupPath))
    if ($hasNAS) {
        $bkpRoot = $cfg.NASBackupPath
        Write-Log "Destino backup: NAS ($bkpRoot)"
    } else {
        $bkpRoot = $cfg.BackupRoot
        Write-Log "Destino backup: local ($bkpRoot)"
    }
    $ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
    $dest    = Join-Path $bkpRoot $ts
    $gpoDestino = $dest
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    $domain = (Get-ADDomain).DNSRoot
    Get-GPO -All -Domain $domain | ForEach-Object {
        try {
            Backup-GPO -Guid $_.Id -Path $dest -Domain $domain | Out-Null
            $gpoRows += [PSCustomObject]@{ GPO = $_.DisplayName; Estado = 'OK'; Detalle = '' }
            $gpoOK++
        }
        catch {
            $gpoRows += [PSCustomObject]@{ GPO = $_.DisplayName; Estado = 'ERROR'; Detalle = $_.Exception.Message }
            $gpoErr++
            Write-Log "Error GPO '$($_.DisplayName)': $($_.Exception.Message)" 'WARN'
        }
    }
    # Rotacion
    Get-ChildItem $bkpRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{8}_' } | Sort-Object Name -Descending |
        Select-Object -Skip ([int]$cfg.MaxBackups) |
        ForEach-Object { Remove-Item $_.FullName -Recurse -Force; Write-Log "Backup antiguo eliminado: $($_.Name)" }
    Write-Log "GPO Backup: OK=$gpoOK, Errores=$gpoErr"
}
catch {
    Write-Log "Error en backup GPO: $($_.Exception.Message)" 'ERROR'
}
#endregion

#region -- MODULO 3: INVENTARIO DE EQUIPOS ------------------------------------
Write-Log "--- Modulo 3: Inventario de Equipos (umbral: $($cfg.DaysInactivePC) dias) ---"
$cutPC = (Get-Date).AddDays(-[int]$cfg.DaysInactivePC)
$allPCs = @()
try {
    $p = @{
        Filter     = if ($cfg.OSFilter) { "OperatingSystem -like '$($cfg.OSFilter)'" } else { '*' }
        Properties = 'OperatingSystem','LastLogonDate','IPv4Address','Enabled'
    }
    if ($cfg.SearchBase) { $p['SearchBase'] = $cfg.SearchBase }
    $allPCs = Get-ADComputer @p | ForEach-Object {
        $d = if ($_.LastLogonDate) { [math]::Round(((Get-Date) - $_.LastLogonDate).TotalDays) } else { 9999 }
        [PSCustomObject]@{
            Nombre           = $_.Name
            IP               = $_.IPv4Address
            SistemaOperativo = $_.OperatingSystem
            Habilitado       = $_.Enabled
            UltimaConexion   = if ($_.LastLogonDate) { $_.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Nunca' }
            DiasInactivo     = $d
            Inactivo         = ($null -eq $_.LastLogonDate -or $_.LastLogonDate -lt $cutPC)
        }
    } | Sort-Object DiasInactivo -Descending
    Write-Log "Equipos: total=$(@($allPCs).Count), inactivos=$(@($allPCs | Where-Object Inactivo).Count)"
}
catch {
    Write-Log "Error en inventario: $($_.Exception.Message)" 'ERROR'
}
#endregion

#region -- MODULO 4: CONTRASENAS POR VENCER ------------------------------------
Write-Log "--- Modulo 4: Contrasenas por Vencer (proximos $daysPassExp dias) ---"
$expUsers = @()
try {
    $maxPwd = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
    if ($maxPwd.TotalDays -gt 0) {
        $p = @{ Filter = '*'; Properties = 'PasswordLastSet','EmailAddress','Department','Enabled','PasswordNeverExpires','LockedOut' }
        if ($cfg.SearchBase) { $p['SearchBase'] = $cfg.SearchBase }
        $expUsers = Get-ADUser @p | Where-Object {
            $_.Enabled -and -not $_.PasswordNeverExpires -and $_.PasswordLastSet
        } | ForEach-Object {
            $expires   = $_.PasswordLastSet + $maxPwd
            $daysLeft  = [math]::Round(($expires - (Get-Date)).TotalDays)
            [PSCustomObject]@{
                Usuario       = $_.SamAccountName
                Nombre        = $_.Name
                Email         = $_.EmailAddress
                Departamento  = $_.Department
                FechaVence    = $expires.ToString('yyyy-MM-dd')
                DiasRestantes = $daysLeft
                Urgencia      = if ($daysLeft -le 3) { 'CRITICO' } elseif ($daysLeft -le 7) { 'Alto' } else { 'Medio' }
            }
        } | Where-Object { $_.DiasRestantes -ge 0 -and $_.DiasRestantes -le $daysPassExp } |
          Sort-Object DiasRestantes
        Write-Log "Contrasenas por vencer (proximos $daysPassExp dias): $(@($expUsers).Count)"
    }
    else {
        Write-Log "Politica de dominio: contrasenas sin expiracion configurada."
    }
}
catch {
    Write-Log "Error en modulo contrasenas: $($_.Exception.Message)" 'ERROR'
}
#endregion

#region -- EXPORTAR CSVs (solo local, no se adjuntan al email) -----------------
$csvUsers = Join-Path $reportDir "Users_$dateStr.csv"
$csvPCs   = Join-Path $reportDir "PCs_$dateStr.csv"
$csvGPO   = Join-Path $reportDir "GPO_$dateStr.csv"
$csvExp   = Join-Path $reportDir "PassExp_$dateStr.csv"

$attachments = @()   # Los CSVs se guardan localmente pero NO se adjuntan al email
if ($inactiveUsers) { $inactiveUsers | Export-Csv $csvUsers -NoTypeInformation -Encoding UTF8 }
if ($allPCs)        { $allPCs        | Export-Csv $csvPCs   -NoTypeInformation -Encoding UTF8 }
if ($gpoRows)       { $gpoRows       | Export-Csv $csvGPO   -NoTypeInformation -Encoding UTF8 }
if ($expUsers)      { $expUsers      | Export-Csv $csvExp   -NoTypeInformation -Encoding UTF8 }
Write-Log "CSVs exportados localmente en: $reportDir"
#endregion

#region -- GENERAR EXCEL (nativo .NET 4.5 - sin modulos externos) --------------
$xlsxPath = $null
if (Get-Command Export-MultiSheetXlsx -ErrorAction SilentlyContinue) {
    $xlsxPath = Join-Path $reportDir "ADReport_$dateStr.xlsx"
    try {
        Remove-Item $xlsxPath -Force -ErrorAction SilentlyContinue

        $sheets = @()
        if (@($inactiveUsers).Count -gt 0) {
            $sheets += @{ Name = 'Usuarios Inactivos';     Data = $inactiveUsers }
        }
        if (@($allPCs | Where-Object Inactivo).Count -gt 0) {
            $sheets += @{ Name = 'Equipos Inactivos';      Data = @($allPCs | Where-Object Inactivo) }
        }
        if (@($expUsers).Count -gt 0) {
            $sheets += @{ Name = 'Contrasenas por Vencer'; Data = $expUsers }
        }
        if (@($gpoRows).Count -gt 0) {
            $sheets += @{ Name = 'Backup GPOs';            Data = $gpoRows }
        }

        if ($sheets.Count -gt 0) {
            Export-MultiSheetXlsx -Path $xlsxPath -Sheets $sheets
            Write-Log "Excel generado: $xlsxPath ($($sheets.Count) hojas)"
            $attachments += $xlsxPath
        }
    }
    catch {
        Write-Log "Error generando Excel: $($_.Exception.Message)" 'WARN'
        $xlsxPath = $null
    }
}
else {
    Write-Log "Funcion Export-MultiSheetXlsx no disponible. Verificar Export-MultiSheetXlsx.ps1" 'WARN'
}
#endregion

#region -- FUNCIONES HTML ------------------------------------------------------
function Build-Table {
    param($data, $cols, [int]$limit = 50)
    $arr = @($data)
    if ($arr.Count -eq 0) { return '<p style="color:#64748b;font-style:italic;font-size:13px">Sin registros.</p>' }
    $hdr  = ($cols | ForEach-Object { "<th>$_</th>" }) -join ''
    $rows = ($arr | Select-Object -First $limit | ForEach-Object {
        $r = $_
        $cells = ($cols | ForEach-Object {
            $v = $r.$_
            $style = ''
            if ($_ -eq 'Urgencia') {
                if ($v -eq 'CRITICO') { $style = ' style="color:#dc2626;font-weight:700"' }
                elseif ($v -eq 'Alto') { $style = ' style="color:#f59e0b;font-weight:600"' }
            }
            if ($_ -eq 'Estado') {
                if ($v -eq 'OK')    { $v = '<span class="badge-ok">OK</span>' }
                if ($v -eq 'ERROR') { $v = '<span class="badge-err">ERROR</span>' }
            }
            "<td$style>$v</td>"
        }) -join ''
        "<tr>$cells</tr>"
    }) -join ''
    $more = if ($arr.Count -gt $limit) { "<p style='font-size:12px;color:#94a3b8'>Mostrando $limit de $($arr.Count). Ver adjunto para lista completa.</p>" } else { '' }
    return "<table><thead><tr>$hdr</tr></thead><tbody>$rows</tbody></table>$more"
}

function Build-PieSVG {
    param(
        [int]$ValueA,
        [int]$ValueB,
        [string]$ColorA  = '#ef4444',
        [string]$ColorB  = '#22c55e',
        [string]$LabelA  = 'Inactivos',
        [string]$LabelB  = 'Activos',
        [string]$Title   = ''
    )
    $total = $ValueA + $ValueB
    if ($total -eq 0) { return '<p style="color:#94a3b8;font-size:12px;text-align:center">Sin datos</p>' }
    $pctA  = [math]::Round($ValueA * 100.0 / $total, 1)
    $pctB  = [math]::Round(100.0 - $pctA, 1)
    $r     = 38
    $cx    = 50; $cy = 50
    $circ  = [math]::Round(2 * [math]::PI * $r, 2)
    $dashA = [math]::Round($ValueA * $circ / $total, 2)
    $dashB = [math]::Round($circ - $dashA, 2)
    $titleHtml = if ($Title) { "<div style='font-size:12px;font-weight:700;color:#334155;margin-bottom:8px;text-transform:uppercase;letter-spacing:.5px'>$Title</div>" } else { '' }
    return @"
<div style="display:inline-block;text-align:center;padding:16px 20px;background:#f8fafc;border-radius:10px;margin:0 8px;min-width:180px;vertical-align:top">
  $titleHtml
  <svg width="100" height="100" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
    <circle cx="$cx" cy="$cy" r="$r" fill="none" stroke="$ColorB" stroke-width="20"/>
    <circle cx="$cx" cy="$cy" r="$r" fill="none" stroke="$ColorA" stroke-width="20"
      stroke-dasharray="$dashA $dashB" transform="rotate(-90 $cx $cy)"/>
    <text x="$cx" y="$([int]$cy - 3)" text-anchor="middle" font-size="15" font-weight="800" font-family="Segoe UI,Arial,sans-serif" fill="#1e293b">$pctA%</text>
    <text x="$cx" y="$([int]$cy + 13)" text-anchor="middle" font-size="9" font-family="Segoe UI,Arial,sans-serif" fill="#64748b">$LabelA</text>
  </svg>
  <div style="font-size:11px;color:#475569;margin-top:8px;line-height:1.8">
    <span style="display:inline-block;width:9px;height:9px;background:$ColorA;border-radius:2px;margin-right:4px"></span><b>$LabelA</b>: $ValueA ($pctA%)<br>
    <span style="display:inline-block;width:9px;height:9px;background:$ColorB;border-radius:2px;margin-right:4px"></span><b>$LabelB</b>: $ValueB ($pctB%)
  </div>
</div>
"@
}
#endregion

#region -- CONSTRUIR EMAIL HTML ------------------------------------------------
$usersCount = @($inactiveUsers).Count
$usersActive = $totalEnabledUsers - $usersCount
$pcInact  = @($allPCs | Where-Object Inactivo).Count
$pcTotal  = @($allPCs).Count
$pcActive = $pcTotal - $pcInact
$expCount = @($expUsers).Count

$tblU   = Build-Table $inactiveUsers @('Usuario','Nombre','Departamento','UltimoLogin','DiasInactivo','UltimoCambioPass')
$tblG   = Build-Table $gpoRows       @('GPO','Estado','Detalle')
$tblP   = Build-Table ($allPCs | Where-Object Inactivo) @('Nombre','IP','SistemaOperativo','UltimaConexion','DiasInactivo')
$tblExp = Build-Table $expUsers      @('Usuario','Nombre','Departamento','FechaVence','DiasRestantes','Urgencia')

$pieUsers = Build-PieSVG -ValueA $usersCount -ValueB $usersActive -ColorA '#ef4444' -ColorB '#22c55e' -LabelA 'Inactivos' -LabelB 'Activos' -Title 'Usuarios'
$piePCs   = Build-PieSVG -ValueA $pcInact    -ValueB $pcActive    -ColorA '#f97316' -ColorB '#3b82f6' -LabelA 'Inactivos' -LabelB 'Activos' -Title 'Equipos'

try { $domainName = (Get-ADDomain).DNSRoot } catch { $domainName = 'N/A' }

$xlsxNote = if ($xlsxPath) { '<span style="background:#dcfce7;color:#15803d;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600">&#128196; Excel adjunto</span>' } else { '' }

$html = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<style>
  body  { font-family: Segoe UI, Arial, sans-serif; background: #f1f5f9; margin: 0; padding: 20px; color: #1e293b; }
  .wrap { max-width: 960px; margin: auto; background: #fff; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 16px rgba(0,0,0,.1); }
  .hdr  { background: linear-gradient(135deg,#0054a6,#0078d4); color: #fff; padding: 28px 36px; }
  .hdr h1 { margin: 0; font-size: 22px; letter-spacing: -.3px; }
  .hdr p  { margin: 6px 0 0; font-size: 13px; opacity: .85; }
  .body { padding: 32px; }
  .cards { display: flex; gap: 14px; margin-bottom: 28px; flex-wrap: wrap; }
  .card  { flex: 1; min-width: 140px; border-radius: 8px; padding: 16px 18px; }
  .card b { display: block; font-size: 28px; font-weight: 800; line-height: 1.1; }
  .card p { margin: 6px 0 0; font-size: 12px; }
  .cr { background: #fee2e2; color: #dc2626; }
  .cb { background: #dbeafe; color: #1d4ed8; }
  .cy { background: #fef3c7; color: #b45309; }
  .cg { background: #dcfce7; color: #15803d; }
  .co { background: #ffedd5; color: #c2410c; }
  .section-charts { background: #f8fafc; border-radius: 10px; padding: 20px 16px; margin-bottom: 28px; text-align: center; }
  h2  { color: #0054a6; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; font-size: 15px; margin-top: 32px; margin-bottom: 12px; }
  table { border-collapse: collapse; width: 100%; font-size: 13px; margin: 8px 0 6px; }
  th    { background: #0054a6; color: #fff; padding: 8px 12px; text-align: left; font-weight: 600; }
  td    { padding: 6px 12px; border-bottom: 1px solid #e2e8f0; }
  tr:nth-child(even) td { background: #f8fafc; }
  .ftr  { background: #f1f5f9; padding: 16px 36px; font-size: 12px; color: #94a3b8; text-align: center; border-top: 1px solid #e2e8f0; }
  .badge-ok  { background: #dcfce7; color: #15803d; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }
  .badge-err { background: #fee2e2; color: #dc2626; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }
</style>
</head>
<body>
<div class="wrap">
  <div class="hdr">
    <h1>AD Toolkit &#8212; Reporte Diario</h1>
    <p>$(Get-Date -Format 'dddd, dd MMMM yyyy HH:mm') &nbsp;&bull;&nbsp; Dominio: $domainName &nbsp;&bull;&nbsp; $xlsxNote</p>
  </div>
  <div class="body">

    <!-- CARDS RESUMEN -->
    <div class="cards">
      <div class="card cr"><b>$usersCount</b><p>Usuarios inactivos (&gt;$($cfg.DaysInactiveUsers) dias)</p></div>
      <div class="card cb"><b>$gpoOK</b><p>GPOs respaldadas ($gpoErr errores)</p></div>
      <div class="card cy"><b>$pcInact</b><p>Equipos inactivos (&gt;$($cfg.DaysInactivePC) dias)</p></div>
      <div class="card cg"><b>$pcTotal</b><p>Total equipos en dominio</p></div>
      <div class="card co"><b>$expCount</b><p>Contrasenas vencen en $daysPassExp dias</p></div>
    </div>

    <!-- GRAFICOS DE TORTA -->
    <div class="section-charts">
      <div style="font-size:13px;font-weight:700;color:#475569;margin-bottom:14px;text-align:left">Distribucion Activos / Inactivos</div>
      $pieUsers
      $piePCs
    </div>

    <!-- CONTRASENAS POR VENCER -->
    <h2>&#128274; Contrasenas por Vencer en los proximos $daysPassExp dias ($expCount usuarios)</h2>
    $tblExp

    <!-- USUARIOS INACTIVOS -->
    <h2>&#128100; Usuarios Inactivos &#8212; $usersCount de $totalEnabledUsers habilitados</h2>
    $tblU

    <!-- BACKUP GPOs -->
    <h2>&#128190; Backup de GPOs &#8212; $gpoOK OK &nbsp;/&nbsp; $gpoErr errores</h2>
    $tblG

    <!-- EQUIPOS INACTIVOS -->
    <h2>&#128187; Equipos Inactivos &#8212; $pcInact de $pcTotal</h2>
    $tblP

  </div>
  <div class="ftr">
    Generado por AD Toolkit &bull; $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') &bull; $PSScriptRoot
  </div>
</div>
</body>
</html>
"@

$htmlFile = Join-Path $reportDir "ADReport_$dateStr.html"
$html | Out-File $htmlFile -Encoding UTF8
Write-Log "Reporte HTML guardado: $htmlFile"

# Intentar convertir a PDF (requiere wkhtmltopdf o Chrome instalado)
$reportAttachment = $htmlFile   # fallback a HTML si no hay conversor PDF
$pdfFile = Join-Path $reportDir "ADReport_$dateStr.pdf"

if ($cfg.PSObject.Properties.Name -contains 'WkhtmltopdfPath' -and $cfg.WkhtmltopdfPath) {
    $wkhtmlPath = $cfg.WkhtmltopdfPath
} else {
    $wkhtmlPath = 'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
}
$chromePaths = @(
    'C:\Program Files\Google\Chrome\Application\chrome.exe',
    'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
)
$chromePath = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (Test-Path $wkhtmlPath) {
    try {
        & $wkhtmlPath '--quiet' '--encoding' 'utf-8' '--enable-local-file-access' $htmlFile $pdfFile 2>$null
        if (Test-Path $pdfFile) {
            $reportAttachment = $pdfFile
            Write-Log "PDF generado con wkhtmltopdf: $pdfFile"
        }
    } catch { Write-Log "Error generando PDF con wkhtmltopdf: $($_.Exception.Message)" 'WARN' }
} elseif ($chromePath) {
    try {
        $args = "--headless --disable-gpu --print-to-pdf=`"$pdfFile`" --no-margins `"$htmlFile`""
        Start-Process -FilePath $chromePath -ArgumentList $args -Wait -WindowStyle Hidden
        if (Test-Path $pdfFile) {
            $reportAttachment = $pdfFile
            Write-Log "PDF generado con Chrome headless: $pdfFile"
        }
    } catch { Write-Log "Error generando PDF con Chrome: $($_.Exception.Message)" 'WARN' }
} else {
    Write-Log "Sin conversor PDF disponible - se adjuntara el HTML. Para PDF: instalar wkhtmltopdf desde https://wkhtmltopdf.org" 'WARN'
}

# Adjuntar: reporte (PDF o HTML) + Excel
$attachments += $reportAttachment
if ($xlsxPath -and (Test-Path $xlsxPath)) { $attachments += $xlsxPath }
#endregion

#region -- CUERPO DEL EMAIL (mensaje simple, el detalle va en adjuntos) --------
# IMPORTANTE: mantener este archivo en ASCII puro (sin tildes ni guiones largos).
# PowerShell lo lee como ANSI/Windows-1252 si no hay BOM, y un caracter no-ASCII
# se interpreta como comilla tipografica que rompe el parseo del resto del script.
$gpoTotal = $gpoOK + $gpoErr
$emailBody = @"
<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f1f5f9">
<div style="max-width:580px;margin:30px auto;background:#fff;border-radius:10px;overflow:hidden">
  <div style="background:linear-gradient(135deg,#0054a6,#0078d4);padding:28px 36px">
    <h1 style="margin:0;font-size:20px;color:#fff;font-weight:700">AD Toolkit - Reporte Diario</h1>
    <p style="margin:6px 0 0;font-size:13px;color:rgba(255,255,255,.85)">$(Get-Date -Format 'dddd, dd MMMM yyyy') &bull; Dominio: $domainName</p>
  </div>
  <div style="padding:28px 36px">
    <p style="margin:0 0 20px;font-size:14px;color:#475569">Se adjunta el reporte diario sobre el estado del Active Directory.</p>
    <table style="width:100%;border-collapse:collapse;font-size:14px">
      <tr>
        <td style="padding:10px 0;border-bottom:1px solid #e2e8f0;color:#64748b">Usuarios inactivos (&gt;$($cfg.DaysInactiveUsers) dias)</td>
        <td style="padding:10px 0;border-bottom:1px solid #e2e8f0;text-align:right;font-weight:700;color:#dc2626;font-size:18px">$usersCount</td>
      </tr>
      <tr>
        <td style="padding:10px 0;border-bottom:1px solid #e2e8f0;color:#64748b">GPOs respaldadas</td>
        <td style="padding:10px 0;border-bottom:1px solid #e2e8f0;text-align:right;font-weight:700;color:#1d4ed8;font-size:18px">$gpoOK / $gpoTotal</td>
      </tr>
      <tr>
        <td style="padding:10px 0;border-bottom:1px solid #e2e8f0;color:#64748b">Equipos inactivos (&gt;$($cfg.DaysInactivePC) dias)</td>
        <td style="padding:10px 0;border-bottom:1px solid #e2e8f0;text-align:right;font-weight:700;color:#b45309;font-size:18px">$pcInact / $pcTotal</td>
      </tr>
      <tr>
        <td style="padding:10px 0;color:#64748b">Contrasenas vencen en $daysPassExp dias</td>
        <td style="padding:10px 0;text-align:right;font-weight:700;color:#c2410c;font-size:18px">$expCount</td>
      </tr>
    </table>
    <p style="margin:24px 0 0;font-size:11px;color:#94a3b8;border-top:1px solid #f1f5f9;padding-top:16px">Generado por AD Toolkit &bull; $(Get-Date -Format 'yyyy-MM-dd HH:mm')</p>
  </div>
</div>
</body></html>
"@
#endregion

#region -- ENVIAR EMAIL --------------------------------------------------------
if ($NoMail) {
    Write-Log "Modo -NoMail: correo omitido."
    Flush-Log; exit 0
}

if (-not $cfg.GraphTenantId -or -not $cfg.GraphClientSecretEncrypted -or -not $cfg.EmailTo) {
    Write-Log "Graph API no configurado en config.json. Configure la GUI y use 'Guardar'." 'WARN'
    Flush-Log; exit 0
}

function Send-GraphMail {
    param($tenantId, $clientId, $clientSecret, $from, $to, $subject, $htmlBody, $attachPaths = @())

    # Obtener token
    $tok = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body @{
            grant_type    = 'client_credentials'
            scope         = 'https://graph.microsoft.com/.default'
            client_id     = $clientId
            client_secret = $clientSecret
        }

    $toArr = @(($to -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) |
               ForEach-Object { @{ emailAddress = @{ address = $_ } } })

    $msg = @{
        subject      = $subject
        body         = @{ contentType = 'HTML'; content = $htmlBody }
        toRecipients = $toArr
    }

    if ($attachPaths -and @($attachPaths).Count -gt 0) {
        $atts = @()
        foreach ($ap in $attachPaths) {
            if (Test-Path $ap) {
                $atts += @{
                    '@odata.type' = '#microsoft.graph.fileAttachment'
                    name          = [System.IO.Path]::GetFileName($ap)
                    contentBytes  = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($ap))
                }
            }
        }
        if (@($atts).Count -gt 0) { $msg['attachments'] = $atts }
    }

    # FIX ENCODING: ConvertTo-Json en PS5.1 escapa Unicode como \uXXXX.
    # Convertimos de vuelta a caracteres UTF-8 reales antes de enviar.
    $jsonRaw = @{ message = $msg; saveToSentItems = $false } | ConvertTo-Json -Depth 10
    $jsonFixed = [regex]::Replace($jsonRaw, '\\u([0-9a-fA-F]{4})', {
        param($m) [char][int]('0x' + $m.Groups[1].Value)
    })
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($jsonFixed)

    Invoke-RestMethod -Method Post `
        -Uri "https://graph.microsoft.com/v1.0/users/$from/sendMail" `
        -Headers @{ Authorization = "Bearer $($tok.access_token)" } `
        -ContentType 'application/json; charset=utf-8' `
        -Body $bodyBytes
}

Write-Log "Enviando reporte a: $($cfg.EmailTo)"
try {
    $subject  = "AD Toolkit $(Get-Date -Format 'dd/MM/yyyy') -- $usersCount inactivos | $expCount pass vencen | $gpoOK GPOs | $pcTotal equipos"
    $cred     = New-Object Management.Automation.PSCredential('x', ($cfg.GraphClientSecretEncrypted | ConvertTo-SecureString))
    $secret   = $cred.GetNetworkCredential().Password

    Send-GraphMail -tenantId $cfg.GraphTenantId -clientId $cfg.GraphClientId `
        -clientSecret $secret -from $cfg.EmailFrom -to $cfg.EmailTo `
        -subject $subject -htmlBody $emailBody -attachPaths $attachments

    Write-Log "Correo enviado exitosamente."
}
catch {
    Write-Log "Error enviando correo: $($_.Exception.Message)" 'ERROR'
}
#endregion

Flush-Log
Write-Log "=== Proceso completado ==="
