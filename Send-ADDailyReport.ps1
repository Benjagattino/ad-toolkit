<#
.SYNOPSIS
    Ejecuta los 3 reportes de AD y envía un email HTML consolidado.
    Diseñado para correr como tarea programada (headless).
.PARAMETER ConfigPath
    Ruta al config.json. Default: mismo directorio que el script.
.PARAMETER LogPath
    Archivo de log. Default: .\Reports\daily_log.txt
.PARAMETER NoMail
    Si se especifica, genera los reportes pero no envía email.
.EXAMPLE
    .\Send-ADDailyReport.ps1
.EXAMPLE
    .\Send-ADDailyReport.ps1 -NoMail   # solo genera archivos
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
$logBuf = New-Object 'System.Collections.Generic.List[string]'

function Write-Log {
    param([string]$Msg, [string]$Level='INFO')
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Msg"
    $logBuf.Add($entry)
    Write-Host $entry
}

function Flush-Log {
    $dir = Split-Path $LogPath
    if ($dir) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $logBuf | Out-File $LogPath -Append -Encoding UTF8
}

#region ── CONFIG ────────────────────────────────────────────────────────────
if (-not (Test-Path $ConfigPath)) {
    Write-Log "config.json no encontrado en: $ConfigPath. Configure la GUI primero." 'ERROR'
    Flush-Log; exit 1
}
$cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
Write-Log "Configuración cargada desde: $ConfigPath"

$reportDir = Join-Path $PSScriptRoot 'Reports'
New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
$dateStr = Get-Date -Format 'yyyyMMdd'
#endregion

#region ── USUARIOS INACTIVOS ────────────────────────────────────────────────
Write-Log "--- Módulo 1: Usuarios Inactivos (umbral: $($cfg.DaysInactiveUsers) días) ---"
$cutU = (Get-Date).AddDays(-[int]$cfg.DaysInactiveUsers)
$inactiveUsers = @()
try {
    $p = @{ Filter='*'; Properties='LastLogonDate','EmailAddress','Department','Enabled','PasswordLastSet' }
    if ($cfg.SearchBase) { $p['SearchBase'] = $cfg.SearchBase }
    $inactiveUsers = Get-ADUser @p | Where-Object {
        $_.Enabled -and ($null -eq $_.LastLogonDate -or $_.LastLogonDate -lt $cutU)
    } | ForEach-Object {
        $d = if ($_.LastLogonDate) { [math]::Round(((Get-Date)-$_.LastLogonDate).TotalDays) } else { 9999 }
        [PSCustomObject]@{
            Usuario     = $_.SamAccountName
            Nombre      = $_.Name
            Email       = $_.EmailAddress
            Departamento= $_.Department
            UltimoLogin = if ($_.LastLogonDate) { $_.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Nunca' }
            DiasInactivo= $d
            UltimoCambioPass = if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString('yyyy-MM-dd') } else { 'Nunca' }
        }
    } | Sort-Object DiasInactivo -Descending
    Write-Log "Usuarios inactivos encontrados: $(@($inactiveUsers).Count)"
} catch {
    Write-Log "Error en usuarios: $($_.Exception.Message)" 'ERROR'
}
#endregion

#region ── BACKUP DE GPOs ────────────────────────────────────────────────────
Write-Log "--- Módulo 2: Backup de GPOs ---"
$gpoOK=0; $gpoErr=0; $gpoRows=@(); $gpoDestino=''
try {
    $bkpRoot = $cfg.BackupRoot
    $ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
    $dest    = Join-Path $bkpRoot $ts
    $gpoDestino = $dest
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    $domain = (Get-ADDomain).DNSRoot
    Get-GPO -All -Domain $domain | ForEach-Object {
        try {
            Backup-GPO -Guid $_.Id -Path $dest -Domain $domain | Out-Null
            $gpoRows += [PSCustomObject]@{ GPO=$_.DisplayName; Estado='OK'; Detalle='' }
            $gpoOK++
        } catch {
            $gpoRows += [PSCustomObject]@{ GPO=$_.DisplayName; Estado='ERROR'; Detalle=$_.Exception.Message }
            $gpoErr++
            Write-Log "Error GPO '$($_.DisplayName)': $($_.Exception.Message)" 'WARN'
        }
    }
    # Rotación
    Get-ChildItem $bkpRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{8}_' } | Sort-Object Name -Descending |
        Select-Object -Skip ([int]$cfg.MaxBackups) |
        ForEach-Object { Remove-Item $_.FullName -Recurse -Force; Write-Log "Backup antiguo eliminado: $($_.Name)" }
    Write-Log "GPO Backup: OK=$gpoOK, Errores=$gpoErr"
} catch {
    Write-Log "Error en backup GPO: $($_.Exception.Message)" 'ERROR'
}
#endregion

#region ── INVENTARIO DE EQUIPOS ────────────────────────────────────────────
Write-Log "--- Módulo 3: Inventario de Equipos (umbral: $($cfg.DaysInactivePC) días) ---"
$cutPC = (Get-Date).AddDays(-[int]$cfg.DaysInactivePC)
$allPCs = @()
try {
    $p = @{
        Filter     = if ($cfg.OSFilter) { "OperatingSystem -like '$($cfg.OSFilter)'" } else { '*' }
        Properties = 'OperatingSystem','LastLogonDate','IPv4Address','Enabled'
    }
    if ($cfg.SearchBase) { $p['SearchBase'] = $cfg.SearchBase }
    $allPCs = Get-ADComputer @p | ForEach-Object {
        $d = if ($_.LastLogonDate) { [math]::Round(((Get-Date)-$_.LastLogonDate).TotalDays) } else { 9999 }
        [PSCustomObject]@{
            Nombre          = $_.Name
            IP              = $_.IPv4Address
            SistemaOperativo= $_.OperatingSystem
            Habilitado      = $_.Enabled
            UltimaConexion  = if ($_.LastLogonDate) { $_.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Nunca' }
            DiasInactivo    = $d
            Inactivo        = ($null -eq $_.LastLogonDate -or $_.LastLogonDate -lt $cutPC)
        }
    } | Sort-Object DiasInactivo -Descending
    Write-Log "Equipos: total=$(@($allPCs).Count), inactivos=$(@($allPCs | Where-Object Inactivo).Count)"
} catch {
    Write-Log "Error en inventario: $($_.Exception.Message)" 'ERROR'
}
#endregion

#region ── EXPORTAR CSVs ────────────────────────────────────────────────────
$csvUsers = Join-Path $reportDir "Users_$dateStr.csv"
$csvPCs   = Join-Path $reportDir "PCs_$dateStr.csv"
$csvGPO   = Join-Path $reportDir "GPO_$dateStr.csv"

$attachments = @()
if ($inactiveUsers) { $inactiveUsers | Export-Csv $csvUsers -NoTypeInformation -Encoding UTF8; $attachments += $csvUsers }
if ($allPCs)        { $allPCs        | Export-Csv $csvPCs   -NoTypeInformation -Encoding UTF8; $attachments += $csvPCs }
if ($gpoRows)       { $gpoRows       | Export-Csv $csvGPO   -NoTypeInformation -Encoding UTF8; $attachments += $csvGPO }
Write-Log "CSVs exportados en: $reportDir"
#endregion

#region ── CONSTRUIR EMAIL HTML ─────────────────────────────────────────────
$usersCount = @($inactiveUsers).Count
$pcInact    = @($allPCs | Where-Object Inactivo).Count
$pcTotal    = @($allPCs).Count

function Build-Table($data, $cols, $limit=50) {
    $arr = @($data)
    if ($arr.Count -eq 0) { return '<p style="color:#64748b;font-style:italic">Sin registros.</p>' }
    $hdr = ($cols | ForEach-Object { "<th>$_</th>" }) -join ''
    $rows = ($arr | Select-Object -First $limit | ForEach-Object {
        $r = $_; $cells = ($cols | ForEach-Object { "<td>$($r.$_)</td>" }) -join ''
        "<tr>$cells</tr>"
    }) -join ''
    $more = if ($arr.Count -gt $limit) { "<p style='font-size:12px;color:#94a3b8'>Mostrando $limit de $($arr.Count). Ver CSV adjunto.</p>" } else { '' }
    return "<table><thead><tr>$hdr</tr></thead><tbody>$rows</tbody></table>$more"
}

$tblU = Build-Table $inactiveUsers @('Usuario','Nombre','Departamento','UltimoLogin','DiasInactivo')
$tblG = Build-Table $gpoRows       @('GPO','Estado','Detalle')
$tblP = Build-Table ($allPCs | Where-Object Inactivo) @('Nombre','IP','SistemaOperativo','UltimaConexion','DiasInactivo')

try { $domainName = (Get-ADDomain).DNSRoot } catch { $domainName = 'N/A' }

$html = @"
<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<style>
  body  {font-family:Segoe UI,Arial,sans-serif;background:#f1f5f9;margin:0;padding:20px;color:#1e293b}
  .wrap {max-width:900px;margin:auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,.1)}
  .hdr  {background:linear-gradient(135deg,#0054a6,#0078d4);color:#fff;padding:28px 36px}
  .hdr h1{margin:0;font-size:22px;letter-spacing:-.3px}
  .hdr p {margin:6px 0 0;font-size:13px;opacity:.85}
  .body {padding:32px}
  .cards{display:flex;gap:14px;margin-bottom:32px;flex-wrap:wrap}
  .card {flex:1;min-width:150px;border-radius:8px;padding:16px 18px}
  .card b{display:block;font-size:28px;font-weight:800;line-height:1.1}
  .card p{margin:6px 0 0;font-size:12px}
  .cr {background:#fee2e2;color:#dc2626} .cb {background:#dbeafe;color:#1d4ed8}
  .cy {background:#fef3c7;color:#b45309} .cg {background:#dcfce7;color:#15803d}
  h2  {color:#0054a6;border-bottom:2px solid #e2e8f0;padding-bottom:8px;font-size:15px;margin-top:32px}
  table{border-collapse:collapse;width:100%;font-size:13px;margin:12px 0 6px}
  th  {background:#0054a6;color:#fff;padding:8px 12px;text-align:left;font-weight:600}
  td  {padding:6px 12px;border-bottom:1px solid #e2e8f0}
  tr:nth-child(even) td{background:#f8fafc}
  .ftr{background:#f1f5f9;padding:16px 36px;font-size:12px;color:#94a3b8;text-align:center;border-top:1px solid #e2e8f0}
  .badge-ok  {background:#dcfce7;color:#15803d;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600}
  .badge-err {background:#fee2e2;color:#dc2626;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600}
</style></head><body>
<div class="wrap">
  <div class="hdr">
    <h1>AD Toolkit — Reporte Diario</h1>
    <p>$(Get-Date -Format 'dddd, dd MMMM yyyy HH:mm') &nbsp;&bull;&nbsp; Dominio: $domainName</p>
  </div>
  <div class="body">
    <div class="cards">
      <div class="card cr"><b>$usersCount</b><p>Usuarios inactivos (&gt;$($cfg.DaysInactiveUsers) días)</p></div>
      <div class="card cb"><b>$gpoOK</b><p>GPOs respaldadas ($gpoErr errores)</p></div>
      <div class="card cy"><b>$pcInact</b><p>Equipos inactivos (&gt;$($cfg.DaysInactivePC) días)</p></div>
      <div class="card cg"><b>$pcTotal</b><p>Total equipos en dominio</p></div>
    </div>

    <h2>&#128100; Usuarios Inactivos ($usersCount)</h2>
    $tblU

    <h2>&#128190; Backup de GPOs &mdash; $gpoOK OK &nbsp;/&nbsp; $gpoErr errores</h2>
    $tblG

    <h2>&#128187; Equipos Inactivos ($pcInact de $pcTotal)</h2>
    $tblP
  </div>
  <div class="ftr">Generado por AD Toolkit &bull; $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') &bull; $($PSScriptRoot)</div>
</div></body></html>
"@

$htmlFile = Join-Path $reportDir "ADReport_$dateStr.html"
$html | Out-File $htmlFile -Encoding UTF8
Write-Log "Reporte HTML guardado: $htmlFile"
if ($cfg.SendAttachments) { $attachments += $htmlFile }
#endregion

#region ── ENVIAR EMAIL ─────────────────────────────────────────────────────
if ($NoMail) {
    Write-Log "Modo -NoMail: correo omitido."
    Flush-Log; exit 0
}

if (-not $cfg.GraphTenantId -or -not $cfg.GraphClientSecretEncrypted -or -not $cfg.EmailTo) {
    Write-Log "Graph API no configurado en config.json. Configure la GUI y use 'Guardar'." 'WARN'
    Flush-Log; exit 0
}

function Send-GraphMail {
    param($tenantId, $clientId, $clientSecret, $from, $to, $subject, $htmlBody, $attachPaths=@())
    $tok = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body @{ grant_type='client_credentials'; scope='https://graph.microsoft.com/.default'
                 client_id=$clientId; client_secret=$clientSecret }
    $toArr = @(($to -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) |
               ForEach-Object { @{ emailAddress=@{ address=$_ } } })
    $msg = @{ subject=$subject; body=@{ contentType='HTML'; content=$htmlBody }; toRecipients=$toArr }
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
    $json = @{ message=$msg; saveToSentItems=$false } | ConvertTo-Json -Depth 10
    Invoke-RestMethod -Method Post `
        -Uri "https://graph.microsoft.com/v1.0/users/$from/sendMail" `
        -Headers @{ Authorization="Bearer $($tok.access_token)" } `
        -ContentType 'application/json' `
        -Body $json
}

Write-Log "Enviando reporte a: $($cfg.EmailTo)"
try {
    $subject = "AD Toolkit $(Get-Date -Format 'dd/MM/yyyy') -- $usersCount usuarios inactivos | $gpoOK GPOs | $pcTotal equipos"
    $cred    = New-Object Management.Automation.PSCredential('x', ($cfg.GraphClientSecretEncrypted | ConvertTo-SecureString))
    $secret  = $cred.GetNetworkCredential().Password
    $attPaths = if ($cfg.SendAttachments) { $attachments } else { @() }
    Send-GraphMail -tenantId $cfg.GraphTenantId -clientId $cfg.GraphClientId `
        -clientSecret $secret -from $cfg.EmailFrom -to $cfg.EmailTo `
        -subject $subject -htmlBody $html -attachPaths $attPaths
    Write-Log "Correo enviado exitosamente."
} catch {
    Write-Log "Error enviando correo: $($_.Exception.Message)" 'ERROR'
}
#endregion

Flush-Log
Write-Log "=== Proceso completado ==="
