<#
.SYNOPSIS
    AD Toolkit  - Interfaz gráfica centralizada.
    Ejecuta, visualiza y exporta los 3 reportes de Active Directory.
.NOTES
    Requiere módulos RSAT: ActiveDirectory, GroupPolicy
    Sin dependencias externas  - solo WinForms nativo de Windows.
#>
trap {
    # Captura cualquier error no manejado y lo muestra antes de cerrar
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    [System.Windows.Forms.MessageBox]::Show(
        "Error al iniciar AD Toolkit:`n`n$($_.Exception.Message)`n`nEn: $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())",
        'AD Toolkit - Error de inicio', 'OK', 'Error'
    ) | Out-Null
    break
}

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
$script:configPath = Join-Path $script:scriptRoot 'config.json'
$scriptRoot = $script:scriptRoot   # alias para compatibilidad con el resto del script
$configPath = $script:configPath

#region -- CONFIG --------------------------------------------------------------
function Load-ADConfig {
    $d = [ordered]@{
        DaysInactiveUsers          = 90;  DaysInactivePC = 90
        SearchBase                 = '';  BackupRoot = (Join-Path $script:scriptRoot 'GPO-Backups')
        MaxBackups                 = 10;  OSFilter = ''
        GraphTenantId              = '';  GraphClientId = ''
        GraphClientSecretEncrypted = '';  EmailFrom = ''
        EmailTo                    = '';  SendAttachments = $true
    }
    if (Test-Path $script:configPath) {
        $loaded = Get-Content $script:configPath -Raw | ConvertFrom-Json
        foreach ($k in @($d.Keys)) { if ($null -ne $loaded.$k) { $d[$k] = $loaded.$k } }
    }
    return [PSCustomObject]$d
}
function Save-ADConfig($c) { $c | ConvertTo-Json -Depth 3 | Out-File $script:configPath -Encoding UTF8 }
$cfg = Load-ADConfig
#endregion

#region -- PALETTE & FONTS -----------------------------------------------------
$C = @{
    Sidebar       = [Drawing.Color]::FromArgb(15,  23,  42)
    SidebarHover  = [Drawing.Color]::FromArgb(30,  41,  59)
    SidebarActive = [Drawing.Color]::FromArgb(0,  100, 210)
    Header        = [Drawing.Color]::FromArgb(0,   84, 166)
    BG            = [Drawing.Color]::FromArgb(244, 247, 251)
    White         = [Drawing.Color]::White
    Green         = [Drawing.Color]::FromArgb(16,  163,  74)
    Red           = [Drawing.Color]::FromArgb(220,  38,  38)
    Yellow        = [Drawing.Color]::FromArgb(202, 138,   4)
    Blue          = [Drawing.Color]::FromArgb(0,  120, 212)
    TextDark      = [Drawing.Color]::FromArgb(15,  23,  42)
    TextMid       = [Drawing.Color]::FromArgb(100, 116, 139)
    Border        = [Drawing.Color]::FromArgb(220, 226, 236)
}
$F = @{
    Sm      = New-Object Drawing.Font('Segoe UI',  8)
    Norm    = New-Object Drawing.Font('Segoe UI',  9)
    Bold    = New-Object Drawing.Font('Segoe UI',  9, [Drawing.FontStyle]::Bold)
    Title   = New-Object Drawing.Font('Segoe UI', 13, [Drawing.FontStyle]::Bold)
    Nav     = New-Object Drawing.Font('Segoe UI', 10)
    Hdr     = New-Object Drawing.Font('Segoe UI', 17, [Drawing.FontStyle]::Bold)
    CardNum = New-Object Drawing.Font('Segoe UI', 26, [Drawing.FontStyle]::Bold)
    Mono    = New-Object Drawing.Font('Consolas',  9)
}
#endregion

#region -- HELPERS -------------------------------------------------------------
function New-Lbl($text, $x, $y, $w=200, $h=20, $font=$F.Norm, $fg=$C.TextDark) {
    [Windows.Forms.Label]@{ Text=$text; Left=$x; Top=$y; Width=$w; Height=$h
        Font=$font; ForeColor=$fg; BackColor=[Drawing.Color]::Transparent }
}
function New-Btn($text, $x, $y, $w=140, $h=30, $bg=$C.Blue, $fg=$C.White) {
    $b = New-Object Windows.Forms.Button
    $b.Text=$text; $b.Left=$x; $b.Top=$y; $b.Width=$w; $b.Height=$h
    $b.Font=$F.Bold; $b.BackColor=$bg; $b.ForeColor=$fg; $b.FlatStyle='Flat'
    $b.Cursor = [Windows.Forms.Cursors]::Hand
    $b.FlatAppearance.BorderSize = 0; return $b
}
function New-Txt($x, $y, $w=200, $val='') {
    [Windows.Forms.TextBox]@{ Left=$x; Top=$y; Width=$w; Height=26
        Font=$F.Norm; Text=$val; BorderStyle='FixedSingle' }
}
function New-Num($x, $y, $min=1, $max=730, $val=90) {
    [Windows.Forms.NumericUpDown]@{ Left=$x; Top=$y; Width=80; Height=26
        Minimum=$min; Maximum=$max; Value=$val; Font=$F.Norm }
}
function New-Chk($text, $x, $y, $chk=$false) {
    [Windows.Forms.CheckBox]@{ Text=$text; Left=$x; Top=$y; AutoSize=$true
        Font=$F.Norm; Checked=$chk; BackColor=[Drawing.Color]::Transparent }
}
function New-Grid($x, $y, $w, $h) {
    $g = New-Object Windows.Forms.DataGridView
    $g.Location=(New-Object System.Drawing.Point -ArgumentList $x, $y); $g.Size=(New-Object System.Drawing.Size -ArgumentList $w, $h)
    $g.ReadOnly=$true; $g.AllowUserToAddRows=$false; $g.RowHeadersVisible=$false
    $g.SelectionMode='FullRowSelect'; $g.BackgroundColor=$C.White
    $g.GridColor=$C.Border; $g.BorderStyle='None'; $g.Font=$F.Sm
    $g.AutoSizeColumnsMode='DisplayedCells'; $g.EnableHeadersVisualStyles=$false
    $g.ColumnHeadersDefaultCellStyle.BackColor=$C.Header
    $g.ColumnHeadersDefaultCellStyle.ForeColor=$C.White
    $g.ColumnHeadersDefaultCellStyle.Font=$F.Bold; $g.ColumnHeadersHeight=30
    $g.AlternatingRowsDefaultCellStyle.BackColor=[Drawing.Color]::FromArgb(248,250,253)
    return $g
}
function Set-GridData($grid, $data) {
    $grid.DataSource = $null
    $arr = @($data)
    if ($arr.Count -eq 0) { return }
    $dt = New-Object System.Data.DataTable
    $arr[0].PSObject.Properties.Name | ForEach-Object { [void]$dt.Columns.Add($_) }
    foreach ($item in $arr) {
        $dr = $dt.NewRow()
        $item.PSObject.Properties | ForEach-Object { $dr[$_.Name] = "$($_.Value)" }
        $dt.Rows.Add($dr)
    }
    $grid.DataSource = $dt
}
function Export-GridCsv($grid, $name) {
    $dlg = [Windows.Forms.SaveFileDialog]@{ Filter='CSV (*.csv)|*.csv'; FileName=$name; Title='Guardar reporte' }
    if ($dlg.ShowDialog() -ne 'OK') { return }
    $csvSb = (New-Object System.Text.StringBuilder)
    $csvSb.AppendLine(($grid.Columns | ForEach-Object { $_.HeaderText }) -join ',') | Out-Null
    foreach ($row in $grid.Rows) {
        $csvSb.AppendLine(($row.Cells | ForEach-Object { '"' + ($_.Value -replace '"','""') + '"' }) -join ',') | Out-Null
    }
    $csvSb.ToString() | Out-File $dlg.FileName -Encoding UTF8
    [Windows.Forms.MessageBox]::Show("Exportado:`n$($dlg.FileName)", 'Exportar', 'OK', 'Information') | Out-Null
}
function Send-GraphMail {
    param($tenantId, $clientId, $clientSecret, $from, $to, $subject, $htmlBody, $attachPaths=@())
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
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

function New-Card($parent, $x, $title, $val, $sub, $color) {
    # Capturar colores como locales antes del closure — PS4 no hereda scope externo en GetNewClosure
    $localAccent = $color
    $localBorder = $C.Border
    $p = New-Object Windows.Forms.Panel
    $p.Location=(New-Object System.Drawing.Point -ArgumentList $x, 46); $p.Size=(New-Object System.Drawing.Size -ArgumentList 198, 115); $p.BackColor=$C.White
    $p.Add_Paint({
        param($s,$e)
        if ($localAccent -and $localBorder) {
            $pen1 = New-Object System.Drawing.Pen -ArgumentList $localAccent, 4
            $pen2 = New-Object System.Drawing.Pen -ArgumentList $localBorder, 1
            $e.Graphics.DrawLine($pen1, 0, 0, 0, $s.Height)
            $e.Graphics.DrawRectangle($pen2, 0, 0, $s.Width-1, $s.Height-1)
            $pen1.Dispose(); $pen2.Dispose()
        }
    }.GetNewClosure())
    $lT = New-Lbl $title 14  8 170 18 $F.Sm   $C.TextMid
    $lV = New-Lbl $val   14 24 170 55 $F.CardNum $color
    $lS = New-Lbl $sub   14 80 170 16 $F.Sm   $C.TextMid
    $p.Controls.AddRange(@($lT,$lV,$lS))
    $parent.Controls.Add($p)
    return @{ Value=$lV; Sub=$lS }
}
#endregion

#region -- AD QUERIES ----------------------------------------------------------
function Get-InactiveUsersData($days, $base, $incDis) {
    $cut = (Get-Date).AddDays(-$days)
    $p = @{ Filter='*'; Properties='LastLogonDate','PasswordLastSet','EmailAddress','Department','Enabled','DistinguishedName','PasswordNeverExpires' }
    if ($base) { $p['SearchBase'] = $base }
    Get-ADUser @p | Where-Object {
        ($incDis -or $_.Enabled) -and ($null -eq $_.LastLogonDate -or $_.LastLogonDate -lt $cut)
    } | ForEach-Object {
        $dias = if ($_.LastLogonDate) { [math]::Round(((Get-Date)-$_.LastLogonDate).TotalDays) } else { 9999 }
        [PSCustomObject]@{
            Usuario          = $_.SamAccountName
            Nombre           = $_.Name
            Email            = $_.EmailAddress
            Departamento     = $_.Department
            Habilitado       = $_.Enabled
            UltimoLogin      = if ($_.LastLogonDate) { $_.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Nunca' }
            DiasInactivo     = $dias
            UltimoCambioPass = if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString('yyyy-MM-dd') } else { 'Nunca' }
            PassNuncaExpira  = $_.PasswordNeverExpires
            OU               = ($_.DistinguishedName -replace '^CN=[^,]+,','')
        }
    } | Sort-Object DiasInactivo -Descending
}

function Invoke-GPOBackup($root, $maxBkp, $rtb) {
    $log = (New-Object 'System.Collections.Generic.List[string]')
    try {
        $ts   = Get-Date -Format 'yyyyMMdd_HHmmss'
        $dest = Join-Path $root $ts
        New-Item -ItemType Directory -Path $dest -Force | Out-Null
        $domain = (Get-ADDomain).DNSRoot
        $log.Add("[INFO] Dominio: $domain | Carpeta: $dest")
        $gpos  = Get-GPO -All -Domain $domain
        $log.Add("[INFO] Exportando $($gpos.Count) GPOs...")
        $ok = 0; $err = 0
        $n   = 0
        foreach ($gpo in $gpos) {
            $n++
            try {
                Backup-GPO -Guid $gpo.Id -Path $dest -Domain $domain | Out-Null
                $line = "  [OK] ($n/$($gpos.Count)) $($gpo.DisplayName)"
                $log.Add($line); $ok++
            } catch {
                $line = "  [ERROR] ($n/$($gpos.Count)) $($gpo.DisplayName) - $($_.Exception.Message)"
                $log.Add($line); $err++
            }
            # Actualizar RichTextBox en tiempo real si se paso como parametro
            if ($rtb) {
                $col = if ($line -match '\[ERROR') { [Drawing.Color]::FromArgb(255,100,100) } else { [Drawing.Color]::FromArgb(100,255,150) }
                $rtb.SelectionColor = $col
                $rtb.AppendText("$line`n")
                $rtb.ScrollToCaret()
            }
            # Liberar el hilo de UI para que la ventana no se tilde
            [System.Windows.Forms.Application]::DoEvents()
        }
        Get-ChildItem $root -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^\d{8}_' } | Sort-Object Name -Descending |
            Select-Object -Skip $maxBkp | ForEach-Object {
                Remove-Item $_.FullName -Recurse -Force
                $log.Add("[INFO] Backup antiguo eliminado: $($_.Name)")
            }
        $log.Add(""); $log.Add("[RESULTADO] Exitosos: $ok  |  Errores: $err")
        $log.Add("[RESULTADO_OK=$ok]")
    } catch {
        $log.Add("[ERROR FATAL] $($_.Exception.Message)")
    }
    return $log
}

function Get-ComputersData($days, $base, $osf) {
    $cut = (Get-Date).AddDays(-$days)
    $p = @{
        Filter     = if ($osf) { "OperatingSystem -like '$osf'" } else { '*' }
        Properties = 'OperatingSystem','OperatingSystemVersion','LastLogonDate','IPv4Address','Enabled','Created','Location','DistinguishedName'
    }
    if ($base) { $p['SearchBase'] = $base }
    Get-ADComputer @p | ForEach-Object {
        $dias = if ($_.LastLogonDate) { [math]::Round(((Get-Date)-$_.LastLogonDate).TotalDays) } else { 9999 }
        $tipo = switch -Wildcard ($_.OperatingSystem) {
            '*Server*'     {'Servidor'} '*Windows 11*' {'W11'} '*Windows 10*' {'W10'}
            '*Windows 7*'  {'W7 (EOL)'} '*Windows 8*' {'W8 (EOL)'} $null {'Sin OS'} default {'Otro'}
        }
        [PSCustomObject]@{
            Nombre         = $_.Name
            IP             = $_.IPv4Address
            SistemaOperativo = $_.OperatingSystem
            Tipo           = $tipo
            Habilitado     = $_.Enabled
            UltimaConexion = if ($_.LastLogonDate) { $_.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Nunca' }
            DiasInactivo   = $dias
            CreadoEn       = $_.Created.ToString('yyyy-MM-dd')
            Ubicacion      = $_.Location
            OU             = ($_.DistinguishedName -replace '^CN=[^,]+,','')
        }
    } | Sort-Object DiasInactivo -Descending
}
#endregion

#region -- MAIN FORM -----------------------------------------------------------
$form = New-Object Windows.Forms.Form
$form.Text='AD Toolkit'; $form.Size=(New-Object System.Drawing.Size -ArgumentList 1180, 760)
$form.StartPosition='CenterScreen'; $form.BackColor=$C.BG
$form.MinimumSize=(New-Object System.Drawing.Size -ArgumentList 1000, 640)
try { $form.Icon=[Drawing.Icon]::ExtractAssociatedIcon([Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) } catch {}

# Header — agregar primero (Top)
$hdr = New-Object Windows.Forms.Panel
$hdr.Dock='Top'; $hdr.Height=58; $hdr.BackColor=$C.Header
$hdr.Controls.Add((New-Lbl 'AD Toolkit' 18 12 300 34 $F.Hdr $C.White))
$lblDom = New-Lbl '' 18 40 500 16 $F.Sm ([Drawing.Color]::FromArgb(170,200,255))
try { $lblDom.Text = "Dominio: $((Get-ADDomain -ErrorAction SilentlyContinue).DNSRoot)" } catch { $lblDom.Text = 'Dominio: (no disponible)' }
$lblClock = New-Lbl (Get-Date -Format 'HH:mm') 1090 18 70 22 $F.Bold $C.White
$lblClock.TextAlign='MiddleRight'
$tClock = New-Object Windows.Forms.Timer; $tClock.Interval=30000
$tClock.Add_Tick({ $lblClock.Text = Get-Date -Format 'HH:mm' }); $tClock.Start()
$hdr.Controls.AddRange(@($lblDom,$lblClock))
$form.Controls.Add($hdr)

# Status bar (Bottom)
$stBar = New-Object Windows.Forms.Panel
$stBar.Dock='Bottom'; $stBar.Height=26; $stBar.BackColor=[Drawing.Color]::FromArgb(230,235,244)
$lblSt = New-Lbl 'Listo.' 10 5 900 18 $F.Sm $C.TextMid
$stBar.Controls.Add($lblSt)
$form.Controls.Add($stBar)

# SplitContainer: divide sidebar (205px) y contenido (resto) — mas confiable que Dock=Left+Fill en PS4
$mainSplit = New-Object Windows.Forms.SplitContainer
$mainSplit.Dock = 'Fill'
$mainSplit.SplitterDistance = 205
$mainSplit.IsSplitterFixed  = $true
$mainSplit.Panel1MinSize    = 205
$mainSplit.Panel2MinSize    = 100
$mainSplit.BorderStyle      = 'None'
$mainSplit.SplitterWidth    = 1
$mainSplit.BackColor        = $C.Sidebar
$mainSplit.Panel1.BackColor = $C.Sidebar
$mainSplit.Panel2.BackColor = $C.BG
$mainSplit.Panel2.Padding   = (New-Object System.Windows.Forms.Padding -ArgumentList 20, 16, 20, 8)
$form.Controls.Add($mainSplit)

# Aliases — el resto del script usa $sb y $ct sin cambios
$sb = $mainSplit.Panel1
$ct = $mainSplit.Panel2

function Set-Status($msg, $col=$C.TextMid) { $lblSt.Text=$msg; $lblSt.ForeColor=$col; $form.Refresh() }
#endregion

#region -- PANELS --------------------------------------------------------------

# -- DASHBOARD
$pDash = New-Object Windows.Forms.Panel; $pDash.Dock='Fill'; $pDash.BackColor=[Drawing.Color]::Transparent
$pDash.Controls.Add((New-Lbl 'Dashboard' 0 0 400 30 $F.Title $C.TextDark))

$script:cUsers = New-Card $pDash   0 'Usuarios Inactivos' '-' "umbral: $($cfg.DaysInactiveUsers)d" $C.Red
$script:cGPO   = New-Card $pDash 210 'GPOs Respaldadas'   '-' 'ultimo backup' $C.Blue
$script:cPCsI  = New-Card $pDash 420 'Equipos Inactivos'  '-' "umbral: $($cfg.DaysInactivePC)d"   $C.Yellow
$script:cPCsT  = New-Card $pDash 630 'Total Equipos'      '-' 'en el dominio'  $C.Green

$btnAll = New-Btn 'Ejecutar Todo' 0 176 180 34 $C.Blue $C.White
$pDash.Controls.Add($btnAll)
$lblLast = New-Lbl '' 0 218 700 18 $F.Sm $C.TextMid; $pDash.Controls.Add($lblLast)

$dashLog = New-Object Windows.Forms.RichTextBox
$dashLog.Location=(New-Object System.Drawing.Point -ArgumentList 0, 245); $dashLog.Size=(New-Object System.Drawing.Size -ArgumentList 900, 360)
$dashLog.ReadOnly=$true; $dashLog.Font=$F.Mono; $dashLog.BorderStyle='None'
$dashLog.BackColor=[Drawing.Color]::FromArgb(18,22,36); $dashLog.ForeColor=[Drawing.Color]::FromArgb(160,240,160)
$pDash.Controls.Add($dashLog)
$ct.Controls.Add($pDash)

# -- USUARIOS
$pUsers = New-Object Windows.Forms.Panel; $pUsers.Dock='Fill'; $pUsers.BackColor=[Drawing.Color]::Transparent; $pUsers.Visible=$false
$pUsers.Controls.Add((New-Lbl 'Usuarios Inactivos' 0 0 400 30 $F.Title $C.TextDark))

$pUsers.Controls.Add((New-Lbl 'Días sin login:' 0 44 110 18 $F.Sm $C.TextMid))
$script:nDaysU = New-Num 0 60 1 730 $cfg.DaysInactiveUsers; $pUsers.Controls.Add($script:nDaysU)
$pUsers.Controls.Add((New-Lbl 'SearchBase (OU):' 100 44 120 18 $F.Sm $C.TextMid))
$script:tBaseU = New-Txt 100 60 310 $cfg.SearchBase; $pUsers.Controls.Add($script:tBaseU)
$script:chkDis = New-Chk 'Incluir deshabilitadas' 430 63; $pUsers.Controls.Add($script:chkDis)

$btnRunU  = New-Btn 'Ejecutar'    0 100 120 28
$btnExpU  = New-Btn 'Exportar CSV' 130 100 140 28 $C.Green $C.White
$lblCntU  = New-Lbl '' 285 106 500 18 $F.Sm $C.TextMid
$pUsers.Controls.AddRange(@($btnRunU,$btnExpU,$lblCntU))

$script:grdU = New-Grid 0 136 900 490; $pUsers.Controls.Add($script:grdU)
$ct.Controls.Add($pUsers)

# -- GPO BACKUP
$pGPO = New-Object Windows.Forms.Panel; $pGPO.Dock='Fill'; $pGPO.BackColor=[Drawing.Color]::Transparent; $pGPO.Visible=$false
$pGPO.Controls.Add((New-Lbl 'Backup de GPOs' 0 0 400 30 $F.Title $C.TextDark))

$pGPO.Controls.Add((New-Lbl 'Carpeta destino:' 0 44 120 18 $F.Sm $C.TextMid))
$script:tBkpRoot = New-Txt 0 60 400 $cfg.BackupRoot; $pGPO.Controls.Add($script:tBkpRoot)
$btnBrowse = New-Btn '...' 410 60 40 26 $C.Border $C.TextDark; $pGPO.Controls.Add($btnBrowse)

$pGPO.Controls.Add((New-Lbl 'Backups a conservar:' 465 44 150 18 $F.Sm $C.TextMid))
$script:nMaxBkp = New-Num 465 60 1 100 $cfg.MaxBackups; $pGPO.Controls.Add($script:nMaxBkp)

$btnRunGPO = New-Btn 'Ejecutar Backup' 0 100 170 28; $pGPO.Controls.Add($btnRunGPO)

$rtbGPO = New-Object Windows.Forms.RichTextBox
$rtbGPO.Location=(New-Object System.Drawing.Point -ArgumentList 0, 140); $rtbGPO.Size=(New-Object System.Drawing.Size -ArgumentList 900, 490)
$rtbGPO.ReadOnly=$true; $rtbGPO.Font=$F.Mono; $rtbGPO.BorderStyle='None'
$rtbGPO.BackColor=[Drawing.Color]::FromArgb(18,22,36); $rtbGPO.ForeColor=[Drawing.Color]::FromArgb(160,240,160)
$pGPO.Controls.Add($rtbGPO)
$ct.Controls.Add($pGPO)

# -- INVENTARIO
$pInv = New-Object Windows.Forms.Panel; $pInv.Dock='Fill'; $pInv.BackColor=[Drawing.Color]::Transparent; $pInv.Visible=$false
$pInv.Controls.Add((New-Lbl 'Inventario de Equipos' 0 0 400 30 $F.Title $C.TextDark))

$pInv.Controls.Add((New-Lbl 'Días inactivo:'  0 44 100 18 $F.Sm $C.TextMid))
$script:nDaysPC = New-Num 0 60 1 730 $cfg.DaysInactivePC; $pInv.Controls.Add($script:nDaysPC)
$pInv.Controls.Add((New-Lbl 'SearchBase:' 100 44 90 18 $F.Sm $C.TextMid))
$script:tBasePC = New-Txt 100 60 270 $cfg.SearchBase; $pInv.Controls.Add($script:tBasePC)
$pInv.Controls.Add((New-Lbl 'Filtro OS:'  385 44 75 18 $F.Sm $C.TextMid))
$script:tOSF = New-Txt 385 60 170 $cfg.OSFilter; $pInv.Controls.Add($script:tOSF)

$btnRunI  = New-Btn 'Ejecutar'     0 100 120 28
$btnExpI  = New-Btn 'Exportar CSV' 130 100 140 28 $C.Green $C.White
$lblCntI  = New-Lbl '' 285 106 500 18 $F.Sm $C.TextMid
$pInv.Controls.AddRange(@($btnRunI,$btnExpI,$lblCntI))

$script:grdI = New-Grid 0 136 900 490; $pInv.Controls.Add($script:grdI)
$ct.Controls.Add($pInv)

# -- CONFIGURACIÓN
$pCfg = New-Object Windows.Forms.Panel; $pCfg.Dock='Fill'; $pCfg.BackColor=[Drawing.Color]::Transparent; $pCfg.Visible=$false
$pCfg.Controls.Add((New-Lbl 'Configuración' 0 0 400 30 $F.Title $C.TextDark))

# Graph API
$pCfg.Controls.Add((New-Lbl 'Microsoft Graph API (correo)' 0 40 400 18 $F.Bold $C.TextDark))

$cfgF = @(
    @{L='Tenant ID';           K='GraphTenantId';     X=  0; Y= 68; W=580; V=$cfg.GraphTenantId}
    @{L='Client ID (App)';     K='GraphClientId';     X=  0; Y=108; W=580; V=$cfg.GraphClientId}
    @{L='Secreto de cliente';  K='GraphClientSecret'; X=  0; Y=148; W=580; V=''; Pass=$true}
    @{L='Remitente (From)';    K='EmailFrom';         X=  0; Y=188; W=260; V=$cfg.EmailFrom}
    @{L='Destinatario (To)';   K='EmailTo';           X=275; Y=188; W=300; V=$cfg.EmailTo}
)
$script:cfgCtrl = @{}
foreach ($f in $cfgF) {
    $pCfg.Controls.Add((New-Lbl $f.L $f.X ($f.Y-16) 200 15 $F.Sm $C.TextMid))
    $tb = New-Txt $f.X $f.Y $f.W $f.V
    if ($f.Pass) { $tb.PasswordChar='*' }
    $script:cfgCtrl[$f.K] = $tb
    $pCfg.Controls.Add($tb)
}
$script:chkAtt = New-Chk 'Adjuntar CSVs' 0 218 $cfg.SendAttachments; $pCfg.Controls.Add($script:chkAtt)

# Defaults
$pCfg.Controls.Add((New-Lbl 'Valores por defecto' 0 252 300 18 $F.Bold $C.TextDark))
$pCfg.Controls.Add((New-Lbl 'Dias inactividad usuarios:'  0 276 175 18 $F.Sm $C.TextMid))
$script:cfgDU = New-Num  0 292 1 730 $cfg.DaysInactiveUsers; $pCfg.Controls.Add($script:cfgDU)
$pCfg.Controls.Add((New-Lbl 'Dias inactividad equipos:'  100 276 175 18 $F.Sm $C.TextMid))
$script:cfgDP = New-Num 275 292 1 730 $cfg.DaysInactivePC;   $pCfg.Controls.Add($script:cfgDP)
$pCfg.Controls.Add((New-Lbl 'Backups GPO a conservar:'    0 322 175 18 $F.Sm $C.TextMid))
$script:cfgMB = New-Num  0 338 1 100 $cfg.MaxBackups;        $pCfg.Controls.Add($script:cfgMB)

$btnSvCfg  = New-Btn 'Guardar'           0 380 140 32
$btnTstMail= New-Btn 'Enviar Reporte'  150 380 140 32 $C.Green $C.White
$btnRegTsk = New-Btn 'Programar tarea' 300 380 180 32 ([Drawing.Color]::FromArgb(80,80,100)) $C.White
$lblCfgSt  = New-Lbl '' 0 424 750 20 $F.Sm $C.TextMid
$pCfg.Controls.AddRange(@($btnSvCfg,$btnTstMail,$btnRegTsk,$lblCfgSt))

# Scheduled task info box
$rtbInfo = New-Object Windows.Forms.RichTextBox
$rtbInfo.Location=(New-Object System.Drawing.Point -ArgumentList 0, 452); $rtbInfo.Size=(New-Object System.Drawing.Size -ArgumentList 700, 158)
$rtbInfo.ReadOnly=$true; $rtbInfo.Font=$F.Sm; $rtbInfo.BackColor=[Drawing.Color]::FromArgb(240,244,255)
$rtbInfo.BorderStyle='FixedSingle'; $rtbInfo.ForeColor=$C.TextDark
$rtbInfo.Text = "Tarea programada ('Programar tarea'):`r`n`r`n" +
    "Registra una tarea en el Programador de Tareas de Windows que ejecuta`r`n" +
    "Send-ADDailyReport.ps1 todos los dias a las 07:00.`r`n`r`n" +
    "El script Send-ADDailyReport.ps1 puede ejecutarse tambien manualmente`r`n" +
    "o desde cualquier scheduler externo (Zabbix, SCOM, etc.)."
$pCfg.Controls.Add($rtbInfo)
$ct.Controls.Add($pCfg)
#endregion

#region -- SIDEBAR NAV ----------------------------------------------------------
$navDef = @(
    @{T='  Dashboard';          P=$pDash}
    @{T='  Usuarios Inactivos'; P=$pUsers}
    @{T='  Backup GPOs';        P=$pGPO}
    @{T='  Inventario Equipos'; P=$pInv}
    @{T='  Configuración';      P=$pCfg}
)
$script:curPnl = $pDash
$script:navBtns = @()
$y = 76
foreach ($nd in $navDef) {
    $nb = New-Object Windows.Forms.Button
    $nb.Text=$nd.T; $nb.Left=0; $nb.Top=$y; $nb.Width=205; $nb.Height=44
    $nb.Font=$F.Nav; $nb.ForeColor=[Drawing.Color]::FromArgb(170,195,220)
    $nb.BackColor=$C.Sidebar; $nb.FlatStyle='Flat'
    $nb.Cursor = [Windows.Forms.Cursors]::Hand
    $nb.TextAlign='MiddleLeft'; $nb.Padding=(New-Object System.Windows.Forms.Padding -ArgumentList 12, 0, 0, 0)
    $nb.FlatAppearance.BorderSize=0
    $nb.FlatAppearance.MouseOverBackColor=$C.SidebarHover
    $nb.FlatAppearance.MouseDownBackColor=$C.SidebarActive
    $script:navBtns += $nb
    $sb.Controls.Add($nb)
    $y += 44
}
$script:navBtns[0].BackColor=$C.SidebarActive; $script:navBtns[0].ForeColor=$C.White

# Paneles en array $script: para que el handler los acceda sin closures
$script:contentPanels = @($pDash, $pUsers, $pGPO, $pInv, $pCfg)
$script:navBg         = $C.Sidebar
$script:navFg         = [Drawing.Color]::FromArgb(170,195,220)
$script:navActiveBg   = $C.SidebarActive
$script:navActiveFg   = $C.White

# Asignar indice como Tag en cada boton — evita closures de loop que fallan en PS4
for ($i=0; $i -lt $navDef.Count; $i++) {
    $script:navBtns[$i].Tag = $i
}

# Handler unico compartido: usa $this.Tag para saber que panel mostrar
$navClickHandler = {
    $idx = [int]$this.Tag
    $script:curPnl.Visible = $false
    $script:contentPanels[$idx].Visible = $true
    $script:curPnl = $script:contentPanels[$idx]
    foreach ($b in $script:navBtns) {
        $b.BackColor = $script:navBg
        $b.ForeColor = $script:navFg
    }
    $this.BackColor = $script:navActiveBg
    $this.ForeColor = $script:navActiveFg
}
foreach ($btn in $script:navBtns) { $btn.Add_Click($navClickHandler) }

$sb.Controls.Add((New-Lbl 'AD Toolkit v1.0' 10 674 185 18 $F.Sm ([Drawing.Color]::FromArgb(60,80,110))))
#endregion

#region -- EVENT HANDLERS ------------------------------------------------------

# -- Dashboard: Ejecutar Todo
$btnAll.Add_Click({
    Set-Status 'Ejecutando todos los módulos...' $C.Blue
    $dashLog.Clear()

    function DLog($msg, $col=[Drawing.Color]::FromArgb(160,240,160)) {
        $dashLog.SelectionColor=$col
        $dashLog.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $msg`n")
        $dashLog.ScrollToCaret()
        $form.Refresh()
    }

    DLog "Iniciando ejecución completa..."

    # Usuarios
    DLog "Consultando usuarios inactivos (umbral: $($script:nDaysU.Value) días)..."
    try {
        $script:lastUsers = Get-InactiveUsersData $script:nDaysU.Value $script:tBaseU.Text $script:chkDis.Checked
        Set-GridData $script:grdU $script:lastUsers
        $cnt = @($script:lastUsers).Count
        $script:cUsers.Value.Text="$cnt"; $script:cUsers.Sub.Text="umbral: $($script:nDaysU.Value)d"
        DLog "Usuarios inactivos encontrados: $cnt" ([Drawing.Color]::FromArgb(100,220,255))
    } catch {
        DLog "ERROR usuarios: $($_.Exception.Message)" ([Drawing.Color]::FromArgb(255,100,100))
    }

    # GPO
    DLog "Ejecutando backup de GPOs (la ventana se actualiza por GPO)..."
    try {
        $log = Invoke-GPOBackup $script:tBkpRoot.Text $script:nMaxBkp.Value $dashLog
        $okN = ($log | Where-Object { $_ -match '\[OK\]' }).Count
        $script:cGPO.Value.Text="$okN"
    } catch {
        DLog "ERROR GPO: $($_.Exception.Message)" ([Drawing.Color]::FromArgb(255,100,100))
    }

    # Inventario
    DLog "Consultando inventario de equipos (umbral: $($script:nDaysPC.Value) días)..."
    try {
        $script:lastInv = Get-ComputersData $script:nDaysPC.Value $script:tBasePC.Text $script:tOSF.Text
        Set-GridData $script:grdI $script:lastInv
        $total  = @($script:lastInv).Count
        $inact  = @($script:lastInv | Where-Object { [int]$_.DiasInactivo -gt $script:nDaysPC.Value }).Count
        $script:cPCsI.Value.Text="$inact"; $script:cPCsI.Sub.Text="umbral: $($script:nDaysPC.Value)d"
        $script:cPCsT.Value.Text="$total"
        DLog "Equipos: total=$total, inactivos=$inact" ([Drawing.Color]::FromArgb(100,220,255))
    } catch {
        DLog "ERROR inventario: $($_.Exception.Message)" ([Drawing.Color]::FromArgb(255,100,100))
    }

    $lblLast.Text = "Última ejecución: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"
    Set-Status "Completado  - $(Get-Date -Format 'HH:mm')" $C.Green
})

# -- Usuarios
$btnRunU.Add_Click({
    Set-Status 'Consultando usuarios...' $C.Blue
    try {
        $script:lastUsers = Get-InactiveUsersData $script:nDaysU.Value $script:tBaseU.Text $script:chkDis.Checked
        Set-GridData $script:grdU $script:lastUsers
        $cnt = @($script:lastUsers).Count
        $lblCntU.Text = "$cnt usuarios encontrados"
        $script:cUsers.Value.Text="$cnt"; $script:cUsers.Sub.Text="umbral: $($script:nDaysU.Value)d"
        Set-Status "Usuarios cargados: $cnt" $C.Green
    } catch {
        Set-Status "Error: $($_.Exception.Message)" $C.Red
        [Windows.Forms.MessageBox]::Show($_.Exception.Message,'Error','OK','Error') | Out-Null
    }
})
$btnExpU.Add_Click({
    if ($script:grdU.Rows.Count -eq 0) { [Windows.Forms.MessageBox]::Show('Sin datos para exportar.','Exportar','OK','Warning') | Out-Null; return }
    Export-GridCsv $script:grdU "InactiveUsers_$(Get-Date -Format 'yyyyMMdd').csv"
})

# -- GPO
$btnBrowse.Add_Click({
    $dlg = New-Object Windows.Forms.FolderBrowserDialog; $dlg.SelectedPath=$script:tBkpRoot.Text
    if ($dlg.ShowDialog() -eq 'OK') { $script:tBkpRoot.Text=$dlg.SelectedPath }
})
$btnRunGPO.Add_Click({
    Set-Status 'Ejecutando backup de GPOs...' $C.Blue; $rtbGPO.Clear()
    try {
        # Pasar $rtbGPO para que el log se actualice en tiempo real (sin tildar la ventana)
        $log = Invoke-GPOBackup $script:tBkpRoot.Text $script:nMaxBkp.Value $rtbGPO
        # Agregar lineas de resumen final (las individuales ya las agrego la funcion)
        foreach ($line in ($log | Where-Object { $_ -match 'RESULTADO|INFO.*eliminado|^$' })) {
            $col = if ($line -match 'RESULT') { [Drawing.Color]::FromArgb(255,230,100) } else { [Drawing.Color]::FromArgb(160,240,160) }
            $rtbGPO.SelectionColor=$col; $rtbGPO.AppendText("$line`n")
        }
        $rtbGPO.ScrollToCaret()
        $okN = ($log | Where-Object { $_ -match '\[OK\]' }).Count
        $script:cGPO.Value.Text="$okN"
        Set-Status "Backup completado ($okN GPOs)" $C.Green
    } catch {
        Set-Status "Error: $($_.Exception.Message)" $C.Red
    }
})

# -- Inventario
$btnRunI.Add_Click({
    Set-Status 'Consultando equipos...' $C.Blue
    try {
        $script:lastInv = Get-ComputersData $script:nDaysPC.Value $script:tBasePC.Text $script:tOSF.Text
        Set-GridData $script:grdI $script:lastInv
        $cnt = @($script:lastInv).Count
        $lblCntI.Text="$cnt equipos encontrados"
        $script:cPCsT.Value.Text="$cnt"
        Set-Status "Equipos cargados: $cnt" $C.Green
    } catch {
        Set-Status "Error: $($_.Exception.Message)" $C.Red
        [Windows.Forms.MessageBox]::Show($_.Exception.Message,'Error','OK','Error') | Out-Null
    }
})
$btnExpI.Add_Click({
    if ($script:grdI.Rows.Count -eq 0) { [Windows.Forms.MessageBox]::Show('Sin datos para exportar.','Exportar','OK','Warning') | Out-Null; return }
    Export-GridCsv $script:grdI "DomainComputers_$(Get-Date -Format 'yyyyMMdd').csv"
})

# -- Config: Guardar
$btnSvCfg.Add_Click({
    $cfg.GraphTenantId    = $script:cfgCtrl['GraphTenantId'].Text
    $cfg.GraphClientId    = $script:cfgCtrl['GraphClientId'].Text
    $cfg.EmailFrom        = $script:cfgCtrl['EmailFrom'].Text
    $cfg.EmailTo          = $script:cfgCtrl['EmailTo'].Text
    $cfg.SendAttachments  = $script:chkAtt.Checked
    $cfg.DaysInactiveUsers= [int]$script:cfgDU.Value
    $cfg.DaysInactivePC   = [int]$script:cfgDP.Value
    $cfg.MaxBackups       = [int]$script:cfgMB.Value
    $cfg.BackupRoot       = $script:tBkpRoot.Text
    $secret = $script:cfgCtrl['GraphClientSecret'].Text
    if ($secret) { $cfg.GraphClientSecretEncrypted = $secret | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString }
    Save-ADConfig $cfg
    $lblCfgSt.Text="Configuracion guardada - $(Get-Date -Format 'HH:mm:ss')"; $lblCfgSt.ForeColor=$C.Green
    Set-Status 'Configuracion guardada.' $C.Green
})

# -- Config: Enviar Reporte
$btnTstMail.Add_Click({
    if (-not $cfg.GraphTenantId -or -not $cfg.GraphClientSecretEncrypted -or -not $cfg.EmailTo) {
        [Windows.Forms.MessageBox]::Show('Complete todos los campos de Graph API y guarde la configuracion antes de enviar.','Enviar Reporte','OK','Warning') | Out-Null; return
    }
    $lblCfgSt.Text='Consultando datos de AD...'; $lblCfgSt.ForeColor=$C.Blue; $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    try {
        # -- Usuarios inactivos
        $rUsers = @(Get-InactiveUsersData $cfg.DaysInactiveUsers $cfg.SearchBase $false)
        [System.Windows.Forms.Application]::DoEvents()
        # -- Inventario equipos
        $rPCs   = @(Get-ComputersData $cfg.DaysInactivePC $cfg.SearchBase $cfg.OSFilter)
        [System.Windows.Forms.Application]::DoEvents()

        # -- Contraseñas por vencer y cuentas sin expiracion
        $lblCfgSt.Text='Consultando politica de contraseñas...'; $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        $expiring      = @()
        $neverExpiresU = 0
        $pcPwdNeverExp = 0
        try {
            $maxPwdAge     = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
            $neverExpiresU = @(Get-ADUser -Filter * -Properties PasswordNeverExpires |
                Where-Object { $_.Enabled -and $_.PasswordNeverExpires }).Count
            $expiring = @(Get-ADUser -Filter * -Properties PasswordLastSet,Department,PasswordNeverExpires |
                Where-Object { $_.Enabled -and -not $_.PasswordNeverExpires -and $_.PasswordLastSet } |
                ForEach-Object {
                    $exp  = $_.PasswordLastSet + $maxPwdAge
                    $days = [math]::Round(($exp - (Get-Date)).TotalDays)
                    if ($days -ge 0 -and $days -le 15) {
                        [PSCustomObject]@{
                            Usuario       = $_.SamAccountName
                            Nombre        = $_.Name
                            Departamento  = $_.Department
                            Expira        = $exp.ToString('yyyy-MM-dd')
                            DiasRestantes = $days
                        }
                    }
                } | Where-Object { $_ } | Sort-Object DiasRestantes)
            $pcPwdNeverExp = @(Get-ADComputer -Filter * -Properties PasswordNeverExpires |
                Where-Object { $_.Enabled -and $_.PasswordNeverExpires }).Count
        } catch { }
        [System.Windows.Forms.Application]::DoEvents()

        $usersCount  = $rUsers.Count
        $expiringCnt = $expiring.Count
        $pcTotal     = $rPCs.Count
        $pcInact     = @($rPCs | Where-Object { $_.DiasInactivo -ne 'Nunca' -and [int]"$($_.DiasInactivo)" -gt [int]$cfg.DaysInactivePC }).Count

        try { $domName = (Get-ADDomain).DNSRoot } catch { $domName = 'N/A' }

        # -- Construir tablas HTML
        function fmtRow($data, $cols, $limit=40) {
            $arr = @($data)
            if ($arr.Count -eq 0) { return '<p style="color:#64748b;font-style:italic">Sin registros.</p>' }
            $hdr  = ($cols | ForEach-Object { "<th>$_</th>" }) -join ''
            $rows = ($arr | Select-Object -First $limit | ForEach-Object {
                $r=$_; $cells=($cols | ForEach-Object { "<td>$($r.$_)</td>" }) -join ''; "<tr>$cells</tr>"
            }) -join ''
            $more = if ($arr.Count -gt $limit) { "<p style='font-size:12px;color:#94a3b8'>Mostrando $limit de $($arr.Count).</p>" } else { '' }
            return "<table><thead><tr>$hdr</tr></thead><tbody>$rows</tbody></table>$more"
        }

        # Tabla contraseñas: coloreada por urgencia (rojo ≤3d, naranja ≤7d, verde el resto)
        $tblExp = if ($expiringCnt -eq 0) {
            '<p style="color:#15803d;font-weight:600">Sin contraseñas por vencer en los proximos 15 dias.</p>'
        } else {
            $hdr  = '<th>Usuario</th><th>Nombre</th><th>Departamento</th><th>Expira</th><th>Dias restantes</th>'
            $rows = ($expiring | ForEach-Object {
                $bg = if ($_.DiasRestantes -le 3) { '#fee2e2' } elseif ($_.DiasRestantes -le 7) { '#fef3c7' } else { '#f0fdf4' }
                "<tr style='background:$bg'><td>$($_.Usuario)</td><td>$($_.Nombre)</td><td>$($_.Departamento)</td><td>$($_.Expira)</td><td><b>$($_.DiasRestantes)</b></td></tr>"
            }) -join ''
            "<table><thead><tr>$hdr</tr></thead><tbody>$rows</tbody></table>"
        }
        $tblU = fmtRow $rUsers @('Usuario','Nombre','Departamento','UltimoLogin','DiasInactivo')
        $tblP = fmtRow ($rPCs | Where-Object { $_.DiasInactivo -ne 'Nunca' -and [int]"$($_.DiasInactivo)" -gt [int]$cfg.DaysInactivePC }) @('Nombre','IP','SistemaOperativo','UltimaConexion','DiasInactivo')

        $html = @"
<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<style>
  body{font-family:Segoe UI,Arial,sans-serif;background:#f1f5f9;margin:0;padding:20px;color:#1e293b}
  .wrap{max-width:900px;margin:auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,.1)}
  .hdr{background:linear-gradient(135deg,#0054a6,#0078d4);color:#fff;padding:28px 36px}
  .hdr h1{margin:0;font-size:22px}.hdr p{margin:6px 0 0;font-size:13px;opacity:.85}
  .body{padding:32px}
  .cards{display:flex;gap:12px;margin-bottom:32px;flex-wrap:wrap}
  .card{flex:1;min-width:130px;border-radius:8px;padding:14px 16px}
  .card b{display:block;font-size:26px;font-weight:800;line-height:1.1}.card p{margin:5px 0 0;font-size:11px}
  .cr{background:#fee2e2;color:#dc2626}.cy{background:#fef3c7;color:#b45309}
  .cg{background:#dcfce7;color:#15803d}.co{background:#ffedd5;color:#c2410c}
  .cp{background:#f3e8ff;color:#7c3aed}
  h2{color:#0054a6;border-bottom:2px solid #e2e8f0;padding-bottom:8px;font-size:15px;margin-top:32px}
  table{border-collapse:collapse;width:100%;font-size:13px;margin:12px 0 6px}
  th{background:#0054a6;color:#fff;padding:8px 12px;text-align:left;font-weight:600}
  td{padding:6px 12px;border-bottom:1px solid #e2e8f0}
  .ftr{background:#f1f5f9;padding:16px 36px;font-size:12px;color:#94a3b8;text-align:center;border-top:1px solid #e2e8f0}
</style></head><body>
<div class="wrap">
  <div class="hdr"><h1>AD Toolkit - Reporte</h1>
    <p>$(Get-Date -Format 'dddd, dd MMMM yyyy HH:mm') &bull; Dominio: $domName</p></div>
  <div class="body">
    <div class="cards">
      <div class="card cr"><b>$usersCount</b><p>Usuarios inactivos (&gt;$($cfg.DaysInactiveUsers) dias)</p></div>
      <div class="card co"><b>$expiringCnt</b><p>Contrasenas por vencer (&le;15 dias)</p></div>
      <div class="card cp"><b>$neverExpiresU</b><p>Usuarios sin expiracion de pass.</p></div>
      <div class="card cy"><b>$pcInact</b><p>Equipos inactivos (&gt;$($cfg.DaysInactivePC) dias)</p></div>
      <div class="card cg"><b>$pcTotal</b><p>Total equipos &bull; $pcPwdNeverExp con pass. fija</p></div>
    </div>
    <h2>Contrasenas por vencer - proximos 15 dias ($expiringCnt)</h2>
    $tblExp
    <h2>Usuarios Inactivos ($usersCount)</h2>
    $tblU
    <h2>Equipos Inactivos ($pcInact de $pcTotal)</h2>
    $tblP
  </div>
  <div class="ftr">Generado por AD Toolkit &bull; $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>
</div></body></html>
"@

        $lblCfgSt.Text='Enviando reporte...'; $lblCfgSt.ForeColor=$C.Blue; $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()

        $cred   = New-Object Management.Automation.PSCredential('x', ($cfg.GraphClientSecretEncrypted | ConvertTo-SecureString))
        $secret = $cred.GetNetworkCredential().Password
        $subj   = "AD Toolkit $(Get-Date -Format 'dd/MM/yyyy') -- $usersCount inactivos | $expiringCnt pass. por vencer | $pcTotal equipos"
        Send-GraphMail -tenantId $cfg.GraphTenantId -clientId $cfg.GraphClientId `
            -clientSecret $secret -from $cfg.EmailFrom -to $cfg.EmailTo `
            -subject $subj -htmlBody $html
        $lblCfgSt.Text="Reporte enviado - $(Get-Date -Format 'HH:mm:ss')"; $lblCfgSt.ForeColor=$C.Green
    } catch {
        $lblCfgSt.Text="Error: $($_.Exception.Message)"; $lblCfgSt.ForeColor=$C.Red
    }
})

# -- Config: Registrar tarea programada
$btnRegTsk.Add_Click({
    $sender = Join-Path $scriptRoot 'Send-ADDailyReport.ps1'
    if (-not (Test-Path $sender)) {
        [Windows.Forms.MessageBox]::Show("No encontrado: $sender",'Error','OK','Error') | Out-Null; return
    }
    $timePick = [Windows.Forms.MessageBox]::Show(
        "Se registrará la tarea 'AD Toolkit Daily Report' para ejecutarse todos los días a las 07:00.`n`n¿Confirmar?",
        'Programar tarea', 'YesNo', 'Question')
    if ($timePick -ne 'Yes') { return }
    try {
        $action   = New-ScheduledTaskAction -Execute 'powershell.exe' `
            -Argument "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$sender`""
        $trigger  = New-ScheduledTaskTrigger -Daily -At '07:00'
        $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable
        Register-ScheduledTask -TaskName 'AD Toolkit Daily Report' `
            -Action $action -Trigger $trigger -Settings $settings `
            -Description 'Reporte diario AD Toolkit' -Force | Out-Null
        $lblCfgSt.Text='Tarea registrada: "AD Toolkit Daily Report" - 07:00 diario'; $lblCfgSt.ForeColor=$C.Green
    } catch {
        $lblCfgSt.Text="Error: $($_.Exception.Message)"; $lblCfgSt.ForeColor=$C.Red
    }
})
#endregion

$form.Add_Shown({ Set-Status "AD Toolkit iniciado - $(Get-Date -Format 'dd/MM/yyyy HH:mm')" })
[void]$form.ShowDialog()
