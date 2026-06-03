<#
.SYNOPSIS
    Genera un archivo .xlsx con multiples hojas sin modulos externos.
    Compatible con PowerShell 4.0 / .NET 4.5 (Windows Server 2012 R2+).
#>

function Export-MultiSheetXlsx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][array]$Sheets
    )

    Add-Type -AssemblyName 'System.IO.Compression'
    Add-Type -AssemblyName 'System.IO.Compression.FileSystem'

    #-- Escape caracteres XML especiales
    function xesc([string]$s) {
        return $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
    }

    #-- Letra de columna (0-based: 0=A, 25=Z, 26=AA)
    function colLetter([int]$i) {
        if ($i -lt 26) { return [string][char](65 + $i) }
        return ([string][char](64 + [math]::Floor($i / 26))) + ([string][char](65 + ($i % 26)))
    }

    #-- Agrega una entrada al ZIP
    function zipAdd([System.IO.Compression.ZipArchive]$zip, [string]$name, [string]$xml) {
        $entry  = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
        $stream = $entry.Open()
        $enc    = New-Object System.Text.UTF8Encoding($false)
        $writer = New-Object System.IO.StreamWriter($stream, $enc)
        $writer.Write($xml)
        $writer.Flush()
        $writer.Close()
        $stream.Dispose()
    }

    #-- Genera el XML de una hoja
    function sheetXml([hashtable]$sheet) {
        $rows = @($sheet.Data)
        if ($rows.Count -eq 0) {
            return '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
        }
        $headers = @($rows[0].PSObject.Properties.Name)
        $sb = New-Object System.Text.StringBuilder
        $null = $sb.Append('<?xml version="1.0" encoding="UTF-8"?>')
        $null = $sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        $null = $sb.Append('<sheetData>')

        # Encabezado (estilo 1 = azul + blanco)
        $null = $sb.Append('<row r="1">')
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $col = colLetter $c
            $v   = xesc $headers[$c]
            $null = $sb.Append('<c r="' + $col + '1" t="inlineStr" s="1"><is><t>' + $v + '</t></is></c>')
        }
        $null = $sb.Append('</row>')

        # Datos
        for ($r = 0; $r -lt $rows.Count; $r++) {
            $rowNum = $r + 2
            $null = $sb.Append('<row r="' + $rowNum + '">')
            for ($c = 0; $c -lt $headers.Count; $c++) {
                $col = colLetter $c
                $val = [string]$rows[$r].($headers[$c])
                $tmp = 0
                if ([int]::TryParse($val, [ref]$tmp)) {
                    $null = $sb.Append('<c r="' + $col + $rowNum + '"><v>' + $val + '</v></c>')
                } else {
                    $v = xesc $val
                    $null = $sb.Append('<c r="' + $col + $rowNum + '" t="inlineStr"><is><t>' + $v + '</t></is></c>')
                }
            }
            $null = $sb.Append('</row>')
        }

        $lastCol = colLetter ($headers.Count - 1)
        $null = $sb.Append('</sheetData>')
        $null = $sb.Append('<autoFilter ref="A1:' + $lastCol + '1"/>')
        $null = $sb.Append('</worksheet>')
        return $sb.ToString()
    }

    #====================================================================
    # Crear ZIP en memoria
    #====================================================================
    $n   = $Sheets.Count
    $mem = New-Object System.IO.MemoryStream
    $zip = New-Object System.IO.Compression.ZipArchive($mem, [System.IO.Compression.ZipArchiveMode]::Create, $true)

    #-- [Content_Types].xml
    $ct = New-Object System.Text.StringBuilder
    $null = $ct.Append('<?xml version="1.0" encoding="UTF-8"?>')
    $null = $ct.Append('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
    $null = $ct.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
    $null = $ct.Append('<Default Extension="xml"  ContentType="application/xml"/>')
    $null = $ct.Append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    $null = $ct.Append('<Override PartName="/xl/styles.xml"   ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    for ($i = 1; $i -le $n; $i++) {
        $null = $ct.Append('<Override PartName="/xl/worksheets/sheet' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    }
    $null = $ct.Append('</Types>')
    zipAdd $zip '[Content_Types].xml' $ct.ToString()

    #-- _rels/.rels
    zipAdd $zip '_rels/.rels' ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')

    #-- xl/workbook.xml
    $wb = New-Object System.Text.StringBuilder
    $null = $wb.Append('<?xml version="1.0" encoding="UTF-8"?>')
    $null = $wb.Append('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    $null = $wb.Append('<sheets>')
    for ($i = 0; $i -lt $n; $i++) {
        $sname = xesc $Sheets[$i].Name
        $null = $wb.Append('<sheet name="' + $sname + '" sheetId="' + ($i+1) + '" r:id="rId' + ($i+1) + '"/>')
    }
    $null = $wb.Append('</sheets></workbook>')
    zipAdd $zip 'xl/workbook.xml' $wb.ToString()

    #-- xl/_rels/workbook.xml.rels
    $wr = New-Object System.Text.StringBuilder
    $null = $wr.Append('<?xml version="1.0" encoding="UTF-8"?>')
    $null = $wr.Append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    for ($i = 0; $i -lt $n; $i++) {
        $null = $wr.Append('<Relationship Id="rId' + ($i+1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + ($i+1) + '.xml"/>')
    }
    $styId = $n + 1
    $null = $wr.Append('<Relationship Id="rId' + $styId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    $null = $wr.Append('</Relationships>')
    zipAdd $zip 'xl/_rels/workbook.xml.rels' $wr.ToString()

    #-- xl/styles.xml (estilo 0=normal, 1=encabezado azul corporativo)
    $styles = '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0054A6"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>'
    zipAdd $zip 'xl/styles.xml' $styles

    #-- xl/worksheets/sheet{N}.xml
    for ($i = 0; $i -lt $n; $i++) {
        zipAdd $zip ('xl/worksheets/sheet' + ($i+1) + '.xml') (sheetXml $Sheets[$i])
    }

    $zip.Dispose()
    [System.IO.File]::WriteAllBytes($Path, $mem.ToArray())
    $mem.Dispose()
}
