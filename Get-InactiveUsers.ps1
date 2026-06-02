#Requires -Module ActiveDirectory
<#
.SYNOPSIS
    Genera un CSV con usuarios de AD que no iniciaron sesión en los últimos X días.
.PARAMETER DaysInactive
    Cantidad de días sin inicio de sesión para considerar un usuario inactivo. Default: 90.
.PARAMETER OutputPath
    Ruta del archivo CSV de salida. Default: InactiveUsers_<fecha>.csv en el directorio actual.
.PARAMETER SearchBase
    OU base de búsqueda. Si se omite, busca en todo el dominio.
.PARAMETER IncludeDisabled
    Si se especifica, incluye también las cuentas deshabilitadas.
.EXAMPLE
    .\Get-InactiveUsers.ps1 -DaysInactive 60 -OutputPath "C:\Reports\inactivos.csv"
.EXAMPLE
    .\Get-InactiveUsers.ps1 -DaysInactive 30 -SearchBase "OU=Empleados,DC=empresa,DC=local"
#>
[CmdletBinding()]
param (
    [Parameter()]
    [int]$DaysInactive = 90,

    [Parameter()]
    [string]$OutputPath = "InactiveUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter()]
    [string]$SearchBase,

    [Parameter()]
    [switch]$IncludeDisabled
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Status {
    param([string]$Message, [string]$Color = 'Cyan')
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $Message" -ForegroundColor $Color
}

Write-Status "Iniciando reporte de usuarios inactivos (umbral: $DaysInactive días)"

$cutoffDate = (Get-Date).AddDays(-$DaysInactive)

$adParams = @{
    Filter     = '*'
    Properties = 'LastLogonDate', 'PasswordLastSet', 'PasswordNeverExpires',
                 'PasswordExpired', 'EmailAddress', 'Department', 'Manager',
                 'Description', 'Enabled', 'Created', 'Modified', 'DistinguishedName'
}
if ($SearchBase) { $adParams['SearchBase'] = $SearchBase }

Write-Status "Consultando Active Directory..."
$allUsers = Get-ADUser @adParams

$filtered = $allUsers | Where-Object {
    ($IncludeDisabled -or $_.Enabled) -and
    (
        $null -eq $_.LastLogonDate -or
        $_.LastLogonDate -lt $cutoffDate
    )
}

Write-Status "Procesando $($filtered.Count) usuarios inactivos encontrados..."

$report = $filtered | ForEach-Object {
    $managerName = ''
    if ($_.Manager) {
        try {
            $mgr = Get-ADUser -Identity $_.Manager -Properties DisplayName
            $managerName = $mgr.DisplayName
        } catch {
            $managerName = $_.Manager
        }
    }

    $daysSinceLogin = if ($_.LastLogonDate) {
        [math]::Round(((Get-Date) - $_.LastLogonDate).TotalDays)
    } else {
        'Nunca'
    }

    [PSCustomObject]@{
        SamAccountName      = $_.SamAccountName
        DisplayName         = $_.Name
        Email               = $_.EmailAddress
        Department          = $_.Department
        Manager             = $managerName
        Habilitado          = $_.Enabled
        UltimoInicioSesion  = if ($_.LastLogonDate) { $_.LastLogonDate.ToString('yyyy-MM-dd HH:mm') } else { 'Nunca' }
        DiasInactivo        = $daysSinceLogin
        UltimoCambioPass    = if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString('yyyy-MM-dd') } else { 'Nunca' }
        PassNuncaExpira     = $_.PasswordNeverExpires
        PassExpirada        = $_.PasswordExpired
        CreadoEn            = $_.Created.ToString('yyyy-MM-dd')
        OU                  = ($_.DistinguishedName -replace '^CN=[^,]+,', '')
        Descripcion         = $_.Description
    }
}

$report | Sort-Object DiasInactivo -Descending |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Status "Reporte generado: $OutputPath" -Color Green
Write-Status "Total usuarios inactivos: $($report.Count)" -Color Yellow

$report | Sort-Object DiasInactivo -Descending | Select-Object -First 5 |
    Format-Table SamAccountName, DisplayName, UltimoInicioSesion, DiasInactivo, Department -AutoSize
