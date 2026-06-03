# AD Toolkit - Automatizacion de Active Directory

![PowerShell](https://img.shields.io/badge/PowerShell-4.0%2B-5391FE?style=flat&logo=powershell&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows%20Server-0078D6?style=flat&logo=windows&logoColor=white)
![Active Directory](https://img.shields.io/badge/Active%20Directory-0078D4?style=flat&logo=microsoft&logoColor=white)
![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph%20API-00A4EF?style=flat&logo=microsoft&logoColor=white)
![WinForms](https://img.shields.io/badge/UI-WinForms-512BD4?style=flat&logo=dotnet&logoColor=white)

Toolkit de PowerShell con interfaz grafica WinForms para automatizar auditorias de Active Directory. Centraliza reportes de usuarios inactivos, backup de GPOs, inventario de equipos y control de contrasenas por vencer, con envio de reportes diarios por email via Microsoft Graph API.

---

## Problema que resuelve

En entornos con Active Directory, tareas de auditoria como detectar cuentas inactivas, respaldar politicas de grupo, inventariar equipos o anticipar vencimientos de contrasenas se realizan manualmente desde distintas consolas, sin registro centralizado ni automatizacion. Este toolkit unifica esas operaciones en una sola interfaz y las programa para ejecutarse de forma autonoma, incluso sin sesion de usuario abierta.

---

## Funcionalidades

### Dashboard
- Resumen en tiempo real de los indicadores clave del dominio
- Log de ejecucion integrado con codigo de colores por severidad
- Boton "Ejecutar Todo" para correr todos los modulos en secuencia

### Usuarios Inactivos
- Detecta cuentas sin login en X dias (configurable)
- Filtra por OU base
- Visualizacion en grilla con exportacion a CSV

### Backup de GPOs
- Respalda todas las GPOs del dominio con timestamp
- Destino configurable: carpeta local o recurso compartido de red (NAS/UNC)
- Rotacion automatica: conserva solo los N backups mas recientes
- Log por cada GPO procesada

### Inventario de Equipos
- Lista todos los equipos con OS, IP, dias de inactividad y estado
- Filtro por sistema operativo (wildcards: `*Server*`, `*Windows 10*`)
- Exportacion a CSV

### Contrasenas por Vencer (nuevo)
- Detecta usuarios cuya contrasena vence en los proximos N dias (configurable, default 15)
- Indicador de urgencia: CRITICO (<3 dias), Alto (<7 dias), Medio
- Incluido en el reporte diario y en el Excel adjunto

### Reporte Diario por Email (Microsoft Graph API)
- **Cuerpo del email**: mensaje simple con los 4 indicadores clave
- **Adjuntos**:
  - `ADReport_YYYYMMDD.xlsx` - Excel con 4 hojas (ver abajo)
  - `ADReport_YYYYMMDD.html` - Reporte visual completo (o PDF si hay conversor instalado)
- Los CSV se generan localmente pero no se adjuntan al email
- Envio via Microsoft Graph API (sin SMTP legacy)
- Asunto dinamico con resumen del dia

### Excel multi-hoja (nativo, sin modulos externos)
- Generado sin necesidad de `Install-Module` ni Office instalado
- 4 hojas en un solo archivo `.xlsx`:

| Hoja | Contenido |
|------|-----------|
| Usuarios Inactivos | Sam, Nombre, Email, Departamento, Ultimo login, Dias inactivo, Ultimo cambio de pass |
| Equipos Inactivos | Nombre, IP, OS, Habilitado, Ultima conexion, Dias inactivo |
| Contrasenas por Vencer | Sam, Nombre, Departamento, Fecha vence, Dias restantes, Urgencia |
| Backup GPOs | Nombre GPO, Estado OK/ERROR, Detalle |

### PDF opcional
- Si `wkhtmltopdf` esta instalado, genera PDF del reporte en lugar de HTML
- Tambien funciona con Chrome instalado (modo headless)
- Si no hay conversor: adjunta el HTML directamente

### Configuracion de la tarea programada
- `Register-DailyReportTask.ps1` registra el reporte como tarea del sistema
- Se ejecuta **sin necesidad de sesion de usuario abierta** (LogonType: Password)
- `StartWhenAvailable`: si el equipo estaba apagado a la hora programada, corre al encenderse
- Reintento automatico hasta 3 veces si falla

---

## Tech Stack

| Componente | Detalle |
|---|---|
| Lenguaje | PowerShell 4.0+ |
| UI | WinForms (nativo, sin dependencias externas) |
| Acceso a AD | RSAT - modulos `ActiveDirectory` y `GroupPolicy` |
| API de email | Microsoft Graph API v1.0 (REST via `Invoke-RestMethod`) |
| Autenticacion | OAuth 2.0 Client Credentials Flow |
| Configuracion | JSON + cifrado DPAPI para el secreto de cliente |
| Excel | Generacion nativa con `System.IO.Compression` (.NET 4.5) - sin modulos externos |

---

## Estructura del proyecto

```
AD-Toolkit/
|-- AD-Toolkit-GUI.ps1          # Interfaz grafica principal (WinForms)
|-- Send-ADDailyReport.ps1      # Reporte headless para Task Scheduler
|-- Export-MultiSheetXlsx.ps1   # Generador de Excel nativo (.NET 4.5, sin modulos)
|-- Register-DailyReportTask.ps1# Registra la tarea en el Programador de Windows
|-- Get-InactiveUsers.ps1       # Script standalone: usuarios inactivos
|-- Backup-GPOs.ps1             # Script standalone: backup de GPOs
|-- Get-DomainComputers.ps1     # Script standalone: inventario de equipos
|-- Launch-ADToolkit.bat        # Launcher con captura de errores
|-- config.json                 # Configuracion local (excluido del repo)
`-- config.example.json         # Plantilla de configuracion
```

---

## Instalacion y uso

### Requisitos

| Requisito | Detalle |
|---|---|
| Windows Server | 2012 R2 o superior |
| PowerShell | 4.0 o superior |
| RSAT | Modulos `ActiveDirectory` y `GroupPolicy` |
| Microsoft 365 | Buzon Exchange Online (puede ser buzon compartido, sin licencia paga) |

### Instalar RSAT (si no esta disponible)

```powershell
# Windows Server
Install-WindowsFeature RSAT-AD-PowerShell
Install-WindowsFeature GPMC
```

### Configurar

```
copy config.example.json config.json
```

Editar `config.json` con los valores del entorno, o completarlo desde la GUI en la seccion **Configuracion**.

### Ejecutar la GUI

```powershell
# Con launcher (recomendado)
Launch-ADToolkit.bat

# Directo
powershell.exe -ExecutionPolicy Bypass -File "AD-Toolkit-GUI.ps1"
```

### Registrar la tarea programada (una sola vez, como Administrador)

```powershell
# Corre a las 8:00, sin necesidad de sesion abierta
.\Register-DailyReportTask.ps1 -Hour 8 -Minute 0

# Verificar que quedo con LogonType: Password (no Interactive)
(Get-ScheduledTask -TaskName "AD Toolkit Daily Report").Principal | Select-Object UserId, LogonType
```

---

## Configuracion de Microsoft Graph API

Para habilitar el envio de reportes por email se necesita una **App Registration en Azure AD**:

1. Ir a **portal.azure.com > Microsoft Entra > Registros de aplicaciones > Nuevo registro**
2. Nombre: `AD Toolkit Mailer` - Tipo de cuenta: solo este directorio
3. Anotar el **Tenant ID** y el **Client ID**
4. **Certificados y secretos > Nuevo secreto de cliente** -> copiar el valor
5. **Permisos de API > Microsoft Graph > Permisos de aplicacion > `Mail.Send`**
6. Clic en **Conceder consentimiento de administrador**

Ingresar los tres valores en la GUI (seccion Configuracion) y hacer clic en **Guardar**.

> El buzon remitente puede ser un **buzon compartido** de Exchange Online: no consume licencia de usuario y el permiso `Mail.Send` de la app puede enviar desde el sin configuracion adicional.

### Archivo de configuracion

`config.json` (generado automaticamente por la GUI):

```json
{
  "DaysInactiveUsers": 90,
  "DaysInactivePC": 90,
  "DaysPasswordExpiry": 15,
  "SearchBase": "",
  "BackupRoot": "C:\\AD-Toolkit\\GPO-Backups",
  "NASBackupPath": "\\\\NAS01\\Backup\\GPO-Backups",
  "MaxBackups": 10,
  "OSFilter": "",
  "GraphTenantId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "GraphClientId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "GraphClientSecretEncrypted": "<cifrado DPAPI>",
  "EmailFrom": "alertas-ad@tudominio.com",
  "EmailTo": "admin@tudominio.com",
  "SendAttachments": true
}
```

**Campos opcionales:**

| Campo | Default | Descripcion |
|-------|---------|-------------|
| `DaysPasswordExpiry` | 15 | Dias de anticipacion para alertar contrasenas por vencer |
| `NASBackupPath` | (ninguno) | Ruta UNC del NAS. Si se define, el backup va ahi en lugar de `BackupRoot` |
| `WkhtmltopdfPath` | Ruta default de instalacion | Ruta al ejecutable de wkhtmltopdf si no esta en la ruta default |

> `config.json` esta en `.gitignore` y nunca se sube al repositorio.

---

## Reporte diario (headless)

`Send-ADDailyReport.ps1` esta disenado para correr desde el Programador de Tareas sin interfaz grafica.

```powershell
# Ejecutar manualmente (con envio de email)
.\Send-ADDailyReport.ps1

# Solo generar archivos locales, sin enviar email (ideal para testear)
.\Send-ADDailyReport.ps1 -NoMail
```

### Archivos que genera por ejecucion

```
Reports/
|-- ADReport_YYYYMMDD.html      # Reporte visual completo
|-- ADReport_YYYYMMDD.xlsx      # Excel con 4 hojas
|-- ADReport_YYYYMMDD.pdf       # Solo si hay conversor PDF instalado
|-- Users_YYYYMMDD.csv          # Usuarios inactivos (local, no se adjunta)
|-- PCs_YYYYMMDD.csv            # Equipos (local, no se adjunta)
|-- GPO_YYYYMMDD.csv            # Resultado del backup (local, no se adjunta)
|-- PassExp_YYYYMMDD.csv        # Contrasenas por vencer (local, no se adjunta)
`-- daily_log.txt               # Log acumulado de todas las ejecuciones
```

---

## PDF opcional (sin licencia)

Si se instala `wkhtmltopdf` (gratuito), el reporte se adjunta como PDF en lugar de HTML:

1. Descargar desde [wkhtmltopdf.org](https://wkhtmltopdf.org/downloads.html)
2. Instalar en `C:\Program Files\wkhtmltopdf\` (ruta que el script detecta automaticamente)
3. El proximo reporte adjuntara `.pdf` en lugar de `.html`

Si Chrome esta instalado en el servidor, tambien se usa automaticamente como conversor.

---

## Scripts standalone

Los scripts de auditoria funcionan de forma independiente:

```powershell
# Usuarios sin login en los ultimos 60 dias
.\Get-InactiveUsers.ps1 -DaysInactive 60 -OutputPath "C:\Reports\usuarios.csv"

# Backup conservando los ultimos 5
.\Backup-GPOs.ps1 -BackupRoot "D:\GPO-Backups" -MaxBackups 5 -GenerateReport

# Inventario filtrando solo servidores
.\Get-DomainComputers.ps1 -OSFilter "*Server*" -GenerateReport
```

---

## Errores comunes

| Error | Causa | Solucion |
|---|---|---|
| "No se reconoce el modulo ActiveDirectory" | RSAT no instalado | `Install-WindowsFeature RSAT-AD-PowerShell` |
| "Invoke-RestMethod: 401 Unauthorized" | Secreto de cliente expirado | Regenerar el secreto en Azure AD y volver a guardar desde la GUI |
| "Invoke-RestMethod: 403 Forbidden" | Falta consentimiento de administrador | En portal Azure, hacer clic en "Conceder consentimiento de administrador" |
| El correo no llega pero no hay error | Buzon remitente sin Exchange Online | Verificar que la cuenta tenga un buzon en Exchange Online |
| La tarea corre pero no genera archivos | Ruta incorrecta en la tarea | Verificar con `(Get-ScheduledTask -TaskName "AD Toolkit Daily Report").Actions` |
| La tarea no corre sin sesion abierta | LogonType: Interactive | Ejecutar `Register-DailyReportTask.ps1` para reconfigurar con contrasena almacenada |
| Error de parseo en PS4 con caracteres especiales | Archivo guardado sin BOM UTF-8 | Guardar el script con UTF-8 BOM (el repo incluye BOM en todos los .ps1) |

---

## Seguridad

- `config.json` esta en `.gitignore` - las credenciales nunca se suben al repositorio
- El secreto de cliente se almacena cifrado con **DPAPI de Windows** (vinculado al usuario y equipo local)
- El permiso `Mail.Send` es el unico concedido a la app - sin acceso de lectura al directorio
- Se recomienda usar un buzon compartido como remitente: alcance minimo, sin MFA, sin licencia
- Los scripts del repositorio se guardan con **UTF-8 BOM** para evitar problemas de encoding en PowerShell 4.0

---

## Licencia

MIT - libre para uso personal y comercial.
