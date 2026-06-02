# 🛡 AD Toolkit — Automatización de Active Directory

![PowerShell](https://img.shields.io/badge/PowerShell-4.0%2B-5391FE?style=flat&logo=powershell&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows%20Server-0078D6?style=flat&logo=windows&logoColor=white)
![Active Directory](https://img.shields.io/badge/Active%20Directory-0078D4?style=flat&logo=microsoft&logoColor=white)
![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph%20API-00A4EF?style=flat&logo=microsoft&logoColor=white)
![WinForms](https://img.shields.io/badge/UI-WinForms-512BD4?style=flat&logo=dotnet&logoColor=white)

Toolkit de PowerShell con interfaz gráfica WinForms para automatizar auditorías de Active Directory. Centraliza reportes de usuarios inactivos, backup de GPOs e inventario de equipos, con envío de reportes diarios por email vía Microsoft Graph API.

---

## 🎯 Problema que resuelve

En entornos con Active Directory, tareas de auditoría como detectar cuentas inactivas, respaldar políticas de grupo o inventariar equipos se realizan manualmente desde distintas consolas, sin registro centralizado ni automatización. Este toolkit unifica esas operaciones en una sola interfaz y las programa para ejecutarse de forma autónoma.

---

## ✅ Funcionalidades

### 📊 Dashboard
- Resumen en tiempo real de los 4 indicadores clave del dominio
- Log de ejecución integrado con código de colores por severidad
- Botón "Ejecutar Todo" para correr los 3 módulos en secuencia

### 👤 Usuarios Inactivos
- Detecta cuentas sin login en X días (configurable)
- Filtra por OU base y opcionalmente incluye cuentas deshabilitadas
- Visualización en grilla con exportación a CSV

### 🗂 Backup de GPOs
- Respalda todas las GPOs del dominio en carpeta con timestamp
- Rotación automática: conserva solo los N backups más recientes
- Log en tiempo real por cada GPO procesada (sin congelar la ventana)

### 💻 Inventario de Equipos
- Lista todos los equipos con OS, IP, días de inactividad y OU
- Filtro por sistema operativo (admite wildcards: `*Server*`, `*Windows 10*`)
- Exportación a CSV

### 📧 Reporte por Email (Microsoft Graph API)
- Reporte HTML completo con:
  - Usuarios inactivos
  - Contraseñas por vencer en los próximos 15 días (coloreadas por urgencia)
  - Usuarios con contraseña sin expiración
  - Equipos inactivos y equipos con contraseña fija
- Envío vía Microsoft Graph API (sin SMTP legacy, sin problemas de autenticación moderna)
- Asunto dinámico con resumen de los indicadores del día

### ⚙️ Configuración
- Panel visual para configurar umbrales, rutas y credenciales
- Registro de tarea en el Programador de Windows con un clic (ejecución diaria 07:00)

---

## 🛠 Tech Stack

| Componente | Detalle |
|---|---|
| Lenguaje | PowerShell 4.0+ |
| UI | WinForms (nativo, sin dependencias externas) |
| Acceso a AD | RSAT — módulos `ActiveDirectory` y `GroupPolicy` |
| API de email | Microsoft Graph API v1.0 (REST via `Invoke-RestMethod`) |
| Autenticación | OAuth 2.0 Client Credentials Flow |
| Configuración | JSON + cifrado DPAPI para el secreto de cliente |

---

## 🏗 Estructura del proyecto

```
AD-Toolkit/
├── AD-Toolkit-GUI.ps1          # Interfaz gráfica principal (WinForms)
├── Send-ADDailyReport.ps1      # Reporte headless para Task Scheduler
├── Get-InactiveUsers.ps1       # Script standalone: usuarios inactivos
├── Backup-GPOs.ps1             # Script standalone: backup de GPOs
├── Get-DomainComputers.ps1     # Script standalone: inventario de equipos
├── Launch-ADToolkit.bat        # Launcher con captura de errores
├── config.json                 # Configuracion local (excluido del repo)
└── config.example.json         # Plantilla de configuracion
```

---

## 🚀 Instalación y uso

### Requisitos

| Requisito | Detalle |
|---|---|
| Windows Server | 2012 R2 o superior |
| PowerShell | 4.0 o superior |
| RSAT | Módulos `ActiveDirectory` y `GroupPolicy` |
| Microsoft 365 | Buzón Exchange Online (puede ser buzón compartido, sin licencia paga) |

### Instalar RSAT (si no está disponible)

```powershell
# Windows Server
Install-WindowsFeature RSAT-AD-PowerShell
Install-WindowsFeature GPMC
```

### Configurar

```
copy config.example.json config.json
```

Editar `config.json` con los valores del entorno, o completarlo desde la GUI en la sección **Configuración**.

### Ejecutar

```powershell
# Con launcher (recomendado — muestra errores si algo falla)
Launch-ADToolkit.bat

# Directo
powershell.exe -ExecutionPolicy Bypass -File "AD-Toolkit-GUI.ps1"
```

---

## ⚙️ Configuración de Microsoft Graph API

Para habilitar el envío de reportes por email se necesita una **App Registration en Azure AD**:

1. Ir a **portal.azure.com → Microsoft Entra → Registros de aplicaciones → Nuevo registro**
2. Nombre: `AD Toolkit Mailer` · Tipo de cuenta: solo este directorio
3. Anotar el **Tenant ID** y el **Client ID**
4. **Certificados y secretos → Nuevo secreto de cliente** → copiar el valor
5. **Permisos de API → Microsoft Graph → Permisos de aplicación → `Mail.Send`**
6. Clic en **Conceder consentimiento de administrador**

Ingresar los tres valores en la GUI (sección Configuración) y hacer clic en **Guardar**.

> El buzón remitente puede ser un **buzón compartido** de Exchange Online: no consume licencia de usuario y el permiso `Mail.Send` de la app puede enviar desde él sin configuración adicional.

### Archivo de configuración

`config.json` (generado automáticamente por la GUI):

```json
{
  "DaysInactiveUsers": 90,
  "DaysInactivePC": 90,
  "SearchBase": "",
  "BackupRoot": "C:\\AD-Toolkit\\GPO-Backups",
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

> `config.json` está en `.gitignore` y nunca se sube al repositorio.

---

## 📅 Reporte diario (headless)

`Send-ADDailyReport.ps1` está diseñado para correr desde el Programador de Tareas sin interfaz gráfica.

```powershell
# Ejecutar manualmente
.\Send-ADDailyReport.ps1

# Solo generar archivos locales, sin enviar email
.\Send-ADDailyReport.ps1 -NoMail
```

Desde la GUI, el botón **Programar tarea** registra automáticamente la ejecución diaria a las 07:00.

---

## 📦 Scripts standalone

Los tres scripts de auditoría funcionan de forma independiente:

```powershell
# Usuarios sin login en los ultimos 60 dias
.\Get-InactiveUsers.ps1 -DaysInactive 60 -OutputPath "C:\Reports\usuarios.csv"

# Backup conservando los ultimos 5, con reporte HTML
.\Backup-GPOs.ps1 -BackupRoot "D:\GPO-Backups" -MaxBackups 5 -GenerateReport

# Inventario filtrando solo servidores
.\Get-DomainComputers.ps1 -OSFilter "*Server*" -GenerateReport
```

---

## ❗ Errores comunes

| Error | Causa | Solución |
|---|---|---|
| "No se reconoce el módulo ActiveDirectory" | RSAT no instalado | `Install-WindowsFeature RSAT-AD-PowerShell` |
| "Invoke-RestMethod: 401 Unauthorized" | Secreto de cliente expirado o incorrecto | Regenerar el secreto en Azure AD y volver a guardar desde la GUI |
| "Invoke-RestMethod: 403 Forbidden" | Falta consentimiento de administrador en el permiso `Mail.Send` | Volver al portal Azure y hacer clic en "Conceder consentimiento de administrador" |
| El correo no llega pero no hay error | El buzón remitente no tiene Exchange Online activo | Verificar que la cuenta tenga un buzón (usuario o compartido) en Exchange Online |
| La GUI se cierra instantáneamente | Error de sintaxis por encoding incorrecto | Guardar el script con UTF-8 BOM |

---

## 🔒 Seguridad

- `config.json` está en `.gitignore` — las credenciales nunca se suben al repositorio
- El secreto de cliente se almacena cifrado con **DPAPI de Windows** (vinculado al usuario y equipo local)
- El permiso `Mail.Send` es el único concedido a la app — sin acceso de lectura al directorio
- Se recomienda usar un buzón compartido como remitente: alcance mínimo, sin MFA, sin licencia

---

## 📋 Licencia

MIT — libre para uso personal y comercial.
