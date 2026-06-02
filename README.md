# AD Toolkit — PowerShell para Active Directory

Conjunto de herramientas PowerShell para automatizar auditorías de Active Directory en entornos Windows Server. Incluye una interfaz gráfica WinForms y envío de reportes por email vía Microsoft Graph API.

---

## Características

- **Usuarios inactivos** — detecta cuentas sin login en X días, con exportación CSV
- **Backup de GPOs** — respalda todas las políticas de grupo con rotación automática
- **Inventario de equipos** — lista todos los equipos del dominio con OS, IP y estado
- **Contraseñas por vencer** — alerta sobre contraseñas que expiran en los próximos 15 días
- **Cuentas sin expiración** — identifica usuarios y equipos con contraseña fija (riesgo de seguridad)
- **Reporte HTML por email** — envío automático diario vía Microsoft Graph API (sin SMTP legacy)
- **GUI WinForms** — interfaz visual que centraliza todas las operaciones
- **Tarea programada** — registro en el Programador de Windows con un clic

## Capturas

> *(proximamente)*

---

## Requisitos

| Componente | Versión mínima |
|---|---|
| Windows Server | 2012 R2 |
| PowerShell | 4.0 |
| Módulo RSAT | ActiveDirectory, GroupPolicy |
| Microsoft 365 | Cuenta con buzón Exchange Online |

---

## Instalación

### 1. Clonar el repositorio

```
git clone https://github.com/tu-usuario/ad-toolkit.git
```

### 2. Copiar el archivo de configuración

```
copy config.example.json config.json
```

Editar `config.json` con los valores del entorno (o usar la GUI — sección Configuración).

### 3. Registrar la App en Azure AD (para envío de emails)

1. Ingresar a [portal.azure.com](https://portal.azure.com)
2. **Microsoft Entra → Registros de aplicaciones → Nuevo registro**
   - Nombre: `AD Toolkit Mailer`
   - Tipo de cuenta: solo este directorio
3. Anotar el **Tenant ID** y el **Client ID**
4. **Certificados y secretos → Nuevo secreto de cliente** → copiar el valor
5. **Permisos de API → Agregar → Microsoft Graph → Permisos de aplicación → `Mail.Send`**
6. Hacer clic en **Conceder consentimiento de administrador**

> El buzón remitente no necesita licencia paga: se puede usar un **buzón compartido** (Shared Mailbox) de Exchange Online.

### 4. Ejecutar

Doble clic en `Launch-ADToolkit.bat` o directamente:

```powershell
powershell.exe -ExecutionPolicy Bypass -File "AD-Toolkit-GUI.ps1"
```

---

## Estructura del proyecto

```
AD-Toolkit/
├── AD-Toolkit-GUI.ps1          # Interfaz gráfica principal (WinForms)
├── Send-ADDailyReport.ps1      # Reporte diario headless (Task Scheduler)
├── Get-InactiveUsers.ps1       # Script standalone: usuarios inactivos
├── Backup-GPOs.ps1             # Script standalone: backup de GPOs
├── Get-DomainComputers.ps1     # Script standalone: inventario de equipos
├── Launch-ADToolkit.bat        # Launcher con manejo de errores
├── config.json                 # Configuracion generada (NO subir al repo)
├── config.example.json         # Plantilla de configuracion
└── .gitignore
```

---

## Uso del reporte diario (headless)

El script `Send-ADDailyReport.ps1` está diseñado para ejecutarse sin interfaz gráfica desde el Programador de Tareas de Windows.

Desde la GUI, sección **Configuración → Programar tarea**, se registra automáticamente la tarea diaria a las 07:00.

Para ejecutarlo manualmente:

```powershell
.\Send-ADDailyReport.ps1

# Solo generar archivos, sin enviar email
.\Send-ADDailyReport.ps1 -NoMail
```

---

## Scripts standalone

Los tres scripts de auditoría funcionan de forma independiente, sin necesidad de la GUI:

```powershell
# Usuarios sin login en los ultimos 60 dias
.\Get-InactiveUsers.ps1 -DaysInactive 60 -OutputPath "C:\Reports\usuarios.csv"

# Backup de GPOs conservando los ultimos 5
.\Backup-GPOs.ps1 -BackupRoot "D:\GPO-Backups" -MaxBackups 5 -GenerateReport

# Inventario filtrando solo servidores
.\Get-DomainComputers.ps1 -OSFilter "*Server*" -GenerateReport
```

---

## Seguridad

- `config.json` está en `.gitignore` — nunca se sube al repositorio
- El secreto de cliente de Azure AD se almacena cifrado con DPAPI de Windows
- El permiso `Mail.Send` es el único concedido a la app — no tiene acceso de lectura al directorio
- Se recomienda usar un buzón compartido como remitente (sin licencia, sin MFA, alcance mínimo)

---

## Compatibilidad

Desarrollado y probado en **Windows Server 2012 R2 / PowerShell 4.0**.  
Compatible con PowerShell 5.1 y superior sin modificaciones.

---

## Licencia

MIT — libre para uso personal y comercial.
