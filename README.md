# Microsoft 365 Administration Scripts

Colección de herramientas interactivas para diagnóstico, análisis y mantenimiento de OneDrive for Business y SharePoint Online.

## Herramientas

| Script | Función |
| --- | --- |
| `OneDrive-Path-Analyzer.zsh` | Analiza rutas largas de OneDrive/SharePoint en macOS y genera un CSV con avisos o hallazgos críticos. Requiere Zsh. |
| `OneDrive-Path-Analyzer.ps1` | Analiza rutas largas de OneDrive/SharePoint en Windows. |
| `M365-Version-History-Cleaner.ps1` | Analiza y limpia versiones históricas de archivos en SharePoint Online y OneDrive for Business. Incluye simulación, revalidación y auditoría. |
| `M365-Legacy-UserId-Cleaner.ps1` | Detecta y repara entradas de usuarios legacy con User ID mismatch en `UserInfoList`. |
| `M365-Recycle-Bin-and-Hold-Cleaner.ps1` | Gestiona la papelera de reciclaje de SharePoint/OneDrive y la Preservation Hold Library. |
| `M365-Permissions-Scope-Manager.ps1` | Cuenta Unique Permission Scopes y permite restablecer la herencia de permisos. |

## Requisitos

- Windows: PowerShell 7.4 o superior.
- macOS: Zsh y las utilidades estándar del sistema.
- Herramientas de Microsoft 365: módulo `PnP.PowerShell` 3.2 o superior.
- Una cuenta con permisos suficientes en el tenant y, según la operación, permisos de administrador de colección de sitios.

## Uso

En Windows, abrir PowerShell 7.4+ en esta carpeta y ejecutar el script deseado:

```powershell
./M365-Version-History-Cleaner.ps1
./M365-Legacy-UserId-Cleaner.ps1
./M365-Recycle-Bin-and-Hold-Cleaner.ps1
./M365-Permissions-Scope-Manager.ps1
./OneDrive-Path-Analyzer.ps1
```

En macOS:

```zsh
chmod +x ./OneDrive-Path-Analyzer.zsh
./OneDrive-Path-Analyzer.zsh
```

Las herramientas destructivas comienzan en modo simulación cuando es posible. Revisa siempre el destino, el resumen y los reportes antes de confirmar una operación real.

## Reportes y configuración

Los scripts de Microsoft 365 guardan configuración y reportes localmente en la carpeta de usuario correspondiente. Los analizadores de rutas generan CSV con los hallazgos detectados.

## Aviso

Estas herramientas pueden modificar permisos, usuarios legacy, versiones, elementos de papelera o contenido retenido. Úsalas primero en modo simulación y valida los resultados en un entorno controlado antes de ejecutarlas en producción.
