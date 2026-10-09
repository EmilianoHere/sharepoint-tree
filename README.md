# SharePoint Online Storage Analyzer

Script **de solo lectura** en PowerShell 7 que analiza el uso de almacenamiento de
SharePoint Online para detectar qué **sitios, bibliotecas y carpetas** consumen más
espacio, cuánto se gasta en **versiones**, **papeleras** y **Preservation Hold Library**,
y qué sitios son candidatos a **Microsoft 365 Archive**.

> El script **no borra, no mueve ni modifica nada**. Solo lee métricas y exporta CSVs.

---

## Requisitos

- **PowerShell 7+** (`pwsh`).
- Módulo **PnP.PowerShell** (última versión):

  ```powershell
  Install-Module PnP.PowerShell -Scope CurrentUser
  ```

- (Opcional) **Microsoft.Online.SharePoint.PowerShell** si preferís el listado de
  sitios con `Get-SPOSite`. El script usa `Get-PnPTenantSite`, así que no es obligatorio.

---

## Autenticación: certificado (app-only) vs interactivo

Hay dos formas de autenticar, con un tradeoff importante:

| | **Certificado (app-only)** — RECOMENDADO | **Interactivo (delegado)** |
|---|---|---|
| Tipo | Permisos de **aplicación** | Permisos **delegados** (los del usuario) |
| Sitios que ve | **Todos** (con `Sites.Read.All` / `Sites.FullControl.All`) | Solo donde el usuario es miembro/propietario |
| Popups | **Ninguno** | Uno por conexión (y el script conecta **por sitio**) |
| Velocidad | Rápido | Lento |
| Ideal para | Auditar **todo el tenant** o un `TopN` alto | Una corrida puntual de pocos sitios |

**Por qué el interactivo falla o es lento:** el script analiza cada sitio con su
propia conexión. En modo interactivo eso abre un **popup de navegador por sitio** y,
como el login es **delegado**, solo alcanza los sitios donde la cuenta tiene acceso;
en los demás SharePoint devuelve **"Attempted to perform an unauthorized operation"**.
En una corrida real del usuario, 3 sitios tardaron **~50 minutos** por esta razón.
Con certificado (app-only) la app lee **todos** los sitios sin popups y mucho más rápido.

> Por eso, para auditorías de tenant completo, usá el **certificado** (Opción B del
> paso 2) y subí el `-TopN`.

---

## 1. Crear / configurar la app registration en Entra ID

### Caso A — app NUEVA (PnP crea app + certificado)

La forma más simple es dejar que PnP cree la app y el certificado por vos. Ejecutá
**una sola vez** (requiere rol de admin para dar consent):

```powershell
Register-PnPEntraIDApp `
  -ApplicationName "SPO-Storage-ReadOnly" `
  -Tenant "contoso.onmicrosoft.com" `
  -Interactive `
  -SharePointApplicationPermissions "Sites.Read.All"
```

Esto:

- Crea la app registration en Entra ID.
- Genera un **certificado autofirmado** y lo instala en tu store local (y en la app).
- Te devuelve el **ClientId** y el **Thumbprint** del certificado.

### Caso B — app EXISTENTE (generar certificado y subirlo a mano)

Si ya tenés la app registration (por ejemplo `ClientId 7dfd260e-...`) y solo querés
pasarla a modo app-only, generá un certificado nuevo con PnP y subí el `.cer` a la app:

```powershell
# 1) Genera un par certificado/clave. Devuelve el .pfx (privado) y el .cer (público)
#    y te muestra el Thumbprint. Guardá el .pfx fuera del repo (está en .gitignore).
New-PnPAzureCertificate `
  -CommonName "SPO-Storage-ReadOnly" `
  -OutPfx "./SPO-Storage-ReadOnly.pfx" `
  -OutCert "./SPO-Storage-ReadOnly.cer" `
  -ValidYears 2
```

Luego, en el portal de Entra ID:

- **App registrations → tu app (ClientId `7dfd260e-...`) → Certificates & secrets →
  Certificates → Upload certificate** y subí el **`.cer`** (nunca el `.pfx`).
- **API permissions**: agregá el permiso de **aplicación** de SharePoint
  (`Sites.Read.All` o `Sites.FullControl.All`) y **otorgá admin consent**.

Con eso ya podés conectar con `-Thumbprint` (el que mostró `New-PnPAzureCertificate`)
o con `-CertificatePath ./SPO-Storage-ReadOnly.pfx`.

### Permisos mínimos

| Objetivo | Permiso de aplicación (SharePoint) |
|---|---|
| Leer sitios y storage metrics | `Sites.Read.All` |
| Además leer/relevar papeleras de forma fiable en todos los sitios | `Sites.FullControl.All` |

`Sites.Read.All` alcanza para métricas de sitios, bibliotecas y carpetas. Si notás
errores de permisos al leer la papelera (`Get-PnPRecycleBinItem`) en sitios donde no
sos admin, volvé a registrar con `Sites.FullControl.All`:

```powershell
Register-PnPEntraIDApp `
  -ApplicationName "SPO-Storage-ReadOnly" `
  -Tenant "contoso.onmicrosoft.com" `
  -Interactive `
  -SharePointApplicationPermissions "Sites.FullControl.All"
```

> Aunque el permiso se llame "FullControl", el script **sigue siendo de solo lectura**:
> nunca invoca operaciones de escritura/borrado.

Después de crear la app, **otorgá admin consent** en el portal de Entra ID
(App registrations → tu app → API permissions → Grant admin consent) si no se otorgó
en el paso interactivo.

---

## 2. Ejecutar el script

### Opción A (RECOMENDADA) — certificado por thumbprint (app-only, tenant completo)

Para auditar todo el tenant: sin popups, lee todos los sitios y es rápido.

```powershell
./Analyze-SPOStorage.ps1 `
  -TenantName "contoso" `
  -ClientId "<TU_CLIENT_ID>" `
  -Thumbprint "<THUMBPRINT_DEL_CERTIFICADO>" `
  -TopN 50
```

### Opción B — certificado desde archivo .pfx (también app-only)

```powershell
$pwd = Read-Host "Password del PFX" -AsSecureString
./Analyze-SPOStorage.ps1 `
  -TenantName "contoso" `
  -ClientId "<TU_CLIENT_ID>" `
  -CertificatePath "./cert.pfx" `
  -CertificatePassword $pwd
```

### Opción C — login interactivo (solo para una corrida puntual de pocos sitios)

Cómodo pero **delegado**: abre un popup por sitio, solo ve los sitios donde la cuenta
tiene acceso y es lento. No recomendado para auditar todo el tenant (ver tabla arriba).

```powershell
./Analyze-SPOStorage.ps1 `
  -TenantName "contoso" `
  -ClientId "<TU_CLIENT_ID>" `
  -Interactive
```

> Precedencia: si pasás un método de certificado (`-Thumbprint`/`-CertificatePath`) y
> además `-Interactive`, el script prioriza el **app-only** (certificado); `-Interactive`
> queda como fallback cuando no hay certificado.

> Tip: para no escribir el ClientId/Thumbprint a mano cada vez, copiá
> `config.example.ps1` a `config.ps1` (ignorado por git) y cargá tus valores con
> dot-sourcing. Ver "Configuración" abajo.

---

## 3. Parámetros

| Parámetro | Default | Descripción |
|---|---|---|
| `-TenantName` | *(requerido)* | Nombre del tenant sin sufijo (`contoso`). |
| `-ClientId` | *(requerido)* | Application (client) ID de la app de Entra ID. |
| `-OutputPath` | `./salida` | Carpeta donde se escriben los CSV. |
| `-MaxDepth` | `3` | Profundidad máxima de carpetas a recorrer. |
| `-TopN` | `20` | Cantidad de sitios (los más pesados) a analizar en detalle. |
| `-MinFolderSizeGB` | `1.0` | Tamaño mínimo de carpeta (GB) para aparecer en `carpetas.csv`. |
| `-ArchiveMonths` | `12` | Umbral de inactividad para candidatos (se evalúan 6 y 12 meses). |
| `-Interactive` | — | Login interactivo en navegador. |
| `-Thumbprint` | — | Thumbprint del certificado instalado. |
| `-CertificatePath` | — | Ruta a un `.pfx`. |
| `-CertificatePassword` | — | Password del `.pfx` (SecureString). |

---

## 4. Salida

Se generan en `-OutputPath` (tamaños en **GB con 2 decimales**):

- **`sitios.csv`** — todos los sitios SPO (sin OneDrive) con uso y cuota, de mayor a menor.
- **`bibliotecas.csv`** — por biblioteca: tamaño total, versión actual, overhead por
  versiones, cantidad de archivos y configuración de versionado (`MajorVersionLimit`, etc.).
- **`carpetas.csv`** — carpetas recursivas (hasta `MaxDepth`) con `TotalSize`,
  `TotalFileStream`, `VersionOverhead`, `TotalFileCount` y `LastModified`.
- **`papeleras.csv`** — papelera de 1ra y 2da etapa + Preservation Hold Library por sitio.
- **`candidatos_archive.csv`** — sitios sin cambios en 6/12 meses.

Además, un **resumen en consola** con:

- Top 10 carpetas más pesadas.
- Top 10 carpetas por `VersionOverhead`.
- Recuperable estimado por versiones y por papeleras.
- Total de sitios candidatos a archivar y su tamaño.
- Costo mensual estimado: **Archive (US$0.05/GB)** vs **Extra Storage (US$0.20/GB)**.

---

## Notas sobre las métricas

- `Get-PnPFolderStorageMetric` devuelve métricas **agregadas** (incluyen subcarpetas):
  - `TotalSize` → incluye **todas las versiones**.
  - `TotalFileStreamSize` → solo la **versión actual**.
  - `VersionOverhead` = `TotalSize - TotalFileStreamSize` → espacio ocupado por versiones.
- Como las métricas son **agregadas** (la carpeta raíz de cada biblioteca ya incluye el
  overhead de todas sus subcarpetas), el **"Recuperable estimado por versiones"** del
  resumen suma **solo las filas raíz por biblioteca** (`Depth = 0`) para no contar el
  mismo overhead varias veces. Los dos "Top 10" de `carpetas.csv` siguen mostrando
  carpetas de cualquier profundidad: son un **drill-down informativo**, no un agregado.
- La **cuota del pool** mostrada es aproximada (suma de cuotas por sitio). Si conocés el
  valor real del pool de tu tenant (ej. 3.36 TB), tomalo del Admin Center de SharePoint
  (Settings → Storage) para comparar contra el uso total que reporta el script.
- `StorageUsageCurrent` viene en MB desde `Get-PnPTenantSite`; el script lo normaliza a bytes.

---

## Seguridad / repo público

- **No hay credenciales en el código.** El `ClientId` y el certificado se pasan por
  parámetro o por un `config.ps1` que está en `.gitignore`.
- Nunca commitees `.pfx`, `.cer`, `config.ps1` ni la carpeta `salida/` (todos están en
  `.gitignore`). A Entra ID se sube **solo el `.cer`** (clave pública); el `.pfx` (clave
  privada) se queda local.
- Usá certificados, no secrets de cliente en texto plano.
