#Requires -Version 7.0
<#
.SYNOPSIS
    Analiza (solo lectura) el uso de almacenamiento de SharePoint Online para
    detectar qué sitios, bibliotecas y carpetas ocupan más espacio.

.DESCRIPTION
    El tenant excede la cuota del pool de SharePoint. Este script releva el uso
    a nivel tenant, baja a los Top N sitios más pesados y detalla bibliotecas,
    carpetas recursivas (con sobrecarga por versiones), Preservation Hold Library,
    papeleras de 1ra y 2da etapa y configuración de versionado. Además marca
    sitios candidatos a Microsoft 365 Archive por inactividad.

    IMPORTANTE: el script es ESTRICTAMENTE de solo lectura. No borra, mueve ni
    modifica absolutamente nada en el tenant.

.NOTES
    Requiere PnP.PowerShell (>= 2.x) y, opcionalmente,
    Microsoft.Online.SharePoint.PowerShell para el listado de sitios a nivel tenant.
    Autenticación mediante app registration de Entra ID (ClientId) con certificado
    o login interactivo. Ver README.md para la creación de la app y permisos.

.EXAMPLE
    ./Analyze-SPOStorage.ps1 -TenantName contoso -ClientId "00000000-0000-0000-0000-000000000000" -Interactive

.EXAMPLE
    ./Analyze-SPOStorage.ps1 -TenantName contoso -ClientId "..." -Thumbprint "ABC123..." -OutputPath ./salida -TopN 30
#>

[CmdletBinding()]
param(
    # Nombre del tenant SIN sufijo (ej: "contoso" para contoso.onmicrosoft.com).
    [Parameter(Mandatory = $true)]
    [string]$TenantName,

    # ClientId (Application ID) de la app registration de Entra ID.
    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    # Carpeta de salida para los CSV. Se crea si no existe.
    [string]$OutputPath = "./salida",

    # Profundidad máxima de carpetas a recorrer dentro de cada biblioteca.
    [int]$MaxDepth = 3,

    # Cantidad de sitios (los más pesados) a analizar en detalle.
    [int]$TopN = 20,

    # Tamaño mínimo de carpeta (en GB) para que aparezca en carpetas.csv.
    [double]$MinFolderSizeGB = 1.0,

    # Umbral de inactividad (meses) para candidatos a archive. Se evalúan 6 y 12.
    [int]$ArchiveMonths = 12,

    # --- Autenticación: elegir UNA de estas opciones ---

    # Login interactivo en el navegador (cómodo para una corrida puntual).
    [switch]$Interactive,

    # Thumbprint del certificado instalado en el store del usuario/máquina.
    [string]$Thumbprint,

    # Ruta a un archivo .pfx (alternativa al Thumbprint).
    [string]$CertificatePath,

    # Password del .pfx como SecureString (si CertificatePath lo requiere).
    [securestring]$CertificatePassword
)

# Cortar ante errores no manejados dentro de nuestras funciones.
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Constantes y helpers de formato
# ---------------------------------------------------------------------------

# Factor para pasar de bytes a GB.
$script:BytesPerGB = 1GB

# Costos de referencia (USD por GB/mes). Ajustar si cambian las tarifas.
$script:CostArchivePerGB = 0.05   # Microsoft 365 Archive
$script:CostExtraStoragePerGB = 0.20  # Office 365 Extra File Storage

# Convierte bytes a GB con 2 decimales.
function ConvertTo-GB {
    param([double]$Bytes)
    return [math]::Round($Bytes / $script:BytesPerGB, 2)
}

# Escribe un mensaje con timestamp y nivel.
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'OK')]
        [string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString('HH:mm:ss')
    $color = switch ($Level) {
        'INFO'  { 'Gray' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        'OK'    { 'Green' }
    }
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $color
}

# ---------------------------------------------------------------------------
# Reintentos con backoff para throttling (HTTP 429 / 503)
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Ejecuta un scriptblock reintentando ante throttling con backoff exponencial.
.DESCRIPTION
    SharePoint Online devuelve 429 (Too Many Requests) o 503 (Server Busy) cuando
    hay throttling, y suele incluir el header "Retry-After". Respetamos ese valor
    si está disponible; si no, aplicamos backoff exponencial con jitter.
#>
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        # Argumentos que se pasan al scriptblock (accesibles como $args[0], $args[1], ...).
        # Necesario porque el scriptblock se ejecuta en el scope de esta función y NO ve
        # las variables locales del llamador.
        [object[]]$ArgumentList = @(),

        [int]$MaxRetries = 5,

        [int]$BaseDelaySeconds = 5,

        [string]$Context = 'operación'
    )

    for ($attempt = 1; $attempt -le ($MaxRetries + 1); $attempt++) {
        try {
            return & $ScriptBlock @ArgumentList
        }
        catch {
            $ex = $_.Exception
            $msg = $ex.Message
            # Intentar extraer el código de estado HTTP de la excepción.
            $statusCode = $null
            if ($ex.Response -and $ex.Response.StatusCode) {
                $statusCode = [int]$ex.Response.StatusCode
            }
            $isThrottle = ($statusCode -in 429, 503) -or
                          ($msg -match '429|503|throttl|Too Many Requests|Server is busy|Server Busy')

            if (-not $isThrottle -or $attempt -gt $MaxRetries) {
                # No es throttling, o agotamos reintentos: propagar.
                throw
            }

            # Respetar Retry-After si viene en la respuesta.
            $retryAfter = $null
            if ($ex.Response -and $ex.Response.Headers) {
                try { $retryAfter = $ex.Response.Headers['Retry-After'] } catch { }
            }

            $parsedRetry = 0
            if ($retryAfter -and [int]::TryParse([string]$retryAfter, [ref]$parsedRetry)) {
                $delay = $parsedRetry
            }
            else {
                # Backoff exponencial con jitter aleatorio para evitar sincronización.
                $delay = [math]::Min([math]::Pow(2, $attempt) * $BaseDelaySeconds, 300)
                $delay = [int]$delay + (Get-Random -Minimum 0 -Maximum 5)
            }

            Write-Log "Throttling en $Context (intento $attempt/$MaxRetries). Esperando $delay s..." 'WARN'
            Start-Sleep -Seconds $delay
        }
    }
}

# ---------------------------------------------------------------------------
# Carga de módulos
# ---------------------------------------------------------------------------

function Initialize-Modules {
    Write-Log "Verificando módulos requeridos..."

    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        throw "Falta el módulo PnP.PowerShell. Instalarlo con: Install-Module PnP.PowerShell -Scope CurrentUser"
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
    Write-Log "PnP.PowerShell cargado." 'OK'
}

# ---------------------------------------------------------------------------
# Conexión
# ---------------------------------------------------------------------------

# URLs base del tenant.
function Get-TenantUrls {
    param([string]$Tenant)
    return [pscustomobject]@{
        Admin = "https://$Tenant-admin.sharepoint.com"
        Root  = "https://$Tenant.sharepoint.com"
    }
}

<#
.SYNOPSIS
    Conecta a una URL de SharePoint con PnP usando el método elegido.
#>
function Connect-SPO {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$Tenant,
        [Parameter(Mandatory = $true)][string]$ClientId
    )

    $connectParams = @{
        Url      = $Url
        ClientId = $ClientId
    }

    # Precedencia: el modo app-only (certificado por Thumbprint o CertificatePath)
    # tiene prioridad SOBRE Interactive aunque ambos estén presentes. Motivo: en una
    # corrida de tenant completo Connect-SPO se invoca una vez por sitio (ver Main);
    # con Interactive eso abriría un popup de navegador por cada sitio y el login
    # delegado solo ve los sitios donde el usuario es miembro (de ahí los errores
    # "unauthorized"). App-only (Sites.Read.All / Sites.FullControl.All) lee todos
    # los sitios sin popups y es mucho más rápido. Interactive queda como fallback
    # cuando no se proveyó ningún certificado. Las ramas de certificado NO setean
    # Interactive. Todos los valores salen de los parámetros (sin hardcodear nada).
    if ($Thumbprint) {
        $connectParams.Thumbprint = $Thumbprint
        $connectParams.Tenant     = "$Tenant.onmicrosoft.com"
    }
    elseif ($CertificatePath) {
        $connectParams.CertificatePath = $CertificatePath
        $connectParams.Tenant          = "$Tenant.onmicrosoft.com"
        if ($CertificatePassword) {
            $connectParams.CertificatePassword = $CertificatePassword
        }
    }
    elseif ($Interactive) {
        $connectParams.Interactive = $true
        # Dirigir el login al directorio concreto (útil si la cuenta pertenece a varios tenants).
        $connectParams.Tenant = "$Tenant.onmicrosoft.com"
    }
    else {
        throw "Debe especificar -Thumbprint, -CertificatePath o -Interactive para autenticar."
    }

    Invoke-WithRetry -Context "conexión a $Url" -ArgumentList @($connectParams) -ScriptBlock {
        param($cp)
        Connect-PnPOnline @cp -ReturnConnection
    }
}

# ---------------------------------------------------------------------------
# Nivel tenant: listado de sitios
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Obtiene todos los sitios del tenant con sus métricas de almacenamiento.
.DESCRIPTION
    Usa Get-PnPTenantSite (requiere conexión al endpoint -admin). Excluye los
    sitios de OneDrive (-my.sharepoint.com) porque no cuentan para el pool de
    SharePoint. Ordena de mayor a menor por StorageUsageCurrent (que viene en MB).
#>
function Get-TenantSites {
    param([object]$AdminConnection)

    Write-Log "Listando sitios del tenant..."

    $sites = Invoke-WithRetry -Context "Get-PnPTenantSite" -ArgumentList @($AdminConnection) -ScriptBlock {
        param($conn)
        Get-PnPTenantSite -Connection $conn -Detailed
    }

    # Excluir OneDrive y filtrar nulos.
    $filtered = $sites | Where-Object {
        $_.Url -and ($_.Url -notmatch '-my\.sharepoint\.com')
    }

    # StorageUsageCurrent viene en MB. Lo normalizamos a bytes para el resto del script.
    $result = foreach ($s in $filtered) {
        [pscustomobject]@{
            Title                     = $s.Title
            Url                       = $s.Url
            Template                  = $s.Template
            StorageUsageCurrentMB     = [double]$s.StorageUsageCurrent
            StorageUsageCurrentBytes  = [double]$s.StorageUsageCurrent * 1MB
            StorageQuotaMB            = [double]$s.StorageMaximumLevel
            LastContentModifiedDate   = $s.LastContentModifiedDate
            LockState                 = $s.LockState
            Owner                     = $s.Owner
        }
    }

    return $result | Sort-Object StorageUsageCurrentBytes -Descending
}

# ---------------------------------------------------------------------------
# Detalle por sitio: bibliotecas
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Devuelve las bibliotecas de documentos de un sitio con tamaño y versionado.
.DESCRIPTION
    BaseTemplate 101 = Document Library. Lee también la config de versionado:
    EnableVersioning, EnableMinorVersions, MajorVersionLimit y, si está disponible,
    el versionado "intelligent"/automático (VersionPolicies, propiedad reciente).
#>
function Get-SiteLibraries {
    param(
        [Parameter(Mandatory = $true)][object]$Connection,
        [Parameter(Mandatory = $true)][string]$SiteUrl
    )

    $lists = Invoke-WithRetry -Context "Get-PnPList $SiteUrl" -ArgumentList @($Connection) -ScriptBlock {
        param($conn)
        Get-PnPList -Connection $conn -Includes `
            RootFolder.ServerRelativeUrl, EnableVersioning, EnableMinorVersions, `
            MajorVersionLimit, MajorWithMinorVersionsLimit, IsSystemList, Hidden
    }

    # Nos quedamos con bibliotecas de documentos visibles.
    $libs = $lists | Where-Object {
        $_.BaseTemplate -eq 101 -and -not $_.Hidden
    }

    # Path absoluto del sitio (ej: "/sites/foo") para convertir URLs server-relative
    # en site-relative, que es lo que espera Get-PnPFolderStorageMetric.
    $siteAbsPath = (New-Object System.Uri $SiteUrl).AbsolutePath
    $siteAbsPathRegex = [regex]::Escape($siteAbsPath)

    $result = foreach ($lib in $libs) {
        # El metric de almacenamiento se obtiene sobre el root folder de la lista.
        $metric = $null
        try {
            $libSiteRel = $lib.RootFolder.ServerRelativeUrl -replace "^$siteAbsPathRegex", ''
            $metric = Invoke-WithRetry -Context "StorageMetric $($lib.Title)" -ArgumentList @($Connection, $libSiteRel) -ScriptBlock {
                param($conn, $rel)
                Get-PnPFolderStorageMetric -Connection $conn -FolderSiteRelativeUrl $rel
            }
        }
        catch {
            Write-Log "No se pudo leer StorageMetric de '$($lib.Title)': $($_.Exception.Message)" 'WARN'
        }

        [pscustomobject]@{
            SiteUrl              = $SiteUrl
            Library              = $lib.Title
            ServerRelativeUrl    = $lib.RootFolder.ServerRelativeUrl
            TotalSizeBytes       = if ($metric) { [double]$metric.TotalSize } else { 0 }
            TotalFileStreamBytes = if ($metric) { [double]$metric.TotalFileStreamSize } else { 0 }
            TotalFileCount       = if ($metric) { [long]$metric.TotalFileCount } else { 0 }
            LastModified         = if ($metric) { $metric.LastModified } else { $null }
            EnableVersioning     = $lib.EnableVersioning
            EnableMinorVersions  = $lib.EnableMinorVersions
            MajorVersionLimit    = $lib.MajorVersionLimit
            MinorVersionLimit    = $lib.MajorWithMinorVersionsLimit
        }
    }

    return $result
}

# ---------------------------------------------------------------------------
# Detalle por sitio: carpetas recursivas con storage metrics
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Recorre carpetas recursivamente hasta MaxDepth leyendo storage metrics.
.DESCRIPTION
    Get-PnPFolderStorageMetric devuelve métricas AGREGADAS de la carpeta (incluye
    subcarpetas), por eso:
      - TotalSize           -> incluye todas las versiones de los archivos.
      - TotalFileStreamSize -> solo la versión actual.
      - VersionOverhead     -> TotalSize - TotalFileStreamSize (sobrecarga por versiones).
    Solo reporta carpetas que superen MinFolderSizeGB para no inundar el CSV.
#>
function Get-FolderMetricsRecursive {
    param(
        [Parameter(Mandatory = $true)][object]$Connection,
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$SiteRelativeUrl,
        [Parameter(Mandatory = $true)][string]$LibraryTitle,
        [int]$Depth = 0,
        [int]$MaxDepth = 3,
        [double]$MinFolderSizeGB = 1.0
    )

    $results = [System.Collections.Generic.List[object]]::new()

    if ($Depth -gt $MaxDepth) { return $results }

    # Métrica agregada de la carpeta actual.
    $metric = $null
    try {
        $metric = Invoke-WithRetry -Context "StorageMetric carpeta $SiteRelativeUrl" -ArgumentList @($Connection, $SiteRelativeUrl) -ScriptBlock {
            param($conn, $rel)
            Get-PnPFolderStorageMetric -Connection $conn -FolderSiteRelativeUrl $rel
        }
    }
    catch {
        Write-Log "Error leyendo métrica de carpeta '$SiteRelativeUrl': $($_.Exception.Message)" 'WARN'
        return $results
    }

    if ($metric) {
        $totalBytes = [double]$metric.TotalSize
        $streamBytes = [double]$metric.TotalFileStreamSize
        $overhead = $totalBytes - $streamBytes

        # Reportar solo si supera el umbral mínimo.
        if ($totalBytes -ge ($MinFolderSizeGB * $script:BytesPerGB)) {
            $results.Add([pscustomobject]@{
                SiteUrl              = $SiteUrl
                Library              = $LibraryTitle
                FolderPath           = $SiteRelativeUrl
                Depth                = $Depth
                TotalSizeBytes       = $totalBytes
                TotalFileStreamBytes = $streamBytes
                VersionOverheadBytes = $overhead
                TotalFileCount       = [long]$metric.TotalFileCount
                LastModified         = $metric.LastModified
            })
        }
    }

    # Descender a subcarpetas si todavía no llegamos al límite de profundidad.
    if ($Depth -lt $MaxDepth) {
        $subFolders = @()
        try {
            $subFolders = Invoke-WithRetry -Context "listar subcarpetas $SiteRelativeUrl" -ArgumentList @($Connection, $SiteRelativeUrl) -ScriptBlock {
                param($conn, $rel)
                Get-PnPFolderItem -Connection $conn -FolderSiteRelativeUrl $rel -ItemType Folder
            }
        }
        catch {
            Write-Log "No se pudieron listar subcarpetas de '$SiteRelativeUrl': $($_.Exception.Message)" 'WARN'
        }

        foreach ($sub in $subFolders) {
            # Ignorar carpetas de sistema típicas.
            if ($sub.Name -in 'Forms', '_catalogs', '_private') { continue }

            $childRel = ($SiteRelativeUrl.TrimEnd('/')) + '/' + $sub.Name
            $childResults = Get-FolderMetricsRecursive -Connection $Connection -SiteUrl $SiteUrl `
                -SiteRelativeUrl $childRel -LibraryTitle $LibraryTitle `
                -Depth ($Depth + 1) -MaxDepth $MaxDepth -MinFolderSizeGB $MinFolderSizeGB
            foreach ($cr in $childResults) { $results.Add($cr) }
        }
    }

    return $results
}

# ---------------------------------------------------------------------------
# Preservation Hold Library
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Devuelve el tamaño de la Preservation Hold Library si existe en el sitio.
.DESCRIPTION
    La PHL retiene copias de contenido bajo retención/eDiscovery y suele ser una
    fuente oculta de consumo. Es una lista oculta llamada "Preservation Hold Library".
#>
function Get-PreservationHoldLibrary {
    param(
        [Parameter(Mandatory = $true)][object]$Connection,
        [Parameter(Mandatory = $true)][string]$SiteUrl
    )

    try {
        $phl = Invoke-WithRetry -Context "PHL $SiteUrl" -ArgumentList @($Connection) -ScriptBlock {
            param($conn)
            Get-PnPList -Connection $conn -Includes RootFolder.ServerRelativeUrl |
                Where-Object { $_.Title -eq 'Preservation Hold Library' }
        }

        if (-not $phl) { return $null }

        $siteAbsPath = (New-Object System.Uri $SiteUrl).AbsolutePath
        $siteAbsPathRegex = [regex]::Escape($siteAbsPath)
        $siteRel = $phl.RootFolder.ServerRelativeUrl -replace "^$siteAbsPathRegex", ''
        $metric = Invoke-WithRetry -Context "StorageMetric PHL" -ArgumentList @($Connection, $siteRel) -ScriptBlock {
            param($conn, $rel)
            Get-PnPFolderStorageMetric -Connection $conn -FolderSiteRelativeUrl $rel
        }

        return [pscustomobject]@{
            SiteUrl        = $SiteUrl
            Library        = 'Preservation Hold Library'
            TotalSizeBytes = [double]$metric.TotalSize
            TotalFileCount = [long]$metric.TotalFileCount
        }
    }
    catch {
        Write-Log "Error leyendo PHL de '$SiteUrl': $($_.Exception.Message)" 'WARN'
        return $null
    }
}

# ---------------------------------------------------------------------------
# Papeleras (1ra y 2da etapa)
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Suma el tamaño de la papelera de reciclaje de 1ra y 2da etapa de un sitio.
.DESCRIPTION
    Get-PnPRecycleBinItem devuelve ítems con ItemState:
      - FirstStageRecycleBin  (papelera del usuario)
      - SecondStageRecycleBin (papelera del administrador del sitio)
    Sumamos la propiedad Size (bytes) de cada ítem. En papeleras grandes puede
    tardar; por eso se envuelve en reintentos.
#>
function Get-RecycleBinSize {
    param(
        [Parameter(Mandatory = $true)][object]$Connection,
        [Parameter(Mandatory = $true)][string]$SiteUrl
    )

    try {
        $items = Invoke-WithRetry -Context "RecycleBin $SiteUrl" -ArgumentList @($Connection) -ScriptBlock {
            param($conn)
            Get-PnPRecycleBinItem -Connection $conn
        }

        $first  = $items | Where-Object { $_.ItemState -eq 'FirstStageRecycleBin' }
        $second = $items | Where-Object { $_.ItemState -eq 'SecondStageRecycleBin' }

        $firstBytes  = ($first  | Measure-Object -Property Size -Sum).Sum
        $secondBytes = ($second | Measure-Object -Property Size -Sum).Sum

        return [pscustomobject]@{
            SiteUrl           = $SiteUrl
            FirstStageBytes   = [double]($firstBytes  ?? 0)
            FirstStageCount   = @($first).Count
            SecondStageBytes  = [double]($secondBytes ?? 0)
            SecondStageCount  = @($second).Count
            TotalBytes        = [double](($firstBytes ?? 0) + ($secondBytes ?? 0))
        }
    }
    catch {
        Write-Log "Error leyendo papelera de '$SiteUrl': $($_.Exception.Message)" 'WARN'
        return [pscustomobject]@{
            SiteUrl = $SiteUrl; FirstStageBytes = 0; FirstStageCount = 0
            SecondStageBytes = 0; SecondStageCount = 0; TotalBytes = 0
        }
    }
}

# ---------------------------------------------------------------------------
# Candidatos a Microsoft 365 Archive
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Marca sitios sin modificaciones recientes como candidatos a archivar.
.DESCRIPTION
    Evalúa LastContentModifiedDate contra umbrales de 6 y 12 meses. Un sitio sin
    cambios por mucho tiempo es buen candidato a Microsoft 365 Archive (contenido
    accesible pero facturado a tarifa de archivo, mucho más barato que Extra Storage).
#>
function Get-ArchiveCandidates {
    param(
        [Parameter(Mandatory = $true)][object[]]$Sites
    )

    $now = Get-Date
    # Diagnóstico de calidad de dato: cuántos sitios traen una fecha de última
    # modificación USABLE. LastContentModifiedDate de Get-PnPTenantSite la actualiza
    # una tarea programada del servicio y puede venir vacía o en DateTime.MinValue
    # (01/01/0001) para muchos sitios; si eso pasa, el conteo de candidatos sale 0
    # no porque no haya sitios inactivos, sino porque falta el dato. Contamos y
    # logueamos para que una corrida futura detecte el problema en vez de callarlo.
    $usableDates = 0
    $totalSites  = @($Sites).Count

    $results = foreach ($s in $Sites) {
        $lastMod = $s.LastContentModifiedDate

        # Saltear fechas inutilizables: $null o DateTime.MinValue (dato ausente).
        if (-not $lastMod -or ($lastMod -is [datetime] -and $lastMod -eq [datetime]::MinValue)) {
            continue
        }
        $usableDates++

        $monthsIdle = [math]::Round((New-TimeSpan -Start $lastMod -End $now).TotalDays / 30.4, 1)
        $inactive6  = $lastMod -lt $now.AddMonths(-6)
        $inactive12 = $lastMod -lt $now.AddMonths(-12)

        if ($inactive6) {
            [pscustomobject]@{
                Title                   = $s.Title
                Url                     = $s.Url
                StorageUsageBytes       = $s.StorageUsageCurrentBytes
                LastContentModifiedDate = $lastMod
                MonthsIdle              = $monthsIdle
                Inactive6Months         = $inactive6
                Inactive12Months        = $inactive12
            }
        }
    }

    Write-Log "Sitios con fecha de última modificación usable: $usableDates/$totalSites" 'INFO'
    if ($usableDates -eq 0 -and $totalSites -gt 0) {
        Write-Log "Ningún sitio trajo fecha usable: el conteo de candidatos será 0 por falta de dato, no por ausencia de sitios inactivos." 'WARN'
    }

    return $results | Sort-Object StorageUsageBytes -Descending
}

# ---------------------------------------------------------------------------
# Exportación a CSV
# ---------------------------------------------------------------------------

# Exporta sitios.csv (nivel tenant).
function Export-SitesCsv {
    param([object[]]$Sites, [string]$Path)
    $rows = $Sites | Select-Object `
        Title, Url, Template, Owner, LockState, `
        @{ N = 'StorageUsageGB'; E = { ConvertTo-GB $_.StorageUsageCurrentBytes } }, `
        @{ N = 'StorageQuotaGB'; E = { [math]::Round($_.StorageQuotaMB / 1024, 2) } }, `
        LastContentModifiedDate
    $rows | Export-Csv -Path (Join-Path $Path 'sitios.csv') -Encoding UTF8
}

# Exporta bibliotecas.csv.
function Export-LibrariesCsv {
    param([object[]]$Libraries, [string]$Path)
    $rows = $Libraries | Select-Object `
        SiteUrl, Library, `
        @{ N = 'TotalSizeGB'; E = { ConvertTo-GB $_.TotalSizeBytes } }, `
        @{ N = 'TotalFileStreamGB'; E = { ConvertTo-GB $_.TotalFileStreamBytes } }, `
        @{ N = 'VersionOverheadGB'; E = { ConvertTo-GB ($_.TotalSizeBytes - $_.TotalFileStreamBytes) } }, `
        TotalFileCount, LastModified, `
        EnableVersioning, EnableMinorVersions, MajorVersionLimit, MinorVersionLimit
    $rows | Export-Csv -Path (Join-Path $Path 'bibliotecas.csv') -Encoding UTF8
}

# Exporta carpetas.csv.
function Export-FoldersCsv {
    param([object[]]$Folders, [string]$Path)
    $rows = $Folders | Sort-Object TotalSizeBytes -Descending | Select-Object `
        SiteUrl, Library, FolderPath, Depth, `
        @{ N = 'TotalSizeGB'; E = { ConvertTo-GB $_.TotalSizeBytes } }, `
        @{ N = 'TotalFileStreamGB'; E = { ConvertTo-GB $_.TotalFileStreamBytes } }, `
        @{ N = 'VersionOverheadGB'; E = { ConvertTo-GB $_.VersionOverheadBytes } }, `
        TotalFileCount, LastModified
    $rows | Export-Csv -Path (Join-Path $Path 'carpetas.csv') -Encoding UTF8
}

# Exporta papeleras.csv (incluye PHL como filas adicionales informativas).
function Export-RecycleBinCsv {
    param([object[]]$Bins, [object[]]$Phls, [string]$Path)
    $rows = foreach ($b in $Bins) {
        # Buscar la PHL correspondiente a este sitio (si existe).
        $phl = $Phls | Where-Object { $_.SiteUrl -eq $b.SiteUrl } | Select-Object -First 1
        $phlGB = if ($phl) { ConvertTo-GB $phl.TotalSizeBytes } else { 0 }

        [pscustomobject]@{
            SiteUrl            = $b.SiteUrl
            FirstStageGB       = ConvertTo-GB $b.FirstStageBytes
            FirstStageCount    = $b.FirstStageCount
            SecondStageGB      = ConvertTo-GB $b.SecondStageBytes
            SecondStageCount   = $b.SecondStageCount
            TotalRecycleGB     = ConvertTo-GB $b.TotalBytes
            PreservationHoldGB = $phlGB
        }
    }
    $rows | Export-Csv -Path (Join-Path $Path 'papeleras.csv') -Encoding UTF8
}

# Exporta candidatos_archive.csv.
function Export-ArchiveCandidatesCsv {
    param([object[]]$Candidates, [string]$Path)
    $rows = $Candidates | Select-Object `
        Title, Url, `
        @{ N = 'StorageUsageGB'; E = { ConvertTo-GB $_.StorageUsageBytes } }, `
        LastContentModifiedDate, MonthsIdle, Inactive6Months, Inactive12Months
    $rows | Export-Csv -Path (Join-Path $Path 'candidatos_archive.csv') -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Resumen en consola
# ---------------------------------------------------------------------------

function Show-Summary {
    param(
        [object[]]$Sites,
        [object[]]$Folders,
        [object[]]$Candidates,
        [object[]]$Bins,
        [double]$TotalUsedBytes,
        [double]$TotalQuotaBytes
    )

    $line = ('=' * 72)
    Write-Host "`n$line" -ForegroundColor Cyan
    Write-Host " RESUMEN DE ALMACENAMIENTO SHAREPOINT ONLINE" -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan

    Write-Host ("`nUso total (sitios SPO): {0} GB" -f (ConvertTo-GB $TotalUsedBytes)) -ForegroundColor White
    if ($TotalQuotaBytes -gt 0) {
        $pct = [math]::Round(($TotalUsedBytes / $TotalQuotaBytes) * 100, 1)
        Write-Host ("Cuota del pool (aprox): {0} GB  ->  uso {1}%" -f (ConvertTo-GB $TotalQuotaBytes), $pct) -ForegroundColor White
    }

    # Top 10 carpetas más pesadas.
    # NOTA: estos dos "Top 10" son listas de drill-down JERÁRQUICO (informativas):
    # muestran carpetas de cualquier profundidad (padre e hijas), por eso NO se
    # suman ni se usan como agregado — hacerlo contaría el mismo tamaño/overhead
    # varias veces. El agregado tenant-wide se calcula más abajo con las filas raíz.
    Write-Host "`n--- Top 10 carpetas por tamaño total ---" -ForegroundColor Yellow
    $Folders | Sort-Object TotalSizeBytes -Descending | Select-Object -First 10 | ForEach-Object {
        Write-Host ("  {0,8} GB  {1}" -f (ConvertTo-GB $_.TotalSizeBytes), $_.FolderPath)
    }

    # Top 10 por sobrecarga de versiones.
    Write-Host "`n--- Top 10 carpetas por VersionOverhead (versiones) ---" -ForegroundColor Yellow
    $Folders | Sort-Object VersionOverheadBytes -Descending | Select-Object -First 10 | ForEach-Object {
        Write-Host ("  {0,8} GB  {1}" -f (ConvertTo-GB $_.VersionOverheadBytes), $_.FolderPath)
    }

    # Recuperable estimado si se limitan versiones.
    # Get-PnPFolderStorageMetric devuelve métricas AGREGADAS: la carpeta raíz de
    # cada biblioteca ya incluye el overhead de TODAS sus subcarpetas. Como el
    # recorrido recursivo emite una fila por carpeta en cada profundidad (Depth
    # 0..MaxDepth), sumar VersionOverheadBytes de TODAS las filas cuenta el mismo
    # overhead en el padre y en cada hija -> doble (triple, etc.) conteo.
    # Por eso el total tenant-wide suma SOLO las filas raíz por biblioteca
    # (Depth -eq 0, una por par SiteUrl+Library), que ya contienen el overhead
    # completo de su árbol. Las filas por carpeta de carpetas.csv quedan intactas
    # (sirven para el drill-down); solo cambia este agregado del resumen.
    $rootFolders = $Folders | Where-Object { $_.Depth -eq 0 }
    $recoverableBytes = ($rootFolders | Measure-Object -Property VersionOverheadBytes -Sum).Sum
    $recoverableBytes = [double]($recoverableBytes ?? 0)
    Write-Host ("`nRecuperable estimado por versiones (carpetas relevadas): {0} GB" -f (ConvertTo-GB $recoverableBytes)) -ForegroundColor Green

    # Papeleras: total recuperable inmediato.
    $binBytes = ($Bins | Measure-Object -Property TotalBytes -Sum).Sum
    $binBytes = [double]($binBytes ?? 0)
    Write-Host ("Recuperable en papeleras (1ra+2da etapa): {0} GB" -f (ConvertTo-GB $binBytes)) -ForegroundColor Green

    # Candidatos a archivar.
    $cand12 = $Candidates | Where-Object { $_.Inactive12Months }
    $candBytes = ($cand12 | Measure-Object -Property StorageUsageBytes -Sum).Sum
    $candBytes = [double]($candBytes ?? 0)
    $candGB = ConvertTo-GB $candBytes

    Write-Host "`n--- Candidatos a Microsoft 365 Archive (sin cambios 12+ meses) ---" -ForegroundColor Yellow
    Write-Host ("  Sitios candidatos: {0}" -f @($cand12).Count)
    Write-Host ("  Tamaño total candidato: {0} GB" -f $candGB)

    # Costos comparativos.
    $costArchive = [math]::Round($candBytes / $script:BytesPerGB * $script:CostArchivePerGB, 2)
    $costExtra   = [math]::Round($candBytes / $script:BytesPerGB * $script:CostExtraStoragePerGB, 2)
    $ahorro      = [math]::Round($costExtra - $costArchive, 2)

    Write-Host "`n--- Costo mensual estimado para ese volumen ---" -ForegroundColor Yellow
    # En strings con comillas dobles se escapa el $ con backtick-dólar (`$) para
    # que imprima "US$" literal y no intente expandir una variable.
    Write-Host ("  Microsoft 365 Archive (US`$0.05/GB): US`$ {0}" -f $costArchive)
    Write-Host ("  Extra File Storage   (US`$0.20/GB): US`$ {0}" -f $costExtra)
    Write-Host ("  Ahorro mensual archivando en vez de comprar: US`$ {0}" -f $ahorro) -ForegroundColor Green

    Write-Host "`n$line`n" -ForegroundColor Cyan
}

# ---------------------------------------------------------------------------
# Orquestador principal
# ---------------------------------------------------------------------------

function Main {
    $swTotal = [System.Diagnostics.Stopwatch]::StartNew()

    Initialize-Modules

    # Preparar carpeta de salida.
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    $OutputPath = (Resolve-Path $OutputPath).Path
    Write-Log "Salida en: $OutputPath" 'OK'

    $urls = Get-TenantUrls -Tenant $TenantName

    # 1) Conectar al admin y listar sitios del tenant.
    Write-Log "Conectando al centro de administración ($($urls.Admin))..."
    $adminConn = Connect-SPO -Url $urls.Admin -Tenant $TenantName -ClientId $ClientId
    $sites = Get-TenantSites -AdminConnection $adminConn
    Write-Log "Sitios SPO (sin OneDrive): $($sites.Count)" 'OK'

    $totalUsedBytes = ($sites | Measure-Object -Property StorageUsageCurrentBytes -Sum).Sum
    $totalUsedBytes = [double]($totalUsedBytes ?? 0)

    # Cuota del pool: aproximada por la suma de cuotas por sitio (ajustar si se conoce el valor real del pool).
    $totalQuotaBytes = (($sites | Measure-Object -Property StorageQuotaMB -Sum).Sum) * 1MB
    $totalQuotaBytes = [double]($totalQuotaBytes ?? 0)

    Export-SitesCsv -Sites $sites -Path $OutputPath
    Write-Log "sitios.csv exportado." 'OK'

    # NOTA: el cálculo de candidatos a Archive (Get-ArchiveCandidates) se difiere a
    # DESPUÉS del loop de Top N. Durante ese loop refinamos la fecha de última
    # modificación de los sitios Top N con Get-PnPWeb (fuente más fiable que el
    # LastContentModifiedDate tenant-wide), así los candidatos y candidatos_archive.csv
    # reflejan el dato refinado para los Top N. La detección global sigue usando el
    # dato tenant-wide (menos fiable) para el resto de los 491 sitios.

    # 2) Analizar en detalle los Top N sitios.
    $topSites = $sites | Select-Object -First $TopN
    Write-Log "Analizando en detalle los Top $($topSites.Count) sitios..."

    $allLibraries = [System.Collections.Generic.List[object]]::new()
    $allFolders   = [System.Collections.Generic.List[object]]::new()
    $allBins      = [System.Collections.Generic.List[object]]::new()
    $allPhls      = [System.Collections.Generic.List[object]]::new()

    $i = 0
    foreach ($site in $topSites) {
        $i++
        $pct = [int](($i / [math]::Max($topSites.Count, 1)) * 100)
        Write-Progress -Activity "Analizando sitios en detalle" `
            -Status "($i/$($topSites.Count)) $($site.Url)" -PercentComplete $pct

        try {
            # Conexión por sitio. Connect-SPO lee el método de auth
            # ($Thumbprint/$CertificatePath/$Interactive) desde el scope del script,
            # así que esta conexión per-site hereda automáticamente el MISMO modo que
            # la conexión al admin: en modo certificado (app-only) no abre popups por
            # sitio; solo con Interactive habría un login por sitio.
            $siteConn = Connect-SPO -Url $site.Url -Tenant $TenantName -ClientId $ClientId

            # 2a) Bibliotecas con tamaño y versionado.
            $libs = Get-SiteLibraries -Connection $siteConn -SiteUrl $site.Url
            foreach ($l in $libs) { $allLibraries.Add($l) }

            # 2b) Carpetas recursivas por biblioteca.
            foreach ($lib in $libs) {
                $siteAbsPath = (New-Object System.Uri $site.Url).AbsolutePath
                $libSiteRel = $lib.ServerRelativeUrl -replace "^$([regex]::Escape($siteAbsPath))", ''
                if ([string]::IsNullOrWhiteSpace($libSiteRel)) { continue }

                $folders = Get-FolderMetricsRecursive -Connection $siteConn -SiteUrl $site.Url `
                    -SiteRelativeUrl $libSiteRel -LibraryTitle $lib.Library `
                    -Depth 0 -MaxDepth $MaxDepth -MinFolderSizeGB $MinFolderSizeGB
                foreach ($f in $folders) { $allFolders.Add($f) }
            }

            # 2c) Preservation Hold Library.
            $phl = Get-PreservationHoldLibrary -Connection $siteConn -SiteUrl $site.Url
            if ($phl) { $allPhls.Add($phl) }

            # 2d) Papeleras.
            $bin = Get-RecycleBinSize -Connection $siteConn -SiteUrl $site.Url
            $allBins.Add($bin)

            # 2e) Refinar la fecha de última modificación de este sitio Top N.
            # Get-PnPWeb -Includes LastItemUserModifiedDate, LastItemModifiedDate es
            # una fuente de SOLO LECTURA más fiable que el LastContentModifiedDate
            # tenant-wide (que una tarea del servicio actualiza con retraso).
            # Preferimos LastItemUserModifiedDate (último cambio hecho por un usuario,
            # ignora toques del sistema) y caemos a LastItemModifiedDate si falta.
            # Como $site es la MISMA referencia que está en $sites, al escribir aquí
            # el dato refinado queda disponible para Get-ArchiveCandidates (que se
            # llama después del loop). El refinamiento solo cubre los Top N.
            try {
                $web = Invoke-WithRetry -Context "Get-PnPWeb $($site.Url)" -ArgumentList @($siteConn) -ScriptBlock {
                    param($conn)
                    Get-PnPWeb -Connection $conn -Includes LastItemUserModifiedDate, LastItemModifiedDate
                }
                $refined = $web.LastItemUserModifiedDate
                if (-not $refined -or ($refined -is [datetime] -and $refined -eq [datetime]::MinValue)) {
                    $refined = $web.LastItemModifiedDate
                }
                if ($refined -and -not ($refined -is [datetime] -and $refined -eq [datetime]::MinValue)) {
                    $site.LastContentModifiedDate = $refined
                }
            }
            catch {
                Write-Log "No se pudo refinar la fecha de '$($site.Url)' con Get-PnPWeb: $($_.Exception.Message)" 'WARN'
            }
        }
        catch {
            # Loguear el error del sitio y seguir con el siguiente (no cortar).
            Write-Log "FALLÓ el análisis de '$($site.Url)': $($_.Exception.Message)" 'ERROR'
        }
    }
    Write-Progress -Activity "Analizando sitios en detalle" -Completed

    # Exportar el resto de los CSV.
    Export-LibrariesCsv -Libraries $allLibraries.ToArray() -Path $OutputPath
    Write-Log "bibliotecas.csv exportado ($($allLibraries.Count) filas)." 'OK'

    Export-FoldersCsv -Folders $allFolders.ToArray() -Path $OutputPath
    Write-Log "carpetas.csv exportado ($($allFolders.Count) filas)." 'OK'

    Export-RecycleBinCsv -Bins $allBins.ToArray() -Phls $allPhls.ToArray() -Path $OutputPath
    Write-Log "papeleras.csv exportado ($($allBins.Count) filas)." 'OK'

    # Candidatos a archive (sobre todos los sitios, no solo Top N). Se calcula aquí,
    # después del loop, para que los Top N ya tengan la fecha refinada por Get-PnPWeb.
    $candidates = Get-ArchiveCandidates -Sites $sites
    Export-ArchiveCandidatesCsv -Candidates $candidates -Path $OutputPath
    Write-Log "candidatos_archive.csv exportado ($($candidates.Count) candidatos)." 'OK'

    # 3) Resumen en consola.
    Show-Summary -Sites $sites -Folders $allFolders.ToArray() -Candidates $candidates `
        -Bins $allBins.ToArray() -TotalUsedBytes $totalUsedBytes -TotalQuotaBytes $totalQuotaBytes

    $swTotal.Stop()
    Write-Log ("Listo en {0:mm\:ss}. CSVs en: $OutputPath" -f $swTotal.Elapsed) 'OK'
}

# Punto de entrada.
Main
