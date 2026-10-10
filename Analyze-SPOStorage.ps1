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

    # Umbral de inactividad (meses) para contar candidatos a archive en el resumen.
    [int]$ArchiveMonths = 6,

    # Cantidad de filas a mostrar en los rankings de consola (sitios y carpetas).
    # Solo afecta la salida en pantalla; los CSV siguen con todas las filas.
    [int]$TopRankings = 30,

    # Antigüedad (en meses) a partir de la cual una versión histórica se considera
    # "recuperable" según el criterio del usuario (versiones más viejas que esto).
    [int]$VersionAgeMonths = 2,

    # Cantidad de carpetas (las de mayor VersionOverhead) sobre las que se hace el
    # análisis PROFUNDO de antigüedad de versiones (archivo por archivo).
    # CUIDADO: ese análisis lee el historial de versiones de CADA archivo y es MUY
    # costoso/lento; por eso se limita a estas Top N carpetas por overhead.
    [int]$DeepVersionTopN = 30,

    # Formato del informe final que se genera además de los CSV:
    #   'html' -> solo escribe el .html (RECOMENDADO). Se abre en cualquier
    #             navegador y se imprime a PDF con Cmd+P -> Guardar como PDF.
    #   'pdf'  -> escribe el .html y lo convierte a .pdf; si la conversión tiene
    #             éxito borra el .html intermedio, pero si falla lo conserva.
    #   'both' -> escribe el .html y además intenta el .pdf, conservando ambos.
    # NOTA: la conversión automática a PDF usa Chrome headless y en macOS puede
    # ser inestable si Chrome está abierto (a veces imprime la pestaña activa en
    # vez del informe). Por eso el default es 'html': el HTML es el entregable
    # confiable y se convierte a PDF perfecto con Cmd+P desde el navegador.
    [ValidateSet('html', 'pdf', 'both')]
    [string]$ReportFormat = 'html',

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
# Tamaño PROPIO (exclusivo) de cada carpeta
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Calcula el tamaño PROPIO (exclusivo) de cada carpeta y lo agrega como OwnSizeBytes.
.DESCRIPTION
    Get-PnPFolderStorageMetric devuelve TotalSize ACUMULADO (la carpeta incluye todas
    sus subcarpetas), por eso el ranking por tamaño total muestra el mismo árbol a cada
    nivel de anidación. El tamaño PROPIO de una carpeta = su TotalSize menos la suma del
    TotalSize de sus carpetas HIJAS DIRECTAS (misma SiteUrl + Library, cuyo FolderPath es
    exactamente el de la carpeta + un segmento más). Se clampa en 0 (nunca negativo).

    LIMITACIÓN: si una carpeta hija cayó por debajo de MinFolderSizeGB no está en la
    lista relevada, así que su tamaño no se puede restar y el tamaño propio queda
    SOBREESTIMADO (aproximado) en ese caso.

    Devuelve la MISMA lista de objetos con una propiedad OwnSizeBytes agregada.
#>
function Add-OwnFolderSize {
    param(
        [Parameter(Mandatory = $true)][object[]]$Folders
    )

    # Índice por clave "SiteUrl|Library|FolderPath normalizado" para ubicar padres/hijas.
    # Normalizamos el path quitando una barra inicial/final para comparar segmentos.
    function Get-NormPath { param([string]$p) return $p.Trim('/') }

    # Agrupar por SiteUrl+Library para que la relación padre/hija sea dentro del mismo árbol.
    $byTree = @{}
    foreach ($f in $Folders) {
        $key = "$($f.SiteUrl)|$($f.Library)"
        if (-not $byTree.ContainsKey($key)) { $byTree[$key] = [System.Collections.Generic.List[object]]::new() }
        $byTree[$key].Add($f)
    }

    foreach ($f in $Folders) {
        $norm = Get-NormPath $f.FolderPath
        $siblings = $byTree["$($f.SiteUrl)|$($f.Library)"]

        # Hijas DIRECTAS: su path normalizado empieza con "$norm/" y no tiene más
        # segmentos adicionales (exactamente un nivel por debajo).
        $childSum = 0.0
        foreach ($cand in $siblings) {
            if ([object]::ReferenceEquals($cand, $f)) { continue }
            $cnorm = Get-NormPath $cand.FolderPath
            # El prefijo de carpeta vacío (raíz "") requiere un manejo distinto.
            if ($norm -eq '') {
                # La raíz: hija directa = un único segmento, sin barras internas.
                if ($cnorm -ne '' -and $cnorm -notmatch '/') { $childSum += [double]$cand.TotalSizeBytes }
            }
            else {
                if ($cnorm.StartsWith("$norm/")) {
                    $rest = $cnorm.Substring($norm.Length + 1)
                    if ($rest -ne '' -and $rest -notmatch '/') { $childSum += [double]$cand.TotalSizeBytes }
                }
            }
        }

        $own = [double]$f.TotalSizeBytes - $childSum
        if ($own -lt 0) { $own = 0.0 }  # Clamp: nunca negativo (ver limitación en el header).
        $f | Add-Member -NotePropertyName OwnSizeBytes -NotePropertyValue $own -Force
    }

    return $Folders
}

# ---------------------------------------------------------------------------
# Análisis profundo de antigüedad de versiones (archivo por archivo)
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Clasifica el overhead de versiones de una carpeta por ANTIGÜEDAD de cada versión.
.DESCRIPTION
    Get-PnPFolderStorageMetric da el overhead TOTAL de versiones pero NO su antigüedad.
    Para saber cuánto overhead corresponde a versiones más viejas que N meses hay que
    bajar a nivel ARCHIVO y leer el historial de versiones de cada uno. Es MUY costoso,
    por eso solo se corre sobre las Top N carpetas por overhead (ver DeepVersionTopN).

    DECISIÓN DOCUMENTADA (recursión): se analizan SOLO los archivos INMEDIATOS de la
    carpeta (NO se baja a subcarpetas). Motivo: recursar multiplica el costo de forma
    brutal (una carpeta tenía 88.499 archivos) y el riesgo de runtime es alto. Además,
    la vista de "tamaño propio" (exclusivo) de la Feature 1 combina justo con esto:
    cada carpeta aporta sus archivos inmediatos, sin solaparse con las hijas. La
    implicancia es que el overhead por antigüedad reportado para una carpeta cubre sus
    archivos directos; las subcarpetas pesadas aparecen como sus propias filas si están
    dentro del Top N por overhead.

    PROPIEDADES usadas del objeto de versión (Microsoft.SharePoint.Client.FileVersion,
    PnP.PowerShell 3.4.1, confirmadas por reflexión del tipo CSOM):
      - Length  (Int64)   -> tamaño en bytes de esa versión (se prefiere sobre Size,
                             que es Int32 y puede desbordar en archivos grandes).
      - Created (DateTime) -> fecha de creación de la versión (para la antigüedad).
      - IsCurrentVersion (Boolean) -> se excluye la versión viva (defensivo; además
                             Get-PnPFileVersion ya devuelve solo versiones históricas).

    Devuelve un objeto con los bytes recientes vs recuperables y los buckets de edad.
#>
function Get-FolderVersionAgeBreakdown {
    param(
        [Parameter(Mandatory = $true)][object]$Connection,
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$Library,
        [Parameter(Mandatory = $true)][string]$FolderSiteRelativeUrl,
        [int]$VersionAgeMonths = 2,
        [int]$FolderIndex = 0,
        [int]$FolderTotal = 0
    )

    $now = Get-Date
    $cutoffRecoverable = $now.AddMonths(-$VersionAgeMonths)  # Más viejo que esto = recuperable.
    $cutoff6  = $now.AddMonths(-6)
    $cutoff12 = $now.AddMonths(-12)

    # Acumuladores (bytes).
    $recentBytes      = 0.0   # Versiones <= N meses (NO recuperable).
    $recoverableBytes = 0.0   # Versiones > N meses (recuperable, suma de los buckets).
    $bucket2to6       = 0.0
    $bucket6to12      = 0.0
    $bucketOver12     = 0.0
    $filesAnalyzed    = 0
    $versionsAnalyzed = 0

    # Listar SOLO los archivos inmediatos de la carpeta (sin recursar).
    $files = @()
    try {
        $files = Invoke-WithRetry -Context "listar archivos $FolderSiteRelativeUrl" -ArgumentList @($Connection, $FolderSiteRelativeUrl) -ScriptBlock {
            param($conn, $rel)
            Get-PnPFolderItem -Connection $conn -FolderSiteRelativeUrl $rel -ItemType File
        }
    }
    catch {
        Write-Log "No se pudieron listar archivos de '$FolderSiteRelativeUrl': $($_.Exception.Message)" 'WARN'
        $files = @()
    }

    $fileTotal = @($files).Count
    $fi = 0
    foreach ($file in $files) {
        $fi++
        # Progreso dentro de la carpeta (barra anidada).
        if ($fileTotal -gt 0) {
            Write-Progress -Id 2 -ParentId 1 -Activity "Analizando versiones (carpeta $FolderIndex/$FolderTotal)" `
                -Status "archivo $fi/$fileTotal" -PercentComplete ([int](($fi / $fileTotal) * 100))
        }

        $fileUrl = $file.ServerRelativeUrl
        if ([string]::IsNullOrWhiteSpace($fileUrl)) { continue }

        # Leer el historial de versiones del archivo (envuelto en reintentos por throttling).
        $versions = @()
        try {
            $versions = Invoke-WithRetry -Context "versiones de $fileUrl" -ArgumentList @($Connection, $fileUrl) -ScriptBlock {
                param($conn, $url)
                Get-PnPFileVersion -Connection $conn -Url $url
            }
        }
        catch {
            Write-Log "No se pudieron leer versiones de '$fileUrl': $($_.Exception.Message)" 'WARN'
            continue
        }

        $filesAnalyzed++

        foreach ($v in $versions) {
            # Excluir la versión viva por si apareciera (defensivo).
            if ($v.PSObject.Properties['IsCurrentVersion'] -and $v.IsCurrentVersion) { continue }

            # Tamaño: preferimos Length (Int64); si no está, caemos a Size.
            $vBytes = 0.0
            if ($v.PSObject.Properties['Length'] -and $v.Length) { $vBytes = [double]$v.Length }
            elseif ($v.PSObject.Properties['Size'] -and $v.Size) { $vBytes = [double]$v.Size }

            $created = $null
            if ($v.PSObject.Properties['Created']) { $created = $v.Created }

            $versionsAnalyzed++

            # Sin fecha usable no podemos clasificar por antigüedad: lo contamos como
            # reciente (criterio conservador: no lo marcamos como recuperable).
            if (-not $created -or ($created -is [datetime] -and $created -eq [datetime]::MinValue)) {
                $recentBytes += $vBytes
                continue
            }

            if ($created -gt $cutoffRecoverable) {
                # Reciente: no recuperable según el criterio del usuario.
                $recentBytes += $vBytes
            }
            else {
                # Recuperable (más viejo que N meses). Lo clasificamos en un bucket.
                $recoverableBytes += $vBytes
                if ($created -gt $cutoff6) { $bucket2to6 += $vBytes }
                elseif ($created -gt $cutoff12) { $bucket6to12 += $vBytes }
                else { $bucketOver12 += $vBytes }
            }
        }
    }
    Write-Progress -Id 2 -ParentId 1 -Completed

    return [pscustomobject]@{
        SiteUrl              = $SiteUrl
        Library             = $Library
        FolderPath          = $FolderSiteRelativeUrl
        FilesAnalyzed       = $filesAnalyzed
        VersionsAnalyzed    = $versionsAnalyzed
        RecentVersionsBytes = $recentBytes
        RecoverableVersionsBytes = $recoverableBytes
        Bucket2to6Bytes     = $bucket2to6
        Bucket6to12Bytes    = $bucket6to12
        BucketOver12Bytes   = $bucketOver12
    }
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
    # OwnSizeGB = tamaño PROPIO (exclusivo, sin subcarpetas). Si la carpeta no tiene
    # OwnSizeBytes calculado, caemos al TotalSize (no debería pasar: Main lo calcula antes).
    $rows = $Folders | Sort-Object TotalSizeBytes -Descending | Select-Object `
        SiteUrl, Library, FolderPath, Depth, `
        @{ N = 'TotalSizeGB'; E = { ConvertTo-GB $_.TotalSizeBytes } }, `
        @{ N = 'OwnSizeGB'; E = { ConvertTo-GB ($_.PSObject.Properties['OwnSizeBytes'] ? $_.OwnSizeBytes : $_.TotalSizeBytes) } }, `
        @{ N = 'TotalFileStreamGB'; E = { ConvertTo-GB $_.TotalFileStreamBytes } }, `
        @{ N = 'VersionOverheadGB'; E = { ConvertTo-GB $_.VersionOverheadBytes } }, `
        TotalFileCount, LastModified
    $rows | Export-Csv -Path (Join-Path $Path 'carpetas.csv') -Encoding UTF8
}

# Exporta versiones_antiguas.csv (resultado del análisis profundo de antigüedad).
function Export-OldVersionsCsv {
    param([object[]]$Breakdowns, [string]$Path)
    # TotalVersionOverheadGB se toma de la métrica de carpeta (cross-check contra la
    # suma de versiones reciente + recuperable que relevó el análisis profundo).
    $rows = $Breakdowns | Sort-Object RecoverableVersionsBytes -Descending | Select-Object `
        SiteUrl, Library, FolderPath, FilesAnalyzed, VersionsAnalyzed, `
        @{ N = 'RecentVersionsGB'; E = { ConvertTo-GB $_.RecentVersionsBytes } }, `
        @{ N = 'RecoverableVersionsGB'; E = { ConvertTo-GB $_.RecoverableVersionsBytes } }, `
        @{ N = 'Bucket_2to6mGB'; E = { ConvertTo-GB $_.Bucket2to6Bytes } }, `
        @{ N = 'Bucket_6to12mGB'; E = { ConvertTo-GB $_.Bucket6to12Bytes } }, `
        @{ N = 'Bucket_over12mGB'; E = { ConvertTo-GB $_.BucketOver12Bytes } }, `
        @{ N = 'TotalVersionOverheadGB'; E = { ConvertTo-GB $_.FolderVersionOverheadBytes } }
    $rows | Export-Csv -Path (Join-Path $Path 'versiones_antiguas.csv') -Encoding UTF8
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

# Devuelve el nombre corto de un sitio: último segmento de la URL (o la URL
# completa si es el sitio raíz). Sirve para distinguir dos bibliotecas homónimas
# que viven en sitios distintos (ej. "Documentos compartidos" en /sites/X y en /).
function Get-SiteShortName {
    param([string]$SiteUrl)
    try {
        $abs = (New-Object System.Uri $SiteUrl).AbsolutePath.Trim('/')
        # Sitio raíz (sin segmento de ruta): etiqueta corta en vez de la URL completa.
        if ([string]::IsNullOrWhiteSpace($abs)) { return '/ (raíz)' }
        return ($abs -split '/')[-1]
    }
    catch {
        return $SiteUrl
    }
}

# Arma la etiqueta legible de una fila de carpeta: "[sitio] Biblioteca/CarpetaPath".
# Normaliza el separador: FolderPath a veces viene con '/' inicial y a veces sin él,
# por eso lo recortamos y unimos con una sola barra (evita "DocumentosDocumentos...").
function Format-FolderLabel {
    param([object]$Folder)
    $short = Get-SiteShortName $Folder.SiteUrl
    $path  = [string]$Folder.FolderPath
    $path  = $path.TrimStart('/')  # quitar barra inicial si la hubiera
    if ([string]::IsNullOrWhiteSpace($path)) {
        return ("[{0}] {1}" -f $short, $Folder.Library)
    }
    return ("[{0}] {1}/{2}" -f $short, $Folder.Library, $path)
}

<#
.SYNOPSIS
    Calcula TODAS las cifras del resumen ejecutivo en un solo lugar.
.DESCRIPTION
    Factor común entre el resumen de consola (Show-Summary) y el informe HTML
    (Export-HtmlReport): ambos consumen el MISMO objeto que devuelve esta función,
    así la pantalla y el PDF/HTML coinciden exactamente (misma aritmética, mismos
    umbrales). No recalcular estas cifras por fuera de esta función.

    Devuelve un pscustomobject con todas las magnitudes en BYTES (y los costos ya
    en USD), más los conteos. La conversión a GB se hace al presentar, con ConvertTo-GB.
#>
function Get-SummaryMetrics {
    param(
        [object[]]$Sites,
        [object[]]$Folders,
        [object[]]$Candidates,
        [object[]]$Bins,
        [double]$TotalUsedBytes,
        [double]$TotalQuotaBytes,
        # Antigüedad (meses) usada como criterio de "recuperable" en el análisis profundo.
        [int]$VersionAgeMonths = 2,
        # Desglose de antigüedad de versiones por carpeta (Get-FolderVersionAgeBreakdown).
        [object[]]$VersionBreakdowns = @(),
        # Umbral de inactividad (meses) para contar candidatos a archivar.
        [int]$ArchiveMonths = 6
    )

    # Overhead TOTAL de versiones (todas las edades). Get-PnPFolderStorageMetric es
    # AGREGADO: la carpeta raíz de cada biblioteca ya incluye el overhead de todo su
    # árbol, por eso se suman SOLO las filas raíz (Depth=0) para evitar doble conteo.
    $rootFolders = $Folders | Where-Object { $_.Depth -eq 0 }
    $versionOverheadBytes = [double](($rootFolders | Measure-Object -Property VersionOverheadBytes -Sum).Sum ?? 0)

    # Recuperable por antigüedad (análisis profundo). Es un LOWER BOUND: solo se
    # drillearon las Top DeepVersionTopN carpetas por overhead (archivos inmediatos).
    $recentTot    = [double](($VersionBreakdowns | Measure-Object -Property RecentVersionsBytes -Sum).Sum ?? 0)
    $deepRecoverableBytes = [double](($VersionBreakdowns | Measure-Object -Property RecoverableVersionsBytes -Sum).Sum ?? 0)
    $b2to6        = [double](($VersionBreakdowns | Measure-Object -Property Bucket2to6Bytes -Sum).Sum ?? 0)
    $b6to12       = [double](($VersionBreakdowns | Measure-Object -Property Bucket6to12Bytes -Sum).Sum ?? 0)
    $bOver12      = [double](($VersionBreakdowns | Measure-Object -Property BucketOver12Bytes -Sum).Sum ?? 0)
    $deepOverheadAnalyzedBytes = $recentTot + $deepRecoverableBytes
    $foldersDeep  = @($VersionBreakdowns).Count

    # Recuperable inmediato en papeleras (1ra + 2da etapa).
    $binBytes = [double](($Bins | Measure-Object -Property TotalBytes -Sum).Sum ?? 0)

    # Candidatos a archive: sitios sin cambios en los últimos $ArchiveMonths meses.
    $cand = $Candidates | Where-Object { $_.MonthsIdle -ge $ArchiveMonths }
    $candBytes = [double](($cand | Measure-Object -Property StorageUsageBytes -Sum).Sum ?? 0)

    # Costos comparativos (USD/mes) para el volumen candidato a archive.
    $costArchive = [math]::Round($candBytes / $script:BytesPerGB * $script:CostArchivePerGB, 2)
    $costExtra   = [math]::Round($candBytes / $script:BytesPerGB * $script:CostExtraStoragePerGB, 2)
    $ahorro      = [math]::Round($costExtra - $costArchive, 2)

    $usagePct = if ($TotalQuotaBytes -gt 0) { [math]::Round(($TotalUsedBytes / $TotalQuotaBytes) * 100, 1) } else { $null }

    return [pscustomobject]@{
        TotalUsedBytes            = $TotalUsedBytes
        TotalQuotaBytes           = $TotalQuotaBytes
        UsagePct                  = $usagePct
        VersionOverheadBytes      = $versionOverheadBytes
        BinBytes                  = $binBytes
        VersionAgeMonths          = $VersionAgeMonths
        DeepOverheadAnalyzedBytes = $deepOverheadAnalyzedBytes
        DeepRecoverableBytes      = $deepRecoverableBytes
        Bucket2to6Bytes           = $b2to6
        Bucket6to12Bytes          = $b6to12
        BucketOver12Bytes         = $bOver12
        FoldersDeepAnalyzed       = $foldersDeep
        ArchiveMonths             = $ArchiveMonths
        CandidateCount            = @($cand).Count
        CandidateBytes            = $candBytes
        CostArchive               = $costArchive
        CostExtra                 = $costExtra
        Ahorro                    = $ahorro
    }
}

function Show-Summary {
    param(
        [object[]]$Sites,
        [object[]]$Folders,
        [object[]]$Candidates,
        [object[]]$Bins,
        [double]$TotalUsedBytes,
        [double]$TotalQuotaBytes,
        # Cantidad de filas a mostrar en cada ranking de consola.
        [int]$TopRankings = 30,
        # Antigüedad (meses) usada como criterio de "recuperable" en el análisis profundo.
        [int]$VersionAgeMonths = 2,
        # Desglose de antigüedad de versiones por carpeta (Get-FolderVersionAgeBreakdown).
        [object[]]$VersionBreakdowns = @(),
        # Umbral de inactividad (meses) para contar candidatos a archivar en el resumen.
        [int]$ArchiveMonths = 6,
        # Métricas ya calculadas por Get-SummaryMetrics (se reutilizan para que la
        # consola y el informe HTML coincidan exactamente). Si no se pasa, se calcula.
        [object]$Metrics = $null
    )

    if (-not $Metrics) {
        $Metrics = Get-SummaryMetrics -Sites $Sites -Folders $Folders -Candidates $Candidates `
            -Bins $Bins -TotalUsedBytes $TotalUsedBytes -TotalQuotaBytes $TotalQuotaBytes `
            -VersionAgeMonths $VersionAgeMonths -VersionBreakdowns $VersionBreakdowns -ArchiveMonths $ArchiveMonths
    }

    $line = ('=' * 72)
    Write-Host "`n$line" -ForegroundColor Cyan
    Write-Host " RESUMEN DE ALMACENAMIENTO SHAREPOINT ONLINE" -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan

    Write-Host ("`nUso total (sitios SPO): {0} GB" -f (ConvertTo-GB $Metrics.TotalUsedBytes)) -ForegroundColor White
    if ($Metrics.TotalQuotaBytes -gt 0) {
        Write-Host ("Cuota del pool (aprox): {0} GB  ->  uso {1}%" -f (ConvertTo-GB $Metrics.TotalQuotaBytes), $Metrics.UsagePct) -ForegroundColor White
    }

    # Top N sitios por tamaño (antes solo estaban en sitios.csv, no en consola).
    Write-Host "`n--- Top $TopRankings sitios por tamaño ---" -ForegroundColor Yellow
    $Sites | Sort-Object StorageUsageCurrentBytes -Descending | Select-Object -First $TopRankings | ForEach-Object {
        Write-Host ("  {0,10} GB  {1}  ({2})" -f (ConvertTo-GB $_.StorageUsageCurrentBytes), $_.Title, $_.Url)
    }

    # VISTA 1 — Top N carpetas por tamaño TOTAL (acumulado, incluye subcarpetas).
    # NOTA: esta vista es un drill-down JERÁRQUICO informativo. Como las métricas son
    # ACUMULADAS, aparece el mismo árbol a cada nivel (padre e hijas); por eso NO se
    # suma ni se usa como agregado. El agregado tenant-wide usa las filas raíz (abajo).
    # Ahora cada fila incluye [sitio] Biblioteca para distinguir homónimos.
    Write-Host "`n--- Top $TopRankings carpetas por tamaño TOTAL (acumulado, incluye subcarpetas) ---" -ForegroundColor Yellow
    $Folders | Sort-Object TotalSizeBytes -Descending | Select-Object -First $TopRankings | ForEach-Object {
        Write-Host ("  {0,10} GB  {1}" -f (ConvertTo-GB $_.TotalSizeBytes), (Format-FolderLabel $_))
    }

    # VISTA 2 — Top N carpetas por tamaño PROPIO (exclusivo, sin subcarpetas).
    # OwnSizeBytes = TotalSize - suma(TotalSize de hijas DIRECTAS). Ver Add-OwnFolderSize.
    # Esta vista evita que una misma rama domine el ranking a cada nivel de anidación.
    Write-Host "`n--- Top $TopRankings carpetas por tamaño PROPIO (exclusivo, sin subcarpetas) ---" -ForegroundColor Yellow
    $Folders |
        Sort-Object -Property @{ E = { if ($_.PSObject.Properties['OwnSizeBytes']) { $_.OwnSizeBytes } else { $_.TotalSizeBytes } } } -Descending |
        Select-Object -First $TopRankings | ForEach-Object {
            $own = if ($_.PSObject.Properties['OwnSizeBytes']) { $_.OwnSizeBytes } else { $_.TotalSizeBytes }
            Write-Host ("  {0,10} GB  {1}" -f (ConvertTo-GB $own), (Format-FolderLabel $_))
        }

    # Top N por sobrecarga de versiones (con la misma etiqueta [sitio] Biblioteca).
    Write-Host "`n--- Top $TopRankings carpetas por VersionOverhead (versiones) ---" -ForegroundColor Yellow
    $Folders | Sort-Object VersionOverheadBytes -Descending | Select-Object -First $TopRankings | ForEach-Object {
        Write-Host ("  {0,10} GB  {1}" -f (ConvertTo-GB $_.VersionOverheadBytes), (Format-FolderLabel $_))
    }

    # Recuperable estimado si se limitan versiones (overhead TOTAL de todas las edades).
    # El cálculo (solo filas raíz Depth=0, para evitar el doble conteo de las métricas
    # AGREGADAS de Get-PnPFolderStorageMetric) vive en Get-SummaryMetrics; aquí solo se
    # presenta la cifra ya calculada, de modo que consola e informe HTML coincidan.
    Write-Host ("`nOverhead TOTAL de versiones (todas las edades, carpetas relevadas): {0} GB" -f (ConvertTo-GB $Metrics.VersionOverheadBytes)) -ForegroundColor Green

    # Sección: recuperable borrando versiones de más de N meses (análisis profundo).
    # Es un LOWER BOUND: solo se analizaron a nivel archivo las Top DeepVersionTopN
    # carpetas por overhead, y solo sus archivos inmediatos (ver Get-FolderVersionAgeBreakdown).
    if ($Metrics.FoldersDeepAnalyzed -gt 0) {
        Write-Host "`n--- Recuperable borrando versiones de más de $($Metrics.VersionAgeMonths) meses (análisis profundo) ---" -ForegroundColor Yellow
        Write-Host ("  De los {0} GB de overhead analizado a fondo, {1} GB son versiones de más de {2} meses (recuperable según criterio)." -f (ConvertTo-GB $Metrics.DeepOverheadAnalyzedBytes), (ConvertTo-GB $Metrics.DeepRecoverableBytes), $Metrics.VersionAgeMonths) -ForegroundColor Green
        Write-Host ("  Desglose por antigüedad:  2-6m: {0} GB   6-12m: {1} GB   >12m: {2} GB" -f (ConvertTo-GB $Metrics.Bucket2to6Bytes), (ConvertTo-GB $Metrics.Bucket6to12Bytes), (ConvertTo-GB $Metrics.BucketOver12Bytes))
        Write-Host ("  Carpetas analizadas a fondo: {0}" -f $Metrics.FoldersDeepAnalyzed)
        Write-Host "  NOTA: es un LOWER BOUND: solo se drillearon las carpetas más pesadas por overhead (archivos inmediatos)." -ForegroundColor Gray
    }

    # Papeleras: total recuperable inmediato.
    Write-Host ("`nRecuperable en papeleras (1ra+2da etapa): {0} GB" -f (ConvertTo-GB $Metrics.BinBytes)) -ForegroundColor Green

    # Candidatos a archivar: sitios sin cambios en los últimos $ArchiveMonths meses.
    Write-Host ("`n--- Candidatos a Microsoft 365 Archive (sin cambios {0}+ meses) ---" -f $Metrics.ArchiveMonths) -ForegroundColor Yellow
    Write-Host ("  Sitios candidatos: {0}" -f $Metrics.CandidateCount)
    Write-Host ("  Tamaño total candidato: {0} GB" -f (ConvertTo-GB $Metrics.CandidateBytes))

    # Costos comparativos (calculados en Get-SummaryMetrics).
    Write-Host "`n--- Costo mensual estimado para ese volumen ---" -ForegroundColor Yellow
    # En strings con comillas dobles se escapa el $ con backtick-dólar (`$) para
    # que imprima "US$" literal y no intente expandir una variable.
    Write-Host ("  Microsoft 365 Archive (US`$0.05/GB): US`$ {0}" -f $Metrics.CostArchive)
    Write-Host ("  Extra File Storage   (US`$0.20/GB): US`$ {0}" -f $Metrics.CostExtra)
    Write-Host ("  Ahorro mensual archivando en vez de comprar: US`$ {0}" -f $Metrics.Ahorro) -ForegroundColor Green

    Write-Host "`n$line`n" -ForegroundColor Cyan
}

# ---------------------------------------------------------------------------
# Informe HTML (auto-convertido a PDF si hay un conversor disponible)
# ---------------------------------------------------------------------------

# Escapa un valor dinámico para incrustarlo en HTML de forma segura.
# Usamos [System.Net.WebUtility]::HtmlEncode (parte de la BCL de .NET, siempre
# disponible en PowerShell 7; confirmado en pwsh-preview con PnP.PowerShell 3.4.1)
# para que caracteres como & < > " y acentos no rompan el documento.
function ConvertTo-HtmlSafe {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

<#
.SYNOPSIS
    Intenta convertir un archivo HTML a PDF con el primer conversor disponible.
.DESCRIPTION
    Detecta en RUNTIME, en orden de prioridad: Chrome headless, Microsoft Edge,
    Chromium (todos vía la misma CLI de Chromium) y, por último, wkhtmltopdf en el
    PATH. Chrome/Edge/Chromium headless necesitan rutas ABSOLUTAS y una URL file://.
    Una falla de conversión NO corta la corrida: se loguea WARN y se conserva el HTML.
    Devuelve $true si generó el PDF, $false si no.
#>
function Convert-HtmlToPdf {
    param(
        [Parameter(Mandatory = $true)][string]$HtmlPath,
        [Parameter(Mandatory = $true)][string]$PdfPath
    )

    # Rutas absolutas obligatorias para los navegadores headless.
    $absHtml = (Resolve-Path $HtmlPath).Path
    $fileUrl = ([System.Uri]$absHtml).AbsoluteUri  # file:///...

    # Candidatos tipo Chromium (misma línea de comando). Se prueba el primero que exista.
    $chromiumCandidates = @(
        '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
        '/Applications/Chromium.app/Contents/MacOS/Chromium',
        '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge'
    )

    # Perfil temporal AISLADO: imprescindible para que Chrome NO reutilice una
    # instancia ya abierta del navegador. Sin esto, si el usuario tiene Chrome
    # abierto, headless se engancha a esa sesión e imprime la pestaña activa
    # (la página de inicio de Google) en vez de nuestro file://. (Bug observado.)
    $tmpProfile = Join-Path ([System.IO.Path]::GetTempPath()) ("spo-chrome-" + [guid]::NewGuid().ToString('N'))

    # Verifica que el PDF exista y NO esté vacío (0 bytes = conversión fallida).
    $pdfValido = {
        param($p)
        (Test-Path $p) -and ((Get-Item $p).Length -gt 0)
    }

    # Ejecuta el navegador headless con un límite de tiempo. En macOS, Chrome
    # headless a veces no cierra su proceso solo tras escribir el PDF; el PDF
    # igual queda bien escrito, así que esperamos a que aparezca el archivo y,
    # pasado el timeout, matamos el proceso para no colgar el script.
    $runHeadless = {
        # OJO: el parámetro NO puede llamarse $Args (colisiona con la variable
        # automática $args de PowerShell y rompe el binding posicional). Usamos $ChromeArgs.
        param([string]$Exe, [string[]]$ChromeArgs, [string]$OutPdf, [int]$TimeoutSec = 60)
        if (Test-Path $OutPdf) { Remove-Item $OutPdf -Force -ErrorAction SilentlyContinue }
        # Nota: -WindowStyle no está soportado por Start-Process en PowerShell en
        # macOS/Linux. Chrome escribe mucho ruido por stderr; lo mandamos a archivos
        # temporales DISTINTOS (Start-Process no admite el mismo destino para ambos).
        $outLog = Join-Path ([System.IO.Path]::GetTempPath()) ("spo-chrome-out-" + [guid]::NewGuid().ToString('N') + '.log')
        $errLog = Join-Path ([System.IO.Path]::GetTempPath()) ("spo-chrome-err-" + [guid]::NewGuid().ToString('N') + '.log')
        $p = Start-Process -FilePath $Exe -ArgumentList $ChromeArgs -PassThru `
            -RedirectStandardOutput $outLog -RedirectStandardError $errLog
        # IMPORTANTE: esperar a que Chrome TERMINE solo (no matarlo apenas aparece
        # el archivo). Chrome escribe primero un PDF parcial y recién al final el
        # contenido correcto; si lo matábamos a los 500 ms, el PDF salía incompleto
        # o con la pestaña equivocada. Usamos WaitForExit con timeout de respaldo.
        if (-not $p.WaitForExit($TimeoutSec * 1000)) {
            try { $p.Kill() } catch { }
        }
        # Limpiar los logs temporales de Chrome.
        foreach ($lg in @($outLog, $errLog)) {
            if ($lg -and (Test-Path $lg)) { Remove-Item $lg -Force -ErrorAction SilentlyContinue }
        }
    }

    try {
        foreach ($exe in $chromiumCandidates) {
            if (-not (Test-Path $exe)) { continue }

            # Modo headless "old": IMPRESCINDIBLE en macOS. El "--headless" moderno
            # (= --headless=new) se CONECTA a una instancia de Chrome ya abierta por
            # el usuario e imprime su pestaña activa (la página de inicio de Google)
            # en vez de nuestro file://. "--headless=old" fuerza el headless clásico
            # aislado, que junto al --user-data-dir temporal garantiza que imprima
            # nuestro HTML. (Bug reproducido y verificado con Chrome abierto.)
            $argsClasico = @(
                '--headless=old'
                '--disable-gpu'
                "--user-data-dir=$tmpProfile"
                '--no-first-run'
                '--no-default-browser-check'
                '--disable-extensions'
                '--no-pdf-header-footer'
                "--print-to-pdf=$PdfPath"
                $fileUrl
            )
            try {
                & $runHeadless $exe $argsClasico $PdfPath 60
                if (& $pdfValido $PdfPath) {
                    Write-Log "PDF generado con '$exe': $PdfPath" 'OK'
                    return $true
                }
            }
            catch {
                Write-Log "Conversor '$exe' (headless clásico) falló: $($_.Exception.Message). Reintentando sin --no-pdf-header-footer..." 'WARN'
            }

            # Fallback: algunas versiones rechazan --no-pdf-header-footer.
            $argsSinFlag = @(
                '--headless=old'
                '--disable-gpu'
                "--user-data-dir=$tmpProfile"
                '--no-first-run'
                '--no-default-browser-check'
                '--disable-extensions'
                "--print-to-pdf=$PdfPath"
                $fileUrl
            )
            try {
                & $runHeadless $exe $argsSinFlag $PdfPath 60
                if (& $pdfValido $PdfPath) {
                    Write-Log "PDF generado con '$exe' (sin --no-pdf-header-footer): $PdfPath" 'OK'
                    return $true
                }
            }
            catch {
                Write-Log "Conversor '$exe' falló también sin el flag: $($_.Exception.Message)." 'WARN'
            }
        }
    }
    finally {
        # Limpiar el perfil temporal de Chrome.
        if (Test-Path $tmpProfile) {
            Remove-Item $tmpProfile -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # Último recurso: wkhtmltopdf si está en el PATH.
    $wk = Get-Command wkhtmltopdf -ErrorAction SilentlyContinue
    if ($wk) {
        try {
            & $wk.Source $absHtml $PdfPath 2>$null
            if (Test-Path $PdfPath) {
                Write-Log "PDF generado con wkhtmltopdf: $PdfPath" 'OK'
                return $true
            }
        }
        catch {
            Write-Log "wkhtmltopdf falló: $($_.Exception.Message)." 'WARN'
        }
    }

    return $false
}

<#
.SYNOPSIS
    Genera un informe HTML autocontenido (y lo convierte a PDF si se puede).
.DESCRIPTION
    Documento HTML con CSS embebido (sin assets externos ni scripts: renderiza
    offline y se puede imprimir). Reutiliza las MISMAS cifras que Show-Summary a
    través de Get-SummaryMetrics (pasadas en -Metrics) para que el informe y la
    consola coincidan exactamente. Todos los valores dinámicos (títulos, URLs,
    rutas de carpeta) se escapan con ConvertTo-HtmlSafe.

    Nombre de archivo con timestamp en formato año-mes-día_hora pedido por el
    usuario: informe_SPO_yyyy-MM-dd_HHmmss (.html / .pdf), escrito en $OutputPath.

    Según $Format:
      'html' -> solo HTML.
      'both' -> HTML + intento de PDF, conservando ambos.
      'pdf'  -> HTML + intento de PDF; si el PDF sale, borra el HTML intermedio;
                si falla, conserva el HTML.
#>
function Export-HtmlReport {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object]$Metrics,
        [object[]]$Sites = @(),
        [object[]]$Folders = @(),
        [object[]]$VersionBreakdowns = @(),
        [string]$TenantName = '',
        [int]$TopRankings = 30,
        [ValidateSet('html', 'pdf', 'both')]
        [string]$Format = 'both'
    )

    # Timestamp humano (es-AR) para el encabezado y timestamp de archivo (año-mes-día_hora).
    $esAR = [System.Globalization.CultureInfo]::GetCultureInfo('es-AR')
    $now  = Get-Date
    $stampHuman = $now.ToString('dddd d \d\e MMMM \d\e yyyy, HH:mm:ss', $esAR)
    $stampFile  = $now.ToString('yyyy-MM-dd_HHmmss')

    $baseName = "informe_SPO_$stampFile"
    $htmlPath = Join-Path $Path "$baseName.html"
    $pdfPath  = Join-Path $Path "$baseName.pdf"

    # Helper local: fila de tabla con una etiqueta ya escapada y un valor numérico GB.
    # (Se arman strings; no se ejecuta nada del contenido dinámico.)

    $tenantSafe = ConvertTo-HtmlSafe $TenantName

    # --- Resumen ejecutivo: cifras desde $Metrics (mismas que la consola) ---
    $usoTotalGB      = ConvertTo-GB $Metrics.TotalUsedBytes
    $papelerasGB     = ConvertTo-GB $Metrics.BinBytes
    $overheadGB      = ConvertTo-GB $Metrics.VersionOverheadBytes
    $deepRecGB       = ConvertTo-GB $Metrics.DeepRecoverableBytes
    $b2to6GB         = ConvertTo-GB $Metrics.Bucket2to6Bytes
    $b6to12GB        = ConvertTo-GB $Metrics.Bucket6to12Bytes
    $bOver12GB       = ConvertTo-GB $Metrics.BucketOver12Bytes
    $candGB          = ConvertTo-GB $Metrics.CandidateBytes

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="es">')
    [void]$sb.AppendLine('<head>')
    [void]$sb.AppendLine('<meta charset="utf-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
    [void]$sb.AppendLine('<title>Informe de almacenamiento SharePoint Online</title>')
    [void]$sb.AppendLine('<style>')
    [void]$sb.AppendLine(@'
  * { box-sizing: border-box; }
  body { font-family: -apple-system, "Segoe UI", Arial, sans-serif; color: #1b1b1b; margin: 24px; font-size: 14px; }
  h1 { font-size: 24px; margin: 0 0 4px 0; color: #0f3d5c; }
  h2 { font-size: 18px; margin: 28px 0 10px 0; color: #0f3d5c; border-bottom: 2px solid #d0dce5; padding-bottom: 4px; }
  .meta { color: #555; font-size: 13px; margin-bottom: 16px; }
  .desc { color: #555; font-size: 12.5px; margin: 0 0 10px 0; font-style: italic; }
  table { border-collapse: collapse; width: 100%; margin: 8px 0 16px 0; }
  th, td { border: 1px solid #cfd8dd; padding: 6px 8px; text-align: left; vertical-align: top; }
  th { background: #eef3f6; }
  td.num, th.num { text-align: right; white-space: nowrap; }
  tr:nth-child(even) td { background: #f8fafb; }
  .cards { display: flex; flex-wrap: wrap; gap: 12px; margin: 8px 0 8px 0; }
  .card { border: 1px solid #cfd8dd; border-radius: 8px; padding: 12px 16px; min-width: 200px; flex: 1 1 200px; background: #f8fafb; }
  .card .label { font-size: 12px; color: #555; text-transform: uppercase; letter-spacing: .03em; }
  .card .value { font-size: 28px; font-weight: 700; color: #0f3d5c; margin-top: 4px; }
  .card.highlight { background: #e9f7ef; border-color: #a6d9bb; }
  .card.highlight .value { color: #176b3a; }
  .buckets { font-size: 12px; color: #555; margin-top: 6px; }
  .footer { margin-top: 28px; font-size: 12px; color: #666; border-top: 1px solid #d0dce5; padding-top: 10px; }
  .ro { font-weight: 700; color: #176b3a; }
'@)
    [void]$sb.AppendLine('</style>')
    [void]$sb.AppendLine('</head>')
    [void]$sb.AppendLine('<body>')

    # --- Título + timestamp + tenant ---
    [void]$sb.AppendLine('<h1>Informe de almacenamiento de SharePoint Online</h1>')
    [void]$sb.AppendLine(('<div class="meta">Generado el {0}{1}</div>' -f (ConvertTo-HtmlSafe $stampHuman), $(if ($tenantSafe) { " &middot; Tenant: <strong>$tenantSafe</strong>" } else { '' })))

    # --- Resumen ejecutivo ---
    [void]$sb.AppendLine('<h2>Resumen ejecutivo</h2>')
    [void]$sb.AppendLine('<p class="desc">Cifras clave del tenant: cuánto espacio se usa en total y cuánto se podría recuperar sin comprar más almacenamiento (vaciando papeleras y limpiando versiones antiguas), más una comparación de costos entre archivar y comprar storage extra.</p>')
    [void]$sb.AppendLine('<div class="cards">')
    [void]$sb.AppendLine(('<div class="card"><div class="label">Uso total (sitios SPO)</div><div class="value">{0} GB</div></div>' -f $usoTotalGB))
    [void]$sb.AppendLine(('<div class="card highlight"><div class="label">Recuperable en papeleras</div><div class="value">{0} GB</div></div>' -f $papelerasGB))
    [void]$sb.AppendLine(('<div class="card"><div class="label">Overhead total de versiones (todas las edades)</div><div class="value">{0} GB</div></div>' -f $overheadGB))
    [void]$sb.AppendLine(('<div class="card highlight"><div class="label">Recuperable borrando versiones &gt; {0} meses</div><div class="value">{1} GB</div><div class="buckets">2-6m: {2} GB &middot; 6-12m: {3} GB &middot; &gt;12m: {4} GB</div></div>' -f $Metrics.VersionAgeMonths, $deepRecGB, $b2to6GB, $b6to12GB, $bOver12GB))
    [void]$sb.AppendLine(('<div class="card"><div class="label">Candidatos a Archive (&ge; {0} meses sin cambios)</div><div class="value">{1}</div><div class="buckets">{2} GB en total</div></div>' -f $Metrics.ArchiveMonths, $Metrics.CandidateCount, $candGB))
    [void]$sb.AppendLine('</div>')

    # Tabla de costos comparativos.
    [void]$sb.AppendLine('<table>')
    [void]$sb.AppendLine('<tr><th>Opción de costo (para el volumen candidato)</th><th class="num">USD / mes</th></tr>')
    [void]$sb.AppendLine(('<tr><td>Microsoft 365 Archive (US$0.05/GB)</td><td class="num">US$ {0}</td></tr>' -f $Metrics.CostArchive))
    [void]$sb.AppendLine(('<tr><td>Office 365 Extra File Storage (US$0.20/GB)</td><td class="num">US$ {0}</td></tr>' -f $Metrics.CostExtra))
    [void]$sb.AppendLine(('<tr><td><strong>Ahorro mensual archivando en vez de comprar storage</strong></td><td class="num"><strong>US$ {0}</strong></td></tr>' -f $Metrics.Ahorro))
    [void]$sb.AppendLine('</table>')

    # --- Top N sitios por tamaño ---
    [void]$sb.AppendLine(('<h2>Top {0} sitios por tamaño</h2>' -f $TopRankings))
    [void]$sb.AppendLine(('<p class="desc">Los {0} sitios de SharePoint que más espacio ocupan, de mayor a menor. Cada "sitio" es un espacio de trabajo (por equipo o área). Muestra a nivel general dónde está repartido el almacenamiento del tenant.</p>' -f $TopRankings))
    [void]$sb.AppendLine('<table>')
    [void]$sb.AppendLine('<tr><th>Título</th><th>URL</th><th class="num">GB</th></tr>')
    $topSites = $Sites | Sort-Object StorageUsageCurrentBytes -Descending | Select-Object -First $TopRankings
    foreach ($s in $topSites) {
        [void]$sb.AppendLine(('<tr><td>{0}</td><td>{1}</td><td class="num">{2}</td></tr>' -f (ConvertTo-HtmlSafe $s.Title), (ConvertTo-HtmlSafe $s.Url), (ConvertTo-GB $s.StorageUsageCurrentBytes)))
    }
    [void]$sb.AppendLine('</table>')

    # --- Top N carpetas por tamaño TOTAL (acumulado) ---
    [void]$sb.AppendLine(('<h2>Top {0} carpetas por tamaño TOTAL (acumulado)</h2>' -f $TopRankings))
    [void]$sb.AppendLine('<p class="desc">Carpetas más pesadas contando TODO lo que hay dentro, incluidas sus subcarpetas. Como una carpeta incluye a sus hijas, acá verás la misma rama repetida a distintos niveles (ej. una carpeta y su subcarpeta). Sirve para ver qué rama completa pesa más. El prefijo [sitio] indica a qué sitio pertenece cada carpeta.</p>')
    [void]$sb.AppendLine('<table>')
    [void]$sb.AppendLine('<tr><th>Carpeta</th><th class="num">GB</th></tr>')
    $topTotal = $Folders | Sort-Object TotalSizeBytes -Descending | Select-Object -First $TopRankings
    foreach ($f in $topTotal) {
        [void]$sb.AppendLine(('<tr><td>{0}</td><td class="num">{1}</td></tr>' -f (ConvertTo-HtmlSafe (Format-FolderLabel $f)), (ConvertTo-GB $f.TotalSizeBytes)))
    }
    [void]$sb.AppendLine('</table>')

    # --- Top N carpetas por tamaño PROPIO (exclusivo) ---
    [void]$sb.AppendLine(('<h2>Top {0} carpetas por tamaño PROPIO (exclusivo)</h2>' -f $TopRankings))
    [void]$sb.AppendLine('<p class="desc">Carpetas más pesadas contando SOLO los archivos que están directamente en ellas, sin sumar sus subcarpetas. Esta vista muestra en qué carpeta puntual vive el peso de verdad (no una rama entera), así es la más útil para decidir dónde actuar.</p>')
    [void]$sb.AppendLine('<table>')
    [void]$sb.AppendLine('<tr><th>Carpeta</th><th class="num">GB</th></tr>')
    $topOwn = $Folders |
        Sort-Object -Property @{ E = { if ($_.PSObject.Properties['OwnSizeBytes']) { $_.OwnSizeBytes } else { $_.TotalSizeBytes } } } -Descending |
        Select-Object -First $TopRankings
    foreach ($f in $topOwn) {
        $own = if ($f.PSObject.Properties['OwnSizeBytes']) { $f.OwnSizeBytes } else { $f.TotalSizeBytes }
        [void]$sb.AppendLine(('<tr><td>{0}</td><td class="num">{1}</td></tr>' -f (ConvertTo-HtmlSafe (Format-FolderLabel $f)), (ConvertTo-GB $own)))
    }
    [void]$sb.AppendLine('</table>')

    # --- Top N carpetas por VersionOverhead ---
    [void]$sb.AppendLine(('<h2>Top {0} carpetas por VersionOverhead (versiones)</h2>' -f $TopRankings))
    [void]$sb.AppendLine('<p class="desc">Carpetas donde más espacio ocupan las VERSIONES anteriores de los archivos (el historial que SharePoint guarda en cada edición). VersionOverhead = tamaño total de la carpeta menos el de la versión actual, es decir, lo que ocupan las copias viejas. Un overhead alto indica una buena candidata para limpiar versiones.</p>')
    [void]$sb.AppendLine('<table>')
    [void]$sb.AppendLine('<tr><th>Carpeta</th><th class="num">GB</th></tr>')
    $topOverhead = $Folders | Sort-Object VersionOverheadBytes -Descending | Select-Object -First $TopRankings
    foreach ($f in $topOverhead) {
        [void]$sb.AppendLine(('<tr><td>{0}</td><td class="num">{1}</td></tr>' -f (ConvertTo-HtmlSafe (Format-FolderLabel $f)), (ConvertTo-GB $f.VersionOverheadBytes)))
    }
    [void]$sb.AppendLine('</table>')

    # --- Versiones recuperables (>N meses) por carpeta ---
    [void]$sb.AppendLine(('<h2>Versiones recuperables (&gt; {0} meses) por carpeta</h2>' -f $Metrics.VersionAgeMonths))
    [void]$sb.AppendLine(('<p class="desc">Análisis fino (archivo por archivo) sobre las carpetas de mayor overhead: cuánto de las versiones antiguas tiene más de {0} meses, es decir, lo que podrías borrar según tu criterio. Las columnas 2-6m / 6-12m / &gt;12m reparten ese espacio por antigüedad. "Archivos analizados" es cuántos archivos inmediatos se revisaron en esa carpeta.</p>' -f $Metrics.VersionAgeMonths))
    [void]$sb.AppendLine('<table>')
    [void]$sb.AppendLine('<tr><th>Carpeta</th><th class="num">Archivos analizados</th><th class="num">Recuperable GB</th><th class="num">2-6m</th><th class="num">6-12m</th><th class="num">&gt;12m</th></tr>')
    $vb = $VersionBreakdowns | Sort-Object RecoverableVersionsBytes -Descending
    foreach ($b in $vb) {
        $label = Format-FolderLabel ([pscustomobject]@{ SiteUrl = $b.SiteUrl; Library = $b.Library; FolderPath = $b.FolderPath })
        [void]$sb.AppendLine(('<tr><td>{0}</td><td class="num">{1}</td><td class="num">{2}</td><td class="num">{3}</td><td class="num">{4}</td><td class="num">{5}</td></tr>' -f `
            (ConvertTo-HtmlSafe $label), $b.FilesAnalyzed, (ConvertTo-GB $b.RecoverableVersionsBytes), `
            (ConvertTo-GB $b.Bucket2to6Bytes), (ConvertTo-GB $b.Bucket6to12Bytes), (ConvertTo-GB $b.BucketOver12Bytes)))
    }
    [void]$sb.AppendLine('</table>')

    # --- Footer ---
    [void]$sb.AppendLine('<div class="footer">')
    [void]$sb.AppendLine(('<p>La cifra de versiones recuperables es un <strong>LOWER BOUND</strong>: solo se analizaron a fondo las carpetas más pesadas por overhead (Top {0}), y solo sus archivos inmediatos.</p>' -f $Metrics.FoldersDeepAnalyzed))
    [void]$sb.AppendLine('<p class="ro">Herramienta de SOLO LECTURA: este informe solo reporta; no borra, mueve ni modifica nada en el tenant.</p>')
    [void]$sb.AppendLine('</div>')

    [void]$sb.AppendLine('</body>')
    [void]$sb.AppendLine('</html>')

    # Escribir el HTML (UTF-8).
    $sb.ToString() | Out-File -FilePath $htmlPath -Encoding utf8
    Write-Log "Informe HTML generado: $htmlPath" 'OK'

    # Si no se pide PDF, terminamos: el HTML es el entregable garantizado.
    if ($Format -eq 'html') {
        Write-Log "Para obtener un PDF: abrí el HTML en el navegador y usá Cmd+P -> Guardar como PDF." 'INFO'
        return
    }

    # Intentar la conversión a PDF (best-effort). Una falla NO corta la corrida.
    $pdfOk = $false
    try {
        $pdfOk = Convert-HtmlToPdf -HtmlPath $htmlPath -PdfPath $pdfPath
    }
    catch {
        Write-Log "La conversión a PDF falló inesperadamente: $($_.Exception.Message). Queda el HTML." 'WARN'
        $pdfOk = $false
    }

    if (-not $pdfOk) {
        Write-Log "No se pudo generar el PDF automáticamente. El informe queda como HTML." 'WARN'
        Write-Log "Para el PDF: abrí '$htmlPath' en el navegador y usá Cmd+P -> Guardar como PDF (sale idéntico y es 100% confiable)." 'INFO'
        return
    }

    # Modo 'pdf': borrar el HTML intermedio SOLO si el PDF se generó bien.
    if ($Format -eq 'pdf') {
        try {
            Remove-Item -Path $htmlPath -Force
            Write-Log "Modo 'pdf': HTML intermedio eliminado tras la conversión exitosa." 'INFO'
        }
        catch {
            Write-Log "No se pudo borrar el HTML intermedio '$htmlPath': $($_.Exception.Message). Se conserva." 'WARN'
        }
    }
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

    # Calcular el tamaño PROPIO (exclusivo) de cada carpeta ANTES de exportar/resumir,
    # para que carpetas.csv (OwnSizeGB) y la vista "tamaño propio" lo tengan disponible.
    $foldersArr = $allFolders.ToArray()
    $foldersArr = Add-OwnFolderSize -Folders $foldersArr

    Export-FoldersCsv -Folders $foldersArr -Path $OutputPath
    Write-Log "carpetas.csv exportado ($($allFolders.Count) filas)." 'OK'

    Export-RecycleBinCsv -Bins $allBins.ToArray() -Phls $allPhls.ToArray() -Path $OutputPath
    Write-Log "papeleras.csv exportado ($($allBins.Count) filas)." 'OK'

    # 2f) Análisis PROFUNDO de antigüedad de versiones sobre las Top DeepVersionTopN
    # carpetas por overhead. Reutiliza la conexión por sitio (una por sitio candidato).
    $versionBreakdowns = [System.Collections.Generic.List[object]]::new()
    $deepTargets = $foldersArr | Sort-Object VersionOverheadBytes -Descending | Select-Object -First $DeepVersionTopN
    $deepCount = @($deepTargets).Count

    if ($deepCount -gt 0) {
        Write-Log "Iniciando análisis PROFUNDO de versiones (archivo por archivo) sobre las Top $deepCount carpetas por overhead. Esto lee el historial de versiones de cada archivo inmediato y PUEDE TARDAR MUCHO (el relevamiento de carpetas de 3 sitios ya tardó ~57 min; esto es más pesado)." 'WARN'

        # Cachear conexiones por sitio para no reconectar en cada carpeta del mismo sitio.
        $connCache = @{}
        $di = 0
        foreach ($folder in $deepTargets) {
            $di++
            $pct = [int](($di / [math]::Max($deepCount, 1)) * 100)
            Write-Progress -Id 1 -Activity "Análisis profundo de versiones" `
                -Status "($di/$deepCount) $($folder.SiteUrl) $($folder.FolderPath)" -PercentComplete $pct

            try {
                if (-not $connCache.ContainsKey($folder.SiteUrl)) {
                    $connCache[$folder.SiteUrl] = Connect-SPO -Url $folder.SiteUrl -Tenant $TenantName -ClientId $ClientId
                }
                $fConn = $connCache[$folder.SiteUrl]

                $bd = Get-FolderVersionAgeBreakdown -Connection $fConn -SiteUrl $folder.SiteUrl `
                    -Library $folder.Library -FolderSiteRelativeUrl $folder.FolderPath `
                    -VersionAgeMonths $VersionAgeMonths -FolderIndex $di -FolderTotal $deepCount

                # Adjuntar el overhead de la métrica de carpeta para el cross-check del CSV.
                $bd | Add-Member -NotePropertyName FolderVersionOverheadBytes -NotePropertyValue ([double]$folder.VersionOverheadBytes) -Force
                $versionBreakdowns.Add($bd)
            }
            catch {
                Write-Log "FALLÓ el análisis profundo de '$($folder.SiteUrl) $($folder.FolderPath)': $($_.Exception.Message)" 'ERROR'
            }
        }
        Write-Progress -Id 1 -Activity "Análisis profundo de versiones" -Completed

        Export-OldVersionsCsv -Breakdowns $versionBreakdowns.ToArray() -Path $OutputPath
        Write-Log "versiones_antiguas.csv exportado ($($versionBreakdowns.Count) filas)." 'OK'
    }

    # Candidatos a archive (sobre todos los sitios, no solo Top N). Se calcula aquí,
    # después del loop, para que los Top N ya tengan la fecha refinada por Get-PnPWeb.
    $candidates = Get-ArchiveCandidates -Sites $sites
    Export-ArchiveCandidatesCsv -Candidates $candidates -Path $OutputPath
    Write-Log "candidatos_archive.csv exportado ($($candidates.Count) candidatos)." 'OK'

    # Calcular las métricas del resumen UNA sola vez y compartirlas entre la consola
    # (Show-Summary) y el informe HTML (Export-HtmlReport), así ambos coinciden exacto.
    $breakdownsArr = $versionBreakdowns.ToArray()
    $binsArr = $allBins.ToArray()
    $metrics = Get-SummaryMetrics -Sites $sites -Folders $foldersArr -Candidates $candidates `
        -Bins $binsArr -TotalUsedBytes $totalUsedBytes -TotalQuotaBytes $totalQuotaBytes `
        -VersionAgeMonths $VersionAgeMonths -VersionBreakdowns $breakdownsArr -ArchiveMonths $ArchiveMonths

    # 3) Resumen en consola.
    Show-Summary -Sites $sites -Folders $foldersArr -Candidates $candidates `
        -Bins $binsArr -TotalUsedBytes $totalUsedBytes -TotalQuotaBytes $totalQuotaBytes `
        -TopRankings $TopRankings -VersionAgeMonths $VersionAgeMonths `
        -VersionBreakdowns $breakdownsArr -ArchiveMonths $ArchiveMonths -Metrics $metrics

    # 4) Informe HTML (y PDF si hay conversor). Se genera DESPUÉS de los CSV y del
    # resumen de consola, para que la pantalla siga mostrando todo igual que antes.
    Export-HtmlReport -Path $OutputPath -Metrics $metrics -Sites $sites -Folders $foldersArr `
        -VersionBreakdowns $breakdownsArr -TenantName $TenantName -TopRankings $TopRankings -Format $ReportFormat

    $swTotal.Stop()
    Write-Log ("Listo en {0:mm\:ss}. CSVs en: $OutputPath" -f $swTotal.Elapsed) 'OK'
}

# Punto de entrada.
Main
