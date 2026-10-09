# =============================================================================
# config.example.ps1
# -----------------------------------------------------------------------------
# Plantilla de configuración. COPIAR a config.ps1 (ignorado por git) y completar.
# NO poner credenciales reales en este archivo de ejemplo.
#
# Uso:
#   cp config.example.ps1 config.ps1   # editar config.ps1 con tus valores
#   . ./config.ps1                      # dot-source para cargar las variables
#   ./Analyze-SPOStorage.ps1 @SPOArgs   # splatting de los parámetros
# =============================================================================

# Hashtable de parámetros para pasar por splatting al script.
$SPOArgs = @{

    # Nombre del tenant SIN sufijo. Ej: "contoso" -> contoso.onmicrosoft.com
    TenantName = "REEMPLAZAR_TENANT"

    # Application (client) ID de la app registration de Entra ID.
    # Lo devuelve Register-PnPEntraIDApp. Ej: 00000000-0000-0000-0000-000000000000
    ClientId   = "REEMPLAZAR_CLIENT_ID"

    # Carpeta de salida de los CSV.
    OutputPath = "./salida"

    # Ajustes de análisis.
    MaxDepth        = 3
    # TopN: sitios más pesados a analizar en detalle. Para una auditoría de tenant
    # completo conviene subirlo (ej. 50 o más) una vez que uses certificado
    # (app-only), que no abre popups por sitio y es mucho más rápido.
    TopN            = 20
    MinFolderSizeGB = 1.0

    # Cantidad de filas en los rankings de CONSOLA (sitios y carpetas). Solo afecta
    # la salida en pantalla; los CSV siempre llevan todas las filas.
    TopRankings     = 30

    # VersionAgeMonths: una versión histórica más vieja que estos meses se considera
    # "recuperable" (criterio del usuario). Las más nuevas no se cuentan como ahorro.
    VersionAgeMonths = 2

    # DeepVersionTopN: cantidad de carpetas (las de mayor overhead) sobre las que se
    # corre el análisis PROFUNDO de antigüedad de versiones, archivo por archivo.
    # CUIDADO PERFORMANCE: ese análisis lee el historial de versiones de CADA archivo
    # inmediato y es MUY lento (el relevamiento de carpetas de 3 sitios ya tardó ~57
    # min; esto es más pesado). Subilo con cautela.
    DeepVersionTopN = 30

    # ArchiveMonths: umbral de inactividad (en meses) para marcar un sitio como
    # candidato a Microsoft 365 Archive. Un sitio sin cambios por más de este tiempo
    # se cuenta como candidato en el resumen (y en candidatos_archive.csv).
    ArchiveMonths   = 6

    # ReportFormat: formato del informe final que se genera además de los CSV, en la
    # misma carpeta de salida (nombre: informe_SPO_yyyy-MM-dd_HHmmss).
    #   'html' -> solo escribe el .html (siempre se puede abrir e imprimir a PDF
    #             desde el navegador).
    #   'pdf'  -> escribe el .html y lo convierte a .pdf; si la conversión sale bien
    #             borra el .html intermedio, pero si falla lo conserva.
    #   'both' -> escribe el .html y además el .pdf, conservando ambos (default).
    # La conversión a PDF detecta en runtime Chrome, Microsoft Edge, Chromium o
    # wkhtmltopdf; si no hay NINGUNO, el informe queda como HTML (sin PDF).
    ReportFormat    = 'both'

    # =========================================================================
    # Autenticación: dejar activa UNA sola opción.
    #
    # Tradeoff:
    #  - CERTIFICADO (app-only, RECOMENDADO para tenant completo): autenticación
    #    de aplicación, desatendida, SIN popups de navegador, lee TODOS los sitios
    #    (con Sites.Read.All / Sites.FullControl.All) y es rápida. Ideal para
    #    auditar los 491 sitios o un TopN alto.
    #  - INTERACTIVO (delegado): cómodo para una corrida puntual, pero abre un
    #    popup por sitio, solo ve los sitios donde el usuario es miembro (de ahí
    #    errores "unauthorized") y es LENTO en corridas grandes.
    #
    # NOTA sobre precedencia: si por error quedaran activos a la vez un método de
    # certificado (Thumbprint/CertificatePath) e Interactive, el script prioriza
    # el modo app-only (certificado). Interactive es el fallback.
    # =========================================================================

    # (Las etiquetas Opción A/B/C coinciden con las del README, paso 2.)

    # ---- Opción A (RECOMENDADA): certificado por thumbprint ----
    # Descomentar y poner el thumbprint del certificado subido a la app.
    # Thumbprint = "REEMPLAZAR_THUMBPRINT"

    # ---- Opción B: certificado desde archivo .pfx (también app-only) ----
    # CertificatePath     = "./cert.pfx"
    # CertificatePassword = (Read-Host "Password del PFX" -AsSecureString)

    # ---- Opción C: login interactivo (comentar las otras) ----
    Interactive = $true
}

# Ejemplo de ejecución tras dot-sourcing este archivo:
#   . ./config.ps1
#   ./Analyze-SPOStorage.ps1 @SPOArgs
