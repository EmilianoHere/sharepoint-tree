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

    # ---- Opción A (RECOMENDADA): certificado por thumbprint ----
    # Descomentar y poner el thumbprint del certificado subido a la app.
    # Thumbprint = "REEMPLAZAR_THUMBPRINT"

    # ---- Opción B: login interactivo (comentar las otras) ----
    Interactive = $true

    # ---- Opción C: certificado desde archivo .pfx ----
    # CertificatePath     = "./cert.pfx"
    # CertificatePassword = (Read-Host "Password del PFX" -AsSecureString)
}

# Ejemplo de ejecución tras dot-sourcing este archivo:
#   . ./config.ps1
#   ./Analyze-SPOStorage.ps1 @SPOArgs
