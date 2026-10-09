# Implementation Plan — Fix bugs + certificate (app-only) auth en Analyze-SPOStorage.ps1

Repo: `/Users/puser/repos/phinx/Arbitrage/emi/sharepoint-tree`
Entorno: macOS, binario PowerShell = `pwsh-preview` (NO `pwsh`), PnP.PowerShell 3.4.1.
NO se puede conectar al tenant real. **Verificación SOLO por parse estático.** Editar archivos IN PLACE (sin worktree).

## Comando de verificación base (parse estático)

Todos los ítems de código verifican con el parser de PowerShell (no ejecuta el script, solo valida sintaxis/AST):

```bash
pwsh-preview -NoProfile -Command '$t=$null;$e=$null;[System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path ./Analyze-SPOStorage.ps1).Path,[ref]$t,[ref]$e)>$null; if($e){$e|ForEach-Object{Write-Host $_.Message}; exit 1}else{Write-Host "PARSE OK"}'
```

Esperado: imprime `PARSE OK` y exit 0. (Confirmado que funciona en este entorno.)

Para `config.ps1` / `config.example.ps1` usar el mismo parser cambiando el path.

## Decisiones de diseño ( grounded en el código leído)

- **BUG 1 — qué filas sumar:** `carpetas.csv` emite una fila por carpeta en cada profundidad (`Depth` 0..MaxDepth). `Get-PnPFolderStorageMetric` ya es agregado (padre incluye hijos). Sumar `VersionOverheadBytes` de TODAS las filas cuenta el mismo overhead en el padre y en cada hijo. Decisión: el total tenant-wide de "recuperable por versiones" debe sumar **solo las filas raíz por biblioteca**, es decir las filas con `Depth -eq 0` (una por par `SiteUrl` + `Library`). Esas raíces ya contienen el overhead de todas sus subcarpetas. Evidencia en `salida/carpetas.csv`: hay dos raíces `Depth=0` (NovaProduccion `/Documentos compartidos` 276,71 y root-site `Documentos compartidos` 550,21) cuya suma de overhead es ~827 GB; sumar todas las filas da los 3095 GB imposibles observados. (Nota: `Depth` se serializa como texto en el CSV, pero en memoria `$allFolders` las filas tienen `Depth` como `[int]`; filtrar en memoria por `-eq 0`.)
- **BUG 1 — "top total size":** El "Uso total (sitios SPO)" ya se calcula desde `$sites` (StorageUsageCurrentBytes), NO desde carpetas, así que ese número no tiene el defecto. El único agregado tenant-wide derivado de carpetas con el defecto de doble conteo es el recuperable por versiones. Las listas "Top 10 carpetas por tamaño" y "Top 10 por VersionOverhead" son drill-down y DEBEN seguir mostrando todas las profundidades (no se tocan). Documentar en comentario que esos Top 10 son informativos/jerárquicos y por eso no se suman.
- **BUG 2 — fuente de fecha:** `Get-PnPTenantSite` llena `LastContentModifiedDate` pero una tarea programada la actualiza y es poco fiable; la fuente confiable por web es `LastItemUserModifiedDate` vía `Get-PnPWeb` ([ref1](https://techcommunity.microsoft.com/discussions/sharepoint_general/lastcontentmodifieddate-using-powershell-not-equal-to-actual-date-in-sharepoint-/1453257), [ref2](https://github.com/pnp/powershell/discussions/166) — contenido reformulado por cumplimiento). En `salida/sitios.csv` las fechas SÍ vienen pobladas (todas recientes, 2026), así que con esos datos de prueba legítimamente hay 0 candidatos; el problema real es la falta de diagnóstico y la fragilidad del dato tenant-wide. Decisión: (a) mantener la detección tenant-wide sobre `LastContentModifiedDate` para los 491 sitios; (b) para los Top N ya conectados, refinar con `Get-PnPWeb -Includes LastItemUserModifiedDate, LastItemModifiedDate` como fuente preferida cuando esté disponible; (c) tratar explícitamente `$null` y `DateTime.MinValue`; (d) loguear cuántos sitios tenían fecha usable. Documentar en comentario en español la limitación (tenant-wide puede ser impreciso; el refinamiento per-site solo cubre Top N).
- **BUG 3 — escape de `$`:** En strings con comillas dobles de PowerShell, `\$` imprime `\$` literal. Usar backtick-dólar `` `$ `` para imprimir `$` limpio. Mantener la aritmética ya correcta.
- **MAIN — app-only:** La forma `Connect-PnPOnline -Url <url> -ClientId <id> -Tenant <name>.onmicrosoft.com -Thumbprint <tp>` es correcta para PnP 3.x ([Connect-PnPOnline](https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html) — reformulado por cumplimiento). El código de `Connect-SPO` ya ramifica Interactive / Thumbprint / CertificatePath correctamente y NO mezcla Interactive con cert. El punto a corregir es la **precedencia**: hoy `if ($Interactive) {...} elseif ($Thumbprint)...`; si el usuario deja `Interactive=$true` en config Y además pasa Thumbprint, gana Interactive. Decisión: dar prioridad al modo app-only (Thumbprint / CertificatePath) por sobre Interactive cuando ambos estén presentes, para que la corrida de tenant completo no abra navegador por sitio. `Connect-SPO` ya se reutiliza en `Main` para cada sitio con los mismos parámetros de script-scope ($Interactive/$Thumbprint/$CertificatePath), así que al corregir la precedencia las conexiones per-site heredan el modo app-only automáticamente. Interactive sigue siendo fallback cuando no hay cert. Sin ClientId/tenant/thumbprint/path hardcodeados (todo por parámetro).

## HARD CONSTRAINTS (aplican a todos los ítems)

- El script sigue siendo **estrictamente de solo lectura**: NO introducir ningún cmdlet `Set-*`, `Remove-*`, `Add-*` de SPO (ni cualquier escritura). Solo lectura de métricas.
- **Todos los comentarios de código en español.**
- Verificar únicamente por parse estático (no hay tenant).

---

- [ ] 1. BUG 4 (LOW) — Quitar `-NoTypeInformation` de todos los `Export-Csv`.
      En las 5 funciones de export (`Export-SitesCsv`, `Export-LibrariesCsv`, `Export-FoldersCsv`, `Export-RecycleBinCsv`, `Export-ArchiveCandidatesCsv`) eliminar `-NoTypeInformation` (obsoleto en PS7, ya es default) y **conservar** `-Encoding UTF8`. Es el mismo patrón repetido 5 veces.
      Files: `Analyze-SPOStorage.ps1`
      Verify: parse estático imprime `PARSE OK`; y `grep -c "NoTypeInformation" Analyze-SPOStorage.ps1` devuelve 0 mientras `grep -c "Encoding UTF8" Analyze-SPOStorage.ps1` sigue en 5 (chequeo auxiliar, no sustituye al parse).

- [ ] 2. BUG 3 (MEDIUM) — Corregir el `US$` literal en `Show-Summary`.
      En las 3 líneas de costo (`Microsoft 365 Archive`, `Extra File Storage`, `Ahorro mensual`) reemplazar cada `US\$` por `` US`$ `` (backtick-dólar) para que imprima `US$` limpio, tanto en las etiquetas con la tarifa (`US\$0.05/GB`, `US\$0.20/GB`) como en los `US\$ {0}`. No cambiar la aritmética (`$costArchive = round(candidateGB*0.05,2)`, `$costExtra = round(candidateGB*0.20,2)`, `$ahorro = round(costExtra-costArchive,2)`) — ya es correcta; solo confirmar que sigue intacta.
      Files: `Analyze-SPOStorage.ps1`
      Verify: parse estático `PARSE OK`; y `grep -n 'US\\\$' Analyze-SPOStorage.ps1` ya no encuentra el backslash-dólar literal (debe devolver 0 coincidencias).

- [ ] 3. BUG 1 (HIGH) — Arreglar el doble conteo del "Recuperable estimado por versiones" en `Show-Summary`.
      Reemplazar `$recoverableBytes = ($Folders | Measure-Object -Property VersionOverheadBytes -Sum).Sum` por una suma que tome **solo las filas raíz por biblioteca** (`Depth -eq 0`), es decir una raíz por par `SiteUrl`+`Library`: filtrar `$Folders | Where-Object { $_.Depth -eq 0 }` y sumar `VersionOverheadBytes` de ese subconjunto (manejar `$null` → 0 como ya se hace). Agregar un comentario en español explicando que `Get-PnPFolderStorageMetric` es agregado (padre incluye hijos) y que por eso solo se suman las filas `Depth=0` para no contar el mismo overhead varias veces. Dejar explícito en comentario que los dos "Top 10" (tamaño y VersionOverhead) son drill-down jerárquico informativo y por eso NO se agregan. NO modificar `Get-FolderMetricsRecursive` ni `Export-FoldersCsv` (las filas por carpeta de `carpetas.csv` se conservan como están).
      Files: `Analyze-SPOStorage.ps1`
      Verify: parse estático `PARSE OK`. (No hay tenant: la correctitud numérica se validó por inspección contra `salida/carpetas.csv` — sumar solo las dos filas `Depth=0` da ~827 GB en vez de los 3095 GB inflados.)

- [ ] 4. BUG 2 (HIGH) — Candidatos a Archive siempre 0: refinar la fuente de fecha y agregar diagnóstico.
      En `Get-ArchiveCandidates`: tratar explícitamente fechas inutilizables — saltear cuando `$lastMod` es `$null` O `= [datetime]::MinValue`, y contar cuántos sitios tenían fecha usable vs. inutilizable, logueando el resultado (`Write-Log "Sitios con fecha de última modificación usable: X/Y" 'INFO'`). Añadir comentario en español documentando que `LastContentModifiedDate` de `Get-PnPTenantSite` la actualiza una tarea programada y puede ser imprecisa.
      En `Main`, para los Top N ya conectados, refinar la fecha por sitio con `Get-PnPWeb -Includes LastItemUserModifiedDate, LastItemModifiedDate` (envuelto en `Invoke-WithRetry`, dentro del `try` por sitio ya existente) y preferir `LastItemUserModifiedDate` sobre el valor tenant-wide cuando esté disponible; dejar comentario en español explicando que el refinamiento per-site solo cubre Top N (la detección global sigue usando el dato tenant-wide, menos fiable). Mantener la evaluación de inactividad a 6 y 12 meses. No romper el orden actual: `Get-ArchiveCandidates` se llama antes del loop de Top N, así que el refinamiento per-site debe actualizar los objetos de sitio (que son los mismos de `$candidates`/`$sites`) o recomputar candidatos tras el loop — elegir recomputar/actualizar de forma que `candidatos_archive.csv` y el resumen reflejen el dato refinado para los Top N. Documentar la elección en comentario español.
      Files: `Analyze-SPOStorage.ps1`
      Verify: parse estático `PARSE OK`. Confirmar por lectura que no se agregó ningún cmdlet de escritura (`Get-PnPWeb` es de lectura).

- [ ] 5. MAIN FEATURE (HIGH) — Priorizar el modo app-only (certificado) sobre Interactive en `Connect-SPO`.
      En `Connect-SPO` reordenar la cadena de autenticación para que, si se proveyó `-Thumbprint` o `-CertificatePath`, se conecte **app-only aunque `-Interactive` también esté presente** (hoy `if ($Interactive)` gana primero): cambiar a `if ($Thumbprint) {...} elseif ($CertificatePath) {...} elseif ($Interactive) {...} else { throw }`. Mantener en cada rama cert el `Tenant = "$Tenant.onmicrosoft.com"` y NO setear `Interactive` en las ramas de certificado. Verificar que la forma app-only coincide con PnP 3.4.1: `Connect-PnPOnline -Url <url> -ClientId <id> -Tenant <name>.onmicrosoft.com -Thumbprint <tp>` (y la variante `-CertificatePath`/`-CertificatePassword`). Agregar comentario en español explicando la precedencia (app-only evita popups de navegador por sitio en corridas de tenant completo; Interactive queda como fallback). Como `Main` ya llama `Connect-SPO` por sitio con los mismos parámetros script-scope, las conexiones per-site heredan el modo app-only sin cambios adicionales — confirmarlo por lectura y dejarlo anotado en comentario. Sin valores hardcodeados: todo sale de los parámetros del script.
      Files: `Analyze-SPOStorage.ps1`
      Verify: parse estático `PARSE OK`. Confirmar por lectura que ninguna rama de certificado setea `Interactive`, que no hay ClientId/tenant/thumbprint/path literales en el cuerpo de `Connect-SPO`, y que el `throw` sigue cubriendo el caso sin ningún método.

- [ ] 6. CONFIG — Actualizar `config.example.ps1` (solo placeholders).
      Reordenar/ampliar el bloque de autenticación para presentar PRIMERO el bloque de certificado/thumbprint como **recomendado** (con `Thumbprint = "REEMPLAZAR_THUMBPRINT"`), luego el bloque interactivo, luego el `.pfx`. Comentarios en español explicando el tradeoff (certificado = desatendido/app-only, sin popups, apto tenant completo; interactivo = cómodo pero delegado, lento y puede fallar por permisos). Solo UNA opción activa a la vez (dejar el certificado comentado y el interactivo activo, o viceversa, pero coherente con las instrucciones: placeholders, sin credenciales reales). Mantener `TenantName="REEMPLAZAR_TENANT"`, `ClientId="REEMPLAZAR_CLIENT_ID"`. Ajustar el comentario de `TopN` sugiriendo subirlo para auditorías completas una vez validado con cert.
      Files: `config.example.ps1`
      Verify: parse estático sobre `config.example.ps1` imprime `PARSE OK`; confirmar por lectura que no hay thumbprint/clientid/tenant reales.

- [ ] 7. CONFIG — Actualizar `config.ps1` (git-ignored, preservar valores reales).
      Preservar `TenantName = "novatechsolutionsar"` y `ClientId = "7dfd260e-3925-438d-b309-0f465d7d07f1"`. Dejar `Interactive = $true` activo por ahora. Agregar una línea `# Thumbprint = "REEMPLAZAR_THUMBPRINT"` comentada (SIN thumbprint real). Agregar comentario en español explicando cómo pasar a app-only (comentar `Interactive`, descomentar `Thumbprint` con el valor del certificado) y una nota sobre subir `TopN` (hoy 3) para auditar más sitios una vez validado el modo certificado. Mantener `OutputPath`, `MaxDepth`, `MinFolderSizeGB`. Verificar que `config.ps1` sigue en `.gitignore` (ya está).
      Files: `config.ps1`
      Verify: parse estático sobre `config.ps1` imprime `PARSE OK`; `git check-ignore config.ps1` confirma que sigue ignorado; confirmar por lectura que NO hay thumbprint real y que TenantName/ClientId quedaron intactos.

- [ ] 8. README — Documentar el modo certificado como recomendado para auditorías de tenant completo.
      En `README.md` (español): marcar la auth por certificado como RECOMENDADA para auditar todo el tenant; incluir el comando exacto `Register-PnPEntraIDApp` (ya presente, con placeholders) para una app nueva, Y el flujo `New-PnPAzureCertificate` + subida manual del `.cer` para una app EXISTENTE (placeholders). Explicar el tradeoff delegado (Interactive) vs aplicación (certificado): por qué el interactivo puede dar "unauthorized" (requiere ser admin del sitio y abre popup por sitio) y es lento (~50 min para 3 sitios, según la corrida real del usuario), mientras que app-only con `Sites.Read.All`/`Sites.FullControl.All` lee todo sin popups. Mantener las notas de seguridad existentes (no commitear `.pfx`/`config.ps1`/`salida/`, usar certificados y no secrets). Actualizar, si corresponde, la mención de `-TopN` default y la guía de ejecución para que la Opción B (thumbprint) figure como la vía recomendada para corridas grandes.
      Files: `README.md`
      Verify: lectura/manual — confirmar que el README documenta ambos flujos (app nueva y app existente), el tradeoff delegado-vs-aplicación, el motivo del "unauthorized"/lentitud, y que todo está en español con placeholders (sin valores reales). (README no se parsea; verificación por revisión.)

## Orden y dependencias

Ítems 1–5 son ediciones independientes del mismo archivo; hacerlas en orden (1→5) deja el script buildeable (parse OK) tras cada paso. Ítems 6–8 (config/README) no dependen del código pero su contenido debe ser coherente con la precedencia app-only del ítem 5, por eso van después. Cada ítem deja el repo en estado parseable.

## Supuestos / gaps

- No hay acceso al tenant: toda verificación numérica de BUG 1/BUG 2 es por inspección de los CSV de muestra y parse estático, no por ejecución real. Queda anotado para el revisor.
- Para BUG 2, los datos de muestra (`sitios.csv`) tienen fechas recientes (2026), por lo que incluso tras el arreglo el conteo de candidatos puede seguir siendo 0 con estos datos; el valor del fix está en el diagnóstico y en la fuente de fecha más fiable para corridas reales. Supuesto razonable, documentado en el comentario del código.
