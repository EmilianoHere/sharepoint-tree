# Nota de verificación — Analyze-SPOStorage.ps1

Entorno: macOS, binario PowerShell = `pwsh-preview` (NO `pwsh`), PnP.PowerShell 3.4.1.
NO se conectó al tenant real. Verificación **solo por parse estático** + inspección.

## 1. Parse estático del script (comando exacto solicitado)

```
pwsh-preview -NoProfile -Command '$t=$null;$e=$null;[System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path /Users/puser/repos/phinx/Arbitrage/emi/sharepoint-tree/Analyze-SPOStorage.ps1).Path,[ref]$t,[ref]$e)|Out-Null; if($e.Count){$e|%{Write-Host ("linea {0}: {1}" -f $_.Extent.StartLineNumber,$_.Message)};exit 1}else{Write-Host "OK parse"}'
```

Resultado: **`OK parse`** (exit 0, sin errores de sintaxis). Se corrió tras cada edición
del `.ps1`; la última corrida (estado final) también imprime `OK parse`.

También parsean sin errores los dos configs:
- `config.example.ps1` → `OK parse`
- `config.ps1` → `OK parse`

## 2. `git check-ignore config.ps1`

```
$ git check-ignore config.ps1 SPO-Storage-ReadOnly.cer SPO-Storage-ReadOnly.pfx
config.ps1
SPO-Storage-ReadOnly.cer
SPO-Storage-ReadOnly.pfx
```

`config.ps1` sigue git-ignored (y también el `.cer`/`.pfx`). `git status --short` NO
lista `config.ps1` ni los certificados, así que no se agregaron ni renombraron por error.

## 3. No se introdujo ningún cmdlet de escritura SPO (Set-/Remove-/Add-)

```
$ grep -nE '\b(Set|Remove|Add)-(PnP|SPO)' Analyze-SPOStorage.ps1
none
```

El script sigue siendo estrictamente de solo lectura. El único cmdlet nuevo agregado
(BUG 2) es `Get-PnPWeb -Includes LastItemUserModifiedDate, LastItemModifiedDate`, que es
de lectura.

## 4. Chequeos auxiliares

- `grep -c "NoTypeInformation" Analyze-SPOStorage.ps1` → **0** (BUG 4 OK).
- `grep -c "Encoding UTF8" Analyze-SPOStorage.ps1` → **5** (se conservó UTF8 en los 5 export).
- `grep -c 'US\$' ...` (backslash-dólar literal) → **0** (BUG 3 OK; ahora usa backtick-dólar).
- `config.example.ps1`: solo placeholders (`REEMPLAZAR_TENANT`, `REEMPLAZAR_CLIENT_ID`,
  `REEMPLAZAR_THUMBPRINT`), sin ClientId/tenant/thumbprint reales.
- `config.ps1`: preserva `TenantName = "novatechsolutionsar"` y
  `ClientId = "7dfd260e-3925-438d-b309-0f465d7d07f1"`, `Interactive = $true` activo,
  línea `# Thumbprint = "REEMPLAZAR_THUMBPRINT"` comentada (SIN thumbprint real).

## 5. Validación numérica de BUG 1 (por inspección de salida/carpetas.csv)

No hay tenant, así que la correctitud se validó leyendo `salida/carpetas.csv`. Las filas
`Depth=0` (raíces por biblioteca) son:
- `/Documentos compartidos` (NovaProduccion) → VersionOverhead 276,71 GB
- `Documentos compartidos` (root site)        → VersionOverhead 550,21 GB

Suma de raíces ≈ **826,92 GB**, en vez de los **3095,52 GB** inflados que daba sumar
todas las profundidades. El fix suma solo `Depth -eq 0`, lo que corrige el doble conteo.
