param(
  [switch]$RemoveVenv
)

$root = Split-Path -Parent $PSScriptRoot

Write-Host "Limpiando temporales en: $root"

# __pycache__
Get-ChildItem -Path $root -Recurse -Directory -Filter "__pycache__" | ForEach-Object {
  try {
    Remove-Item -Recurse -Force $_.FullName
    Write-Host "Borrado: $($_.FullName)"
  } catch {
    Write-Host "No se pudo borrar: $($_.FullName)"
  }
}

# exports_tmp
$exports = Join-Path $root "exports_tmp"
if (Test-Path $exports) {
  try {
    Remove-Item -Recurse -Force $exports
    Write-Host "Borrado: $exports"
  } catch {
    Write-Host "No se pudo borrar: $exports"
  }
}

# legend_error.log
$log = Join-Path $root "legend_error.log"
if (Test-Path $log) {
  try {
    Remove-Item -Force $log
    Write-Host "Borrado: $log"
  } catch {
    Write-Host "No se pudo borrar: $log"
  }
}

# tmp Excel files
$tmpDir = Join-Path $root "data\\legend\\plantillas"
if (Test-Path $tmpDir) {
  Get-ChildItem -Path $tmpDir -Filter "_tmp*.xlsx" | ForEach-Object {
    try {
      Remove-Item -Force $_.FullName
      Write-Host "Borrado: $($_.FullName)"
    } catch {
      Write-Host "No se pudo borrar: $($_.FullName)"
    }
  }
  Get-ChildItem -Path $tmpDir -Filter "~$*.xlsx" | ForEach-Object {
    try {
      Remove-Item -Force $_.FullName
      Write-Host "Borrado: $($_.FullName)"
    } catch {
      Write-Host "No se pudo borrar: $($_.FullName)"
    }
  }
}

if ($RemoveVenv) {
  $venv = Join-Path $root ".venv"
  if (Test-Path $venv) {
    try {
      Remove-Item -Recurse -Force $venv
      Write-Host "Borrado: $venv"
    } catch {
      Write-Host "No se pudo borrar: $venv"
    }
  }
} else {
  Write-Host "Nota: .venv NO se borró. Usa -RemoveVenv si quieres eliminarlo."
}

Write-Host "Limpieza terminada."
