# del_move.ps1 — renomme le .pbiviz fraîchement généré en DynamicHeaderTable.pbiviz

# Le script s'exécute depuis le dossier où il se trouve (dist/)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $scriptDir

$target = "DynamicHeaderTable.pbiviz"
$prefix = "DynamicHeaderTable1A2B3C4D5E6F"

# Cherche tous les .pbiviz commençant par le prefix (peu importe la version)
$candidates = Get-ChildItem -Path $scriptDir -Filter "$prefix*.pbiviz" -File

if ($candidates.Count -eq 0) {
    Write-Error "No file '$prefix*.pbiviz' found in $scriptDir"
    exit 1
}

# Prend le plus récent (au cas où il y aurait plusieurs versions)
$source = ($candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1).Name

# Supprime l'ancien fichier cible s'il existe
if (Test-Path $target) {
    Remove-Item $target -Force
    Write-Host "Deleted : $target"
}

# Renomme le source en target
Rename-Item -Path $source -NewName $target
Write-Host "Renamed : $source -> $target"

