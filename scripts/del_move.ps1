# =====================================================================
# del_move.ps1 - renomme le paquet produit par pbiviz en <name>.pbiviz
#
# Aucun nom ni GUID code en dur : les deux sont lus dans pbiviz.json,
# la seule source de verite du projet. Le script ne se perime donc pas
# quand un GUID ou un nom change.
#
# Fonctionne quel que soit son propre emplacement : il remonte
# l'arborescence jusqu'au dossier contenant pbiviz.json, puis travaille
# dans dist\. Utilisable depuis scripts\ comme depuis dist\.
#
# Appele par : npm run package
# =====================================================================

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# --- racine du projet : premier dossier parent contenant pbiviz.json ---
$racine = $scriptDir
while (-not (Test-Path (Join-Path $racine "pbiviz.json"))) {
    $parent = Split-Path -Parent $racine
    if ([string]::IsNullOrEmpty($parent) -or $parent -eq $racine) {
        Write-Error "pbiviz.json introuvable en remontant depuis $scriptDir"
        exit 1
    }
    $racine = $parent
}

# --- identite lue dans pbiviz.json -------------------------------------
# Trim() absorbe tout espace ou retour a la ligne parasite dans le JSON.
$pb   = Get-Content (Join-Path $racine "pbiviz.json") -Raw | ConvertFrom-Json
$nom  = ([string]$pb.visual.name).Trim()
$guid = ([string]$pb.visual.guid).Trim()

if ([string]::IsNullOrEmpty($nom) -or [string]::IsNullOrEmpty($guid)) {
    Write-Error "name ou guid absent de pbiviz.json"
    exit 1
}

$dist  = Join-Path $racine "dist"
$cible = Join-Path $dist ($nom + ".pbiviz")

if (-not (Test-Path $dist)) {
    Write-Error "Dossier dist introuvable : $dist"
    exit 1
}

# --- paquet a renommer -------------------------------------------------
$candidats = @(Get-ChildItem -Path $dist -Filter ($guid + "*.pbiviz") -File -ErrorAction SilentlyContinue)
if ($candidats.Count -eq 0) {
    Write-Error "Aucun fichier '$guid*.pbiviz' dans $dist"
    exit 1
}

# le plus recent, au cas ou plusieurs versions coexistent
$source = $candidats | Sort-Object LastWriteTime -Descending | Select-Object -First 1

# --- renommage ---------------------------------------------------------
if (Test-Path $cible) {
    Remove-Item $cible -Force
    Write-Host ("Supprime : " + $nom + ".pbiviz")
}

Rename-Item -Path $source.FullName -NewName ($nom + ".pbiviz")
Write-Host ("Renomme  : " + $source.Name + "  ->  " + $nom + ".pbiviz")
