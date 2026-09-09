# <DynamicHeaderTable>

Visuel Power BI custom (`.pbiviz`) — <une ligne de description>.

## Prérequis
- Node.js LTS
- PowerShell 7 (`pwsh`)
- `powerbi-visuals-tools` v7 : `npm i -g powerbi-visuals-tools`

## Développement
npm install
pbiviz start

## Build
pbiviz package   # génère dist/DynamicHeaderTable.pbiviz

## Structure
- src/          code TypeScript
- capabilities.json   dataRoles, dataViewMappings, objects
- pbiviz.json         métadonnées du visuel