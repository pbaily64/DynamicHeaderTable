"use strict";

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
const fm = formattingSettings;

// ════════════════════════════════════════════════════════════════════════════
//  CARTE : Style du tableau  (SimpleCard — inchangée)
// ════════════════════════════════════════════════════════════════════════════
class TableStyleCard extends fm.SimpleCard {
  name = "tableStyle";
  displayName = "Table style";

  headerBackground = new fm.ColorPicker({ name: "headerBackground", displayName: "Header background",           value: { value: "#C55A11" } });
  headerFontColor  = new fm.ColorPicker({ name: "headerFontColor",  displayName: "Header font color",  value: { value: "#ffffff" } });
  headerFontSize   = new fm.NumUpDown(  { name: "headerFontSize",   displayName: "Header font size",  value: 12 });
  rowBackground    = new fm.ColorPicker({ name: "rowBackground",    displayName: "Lines background (even)",    value: { value: "#ffffff" } });
  rowAltBackground = new fm.ColorPicker({ name: "rowAltBackground", displayName: "Lines background (odd)", value: { value: "#f2f2f2" } });
  rowFontColor     = new fm.ColorPicker({ name: "rowFontColor",     displayName: "Row font color",    value: { value: "#000000" } });
  rowFontSize      = new fm.NumUpDown(  { name: "rowFontSize",      displayName: "Row font size",    value: 11 });
  borderColor      = new fm.ColorPicker({ name: "borderColor",      displayName: "Border color",        value: { value: "#cccccc" } });
  gridLines        = new fm.ToggleSwitch({ name: "gridLines",       displayName: "Display grid lines",    value: true });

  slices = [
    this.headerBackground, this.headerFontColor, this.headerFontSize,
    this.rowBackground, this.rowAltBackground, this.rowFontColor, this.rowFontSize,
    this.borderColor, this.gridLines
  ];
}

// ════════════════════════════════════════════════════════════════════════════
//  CARTE : Pagination  (SimpleCard — inchangée)
// ════════════════════════════════════════════════════════════════════════════
class PaginationCard extends fm.SimpleCard {
  name = "pagination";
  displayName = "Pagination";

  enabled     = new fm.ToggleSwitch({ name: "enabled",     displayName: "Enable pagination", value: true });
  rowsPerPage = new fm.NumUpDown(   { name: "rowsPerPage", displayName: "Row per page",        value: 25 });

  topLevelSlice = this.enabled;
  slices = [ this.rowsPerPage ];
}

// ════════════════════════════════════════════════════════════════════════════
//  CARTE : Tri  (SimpleCard — inchangée)
// ════════════════════════════════════════════════════════════════════════════
class SortingCard extends fm.SimpleCard {
  name = "sorting";
  displayName = "Sorting";

  enabled = new fm.ToggleSwitch({ name: "enabled", displayName: "Enable column sorting", value: true });

  topLevelSlice = this.enabled;
  slices = [ ];
}

// ════════════════════════════════════════════════════════════════════════════
//  CARTE COMPOSITE : Gestion des colonnes
//  ┌─ Groupe « Ordre des colonnes »      (objet capabilities : columnOrder)
//  └─ Groupe « Formatage conditionnel »  (objet capabilities : conditionalFormatting)
//  Note : les libellés d'en-têtes restent pilotés par la mesure DAX `headerLabels`
//         (rôle de données), il n'y a donc pas de propriété Format associée.
// ════════════════════════════════════════════════════════════════════════════
class ColumnOrderGroup extends fm.Group {
  name = "columnOrder";
  displayName = "Column order";

  // Libellés d'en-têtes — exposé côté Format avec bouton fx pour y brancher
  // la mesure DAX (JSON {"colonne":"libellé"}). Reste lu via le rôle de données
  // headerLabels dans visual.ts ; ce champ sert de repli / point d'entrée fx.
  headerLabels = new fm.TextInput({
    name: "headerLabels",
    displayName: "Header labels (JSON / fx)",
    value: "",
    placeholder: "{\"Col1\":\"Label 1\"}",
    instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
  });

  // Ordre des colonnes — texte saisissable OU mesure DAX via bouton fx.
  orderJson = new fm.TextInput({
    name: "orderJson",
    displayName: "Column order (JSON / fx)",
    value: "[]",
    placeholder: "[\"Col1\",\"Col2\"]",
    instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
  });

  // Colonnes à masquer dans la grille — texte OU mesure DAX via bouton fx.
  // Ces colonnes restent dans le rôle "Colonnes" (données disponibles pour
  // filtrage / modale) mais ne sont pas rendues dans le tableau.
  hiddenColumns = new fm.TextInput({
    name: "hiddenColumns",
    displayName: "Hidden columns in the grid (JSON / fx)",
    value: "[]",
    placeholder: "[\"Col1\",\"Col2\"]",
    instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
  });

  // Format des colonnes — JSON {colonne:{type,format,decimals,minWidth,...}}.
  // Texte OU mesure DAX via bouton fx.
  columnFormats = new fm.TextInput({
    name: "columnFormats",
    displayName: "Column Format (JSON / fx)",
    value: "{}",
    placeholder: "{\"Col1\":{\"type\":\"number\",\"decimals\":2}}",
    instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
  });

  // Tri par défaut multi-colonnes — texte OU mesure DAX via bouton fx.
  // Appliqué tant qu'aucun tri manuel (clic en-tête) n'est actif.
  defaultSortJson = new fm.TextInput({
    name: "defaultSortJson",
    displayName: "Default sort (JSON / fx)",
    value: "[]",
    placeholder: "[{\"column\":\"Col1\",\"dir\":\"DESC\"}]",
    instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
  });

  slices = [ this.headerLabels, this.orderJson, this.hiddenColumns, this.columnFormats, this.defaultSortJson ];
}

class ConditionalFormattingGroup extends fm.Group {
  name = "conditionalFormatting";
  displayName = "Conditional formatting";

  enabled = new fm.ToggleSwitch({
    name: "cfEnabled",
    displayName: "Enable conditional formatting",
    value: false
  });

  rulesJson = new fm.TextInput({
    name: "cfRulesJson",
    displayName: "Rules (JSON / fx)",
    value: "[]",
    placeholder: "[{\"column\":\"...\",\"value\":\"...\",\"color\":\"#...\"}]",
    instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
  });

  topLevelSlice = this.enabled;
  slices = [ this.rulesJson ];
}

class ColumnManagementCard extends fm.CompositeCard {
  name = "columnManagement";
  displayName = "Column management";

  columnOrder           = new ColumnOrderGroup();
  conditionalFormatting = new ConditionalFormattingGroup();

  groups = [ this.columnOrder, this.conditionalFormatting ];
}

// ════════════════════════════════════════════════════════════════════════════
//  CARTE COMPOSITE : Détail de la Modale
//  ┌─ Groupe « Modale »            (titre, dimensions, disposition)
//  └─ Groupe « Ordre des champs »  (ordre des champs via texte / fx)
//  Note : titre, ordre des champs et libellés sont désormais pilotés côté
//         Format (texte OU mesure DAX via bouton fx), plus par rôle de données.
// ════════════════════════════════════════════════════════════════════════════
class ModalGroup extends fm.Group {
  name = "modal";
  displayName = "Modal";

  enabled   = new fm.ToggleSwitch({ name: "enabled",   displayName: "Enable the click-to-click modal", value: true });

  // Titre de la modale — texte saisissable OU mesure DAX via bouton fx.
  // Si vide, le titre reprend la valeur de la première colonne de la ligne.
  titleText = new fm.TextInput({
    name: "titleText",
    displayName: "Modal title (text / fx)",
    value: "",
    placeholder: "Detail",
    instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
  });

  // Taille du dialogue natif Power BI (openModalDialog), en pixels absolus.
  // Pas de "% de la fenêtre" : le visuel n'a aucun accès fiable à la vraie
  // taille de la fenêtre du navigateur (window.innerWidth/Height, lu depuis
  // le visuel, renvoie la taille de SON PROPRE iframe sur le canevas — pas
  // celle du rapport). 0 = demande une taille généreuse, ramenée de toute
  // façon par l'hôte à son vrai plafond (90% de la fenêtre réelle).
  width     = new fm.NumUpDown(   { name: "width",     displayName: "Width (px, 0 = maximum autorisé par l'hôte)",  value: 0 });
  height    = new fm.NumUpDown(   { name: "height",    displayName: "Height (px, 0 = maximum autorisé par l'hôte)", value: 0 });

  // Temporaire — à désactiver une fois le diagnostic terminé. Affiche un
  // panneau en tête de la modale avec les dimensions demandées vs mesurées.
  debugMode = new fm.ToggleSwitch({ name: "debugMode", displayName: "Debug mode (afficher les dimensions)", value: false });

  topLevelSlice = this.enabled;
  slices = [ this.titleText, this.width, this.height, this.debugMode ];
}

// Disposition positionnelle des champs de la modale : remplace entièrement
// l'ancien trio Field layout / Field direction / Field order (JSON).
// Un seul JSON obligatoire décrit à la fois la liste des champs affichés
// (exhaustive — seuls les champs listés apparaissent), leur libellé, et leur
// position en grille (row/col/colspan/rowspan). Voir validateModalLayout()
// dans visual.ts pour le schéma complet et les règles de validation.
class ModalLayoutGroup extends fm.Group {
  name = "modalLayout";
  displayName = "Layout";

  layoutJson = new fm.TextInput({
    name: "layoutJson",
    displayName: "Layout (JSON / fx)",
    value: "[]",
    placeholder: "[{\"field\":\"Col1\",\"labelField\":\"Label_Col1\",\"row\":1,\"col\":1,\"colspan\":2}]",
    instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
  });

  // Format des champs spécifique à la modale (indépendant de columnFormats).
  // Texte JSON OU mesure DAX via bouton fx.
  modalFieldFormats = new fm.TextInput({
    name: "modalFieldFormats",
    displayName: "Field format of the modal (JSON / fx)",
    value: "{}",
    placeholder: "{\"Col1\":{\"type\":\"date\",\"format\":\"dd/MM/yyyy\"}}",
    instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
  });

  slices = [ this.layoutJson, this.modalFieldFormats ];
}

class ModalDetailCard extends fm.CompositeCard {
  name = "modalDetail";
  displayName = "Modal detail";

  modal       = new ModalGroup();
  modalLayout = new ModalLayoutGroup();

  groups = [ this.modal, this.modalLayout ];
}

// ════════════════════════════════════════════════════════════════════════════
//  CARTE COMPOSITE : Sélection de ligne
//  ┌─ Groupe « Apparence »  : activation + style visuel de la ligne sélectionnée
//  └─ Groupe « Filtrage »   : config du cross-filtering (table/champs cible)
//
//  Si la table cible (groupe Filtrage) est vide, on reste sur le SelectionManager
//  natif (cross-filter Power BI via relations). Sinon, on applique un filtre
//  JSON explicite sur la table cible — utile quand aucune relation native
//  n'existe ou que les champs visibles ne sont pas les bons pour le filtrage.
// ════════════════════════════════════════════════════════════════════════════
class RowSelectionAppearanceGroup extends fm.Group {
  name = "appearance";
  displayName = "Appearance";

  enabled = new fm.ToggleSwitch({
    name: "enabled",
    displayName: "Enable line selection",
    value: true
  });

  borderColor = new fm.ColorPicker({
    name: "borderColor",
    displayName: "Border color",
    value: { value: "#C55A11" }
  });

  borderWidth = new fm.NumUpDown({
    name: "borderWidth",
    displayName: "Border width (px)",
    value: 3
  });

  backgroundColor = new fm.ColorPicker({
    name: "backgroundColor",
    displayName: "Background color",
    value: { value: "#C55A11" }
  });

  backgroundOpacity = new fm.NumUpDown({
    name: "backgroundOpacity",
    displayName: "Background opacity (0-100 %)",
    value: 15
  });

  topLevelSlice = this.enabled;
  slices = [ this.borderColor, this.borderWidth, this.backgroundColor, this.backgroundOpacity ];
}

class RowSelectionFilteringGroup extends fm.Group {
  name = "filtering";
  displayName = "Filtering";

  // Colonne dans le tableau dont la valeur sert de clé de filtre (peut être masquée).
  filterKey = new fm.TextInput({
    name: "filterKey",
    displayName: "Filter key (source)",
    value: "",
    placeholder: "[]",
    instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
  });

  // Table Power BI à filtrer (nom exact tel qu'il apparaît dans le modèle).
  targetTable = new fm.TextInput({
    name: "targetTable",
    displayName: "Target table",
    value: "",
    placeholder: "[]",
    instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
  });

  // Colonne dans la table cible sur laquelle appliquer le filtre.
  targetColumn = new fm.TextInput({
    name: "targetColumn",
    displayName: "Target column",
    value: "",
    placeholder: "[]",
    instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
  });

  slices = [ this.filterKey, this.targetTable, this.targetColumn ];
}

class RowSelectionRowFilterGroup extends fm.Group {
  name = "rowFilter";
  displayName = "Row filter";

  // Filtre d'affichage uniquement : les lignes exclues disparaissent de la
  // grille, du tri et de la pagination — aucun filtre n'est envoyé au modèle.
  enabled = new fm.ToggleSwitch({
    name: "rowFilterEnabled",
    displayName: "Enable row filter",
    value: false
  });

  // Règles — texte JSON OU mesure DAX via bouton fx. Combinées en AND.
  // La colonne testée peut être masquée dans la grille (hiddenColumns).
  rulesJson = new fm.TextInput({
    name: "rowFilterJson",
    displayName: "Filter rules (JSON / fx)",
    value: "[]",
    placeholder: "[{\"column\":\"Status\",\"mode\":\"exclude\",\"values\":[\"Deleted\"]}]",
    instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
  });

  topLevelSlice = this.enabled;
  slices = [ this.rulesJson ];
}

class RowSelectionCard extends fm.CompositeCard {
  name = "rowSelection";
  displayName = "Row selection";

  appearance = new RowSelectionAppearanceGroup();
  filtering  = new RowSelectionFilteringGroup();
  rowFilter  = new RowSelectionRowFilterGroup();

  groups = [ this.appearance, this.filtering, this.rowFilter ];
}

// ════════════════════════════════════════════════════════════════════════════
//  MODÈLE GLOBAL
// ════════════════════════════════════════════════════════════════════════════
export class VisualFormattingSettingsModel extends fm.Model {
  tableStyle       = new TableStyleCard();
  pagination       = new PaginationCard();
  sorting          = new SortingCard();
  columnManagement = new ColumnManagementCard();
  modalDetail      = new ModalDetailCard();
  rowSelection     = new RowSelectionCard();

  cards = [
    this.tableStyle,
    this.pagination,
    this.sorting,
    this.columnManagement,
    this.modalDetail,
    this.rowSelection
  ];
}
