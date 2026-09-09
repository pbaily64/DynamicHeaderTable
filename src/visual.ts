"use strict";

import powerbi from "powerbi-visuals-api";
import DataView = powerbi.DataView;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.extensibility.ISelectionId;

import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { VisualFormattingSettingsModel } from "./settings";
import DialogAction = powerbi.DialogAction;
import { DetailDialog, ModalDialogState, ModalDialogGridItem, ModalDialogInitialState, ModalDialogDebugInfo } from "./detailDialog";
import "./detailDialog"; // side-effect : enregistre DetailDialog dans globalThis.dialogRegistry

import "./../style/visual.less";

// ─── Interfaces ───────────────────────────────────────────────────────────────

interface CondRule {
  column: string; min?: number; max?: number;
  value?: string | number; color: string; fontColor?: string;
}

// Format par colonne, piloté par le champ "Format des colonnes" (JSON / fx)
interface ColFormat {
  type?: "text" | "date" | "number" | "boolean" | "phone";
  kind?: "fixed" | "mobile"; // phone : fixe (9 chiffres) ou GSM (10 chiffres)
  format?:         string;   // masque date (dd/MM/yyyy HH:mm:ss) ou symbole devise (€, $)
  utc?:            boolean;  // true = lire la date en UTC (aucun décalage de fuseau)
  autoTime?:       boolean;  // false = toujours rendre la partie horaire du masque (défaut : true)
  decimals?:       number;   // nombre de décimales pour les nombres
  width?:          number;   // largeur fixe en px
  minWidth?:       number;   // largeur minimale en px
  trueLabel?:      string;   // libellé affiché pour true (boolean)
  falseLabel?:     string;   // libellé affiché pour false (boolean)
  trueColor?:      string;   // fond cellule si true
  falseColor?:     string;   // fond cellule si false
  trueFontColor?:  string;   // couleur texte si true
  falseFontColor?: string;   // couleur texte si false
  multiline?:      boolean;  // true = respecte les sauts de ligne (UNICHAR(10)) et autorise le retour auto
}

interface ColDef {
  displayName: string;
  queryName:   string;
  format:      string;
  isHidden:    boolean;
  isTitle:     boolean;
  isMeasure:   boolean;       // distingue mesure (values) vs colonne (categories)
  values:      any[];
}

// Spécification de tri par défaut (champ "Default sort" / fx)
// [{ "column": "Col1", "dir": "DESC" }, { "column": "Col2", "dir": "ASC" }]
interface SortSpec {
  column: string;
  dir?:   string;             // "ASC" (défaut) ou "DESC", insensible à la casse
}

// Règle de filtre d'affichage des lignes (champ "Filter rules" / fx)
// [{ "column": "Status", "mode": "exclude", "values": ["Deleted","Cancelled"] }]
interface RowFilterRule {
  column: string;
  mode?:  "exclude" | "include";   // défaut : "exclude"
  values: any[];
}

// ── Layout positionnel de la modale (champ "Layout (JSON / fx)") ─────────────
// Trois types d'entrée dans le même tableau :
//   - champ normal ("type" absent, implicite)  : positionné via field/row/col.
//   - séparateur  ("type": "divider")          : occupe TOUTE sa "row", avec
//     un "label" de section optionnel. Pas de field/col/colspan/rowspan.
//   - liste       ("type": "list")              : sous-table répétée occupant
//     TOUTE sa "row" — pour une relation un-à-plusieurs (ex. plusieurs
//     propriétaires). Chaque colonne référence un champ retournant une
//     chaîne délimitée (convention CONCATENATEX/FILTER habituelle) ; le
//     visuel découpe et aligne chaque colonne ligne par ligne.
interface ModalLayoutFieldSpec {
  field:       string;
  labelField?: string;   // nom d'une colonne/mesure (ex. déposée en Hidden columns) fournissant un libellé résolu ligne par ligne — multilingue via DAX
  label?:      string;   // libellé littéral, statique (repli si labelField absent/vide pour cette ligne)
  row:         number;
  col:         number;
  colspan?:    number;
  rowspan?:    number;
}

// Colonne d'une entrée "list" — mêmes règles de résolution de libellé que
// pour un champ normal, mais sans positionnement propre (l'ordre du tableau
// "columns" détermine l'ordre d'affichage gauche → droite).
interface ModalLayoutListColumnSpec {
  field:       string;
  labelField?: string;
  label?:      string;
}

// Entrée résolue (field/labelField → ColDef réels) après validation réussie.
// Union discriminée : un champ positionné, un séparateur pleine largeur, ou
// une liste (sous-table pleine largeur).
type ResolvedModalLayoutItem =
  | {
      kind:           "field";
      fieldCol:       ColDef;
      labelFieldCol?: ColDef;
      label?:         string;
      row:            number;
      col:            number;
      colspan:        number;
      rowspan:        number;
    }
  | {
      kind:  "divider";
      row:   number;
      label?: string;
    }
  | {
      kind:      "list";
      row:       number;
      delimiter: string;
      columns:   { fieldCol: ColDef; labelFieldCol?: ColDef; label?: string }[];
    };

// Résultat de la validation : un seul des trois états, jamais un mélange.
type ModalLayoutValidation =
  | { kind: "empty" }
  | { kind: "error"; errors: string[] }
  | { kind: "ok"; columns: number; entries: ResolvedModalLayoutItem[] };

// ─── Visual ───────────────────────────────────────────────────────────────────

export class DynamicHeaderTable implements IVisual {
  private host:         IVisualHost;
  private container:    HTMLElement;
  private events:       IVisualEventService;
  private fmService:    FormattingSettingsService;
  private fmModel:      VisualFormattingSettingsModel;

  private sortCol:      number  = -1;
  private sortAsc:      boolean = true;
  private currentPage:  number  = 0;

  private allCols:      ColDef[] = [];
  private rowCount:     number   = 0;

  // ── CACHE pour le correctif 1.a ───────────────────────────────────────────
  // Le champ Format "Libellés d'en-têtes" (fx) est lu à chaque update. Si sa
  // valeur se retrouve momentanément vide, on rejoue le dernier mapping connu.
  private lastLabelsMap:  Record<string, string> = {};   // mapping JSON résolu
  private lastLabelByKey: Record<string, string> = {};   // queryName → libellé final

  // Map des formats par colonne (champ "Format des colonnes" / fx)
  private colFormatsMap: Record<string, ColFormat> = {};

  // Map des formats spécifiques à la modale (champ "Format des champs modale" / fx)
  private modalFormatsMap: Record<string, ColFormat> = {};

  // ── Sélection de ligne ────────────────────────────────────────────────────────
  // selectionManager propage la sélection au reste du rapport (cross-filtering).
  // rowSelectionIds[rowIdx] = SelectionId construit pour cette ligne.
  // selectedRowIdx = index de la ligne actuellement sélectionnée (-1 = aucune).
  private selectionManager: ISelectionManager;
  private rowSelectionIds:  ISelectionId[] = [];
  private selectedRowIdx:   number = -1;


  // ── Constructor ──────────────────────────────────────────────────────────────
  constructor(options: VisualConstructorOptions) {
    this.host      = options.host;
    this.container = options.element;
    this.events    = options.host.eventService;
    this.fmService = new FormattingSettingsService();
    this.fmModel   = new VisualFormattingSettingsModel();

    // SelectionManager : propage la sélection vers les autres visuels du rapport.
    this.selectionManager = this.host.createSelectionManager();
    // Power BI rappelle ce callback quand la sélection externe change (par ex.
    // clic en dehors du visuel → restoreSelection). On nettoie l'état visuel
    // si plus aucune sélection n'est active.
    try {
      this.selectionManager.registerOnSelectCallback(() => {
        const current = this.selectionManager.getSelectionIds();
        if (!current || current.length === 0) {
          if (this.selectedRowIdx !== -1) {
            this.selectedRowIdx = -1;
            this.render();
          }
        }
      });
    } catch { /* API absente sur très vieux runtimes — ignoré */ }

    this.container.style.cssText =
      "overflow:hidden;font-family:'Segoe UI',sans-serif;position:relative;";
  }

  // ── Utility ───────────────────────────────────────────────────────────────────
  private clearElement(el: HTMLElement): void {
    while (el.firstChild) el.removeChild(el.firstChild);
  }

  private getColor(picker: any): string {
    return picker?.value?.value ?? picker?.value ?? "#000000";
  }

  // Conversion hex (#RRGGBB ou #RGB) + alpha (0..1) vers rgba()
  private hexToRgba(hex: string, alpha: number): string {
    if (!hex) return "rgba(0,0,0," + alpha + ")";
    let h = hex.replace("#", "").trim();
    if (h.length === 3) h = h.split("").map(c => c + c).join("");
    if (h.length !== 6) return "rgba(0,0,0," + alpha + ")";
    const r = parseInt(h.substring(0, 2), 16);
    const g = parseInt(h.substring(2, 4), 16);
    const b = parseInt(h.substring(4, 6), 16);
    return "rgba(" + r + "," + g + "," + b + "," + alpha + ")";
  }

  // ── Sélection de ligne ────────────────────────────────────────────────────────
  //
  //  Logique :
  //    1. Re-clic sur la ligne sélectionnée → déselection (suppression du filtre).
  //    2. Config incomplète (filterKey / targetTable / targetColumn vides)
  //       → fallback sur le SelectionManager natif (cross-filter via relations).
  //    3. Config complète → filtre Basic JSON sur la colonne clé, canal "general"
  //       (niveau visuel/page — seul canal honoré par l'hôte pour un visuel custom).
  //       La propagation vers d'autres pages se fait nativement via
  //       "Sync slicers" (rendue possible par supportsSynchronizingFilterState).
  //
  private handleRowClick(origIdx: number): void {
    const filt        = this.fmModel.rowSelection.filtering;
    const filterKey   = this.cleanName(filt.filterKey.value    as string);
    const targetTable = this.cleanName(filt.targetTable.value  as string);
    const targetCol   = this.cleanName(filt.targetColumn.value as string);

    // Re-clic même ligne → désélection
    if (this.selectedRowIdx === origIdx) {
      this.clearActiveFilter(targetTable);
      this.selectedRowIdx = -1;
      this.render();
      return;
    }

    // Config incomplète → sélection native (SelectionManager)
    if (!filterKey || !targetTable || !targetCol) {
      const id = this.rowSelectionIds[origIdx];
      if (id) {
        try { this.selectionManager.select(id, false); } catch { /* noop */ }
      }
      this.selectedRowIdx = origIdx;
      this.render();
      return;
    }

    // Lecture de la valeur clé (colonne visible ou masquée dans la grille)
    const keyCol = this.findCol(filterKey);
    if (!keyCol) return;
    const rawVal = keyCol.values[origIdx];
    if (rawVal == null) return;

    try { this.selectionManager.clear(); } catch { /* noop */ }

    const filter = this.buildBasicFilter(targetTable, targetCol, String(rawVal));
    this.host.applyJsonFilter(filter, "general", "filter", powerbi.FilterAction.merge);

    this.selectedRowIdx = origIdx;
    this.render();
  }

  // Suppression propre du filtre actif (JSON ou natif selon la config)
  private clearActiveFilter(targetTable: string): void {
    if (targetTable) {
      this.host.applyJsonFilter(null, "general", "filter", powerbi.FilterAction.remove);
    }
    try { this.selectionManager.clear(); } catch { /* noop */ }
  }

  // Normalise un nom saisi dans le Format pane : retire crochets DAX [Nom],
  // apostrophes 'Nom' et guillemets "Nom", puis trim. Rend la config tolérante
  // à la notation DAX habituelle.
  private cleanName(raw: string): string {
    let s = (raw || "").trim();
    if (s.startsWith("[") && s.endsWith("]")) s = s.slice(1, -1).trim();
    if ((s.startsWith("'")  && s.endsWith("'")) ||
        (s.startsWith("\"") && s.endsWith("\""))) {
      s = s.slice(1, -1).trim();
    }
    return s;
  }

  // Correspondance souple : tolère displayName exact, queryName exact,
  // nom court sans préfixe table ("Table.Col" → "Col"), insensible à la casse.
  private findCol(name: string): ColDef | undefined {
    const lo = name.toLowerCase();
    return this.allCols.find(c =>
      c.displayName === name ||
      c.queryName   === name ||
      c.queryName.split(".").pop()?.toLowerCase() === lo ||
      c.displayName.toLowerCase()                 === lo
    );
  }

  // Filtre Basic Power BI (1 table, 1 colonne, 1 valeur, opérateur "In")
  // La valeur est envoyée dans son type natif : number si numérique, string sinon.
  private buildBasicFilter(table: string, column: string, value: string): any {
    const num = Number(value);
    const typedValue = (!isNaN(num) && value.trim() !== "") ? num : value;
    return {
      $schema:    "https://powerbi.com/product/schema#basic",
      target:     { table, column },
      operator:   "In",
      values:     [typedValue],
      filterType: 1
    };
  }

  // ── Extract data from categorical dataView ────────────────────────────────────
  private extractData(dv: DataView): void {
    this.allCols       = [];
    this.rowCount      = 0;

    if (!dv?.categorical) return;
    const cat = dv.categorical;

    // Liste des colonnes à masquer dans la grille (champ Format / fx)
    const hiddenSet = this.getHiddenColumnsSet();

    // ── Categories → colonnes Grouping (rôle "columns") ───────────────────────
    if (cat.categories) {
      for (const c of cat.categories) {
        const roles = c.source?.roles || {};
        if (roles["columns"]) {
          const dn = c.source.displayName;
          const qn = c.source.queryName || dn;
          this.allCols.push({
            displayName: dn,
            queryName:   qn,
            format:      c.source.format || "",
            isHidden:    hiddenSet.has(dn) || hiddenSet.has(qn),
            isTitle:     false, isMeasure: false,
            values:      c.values as any[],
          });
          this.rowCount = Math.max(this.rowCount, c.values.length);
        }
      }
    }

    // ── Values → mesures (rôle "columns") ─────────────────────────────────────
    if (cat.values) {
      for (const col of cat.values) {
        const roles = col.source?.roles || {};
        const vals  = col.values as any[];

        if (roles["columns"]) {
          const dn = col.source.displayName;
          const qn = col.source.queryName || dn;
          this.allCols.push({
            displayName: dn,
            queryName:   qn,
            format:      col.source.format || "",
            isHidden:    hiddenSet.has(dn) || hiddenSet.has(qn),
            isTitle:     false, isMeasure: true,
            values:      vals,
          });
          this.rowCount = Math.max(this.rowCount, vals.length);
        }
      }
    }

    // ── SelectionId par ligne (cross-filtering) ───────────────────────────────
    // On construit un SelectionId pour chaque ligne en s'appuyant sur la
    // première catégorie Grouping disponible. Si aucune catégorie Grouping
    // (modèle uniquement composé de mesures), la sélection ne pourra pas se
    // propager : c'est une limite native de l'API Power BI.
    this.rowSelectionIds = [];
    const groupingCat = cat.categories?.find(c => c.source?.roles?.["columns"]);
    if (groupingCat) {
      for (let i = 0; i < this.rowCount; i++) {
        const id = this.host.createSelectionIdBuilder()
          .withCategory(groupingCat, i)
          .createSelectionId();
        this.rowSelectionIds.push(id);
      }
    }
  }

  // ── Colonnes à masquer dans la grille (depuis le champ Format / fx) ───────────
  private getHiddenColumnsSet(): Set<string> {
    const set = new Set<string>();
    const raw = (this.fmModel.columnManagement.columnOrder.hiddenColumns.value as string) || "[]";
    let list: string[] = [];
    try { list = JSON.parse(raw || "[]"); } catch { list = []; }
    if (Array.isArray(list)) {
      for (const name of list) {
        if (name !== null && name !== undefined && String(name).trim() !== "") {
          set.add(String(name));
        }
      }
    }
    return set;
  }

  // ── Résolution du mapping libellés AVEC cache (correctif 1.a) ─────────────────
  // ── Résolution du mapping libellés AVEC cache (correctif 1.a) ─────────────────
  // Source unique : champ Format "Libellés d'en-têtes" (constante OU mesure via
  // fx). Comme ce champ est lu à chaque update indépendamment du contexte de
  // données, il reste disponible même quand le tableau est vide. Le cache sert
  // de filet de sécurité si la valeur fx se retrouve momentanément vide.
  private resolveLabelsMap(): Record<string, string> {
    let parsed: Record<string, string> = {};

    const fmtVal = (this.fmModel.columnManagement.columnOrder.headerLabels.value as string) || "";
    if (fmtVal.trim() !== "") {
      try {
        const fmtParsed = JSON.parse(fmtVal);
        if (fmtParsed && typeof fmtParsed === "object") parsed = fmtParsed;
      } catch { /* valeur non-JSON ignorée */ }
    }

    const hasContent = parsed && Object.keys(parsed).length > 0;

    if (hasContent) {
      this.lastLabelsMap = parsed;
      return parsed;
    }
    return this.lastLabelsMap;
  }

  // ── Résolution du libellé d'UNE colonne, AVEC cache par queryName ─────────────
  // Priorité : mapping DAX (courant ou caché) → cache par queryName → displayName.
  // Pour une mesure dont le contexte est vide, displayName == nom technique :
  // on évite donc de l'afficher si un libellé caché existe.
  private resolveLabel(col: ColDef, map: Record<string, string>): string {
    const fromMap = map[col.displayName] ?? map[col.queryName];
    if (fromMap !== undefined && fromMap !== null && String(fromMap).trim() !== "") {
      this.lastLabelByKey[col.queryName] = fromMap;   // alimente le cache
      return fromMap;
    }

    // Pas de libellé dans le mapping. Pour une mesure, le displayName est le
    // nom technique → on préfère le dernier libellé connu si disponible.
    if (col.isMeasure) {
      const cached = this.lastLabelByKey[col.queryName];
      if (cached !== undefined && cached !== null && String(cached).trim() !== "") {
        return cached;
      }
    }

    // Colonne Grouping (ou aucun cache) : le displayName est fiable.
    return col.displayName;
  }

  // ── Résolution de labelField (layout JSON de la modale) ───────────────────────
  // Contrairement à resolveLabel (mapping headerLabels, identique pour toutes
  // les lignes), labelField pointe vers une colonne/mesure dont la VALEUR à la
  // ligne cliquée sert de libellé — c'est là que vit la logique multilingue
  // (_langue_active) côté DAX. Valeur nulle/vide → undefined (repli au niveau
  // supérieur, vers label littéral puis headerLabels).
  private resolveLabelFieldValue(col: ColDef | undefined, rowIdx: number): string | undefined {
    if (!col) return undefined;
    const v = col.values[rowIdx];
    if (v === null || v === undefined) return undefined;
    const s = String(v);
    return s.trim() !== "" ? s : undefined;
  }

  // ── Apply column order (table) ────────────────────────────────────────────────
  private applyColumnOrder(cols: ColDef[], orderJson: string): ColDef[] {
    let orderList: string[] = [];
    try { orderList = JSON.parse(orderJson || "[]"); } catch { /* noop */ }
    if (!orderList.length) return cols;

    const ordered: ColDef[] = [];
    for (const name of orderList) {
      const found = cols.find(c => c.displayName === name || c.queryName === name);
      if (found && !ordered.includes(found)) ordered.push(found);
    }
    for (const c of cols) { if (!ordered.includes(c)) ordered.push(c); }
    return ordered;
  }

  // ── Format des colonnes (grille) et des champs modale (sources indépendantes) ─
  private loadColumnFormats(): void {
    this.colFormatsMap = {};
    this.modalFormatsMap = {};

    const rawGrid = (this.fmModel.columnManagement.columnOrder.columnFormats.value as string) || "{}";
    try {
      const parsed = JSON.parse(rawGrid || "{}");
      if (parsed && typeof parsed === "object") this.colFormatsMap = parsed;
    } catch { /* JSON invalide ignoré */ }

    const rawModal = (this.fmModel.modalDetail.modalLayout.modalFieldFormats.value as string) || "{}";
    try {
      const parsed = JSON.parse(rawModal || "{}");
      if (parsed && typeof parsed === "object") this.modalFormatsMap = parsed;
    } catch { /* JSON invalide ignoré */ }
  }

  private resolveFormat(displayName: string, queryName: string): ColFormat | undefined {
    return this.colFormatsMap[displayName] ?? this.colFormatsMap[queryName];
  }

  // Format pour la modale : strictement la map modale, pas de fallback grille.
  private resolveModalFormat(displayName: string, queryName: string): ColFormat | undefined {
    return this.modalFormatsMap[displayName] ?? this.modalFormatsMap[queryName];
  }

// ── Téléphonie belge ─────────────────────────────────────────────────────────
  // Normalise puis met en forme un numéro belge. Le "kind" est TOUJOURS déclaré
  // explicitement par la mesure DAX — aucune déduction sur la longueur.
  // fixed : 9 chiffres → 02/218.42.16 (zone 2 chiffres : 02, 03, 04, 09)
  // 071/12.34.56 (zone 3 chiffres : tout le reste)
  // mobile : 10 chiffres → 0479/60.73.91
  // Le 0 initial est réintroduit s'il manque (cas fréquent en base).
  // Toute valeur non conforme est renvoyée telle quelle : une anomalie de
  // données doit rester visible, pas être masquée par un formatage silencieux.
  private formatPhone(cell: any, kind: string | undefined): string {
    const digits = String(cell).replace(/\D/g, "");
    if (digits === "") return String(cell);

    // Préfixe international éventuel (0032 / 32)
    let base = digits;
    if (base.startsWith("0032")) base = base.slice(4);
    else if (base.startsWith("32") && (base.length === 10 || base.length === 11)) base = base.slice(2);

    const n = base.startsWith("0") ? base : "0" + base;

    if (kind === "mobile") {
      if (n.length !== 10 || !n.startsWith("04")) return String(cell);
      return n.slice(0, 4) + "/" + n.slice(4, 6) + "." + n.slice(6, 8) + "." + n.slice(8, 10);
    }

    if (kind === "fixed") {
      if (n.length !== 9) return String(cell);
      const shortZone = ["02", "03", "04", "09"].indexOf(n.slice(0, 2)) >= 0;
      const z = shortZone ? 2 : 3;
      return n.slice(0, z) + "/" + n.slice(z, z + 3) + "." + n.slice(z + 3, z + 5) + "." + n.slice(z + 5, 9);
    }

    return String(cell);
  }


  // ── Formatage de cellule (avec ColFormat optionnel) ───────────────────────────
  private formatCell(cell: any, fallbackFormat: string, fmt?: ColFormat): string {
    if (cell === null || cell === undefined) return "";

    const type = fmt?.type;

    // Booléen → libellés personnalisés
    if (type === "boolean") {
      const truthy = cell === true || cell === 1 || String(cell).toLowerCase() === "true";
      return truthy ? (fmt?.trueLabel ?? "✓") : (fmt?.falseLabel ?? "✗");
    }

    // Date → masque piloté par le JSON de format (jetons date ET heure)
    if (type === "date") {
      const d = (cell instanceof Date) ? cell : new Date(cell);
      if (isNaN(d.getTime())) return String(cell);
      const utc = fmt?.utc === true;
      let mask  = fmt?.format || "dd/MM/yyyy";
      // Le masque contient une partie horaire mais la valeur n'en a pas → on la retire
      if (fmt?.autoTime !== false && !this.hasTimeComponent(cell, utc)) {
        mask = this.stripTimeTokens(mask);
      }
      return this.formatDate(d, mask, utc);
    }

    // Téléphone → mise en forme belge (kind explicite : "fixed" | "mobile")
    if (type === "phone") {
      return this.formatPhone(cell, fmt?.kind);
    }

    // Nombre → décimales + symbole devise / pourcentage
    if (type === "number" || typeof cell === "number") {
      const num = typeof cell === "number" ? cell : parseFloat(String(cell));
      if (isNaN(num)) return String(cell);
      const dec = fmt?.decimals;
      const f   = fmt?.format ?? fallbackFormat ?? "";
      try {
        if (f.includes("%")) {
          const opts: any = dec != null ? { minimumFractionDigits: dec, maximumFractionDigits: dec } : {};
          return (num * 100).toLocaleString("fr-BE", opts) + " %";
        }
        if (f.includes("€") || f.includes("$")) {
          const opts: any = {
            style: "currency",
            currency: f.includes("€") ? "EUR" : "USD"
          };
          if (dec != null) { opts.minimumFractionDigits = dec; opts.maximumFractionDigits = dec; }
          return num.toLocaleString("fr-BE", opts);
        }
      } catch { /* noop */ }
      const opts: any = dec != null ? { minimumFractionDigits: dec, maximumFractionDigits: dec } : {};
      return num.toLocaleString("fr-BE", opts);
    }

    // Date native sans type explicite
    if (cell instanceof Date) return cell.toLocaleDateString("fr-BE");

    return String(cell);
  }

  // Formatage de date selon masque : yyyy, yy, MM, M, dd, d, HH, H, mm, m, ss, s
  // Passe unique et sensible à la casse : un jeton déjà substitué ne peut pas être
  // re-matché, et MM (mois) ne peut pas être confondu avec mm (minutes).
  private formatDate(d: Date, mask: string, utc: boolean = false): string {
    const p2 = (n: number) => String(n).padStart(2, "0");
    const Y  = utc ? d.getUTCFullYear() : d.getFullYear();
    const Mo = (utc ? d.getUTCMonth()   : d.getMonth()) + 1;
    const D  = utc ? d.getUTCDate()     : d.getDate();
    const H  = utc ? d.getUTCHours()    : d.getHours();
    const Mi = utc ? d.getUTCMinutes()  : d.getMinutes();
    const S  = utc ? d.getUTCSeconds()  : d.getSeconds();

    const map: Record<string, string> = {
      yyyy: String(Y), yy: String(Y).slice(-2),
      MM: p2(Mo), M: String(Mo),
      dd: p2(D),  d: String(D),
      HH: p2(H),  H: String(H),
      mm: p2(Mi), m: String(Mi),
      ss: p2(S),  s: String(S)
    };
    return mask.replace(/yyyy|yy|MM|M|dd|d|HH|H|mm|m|ss|s/g, (t) => map[t]);
  }

  // Une heure est-elle réellement présente dans la valeur source ?
  private hasTimeComponent(cell: any, utc: boolean): boolean {
    if (cell instanceof Date) {
      return utc
        ? (cell.getUTCHours() || cell.getUTCMinutes() || cell.getUTCSeconds() || cell.getUTCMilliseconds()) !== 0
        : (cell.getHours()    || cell.getMinutes()    || cell.getSeconds()    || cell.getMilliseconds())    !== 0;
    }
    if (typeof cell === "number") return cell % 1 !== 0;              // date sérialisée OLE
    const s = String(cell);
    if (!/\d{1,2}:\d{2}/.test(s)) return false;                       // "2006-10-09" → pas d'heure
    return !/[T ]00:00:00(\.0+)?(Z|[+-]\d{2}:?\d{2})?$/.test(s);      // minuit pile → date seule
  }

  // Coupe le masque avant le premier jeton horaire et nettoie le séparateur restant
  private stripTimeTokens(mask: string): string {
    const i = mask.search(/HH|H|mm|m|ss|s/);
    if (i < 0) return mask;
    return mask.slice(0, i).replace(/[\s:/.,\-]+$/, "");
  }

  private applyCondFormatting(td: HTMLTableCellElement, colName: string, value: any, rules: CondRule[]): void {
    for (const rule of rules) {
      if (rule.column !== colName) continue;
      const num = parseFloat(value);
      let match = false;
      if (rule.min !== undefined && rule.max !== undefined) match = !isNaN(num) && num >= rule.min && num <= rule.max;
      else if (rule.value !== undefined) match = String(value) === String(rule.value);
      if (match) {
        td.style.backgroundColor = rule.color;
        if (rule.fontColor) td.style.color = rule.fontColor;
        break;
      }
    }
  }

  // ── Filtre d'affichage des lignes (groupe "Row filter" / fx) ─────────────────
  // Filtre purement côté rendu : les lignes exclues disparaissent de la grille,
  // du tri et de la pagination, mais AUCUN filtre n'est envoyé au modèle.
  // Règles combinées en AND ; comparaison en chaîne (trim), la colonne testée
  // peut être masquée dans la grille (hiddenColumns).
  private filteredRowIndices(): number[] {
    const indices = Array.from({ length: this.rowCount }, (_, i) => i);

    const rf = this.fmModel.rowSelection.rowFilter;
    if (!(rf.enabled.value as boolean)) return indices;

    let rules: RowFilterRule[] = [];
    try {
      const parsed = JSON.parse((rf.rulesJson.value as string) || "[]");
      if (Array.isArray(parsed)) rules = parsed;
    } catch { /* JSON invalide ignoré → aucune ligne filtrée */ }
    if (!rules.length) return indices;

    // Pré-résolution des colonnes et des ensembles de valeurs (perf)
    const compiled = rules
      .map(r => {
        const col = r?.column ? this.findCol(this.cleanName(String(r.column))) : undefined;
        if (!col || !Array.isArray(r.values)) return null;
        const set = new Set(r.values.map(v => String(v).trim()));
        const include = String(r.mode || "exclude").toLowerCase() === "include";
        return { col, set, include };
      })
      .filter((r): r is { col: ColDef; set: Set<string>; include: boolean } => r !== null);
    if (!compiled.length) return indices;

    return indices.filter(i => {
      for (const rule of compiled) {
        const raw     = rule.col.values[i];
        const inSet   = raw != null && rule.set.has(String(raw).trim());
        // include → la valeur DOIT être dans la liste ; exclude → ne DOIT PAS l'être
        if (rule.include ? !inSet : inSet) return false;
      }
      return true;
    });
  }

// ── Tri ──────────────────────────────────────────────────────────────────────
  // Comparateur générique : numérique si les deux valeurs sont INTÉGRALEMENT
  // numériques, sinon localeCompare (option numeric). Nulls toujours en fin.
  //
  // Le test strict est indispensable : parseFloat() lit un préfixe et renvoie
  // 2006 pour "2006-10-09T12:50:04.000Z", ce qui annulait le tri des colonnes
  // date reçues en chaîne ISO. Avec le repli localeCompare, l'ISO 8601 se trie
  // correctement en ordre lexicographique.
  private isNumericString(s: string): boolean {
    return /^\s*[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?\s*$/.test(s);
  }

  private compareValues(va: any, vb: any): number {
    if (va == null && vb == null) return 0;
    if (va == null) return 1; if (vb == null) return -1;
    if (va instanceof Date && vb instanceof Date) return va.getTime() - vb.getTime();

    const sa = String(va), sb = String(vb);
    if (typeof va === "number" && typeof vb === "number") return va - vb;
    if (this.isNumericString(sa) && this.isNumericString(sb)) return Number(sa) - Number(sb);

    return sa.localeCompare(sb, undefined, { numeric: true });
  }

  // Lecture de la spécification de tri par défaut (champ Format / fx)
  private getDefaultSortSpecs(): { col: ColDef; asc: boolean }[] {
    const raw = (this.fmModel.columnManagement.columnOrder.defaultSortJson.value as string) || "[]";
    let specs: SortSpec[] = [];
    try {
      const parsed = JSON.parse(raw || "[]");
      if (Array.isArray(parsed)) specs = parsed;
    } catch { /* JSON invalide ignoré */ }

    return specs
      .map(s => {
        const col = s?.column ? this.findCol(this.cleanName(String(s.column))) : undefined;
        if (!col) return null;
        const asc = String(s.dir || "ASC").toUpperCase() !== "DESC";
        return { col, asc };
      })
      .filter((s): s is { col: ColDef; asc: boolean } => s !== null);
  }

  // Tri des indices de lignes (après filtre d'affichage) :
  //   1. Tri manuel (clic en-tête) actif → tri mono-colonne existant.
  //   2. Sinon, tri par défaut multi-colonnes si défini (Default sort / fx).
  //   3. Sinon, ordre naturel des données.
  private sortedRowIndices(visibleCols: ColDef[]): number[] {
    const indices = this.filteredRowIndices();

    // 1. Tri manuel prioritaire
    if (this.sortCol >= 0 && this.sortCol < visibleCols.length) {
      const col = visibleCols[this.sortCol];
      return indices.sort((a, b) => {
        const cmp = this.compareValues(col.values[a], col.values[b]);
        return this.sortAsc ? cmp : -cmp;
      });
    }

    // 2. Tri par défaut multi-colonnes
    const specs = this.getDefaultSortSpecs();
    if (specs.length) {
      return indices.sort((a, b) => {
        for (const s of specs) {
          const cmp = this.compareValues(s.col.values[a], s.col.values[b]);
          if (cmp !== 0) return s.asc ? cmp : -cmp;
        }
        return 0;
      });
    }

    // 3. Ordre naturel
    return indices;
  }

  // ── Validation du layout JSON positionnel de la modale ────────────────────────
  //
  // Deux types d'entrée dans le même tableau :
  //
  //   Champ (implicite, "type" absent) :
  //     { "field": "Col1", "labelField": "Label_Col1", "label": "Repli",
  //       "row": 1, "col": 1, "colspan": 2, "rowspan": 1 }
  //     - field    (string, obligatoire) : résolu via findCol/cleanName, tolère
  //                la notation DAX. Peut référencer N'IMPORTE QUELLE colonne du
  //                rôle "columns", visible ou masquée dans la grille (hiddenColumns).
  //     - labelField (string, optionnel) : nom d'une colonne/mesure dont la
  //                VALEUR à la ligne cliquée sert de libellé — mécanisme
  //                multilingue (priorité la plus haute).
  //     - label    (string, optionnel)   : libellé littéral statique (repli
  //                si labelField absent/vide). Priorité : labelField > label >
  //                headerLabels > displayName.
  //     - row/col  (integer ≥ 1, obligatoires).
  //     - colspan/rowspan (integer ≥ 1, optionnels, défaut 1).
  //
  //   Séparateur ("type": "divider") :
  //     { "type": "divider", "row": 4, "label": "Coordonnées" }
  //     - row      (integer ≥ 1, obligatoire) : occupe TOUTE la largeur de
  //                cette ligne — aucun champ ne peut partager la même row.
  //     - label    (string, optionnel) : titre de section affiché avant le trait.
  //
  // Règle d'exhaustivité : SEULS les champs listés sont affichés — aucun repli
  // automatique sur les champs non mentionnés.
  //
  // Trois états, jamais mélangés :
  //   "empty" → JSON absent/vide ("" ou "[]")            → message neutre.
  //   "error" → JSON présent mais invalide (voir règles)  → panneau rouge.
  //   "ok"    → JSON valide                               → grille positionnée.
  private validateModalLayout(raw: string): ModalLayoutValidation {
    const trimmed = (raw || "").trim();
    if (trimmed === "" || trimmed === "[]") return { kind: "empty" };

    let parsed: any;
    try {
      parsed = JSON.parse(trimmed);
    } catch (e) {
      return { kind: "error", errors: ["JSON invalide : " + (e as Error).message] };
    }

    if (!Array.isArray(parsed)) {
      return { kind: "error", errors: ["Le layout doit être un tableau JSON (reçu : " + typeof parsed + ")."] };
    }

    const errors: string[] = [];
    const items: ResolvedModalLayoutItem[] = [];
    const occupied    = new Map<string, number>();   // "row,col" → n° d'entrée (chevauchement entre champs)
    const reservedRows = new Map<number, number>();  // row → n° d'entrée réservant TOUTE la ligne (divider ou list)
    const fieldRows    = new Map<number, number[]>();  // row → n°s d'entrées champ (conflit avec un séparateur/liste)

    parsed.forEach((raw: any, idx: number) => {
      const n = idx + 1;

      if (typeof raw !== "object" || raw === null || Array.isArray(raw)) {
        errors.push("Entrée #" + n + " : doit être un objet.");
        return;
      }

      const row = raw.row;
      if (!Number.isInteger(row) || row < 1) {
        errors.push("Entrée #" + n + " : \"row\" doit être un entier ≥ 1 (valeur reçue : " + JSON.stringify(row) + ").");
        return;
      }

      // ── Séparateur ────────────────────────────────────────────────────────
      if (raw.type === "divider") {
        const label = (typeof raw.label === "string" && raw.label.trim() !== "") ? raw.label : undefined;

        const prevReserved = reservedRows.get(row);
        if (prevReserved !== undefined) {
          errors.push("Entrée #" + n + " : la ligne " + row + " est déjà occupée par l'entrée #" + prevReserved + " (séparateur ou liste).");
          return;
        }
        reservedRows.set(row, n);
        items.push({ kind: "divider", row, label });
        return;
      }

      // ── Liste (sous-table répétée, relation un-à-plusieurs) ────────────────
      if (raw.type === "list") {
        const prevReserved = reservedRows.get(row);
        if (prevReserved !== undefined) {
          errors.push("Entrée #" + n + " : la ligne " + row + " est déjà occupée par l'entrée #" + prevReserved + " (séparateur ou liste).");
          return;
        }

        const rawColumns = raw.columns;
        if (!Array.isArray(rawColumns) || rawColumns.length === 0) {
          errors.push("Entrée #" + n + " : \"columns\" doit être un tableau non vide.");
          return;
        }

        const delimiter = (typeof raw.delimiter === "string" && raw.delimiter !== "") ? raw.delimiter : ";";

        const resolvedColumns: { fieldCol: ColDef; labelFieldCol?: ColDef; label?: string }[] = [];
        let hasError = false;

        rawColumns.forEach((rc: any, colIdx: number) => {
          const cn = colIdx + 1;
          if (typeof rc !== "object" || rc === null || Array.isArray(rc)) {
            errors.push("Entrée #" + n + ", colonne #" + cn + " : doit être un objet.");
            hasError = true;
            return;
          }
          const fieldName = rc.field;
          if (!fieldName || typeof fieldName !== "string" || fieldName.trim() === "") {
            errors.push("Entrée #" + n + ", colonne #" + cn + " : la propriété \"field\" est manquante ou vide.");
            hasError = true;
            return;
          }
          const fieldCol = this.findCol(this.cleanName(fieldName));
          if (!fieldCol) {
            errors.push("Entrée #" + n + ", colonne #" + cn + " : le champ \"" + fieldName + "\" n'existe dans aucune colonne du visuel.");
            hasError = true;
            return;
          }
          let labelFieldCol: ColDef | undefined;
          if (rc.labelField !== undefined && rc.labelField !== null && String(rc.labelField).trim() !== "") {
            const lfName = String(rc.labelField);
            labelFieldCol = this.findCol(this.cleanName(lfName));
            if (!labelFieldCol) {
              errors.push("Entrée #" + n + ", colonne #" + cn + " (\"" + fieldName + "\") : le champ de libellé \"" + lfName + "\" (labelField) n'existe dans aucune colonne du visuel.");
              hasError = true;
              return;
            }
          }
          const label = (typeof rc.label === "string" && rc.label.trim() !== "") ? rc.label : undefined;
          resolvedColumns.push({ fieldCol, labelFieldCol, label });
        });

        if (hasError) return;

        reservedRows.set(row, n);
        items.push({ kind: "list", row, delimiter, columns: resolvedColumns });
        return;
      }

      // ── Champ normal ──────────────────────────────────────────────────────
      const fieldName = raw.field;
      if (!fieldName || typeof fieldName !== "string" || fieldName.trim() === "") {
        errors.push("Entrée #" + n + " : la propriété \"field\" est manquante ou vide.");
        return;
      }

      const fieldCol = this.findCol(this.cleanName(fieldName));
      if (!fieldCol) {
        errors.push("Entrée #" + n + " : le champ \"" + fieldName + "\" n'existe dans aucune colonne du visuel.");
        return;
      }

      const colPos = raw.col;
      if (!Number.isInteger(colPos) || colPos < 1) {
        errors.push("Entrée #" + n + " (\"" + fieldName + "\") : \"col\" doit être un entier ≥ 1 (valeur reçue : " + JSON.stringify(colPos) + ").");
        return;
      }

      let colspan = 1;
      if (raw.colspan !== undefined) {
        if (!Number.isInteger(raw.colspan) || raw.colspan < 1) {
          errors.push("Entrée #" + n + " (\"" + fieldName + "\") : \"colspan\" doit être un entier ≥ 1.");
          return;
        }
        colspan = raw.colspan;
      }

      let rowspan = 1;
      if (raw.rowspan !== undefined) {
        if (!Number.isInteger(raw.rowspan) || raw.rowspan < 1) {
          errors.push("Entrée #" + n + " (\"" + fieldName + "\") : \"rowspan\" doit être un entier ≥ 1.");
          return;
        }
        rowspan = raw.rowspan;
      }

      const label = (typeof raw.label === "string" && raw.label.trim() !== "") ? raw.label : undefined;

      // labelField : optionnel, référence une colonne/mesure (souvent déposée en
      // Hidden columns) fournissant un libellé résolu PAR LIGNE — c'est le
      // mécanisme multilingue : la mesure DAX applique elle-même _langue_active.
      let labelFieldCol: ColDef | undefined;
      if (raw.labelField !== undefined && raw.labelField !== null && String(raw.labelField).trim() !== "") {
        const lfName = String(raw.labelField);
        labelFieldCol = this.findCol(this.cleanName(lfName));
        if (!labelFieldCol) {
          errors.push("Entrée #" + n + " (\"" + fieldName + "\") : le champ de libellé \"" + lfName + "\" (labelField) n'existe dans aucune colonne du visuel.");
          return;
        }
      }

      // Chevauchement entre champs : chaque cellule couverte par colspan/rowspan.
      for (let r = row; r < row + rowspan; r++) {
        for (let c = colPos; c < colPos + colspan; c++) {
          const key = r + "," + c;
          const prev = occupied.get(key);
          if (prev !== undefined) {
            errors.push("Chevauchement : entrée #" + prev + " et entrée #" + n + " (\"" + fieldName + "\") occupent toutes deux la position ligne " + r + ", colonne " + c + ".");
          } else {
            occupied.set(key, n);
          }
        }
        const list = fieldRows.get(r) || [];
        list.push(n);
        fieldRows.set(r, list);
      }

      items.push({ kind: "field", fieldCol, labelFieldCol, label, row, col: colPos, colspan, rowspan });
    });

    // Conflit champ / séparateur : un séparateur réserve TOUTE sa ligne.
    for (const [reservedRow, reservedN] of reservedRows) {
      const conflicting = fieldRows.get(reservedRow);
      if (conflicting && conflicting.length) {
        for (const fieldN of conflicting) {
          errors.push("Chevauchement : l'entrée #" + fieldN + " (champ) et l'entrée #" + reservedN + " (séparateur ou liste) occupent toutes deux la ligne " + reservedRow + ".");
        }
      }
    }

    if (errors.length) return { kind: "error", errors };

    const columns = items.reduce((max, e) => e.kind === "field" ? Math.max(max, e.col + e.colspan - 1) : max, 1);
    return { kind: "ok", columns, entries: items };
  }

  // ── Taille du dialogue natif (relative à la fenêtre du NAVIGATEUR — contrainte
  //    de l'API openModalDialog, pas du visuel). Bornes imposées par Power BI :
  //    min 240×210 px, max 90 % de la fenêtre. ────────────────────────────────
  // window.innerWidth/innerHeight, lu ICI, renvoie la taille de l'IFRAME DU
  // VISUEL LUI-MÊME sur le canevas (ex. 147px de haut pour un DHT compact),
  // PAS celle du rapport/navigateur — le visuel n'a aucun accès fiable à
  // cette dernière (frontière cross-origin). Tenter de calculer un plafond
  // "% de la fenêtre" avec cette donnée était le bug de la v2.2.x : la
  // modale plafonnait à une hauteur ridicule quel que soit le réglage
  // Format pane. Power BI applique déjà, côté hôte, le vrai plafond (90 %
  // de la vraie fenêtre du navigateur, qu'il connaît) et le vrai plancher
  // (min ~210×240px) — inutile et dangereux de le refaire nous-mêmes avec
  // une donnée fausse.
  private computeDialogSize(modalGroup: { width: any; height: any }): { width: number; height: number } {
    const GENEROUS_FALLBACK = 4000; // "0 = maximum" : l'hôte ramènera de toute façon à son vrai 90%

    const fixedWidth  = modalGroup.width.value as number;
    const fixedHeight = modalGroup.height.value as number;

    return {
      width:  fixedWidth  > 0 ? fixedWidth  : GENEROUS_FALLBACK,
      height: fixedHeight > 0 ? fixedHeight : GENEROUS_FALLBACK
    };
  }

  // ── Ouverture de la modale (dialogue natif Power BI) ──────────────────────────
  // Toute la logique métier (validation, libellés, formatage) est résolue ICI ;
  // le dialogue (detailDialog.ts) ne fait que rendre l'état déjà calculé.
  private openModal(rowIdx: number, labelsMap: Record<string, string>): void {
    // Environnements ne supportant pas les dialogues (Embedded, Publish to web,
    // Dashboards) : on ignore l'appel plutôt que de provoquer une erreur.
    if (this.host.hostCapabilities && this.host.hostCapabilities.allowModalDialog === false) return;

    const modalGroup  = this.fmModel.modalDetail.modal;
    const layoutGroup = this.fmModel.modalDetail.modalLayout;
    const headerBg    = this.getColor(this.fmModel.tableStyle.headerBackground);
    const rowFs       = this.fmModel.tableStyle.rowFontSize.value as number;
    const headerFs    = this.fmModel.tableStyle.headerFontSize.value as number;

    // Titre : champ Format "Titre de la modale" (texte OU mesure fx) en priorité ;
    // si vide, repli sur la valeur de la première colonne de la ligne.
    let title = "Détail";
    const fmtTitle = (modalGroup.titleText.value as string) || "";
    if (fmtTitle.trim() !== "") {
      title = fmtTitle;
    } else {
      const firstVisible = this.allCols.find(c => !c.isHidden);
      if (firstVisible) { const v = firstVisible.values[rowIdx]; if (v != null) title = String(v); }
    }

    const validation = this.validateModalLayout((layoutGroup.layoutJson.value as string) || "[]");

    let state: ModalDialogState;
    if (validation.kind === "empty") {
      state = { kind: "empty" };
    } else if (validation.kind === "error") {
      state = { kind: "error", errors: validation.errors };
    } else {
      const entries: ModalDialogGridItem[] = validation.entries.map(e => {
        if (e.kind === "divider") {
          return { kind: "divider", row: e.row, label: e.label };
        }
        if (e.kind === "list") {
          // Résout chaque colonne : libellé (même priorité que les champs
          // normaux) + découpage de la valeur brute selon le délimiteur.
          const colData = e.columns.map(c => {
            const label = this.resolveLabelFieldValue(c.labelFieldCol, rowIdx) ?? c.label ?? this.resolveLabel(c.fieldCol, labelsMap);
            const raw    = c.fieldCol.values[rowIdx];
            const rawStr = (raw == null ? "" : String(raw)).trim();
            const items  = rawStr === "" ? [] : rawStr.split(e.delimiter).map(s => s.trim());
            return { label, items };
          });
          const rowCount = colData.reduce((max, c) => Math.max(max, c.items.length), 0);
          const rows: string[][] = [];
          for (let i = 0; i < rowCount; i++) {
            rows.push(colData.map(c => (c.items[i] !== undefined && c.items[i] !== "") ? c.items[i] : "\u2014"));
          }
          return { kind: "list", row: e.row, headers: colData.map(c => c.label), rows };
        }
        const label = this.resolveLabelFieldValue(e.labelFieldCol, rowIdx) ?? e.label ?? this.resolveLabel(e.fieldCol, labelsMap);
        const raw   = e.fieldCol.values[rowIdx];
        const fmt   = this.resolveModalFormat(e.fieldCol.displayName, e.fieldCol.queryName);
        const value = raw == null ? "\u2014" : this.formatCell(raw, e.fieldCol.format, fmt);
        return { kind: "field", label, value, row: e.row, col: e.col, colspan: e.colspan, rowspan: e.rowspan };
      });
      state = { kind: "grid", columns: validation.columns, entries };
    }

    const size = this.computeDialogSize(modalGroup);

    const initialState: ModalDialogInitialState = {
      accentColor:    headerBg,
      bodyFontSize:   rowFs,
      headerFontSize: headerFs,
      state,
      debug: (modalGroup.debugMode.value as boolean) ? {
        requestedWidth:    size.width,
        requestedHeight:   size.height,
        formatWidth:       modalGroup.width.value as number,
        formatHeight:      modalGroup.height.value as number,
        visualIframeWidth: window.innerWidth,
        visualIframeHeight:window.innerHeight
      } as ModalDialogDebugInfo : undefined
    };

    const dialogOptions = {
      title,
      size,
      actionButtons: [DialogAction.Close]
    };

    this.host.openModalDialog(DetailDialog.id, dialogOptions, initialState)
      .catch(() => { /* dialogue fermé/refusé — lecture seule, rien à faire */ });
  }

  // ── Render ────────────────────────────────────────────────────────────────────
  private render(): void {
    const existing = this.container.querySelector(".dht-wrapper") as HTMLElement;
    if (existing) this.container.removeChild(existing);

    // Mapping libellés effectif (avec cache pour le correctif 1.a)
    const labelsMap = this.resolveLabelsMap();

    const rawVisible = this.allCols.filter(c => !c.isHidden && !c.isTitle);
    const colOrderJson = this.fmModel.columnManagement.columnOrder.orderJson.value as string;
    const visibleCols  = this.applyColumnOrder(rawVisible, colOrderJson);

    if (!visibleCols.length) {
      const msg = document.createElement("div");
      msg.className = "dht-wrapper";
      msg.style.cssText = "padding:16px;color:#666;font-size:13px;";
      msg.textContent = "Ajoutez des données dans la zone Colonnes.";
      this.container.appendChild(msg);
      return;
    }

    // Raccourcis settings
    const ts          = this.fmModel.tableStyle;
    const headerBg    = this.getColor(ts.headerBackground);
    const headerFc    = this.getColor(ts.headerFontColor);
    const headerFs    = ts.headerFontSize.value as number;
    const rowBg       = this.getColor(ts.rowBackground);
    const rowAltBg    = this.getColor(ts.rowAltBackground);
    const rowFc       = this.getColor(ts.rowFontColor);
    const rowFs       = ts.rowFontSize.value as number;
    const borderColor = this.getColor(ts.borderColor);
    const gridLines   = ts.gridLines.value as boolean;
    const pagEnabled  = this.fmModel.pagination.enabled.value as boolean;
    const pageSize    = pagEnabled ? Math.max(1, this.fmModel.pagination.rowsPerPage.value as number) : 99999;
    const sortEnabled = this.fmModel.sorting.enabled.value as boolean;
    const cfGroup     = this.fmModel.columnManagement.conditionalFormatting;
    const condEnabled = cfGroup.enabled.value as boolean;
    const modalEnabled= this.fmModel.modalDetail.modal.enabled.value as boolean;

    // ── Sélection de ligne (paramètres d'apparence) ───────────────────────────
    const rs            = this.fmModel.rowSelection.appearance;
    const selEnabled    = rs.enabled.value as boolean;
    const selBorderC    = this.getColor(rs.borderColor);
    const selBorderW    = Math.max(0, rs.borderWidth.value as number);
    const selBgC        = this.getColor(rs.backgroundColor);
    const selBgOpacity  = Math.min(100, Math.max(0, rs.backgroundOpacity.value as number));
    // Conversion couleur hex + opacité 0-100 vers rgba()
    const selBgRgba     = this.hexToRgba(selBgC, selBgOpacity / 100);

    let condRules: CondRule[] = [];
    if (condEnabled) { try { condRules = JSON.parse((cfGroup.rulesJson.value as string) || "[]"); } catch { /* noop */ } }

    // Spécifications du tri par défaut (résolu une fois par rendu, aussi
    // utilisé pour les indicateurs d'en-tête)
    const defaultSortSpecs = this.getDefaultSortSpecs();

    const sortedIdx  = this.sortedRowIndices(visibleCols);
    const total      = sortedIdx.length;
    const maxPage    = Math.max(0, Math.ceil(total / pageSize) - 1);
    this.currentPage = Math.min(this.currentPage, maxPage);
    const pageIdx    = sortedIdx.slice(this.currentPage * pageSize, (this.currentPage + 1) * pageSize);

    const wrapper = document.createElement("div");
    wrapper.className = "dht-wrapper";
    wrapper.style.cssText = "display:flex;flex-direction:column;height:100%;width:100%;";

    const tableWrapper = document.createElement("div");
    tableWrapper.style.cssText = "flex:1;overflow:auto;";

    const table = document.createElement("table");
    table.style.cssText = "border-collapse:collapse;width:100%;font-size:" + rowFs + "px;color:" + rowFc + ";";

    // ── En-tête ─────────────────────────────────────────────────────────────────
    const thead = document.createElement("thead");
    const hrow  = document.createElement("tr");

    if (modalEnabled) {
      const thBtn = document.createElement("th");
      thBtn.style.cssText = "background:" + headerBg + ";padding:8px 6px;width:32px;position:sticky;top:0;z-index:10;" + (gridLines ? "border:1px solid " + borderColor + ";" : "");
      hrow.appendChild(thBtn);
    }

    visibleCols.forEach((col, idx) => {
      const th = document.createElement("th");
      th.textContent = this.resolveLabel(col, labelsMap);   // correctif 1.a appliqué ici
      const fmt      = this.resolveFormat(col.displayName, col.queryName);
      const colWidth = fmt?.width    ? fmt.width    + "px" : "";
      const colMinW  = fmt?.minWidth ? fmt.minWidth + "px" : "";
      th.style.cssText = "background:" + headerBg + ";color:" + headerFc + ";font-size:" + headerFs + "px;" +
        "padding:8px 10px;text-align:left;position:sticky;top:0;z-index:10;white-space:nowrap;user-select:none;" +
        (colWidth ? "width:" + colWidth + ";min-width:" + colWidth + ";" : (colMinW ? "min-width:" + colMinW + ";" : "")) +
        (gridLines ? "border:1px solid " + borderColor + ";" : "");
      if (sortEnabled) {
        th.style.cursor = "pointer";
        const ind = document.createElement("span");
        ind.style.cssText = "margin-left:6px;opacity:0.6;font-size:10px;";
        // Indicateur : tri manuel actif → flèche pleine sur la colonne triée ;
        // sinon, flèche estompée sur les colonnes du tri par défaut.
        const defSpec = this.sortCol === -1
          ? defaultSortSpecs.find(s => s.col === col)
          : undefined;
        if (this.sortCol === idx) {
          ind.textContent = this.sortAsc ? "▲" : "▼";
          ind.style.opacity = "1";
        } else if (defSpec) {
          ind.textContent = defSpec.asc ? "▲" : "▼";
          ind.style.opacity = "0.45";
        } else {
          ind.textContent = "⇅";
        }
        th.appendChild(ind);
        th.addEventListener("click", () => {
          if (this.sortCol === idx) this.sortAsc = !this.sortAsc;
          else { this.sortCol = idx; this.sortAsc = true; }
          this.currentPage = 0;
          this.render();
        });
      }
      hrow.appendChild(th);
    });

    thead.appendChild(hrow);
    table.appendChild(thead);

    // ── Corps ───────────────────────────────────────────────────────────────────
    const tbody = document.createElement("tbody");

    pageIdx.forEach((origIdx, rowIdx) => {
      const tr = document.createElement("tr");
      const isSelected = selEnabled && origIdx === this.selectedRowIdx;
      tr.style.background = rowIdx % 2 === 0 ? rowBg : rowAltBg;
      tr.addEventListener("mouseenter", () => { tr.style.filter = "brightness(0.93)"; });
      tr.addEventListener("mouseleave", () => { tr.style.filter = ""; });

      // ── Sélection de ligne : surlignage si sélectionnée ─────────────────────
      if (isSelected) {
        // Fond teinté (surcharge l'alternance) + bordure gauche colorée.
        // box-shadow inset évite de décaler le contenu (contrairement à border).
        tr.style.background  = selBgRgba;
        tr.style.boxShadow   = "inset " + selBorderW + "px 0 0 0 " + selBorderC;
      }

      // ── Clic sur la ligne : bascule la sélection (cross-filtering) ──────────
      if (selEnabled) {
        tr.style.cursor = "pointer";
        tr.addEventListener("click", (e) => {
          // Ignorer les clics qui viennent du bouton modale (e.stopPropagation
          // est déjà fait dans son handler, mais ceinture-bretelles).
          const target = e.target as HTMLElement;
          if (target && target.tagName === "BUTTON") return;

          this.handleRowClick(origIdx);
        });
      }

      if (modalEnabled) {
        // Dialogues natifs indisponibles dans certains contextes (Embedded,
        // Publish to web, Dashboards, ou "ne plus afficher" côté utilisateur) :
        // le bouton reste visible mais désactivé, avec une info-bulle explicite.
        const dialogAllowed = !this.host.hostCapabilities || this.host.hostCapabilities.allowModalDialog !== false;

        const tdBtn = document.createElement("td");
        tdBtn.style.cssText = "padding:4px 6px;text-align:center;" +
          (gridLines ? "border:1px solid " + borderColor + ";" : "border-bottom:1px solid " + borderColor + ";");
        const btn = document.createElement("button");
        btn.textContent = "\u22EF";
        btn.title = dialogAllowed ? "Voir le détail" : "Fenêtres modales non disponibles dans cet environnement";
        btn.disabled = !dialogAllowed;
        btn.style.cssText = "background:" + (dialogAllowed ? headerBg : "#999") + ";color:#fff;border:none;border-radius:4px;" +
          "padding:2px 7px;cursor:" + (dialogAllowed ? "pointer" : "not-allowed") + ";font-size:13px;font-weight:bold;line-height:1.4;";
        btn.addEventListener("mouseenter", () => { if (dialogAllowed) btn.style.opacity = "0.8"; });
        btn.addEventListener("mouseleave", () => { btn.style.opacity = "1"; });
        btn.addEventListener("click", (e) => { e.stopPropagation(); if (dialogAllowed) this.openModal(origIdx, labelsMap); });
        tdBtn.appendChild(btn);
        tr.appendChild(tdBtn);
      }

      visibleCols.forEach((col) => {
        const cell = col.values[origIdx];
        const td   = document.createElement("td");
        const fmt  = this.resolveFormat(col.displayName, col.queryName);
        const disp = this.formatCell(cell, col.format, fmt);
        td.textContent = disp;
        td.title       = disp;

        const isNumeric = fmt ? fmt.type === "number" : typeof cell === "number";
        const isBoolean = fmt ? fmt.type === "boolean" : false;
        const colWidth  = fmt?.width    ? fmt.width    + "px" : "";
        const colMinW   = fmt?.minWidth ? fmt.minWidth + "px" : "";
        const isMulti   = fmt?.multiline !== undefined
                        ? !!fmt.multiline
                        : /[\r\n]/.test(disp);

        td.style.cssText = "padding:6px 10px;vertical-align:top;" +
          (gridLines ? "border:1px solid " + borderColor + ";" : "border-bottom:1px solid " + borderColor + ";") +
          (colWidth ? "width:" + colWidth + ";min-width:" + colWidth + ";"
                    : (colMinW ? "min-width:" + colMinW + ";max-width:250px;" : "max-width:250px;")) +
          //"white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" +
          (isMulti
            ? "white-space:pre-line;word-break:break-word;overflow:hidden;line-height:1.35;"
            : "white-space:nowrap;overflow:hidden;text-overflow:ellipsis;") +

          (isNumeric ? "text-align:right;" : isBoolean ? "text-align:center;" : "");

        // Couleurs spécifiques booléen (trueColor / falseColor)
        if (isBoolean && fmt) {
          const truthy = cell === true || cell === 1 || String(cell).toLowerCase() === "true";
          const bg = truthy ? fmt.trueColor : fmt.falseColor;
          const fc = truthy ? fmt.trueFontColor : fmt.falseFontColor;
          if (bg) td.style.backgroundColor = bg;
          if (fc) td.style.color = fc;
        }

        if (condEnabled && condRules.length) this.applyCondFormatting(td, col.displayName, cell, condRules);
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tableWrapper.appendChild(table);
    wrapper.appendChild(tableWrapper);

    // ── Pagination ─────────────────────────────────────────────────────────────
    if (pagEnabled && total > pageSize) {
      const bar = document.createElement("div");
      bar.style.cssText = "display:flex;align-items:center;justify-content:space-between;padding:6px 10px;" +
        "background:" + headerBg + "22;border-top:1px solid " + borderColor + ";font-size:12px;flex-shrink:0;";
      const info = document.createElement("span");
      info.textContent = (this.currentPage * pageSize + 1) + "–" +
        Math.min((this.currentPage + 1) * pageSize, total) + " sur " + total + " lignes";
      info.style.color = "#555";
      const nav = document.createElement("div");
      nav.style.cssText = "display:flex;gap:6px;align-items:center;";
      const bs = "padding:3px 10px;border:1px solid " + borderColor + ";border-radius:3px;background:white;cursor:pointer;font-size:12px;";
      const mk = (t: string, fn: () => void, dis: boolean) => {
        const b = document.createElement("button"); b.textContent = t; b.style.cssText = bs; b.disabled = dis;
        b.addEventListener("click", () => { fn(); this.render(); }); return b;
      };
      const pg = document.createElement("span");
      pg.textContent = "Page " + (this.currentPage + 1) + " / " + (maxPage + 1);
      pg.style.cssText = "font-size:12px;color:#333;min-width:80px;text-align:center;";
      nav.append(
        mk("«", () => { this.currentPage = 0; },       this.currentPage === 0),
        mk("‹", () => { this.currentPage--; },         this.currentPage === 0),
        pg,
        mk("›", () => { this.currentPage++; },         this.currentPage >= maxPage),
        mk("»", () => { this.currentPage = maxPage; }, this.currentPage >= maxPage)
      );
      bar.append(info, nav);
      wrapper.appendChild(bar);
    }

    this.container.appendChild(wrapper);
  }

  // ── Restauration de la sélection depuis les filtres actifs ───────────────────
  // Appelé à chaque rechargement de données (VisualUpdateType.Data).
  // Lit options.jsonFilters pour retrouver la valeur active sur targetTable/targetColumn,
  // puis cherche la ligne correspondante dans allCols via filterKey.
  // Retourne l'index de la ligne restaurée, ou -1 si aucun filtre actif.
  private restoreSelectionFromFilters(options: VisualUpdateOptions): number {
    try {
      const filt        = this.fmModel.rowSelection.filtering;
      const filterKey   = this.cleanName(filt.filterKey.value    as string);
      const targetTable = this.cleanName(filt.targetTable.value  as string);
      const targetCol   = this.cleanName(filt.targetColumn.value as string);

      if (!filterKey || !targetTable || !targetCol) return -1;

      const filters = options.jsonFilters as any[];
      if (!filters || !filters.length) return -1;

      // Chercher le filtre Basic actif sur targetTable.targetColumn
      let activeValue: any = null;
      for (const f of filters) {
        if (
          f?.target?.table  === targetTable &&
          f?.target?.column === targetCol   &&
          Array.isArray(f.values) && f.values.length > 0
        ) {
          activeValue = f.values[0];
          break;
        }
      }
      if (activeValue == null) return -1;

      // Chercher la ligne dont filterKey correspond à activeValue
      const keyCol = this.findCol(filterKey);
      if (!keyCol) return -1;

      const strActive = String(activeValue);
      for (let i = 0; i < this.rowCount; i++) {
        const raw = keyCol.values[i];
        if (raw != null && String(raw) === strActive) return i;
      }
    } catch { /* noop */ }
    return -1;
  }

  // ── Update ────────────────────────────────────────────────────────────────────
  public update(options: VisualUpdateOptions): void {
    this.events.renderingStarted(options);

    const dv = options.dataViews?.[0];
    if (!dv) {
      const existing = this.container.querySelector(".dht-wrapper") as HTMLElement;
      if (existing) this.container.removeChild(existing);
      const d = document.createElement("div");
      d.className = "dht-wrapper";
      d.style.cssText = "padding:16px;color:#666;";
      d.textContent = "Aucune donnée.";
      this.container.appendChild(d);
      this.events.renderingFinished(options);
      return;
    }

    this.fmModel = this.fmService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dv);
    this.loadColumnFormats();
    this.extractData(dv);

    // Restauration de la sélection visuelle depuis les filtres actifs (navigation retour).
    // On ne touche à selectedRowIdx que si les données viennent d'être rechargées
    // (VisualUpdateType.Data) — pas lors d'un simple resize ou changement de format.
    if (options.type & powerbi.VisualUpdateType.Data) {
      this.currentPage   = 0;
      this.selectedRowIdx = this.restoreSelectionFromFilters(options);
    }

    this.render();
    this.events.renderingFinished(options);
  }

  // ── FormattingModel API ───────────────────────────────────────────────────────
  public getFormattingModel(): powerbi.visuals.FormattingModel {
    return this.fmService.buildFormattingModel(this.fmModel);
  }
}
