"use strict";

import powerbi from "powerbi-visuals-api";
import DialogConstructorOptions = powerbi.extensibility.visual.DialogConstructorOptions;

// ═══════════════════════════════════════════════════════════════════════════
//  Contrat visual.ts → dialogue
//
//  Le visuel (visual.ts) résout TOUT avant l'ouverture : validation du JSON
//  de layout, résolution des libellés (headerLabels / label du JSON), et
//  formatage des valeurs (colFormatsMap / modalFormatsMap). Ce fichier ne
//  contient donc aucune logique métier — uniquement du rendu d'un état déjà
//  calculé, reçu en tant qu'objet simple (JSON-serialisable) via initialState.
//
//  Raison : le dialogue s'exécute dans un document/iframe séparé du visuel ;
//  on ne peut pas lui faire traverser des instances de classes (ColDef, etc.),
//  seulement des données brutes.
// ═══════════════════════════════════════════════════════════════════════════

export interface ModalDialogFieldEntry {
  kind:    "field";
  label:   string;
  value:   string;
  row:     number;
  col:     number;
  colspan: number;
  rowspan: number;
}

export interface ModalDialogDividerEntry {
  kind:   "divider";
  row:    number;
  label?: string;
}

export interface ModalDialogListEntry {
  kind:    "list";
  row:     number;
  headers: string[];
  rows:    string[][];   // rows[i][j] = valeur (ou "—") de la colonne j, élément i
}

export type ModalDialogGridItem = ModalDialogFieldEntry | ModalDialogDividerEntry | ModalDialogListEntry;

export type ModalDialogState =
  | { kind: "empty" }
  | { kind: "error"; errors: string[] }
  | { kind: "grid"; columns: number; entries: ModalDialogGridItem[] };

export interface ModalDialogDebugInfo {
  requestedWidth:     number;
  requestedHeight:    number;
  formatWidth:        number;
  formatHeight:       number;
  visualIframeWidth:  number;   // taille de l'iframe DU VISUEL sur le canevas — PAS celle du rapport
  visualIframeHeight: number;
}

export interface ModalDialogInitialState {
  accentColor:    string;
  bodyFontSize:   number;
  headerFontSize: number;
  state:          ModalDialogState;
  debug?:         ModalDialogDebugInfo;
}

export class DetailDialog {
  static id = "DetailDialog";

  constructor(options: DialogConstructorOptions, initialState: object) {
    const state = initialState as ModalDialogInitialState;
    const root  = options.element;

    // Base de référence explicite : sans ceci, "height:100%" sur .ddlg-body
    // n'a rien de concret à quoi se rapporter si root n'est pas littéralement
    // <body> — le contenu se dimensionne alors à sa hauteur naturelle,
    // indépendamment de la taille demandée à openModalDialog (bug corrigé
    // en v2.2.3.0 : Height (px) n'avait aucun effet observable).
    // Layout en flex-column : le panneau debug (le cas échéant) garde sa
    // hauteur naturelle, le contenu prend le reste et scrolle localement.
    root.style.cssText =
      "width:100%;height:100%;overflow:hidden;box-sizing:border-box;" +
      "display:flex;flex-direction:column;";

    const styleEl = document.createElement("style");
    styleEl.textContent = this.buildCss();
    root.appendChild(styleEl);

    if (state.debug) {
      root.appendChild(this.buildDebugPanel(root, state.debug));
    }

    const body = document.createElement("div");
    body.className = "ddlg-body";
    body.style.flex      = "1";
    body.style.minHeight = "0";   // indispensable pour que overflow:auto fonctionne dans un enfant flex
    root.appendChild(body);

    switch (state.state.kind) {
      case "empty":
        this.renderEmpty(body, state.bodyFontSize);
        break;
      case "error":
        this.renderError(body, state.state.errors, state.bodyFontSize);
        break;
      case "grid":
        this.renderGrid(body, state.state, state.accentColor, state.bodyFontSize, state.headerFontSize);
        break;
    }
  }

  // ── Styles de base (le dialogue a son propre document, aucun CSS du
  //    visuel n'est hérité — tout est injecté ici) ──────────────────────────
  private buildCss(): string {
    return (
      "* { box-sizing: border-box; } " +
      "html, body { margin:0; padding:0; font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; " +
      "  height:100%; } " +
      ".ddlg-body { padding:18px; overflow:auto; height:100%; }"
    );
  }

  // ── Panneau de diagnostic (Debug mode) ────────────────────────────────────────
  // Affiche : (1) ce qui a été demandé à openModalDialog, (2) les valeurs
  // brutes du Format pane, (3) ce qui est RÉELLEMENT mesuré à l'intérieur du
  // dialogue une fois construit. Si (1) et (3) diffèrent, l'hôte Power BI
  // n'honore pas la taille demandée. Si (1) et (3) concordent mais que le
  // contenu reste tronqué, le problème est ailleurs (CSS de contenu).
  private buildDebugPanel(root: HTMLElement, debug: ModalDialogDebugInfo): HTMLElement {
    const panel = document.createElement("div");
    panel.style.cssText =
      "flex-shrink:0;background:#fff8e1;border:2px solid #f0ad4e;border-radius:6px;" +
      "padding:10px 12px;margin-bottom:10px;font-family:Consolas,'Courier New',monospace;" +
      "font-size:11px;color:#5c4300;white-space:pre-wrap;line-height:1.6;";

    const title = document.createElement("div");
    title.textContent = "⚙ DEBUG — dimensions du dialogue";
    title.style.cssText = "font-weight:700;margin-bottom:6px;font-family:'Segoe UI',sans-serif;font-size:12px;";
    panel.appendChild(title);

    const pre = document.createElement("div");
    // Lecture différée (prochaine frame) pour être sûr que le layout est posé
    // avant de mesurer — sinon clientWidth/clientHeight peuvent encore valoir 0.
    const fill = () => {
      const rootRect = root.getBoundingClientRect();
      const lines = [
        "Demandé à openModalDialog : width=" + debug.requestedWidth + "px  height=" + debug.requestedHeight + "px",
        "Format pane (bruts)       : width=" + debug.formatWidth + "  height=" + debug.formatHeight,
        "Iframe DU VISUEL (⚠ pas le rapport) : " + debug.visualIframeWidth + " x " + debug.visualIframeHeight,
        "Mesuré DANS le dialogue   : root=" + Math.round(rootRect.width) + "x" + Math.round(rootRect.height) +
          "  document=" + document.documentElement.clientWidth + "x" + document.documentElement.clientHeight +
          "  window=" + window.innerWidth + "x" + window.innerHeight
      ];
      pre.textContent = lines.join("\n");
    };
    fill();
    requestAnimationFrame(fill);   // seconde passe, au cas où le premier rendu était prématuré
    panel.appendChild(pre);

    return panel;
  }

  private renderEmpty(body: HTMLElement, fontSize: number): void {
    const msg = document.createElement("div");
    msg.style.cssText =
      "background:#eef4fa;border:1px solid #b8d4ea;border-radius:6px;" +
      "padding:14px 16px;color:#2c5a7a;font-size:" + fontSize + "px;line-height:1.5;";
    msg.textContent = "Configurez le champ « Layout (JSON / fx) » (carte Modal detail) pour afficher le détail de cette fiche.";
    body.appendChild(msg);
  }

  private renderError(body: HTMLElement, errors: string[], fontSize: number): void {
    const box = document.createElement("div");
    box.style.cssText =
      "background:#fdecea;border:1px solid #f5b5ad;border-radius:6px;padding:14px 16px;color:#611a15;";

    const title = document.createElement("div");
    title.textContent = "⚠ Le layout de la modale contient des erreurs :";
    title.style.cssText = "font-weight:700;margin-bottom:8px;font-size:" + fontSize + "px;";
    box.appendChild(title);

    const ul = document.createElement("ul");
    ul.style.cssText = "margin:0;padding-left:20px;font-size:" + (fontSize - 1) + "px;" +
      "font-family:Consolas,'Courier New',monospace;line-height:1.6;";
    for (const err of errors) {
      const li = document.createElement("li");
      li.textContent = err;
      ul.appendChild(li);
    }
    box.appendChild(ul);
    body.appendChild(box);
  }

  private renderGrid(
    body: HTMLElement,
    gridState: { columns: number; entries: ModalDialogGridItem[] },
    accentColor: string,
    fontSize: number,
    headerFontSize: number
  ): void {
    body.style.display             = "grid";
    body.style.gridTemplateColumns = "repeat(" + gridState.columns + ",1fr)";
    body.style.gap                 = "10px 16px";
    body.style.alignContent        = "start";

    for (const entry of gridState.entries) {
      if (entry.kind === "divider") {
        const row = document.createElement("div");
        row.style.cssText =
          "grid-column:1 / -1;grid-row:" + entry.row + ";" +
          "display:flex;align-items:center;gap:10px;min-width:0;padding:4px 0;";

        if (entry.label) {
          const labelEl = document.createElement("span");
          labelEl.textContent = entry.label;
          labelEl.style.cssText =
            "font-size:" + Math.max(9, fontSize - 1) + "px;font-weight:600;color:" + accentColor + ";" +
            "text-transform:uppercase;letter-spacing:0.04em;white-space:nowrap;flex-shrink:0;" +
            "white-space:pre-line;";
          row.appendChild(labelEl);
        }

        const line = document.createElement("div");
        line.style.cssText = "flex:1;height:1px;background:" + accentColor + ";min-width:0;";
        row.appendChild(line);

        body.appendChild(row);
        continue;
      }

      if (entry.kind === "list") {
        const cols = entry.headers.length;
        const wrap = document.createElement("div");
        wrap.style.cssText =
          "grid-column:1 / -1;grid-row:" + entry.row + ";" +
          "display:flex;flex-direction:column;gap:4px;min-width:0;";

        const headerRow = document.createElement("div");
        headerRow.style.cssText = "display:grid;grid-template-columns:repeat(" + cols + ",1fr);gap:10px 16px;";
        for (const h of entry.headers) {
          const hEl = document.createElement("span");
          hEl.textContent = h;
          hEl.style.cssText =
            "font-size:" + Math.max(9, fontSize - 1) + "px;font-weight:600;color:" + accentColor + ";" +
            "text-transform:uppercase;letter-spacing:0.04em;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;";
          headerRow.appendChild(hEl);
        }
        wrap.appendChild(headerRow);

        for (const rowValues of entry.rows) {
          const dataRow = document.createElement("div");
          dataRow.style.cssText = "display:grid;grid-template-columns:repeat(" + cols + ",1fr);gap:10px 16px;";
          for (const v of rowValues) {
            const vEl = document.createElement("span");
            vEl.textContent = v;
            vEl.style.cssText =
              "font-size:" + fontSize + "px;color:#1a1a2e;padding:5px 8px;background:#f4f6f9;" +
              "border-radius:4px;border-left:3px solid " + accentColor + ";word-break:break-word;min-width:0;" +
              "white-space:pre-line;";
            dataRow.appendChild(vEl);
          }
          wrap.appendChild(dataRow);
        }

        body.appendChild(wrap);
        continue;
      }

      const pair = document.createElement("div");
      pair.style.cssText =
        "display:flex;flex-direction:column;gap:2px;min-width:0;" +
        "grid-column:" + entry.col + " / span " + entry.colspan + ";" +
        "grid-row:"    + entry.row + " / span " + entry.rowspan + ";";

      const labelEl = document.createElement("span");
      labelEl.textContent = entry.label;
      labelEl.style.cssText =
        "font-size:" + Math.max(9, fontSize - 1) + "px;font-weight:600;color:" + accentColor + ";" +
        "text-transform:uppercase;letter-spacing:0.04em;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;";

      const valueEl = document.createElement("span");
      valueEl.textContent = entry.value;
      valueEl.style.cssText =
        "font-size:" + fontSize + "px;color:#1a1a2e;padding:5px 8px;background:#f4f6f9;" +
        "border-radius:4px;border-left:3px solid " + accentColor + ";word-break:break-word;min-width:0;";

      pair.appendChild(labelEl);
      pair.appendChild(valueEl);
      body.appendChild(pair);
    }
  }
}

// Enregistrement obligatoire — pattern imposé par l'API Power BI (create-display-dialog-box).
globalThis.dialogRegistry = globalThis.dialogRegistry || {};
(globalThis.dialogRegistry as any)[DetailDialog.id] = DetailDialog;
