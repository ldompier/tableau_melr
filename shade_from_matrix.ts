import { promises as fs } from "fs";
import * as path from "path";
import JSZip = require("jszip");
import { DOMParser, XMLSerializer } from "xmldom";

/** WordprocessingML namespace */
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/* ============ Helpers CSV ============ */
function stripBom(s: string) { return s.replace(/^\uFEFF/, ""); }
function splitSemicolon(line: string): string[] { return (line ?? "").split(";"); }
async function readCsvSemicolon(file: string): Promise<string[][]> {
  const raw = stripBom(await fs.readFile(file, "utf8"));
  const lines = raw.split(/\r?\n/).map(l => l.trimEnd());
  // on garde les lignes vides de fin ? ici on filtre uniquement le trailing vide
  while (lines.length && lines[lines.length - 1] === "") lines.pop();
  const rows: string[][] = [];
  for (const l of lines) rows.push(splitSemicolon(l));
  return rows;
}

/* ============ XML helpers sans préfixe ============ */
function findAllByLocalName(node: Node, local: string, out: Element[] = []): Element[] {
  if ((node as Element).localName === local) out.push(node as Element);
  for (let ch = node.firstChild; ch; ch = ch.nextSibling) findAllByLocalName(ch, local, out);
  return out;
}
function childrenByLocalName(el: Element, local: string): Element[] {
  const res: Element[] = [];
  for (let ch = el.firstChild; ch; ch = ch.nextSibling) {
    if ((ch as Element).localName === local) res.push(ch as Element);
  }
  return res;
}
function firstChildByLocalName(el: Element, local: string): Element | null {
  for (let ch = el.firstChild; ch; ch = ch.nextSibling) {
    if ((ch as Element).localName === local) return ch as Element;
  }
  return null;
}

/* ============ Shading util ============ */
function ensureTcPr(tc: Element, doc: Document): Element {
  let tcPr = firstChildByLocalName(tc, "tcPr");
  if (!tcPr) {
    tcPr = doc.createElementNS(W_NS, "w:tcPr");
    tc.insertBefore(tcPr, tc.firstChild);
  }
  return tcPr;
}
function clearShading(tcPr: Element) {
  const toRemove = childrenByLocalName(tcPr, "shd");
  for (const n of toRemove) tcPr.removeChild(n);
}
function shadeCell(tc: Element, doc: Document, fillHex: string) {
  const tcPr = ensureTcPr(tc, doc);
  clearShading(tcPr);
  const shd = doc.createElementNS(W_NS, "w:shd");
  shd.setAttributeNS(W_NS, "w:val", "clear");
  shd.setAttributeNS(W_NS, "w:color", "auto");
  shd.setAttributeNS(W_NS, "w:fill", fillHex.toUpperCase());
  tcPr.appendChild(shd);
}

/* ============ Main ============ */
/**
 * CSV attendu (exemple) :
 * eleve;col1;col2;col3;...
 * 1;1;;2;...
 * 2;;2;;...
 * -> on colorie la ligne "élève N" (avec offset) et chaque cellule de mot selon valeur:
 *   "1" -> vert (#00FF00), "2" -> rouge (#FF0000), "" -> jaune (#FFFF00)
 */
async function main() {
  const args = process.argv.slice(2);
  const getArg = (name: string) => { const i = args.indexOf(name); return i >= 0 ? args[i + 1] : undefined; };

  const srcPath  = getArg("--src");                       // ex: compilationdictee_original.docx
  const csvPath  = getArg("--csv");                       // ex: compilation_résultats.csv
  const outPath  = getArg("--out") ?? "Tableau_couleurs.docx";

  // indices/offsets
  const tableIndex        = Number(getArg("--table-index") ?? "0");
  const headerRowIndex    = Number(getArg("--header-row-index") ?? "1");  // ligne des mots (souvent 1)
  const studentRowOffset  = Number(getArg("--student-row-offset") ?? "1"); // élève 1 -> tr index 2 => +1
  const firstDataColIndex = Number(getArg("--first-data-col-index") ?? "1"); // 1 si col 0 = numéro d'élève

  // couleurs (modifiable via CLI si tu veux)
  const colorFor1 = (getArg("--color-1") ?? "00FF00").toUpperCase(); // vert
  const colorFor2 = (getArg("--color-2") ?? "FF0000").toUpperCase(); // rouge
  const colorForEmpty = (getArg("--color-empty") ?? "FFFF00").toUpperCase(); // jaune

  if (!srcPath || !csvPath) {
    console.error("Usage:");
    console.error("  npx ts-node shade_from_matrix.ts --src compilationdictee_original.docx --csv compilation_resultats.csv --out Tableau_couleurs.docx");
    console.error("Options: --table-index 0 --header-row-index 1 --student-row-offset 1 --first-data-col-index 1");
    console.error("Couleurs: --color-1 00FF00 --color-2 FF0000 --color-empty FFFF00");
    process.exit(1);
  }

  // 1) Lire CSV matrice
  const rowsCsv = await readCsvSemicolon(csvPath);
  if (rowsCsv.length < 2) throw new Error("CSV: il faut au moins une ligne d'entête + une ligne d'élève.");
  const header = rowsCsv[0]; // ["eleve", "1", "2", "3a", ...] ou juste noms de colonnes
  if (!header.length) throw new Error("CSV: entête vide.");

  // 2) Ouvrir DOCX
  const zip = await new JSZip().loadAsync(await fs.readFile(srcPath));
  const docXmlPath = "word/document.xml";
  const xmlBuf = await zip.file(docXmlPath)!.async("nodebuffer");
  const xmlStr = xmlBuf.toString("utf8");

  const dom = new DOMParser().parseFromString(xmlStr, "application/xml");
  const documentElement = dom.documentElement;

  // 3) Sélection du tableau
  const tbls = findAllByLocalName(documentElement, "tbl");
  if (!tbls.length) throw new Error("Aucune table trouvée dans le document.");
  const tbl = tbls[tableIndex] ?? tbls[0];

  const trRows = childrenByLocalName(tbl, "tr");
  if (trRows.length <= headerRowIndex) throw new Error("Index de la ligne des mots hors bornes.");

  // 4) Pour chaque élève (une ligne CSV)
  for (let r = 1; r < rowsCsv.length; r++) {
    const csvRow = rowsCsv[r];
    if (!csvRow || csvRow.length === 0) continue;

    // Col 0 = numéro d'élève (obligatoire)
    const eleveNum = Number((csvRow[0] ?? "").trim());
    if (!Number.isFinite(eleveNum)) continue;

    const rowIndex = eleveNum + studentRowOffset;
    if (rowIndex >= trRows.length) {
      console.warn(`⚠️ Ignoré: élève ${eleveNum} (row index ${rowIndex}) > nb lignes (${trRows.length}).`);
      continue;
    }

    const tr = trRows[rowIndex];
    const tcs = childrenByLocalName(tr, "tc");

    // Les colonnes CSV suivantes correspondent 1:1 aux colonnes "mots" du tableau,
    // alignées à partir de firstDataColIndex dans le DOCX.
    for (let c = 1; c < csvRow.length; c++) {
      const v = (csvRow[c] ?? "").trim();
      const tcIndex = firstDataColIndex + (c - 1);
      if (tcIndex >= tcs.length) break;

      // Règle de couleur: "1" => vert, "2" => rouge, "" => jaune, sinon on ne touche pas
      if (v === "1") {
        shadeCell(tcs[tcIndex], dom as unknown as Document, colorFor1);
      } else if (v === "2") {
        shadeCell(tcs[tcIndex], dom as unknown as Document, colorFor2);
      } else if (v === "") {
        shadeCell(tcs[tcIndex], dom as unknown as Document, colorForEmpty);
      } else {
        // valeur inattendue: on ignore (ou tu peux décider de nettoyer/neutraliser)
        // pour neutraliser: shadeCell(tcs[tcIndex], dom as unknown as Document, "FFFFFF");
      }
    }
  }

  // 5) Écrire le DOCX
  const updatedXml = new XMLSerializer().serializeToString(dom);
  zip.file(docXmlPath, Buffer.from(updatedXml, "utf8"));
  const outBuf = await zip.generateAsync({ type: "nodebuffer" });
  await fs.writeFile(path.resolve(outPath), outBuf);
  console.log(`✅ Écrit: ${outPath}`);
}

main().catch(e => { console.error(e); process.exit(1); });
