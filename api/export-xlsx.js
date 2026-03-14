// api/export-xlsx.js
// Robust: Mapping beliebiger Template-Header → Persona-Variablen + Schreiben in die Vorlage.
// Features:
// - Heuristisches Spalten-Mapping (Q2/Q3/Q7/Q9/Q10/Q11)
// - Q10-Duplikate → 1.=Q10_Marke, 2.=Q10_Marke_Andere
// - Sonderregel: wenn Q10_Marke !== "10" → Q10_Marke_Andere = ""
// - Preview (?preview=1) zeigt columns/firstRow/count (ohne XLSX-Schreiben)
// - Writer: liest Header über Scoring oder fällt auf Codebook-Reihenfolge zurück
// - "No./Anzahl/Nr" wird erkannt und 1..N durchnummeriert

/* -------------------------- Normalisierung & Heuristik -------------------------- */

function normalize(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

// Heuristik: Template-Spalte → Variablen-Key im Persona-Datensatz
function guessVarForColumn(rawName, personaKeys) {
  const col = normalize(rawName);

  // 1) Exakte/normalisierte Übereinstimmung mit vorhandenen Persona-Keys
  if (personaKeys?.has?.(rawName)) return rawName;
  if (personaKeys) {
    const byNorm = [...personaKeys].find(k => normalize(k) === col);
    if (byNorm) return byNorm;
  }

  // 2) Häufige Familien/Synonyme
  if (col.includes("q10") && (col.includes("andere") || col.includes("andern") || col.includes("other"))) {
    return "Q10_Marke_Andere";
  }
  if (col.includes("q10")) return "Q10_Marke";

  if (col.includes("fett")) return "Q11_Fettgehalt";
  if (col.includes("ablehn") || col.includes("keine")) return "Q9_Ablehnung";

  if (col.includes("frisch")) return "Q7_Essverhalten_Frischkaese";
  if (col.includes("gouda")) return "Q7_Essverhalten_Gouda";
  if (col.includes("butter")) return "Q7_Essverhalten_Butterkaese";
  if (col.includes("camembert")) return "Q7_Essverhalten_Camembert";

  if (col.includes("gender") || col.includes("geschlecht")) return "Q2_Gender";
  if (col.includes("age") || col.includes("alter") || col.includes("altersgruppe")) return "Q3_Age";

  if (col === "no" || col.includes("anzahl") || col === "nr") return "No";

  // 3) Fallback
  return null;
}

function buildColumnMapping(codebook, samplePersonaRow) {
  const columns =
    Array.isArray(codebook?.column_order_raw) && codebook.column_order_raw.length
      ? [...codebook.column_order_raw]
      : Array.isArray(codebook?.variables)
        ? codebook.variables.map(v => v.raw || v.var).filter(Boolean)
        : [];

  const personaKeys = new Set(Object.keys(samplePersonaRow || {}));

  const mapping = new Map(); // raw -> varKey|null
  for (const raw of columns) {
    mapping.set(raw, guessVarForColumn(raw, personaKeys));
  }

  // Duplikate gruppieren
  const groups = new Map();
  for (const raw of columns) {
    const nk = normalize(raw);
    const base =
      nk.includes("q10") ? "Q10" :
      nk.includes("q11") ? "Q11" :
      nk.includes("q9")  ? "Q9"  :
      nk.includes("q7")  ? "Q7"  :
      nk;
    if (!groups.has(base)) groups.set(base, []);
    groups.get(base).push(raw);
  }

  // Q10: genau 2 Spalten → 1. Marke, 2. Andere
  if (groups.has("Q10")) {
    const q10cols = groups.get("Q10");
    if (q10cols.length === 2) {
      const [c1, c2] = q10cols;
      mapping.set(c1, mapping.get(c1) ?? "Q10_Marke");
      mapping.set(c2, mapping.get(c2) ?? "Q10_Marke_Andere");
    }
  }

  return mapping;
}

function buildTemplateRow(personaRow, mapping) {
  const row = { ...personaRow };

  // Sonderregel: "Andere" nur wenn Marke == "10"
  if (String(row["Q10_Marke"] ?? "") !== "10") {
    row["Q10_Marke_Andere"] = "";
  }

  const out = {};
  for (const [raw, vKey] of mapping.entries()) {
    if (!vKey) { out[raw] = ""; continue; }
    const val = row[vKey];
    out[raw] = (val === null || val === undefined) ? "" : String(val);
  }
  return out;
}

/* ------------------------------- XLSX-Writer ---------------------------------- */

async function buildXlsxFromTemplate(templateBase64, rows, qa, prefix, fallzahl) {
  const ExcelJS = (await import('exceljs')).default;

  const tplBuf = Buffer.from(templateBase64, 'base64');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(tplBuf);

  const ws = workbook.worksheets[0] || workbook.addWorksheet('Daten');

  // 1) Start der Daten (erste vollständig leere Zeile)
  let startRow = findFirstFreeRow(ws, 300, 1);
  // Optional Abstand: startRow += 1;

  // 2) Kopfzeile mit Scoring finden (beste Übereinstimmung mit Codebook-Variablen)
  const codebookNorm = new Set(
    rows.length ? Object.keys(rows[0]).map(k => normalize(k)) : []
  );
  const headerRowIdx = findBestHeaderRow(ws, startRow - 1, codebookNorm, 400);

  // 3) Wenn wir eine gute Kopfzeile haben → spaltenexakt unter Header schreiben
  if (headerRowIdx) {
    const header = readHeaderCells(ws, headerRowIdx, 400); // [{col, raw, norm}, ...]
    const noCol = findNoColumn(header) || 1;

    let r = startRow;
    for (let i = 0; i < rows.length; i++) {
      const obj = rows[i];                 // <- mappedRows (Keys: Codebook-raw)
      const personaKeys = new Set(Object.keys(obj));
      const row = ws.getRow(r);

      // Laufende Nummer in "No."-Spalte
      row.getCell(noCol).value = i + 1;

      // Für jede Kopfspalte den passenden Wert finden (exakt → Heuristik → normalisiert)
      for (const h of header) {
        if (!h.raw) continue;
        if (isNoLike(h.norm)) continue;
        const v = getValueForHeader(obj, h, personaKeys);
        if (v !== undefined && v !== null && v !== '') {
          row.getCell(h.col).value = String(v);
        }
      }

      row.commit();
      r++;
    }
  } else {
    // 4) Fallback: Kein verlässlicher Header gefunden
    //    → Schreibe No. in Spalte A, alle Variablen in Codebook-Reihenfolge ab Spalte B
    const columnsOrder = Array.from(
      rows.reduce((set, r) => { Object.keys(r).forEach(k => set.add(k)); return set; }, new Set())
    );
    const startCol = 2;  // B
    let r = startRow;

    for (let i = 0; i < rows.length; i++) {
      const obj = rows[i];
      const row = ws.getRow(r);

      // No. in A
      row.getCell(1).value = i + 1;

      // Daten ab B in Codebook-Reihenfolge
      columnsOrder.forEach((key, idx) => {
        const v = obj[key];
        if (v !== undefined && v !== null && v !== '') {
          row.getCell(startCol + idx).value = String(v);
        }
      });

      row.commit();
      r++;
    }
  }

  // 5) QA in 2. Sheet (optional)
  if (qa && typeof qa === 'object' && Object.keys(qa).length) {
    const qaWs = workbook.addWorksheet('QA');
    let i = 1;
    for (const [k, v] of Object.entries(qa)) {
      qaWs.getCell(i, 1).value = k;
      qaWs.getCell(i, 2).value = (typeof v === 'object') ? JSON.stringify(v) : String(v ?? '');
      i++;
    }
  }

  const out = await workbook.xlsx.writeBuffer();
  return Buffer.from(out).toString('base64');
}

/* ----------------------------- Writer-Helfer ---------------------------------- */

// Header passend befüllen (exakt → Heuristik → normalisiert)
function getValueForHeader(obj, headerCell, personaKeys) {
  const raw = headerCell.raw;
  const norm = headerCell.norm;

  // 1) exakter Headertitel
  if (obj.hasOwnProperty(raw) && obj[raw] != null && obj[raw] !== '') return obj[raw];

  // 2) Heuristischer Var-Key
  const varKey = guessVarForColumn(raw, personaKeys);
  if (varKey && obj.hasOwnProperty(varKey) && obj[varKey] != null && obj[varKey] !== '') return obj[varKey];

  // 3) Normalisierte Übereinstimmung
  const match = Object.keys(obj).find(k => normalize(k) === norm);
  if (match && obj[match] != null && obj[match] !== '') return obj[match];

  return '';
}

function findBestHeaderRow(ws, fromRow, codebookNormSet, maxCols = 200, window = 50) {
  // Suche bis zu 'window' Zeilen oberhalb von fromRow die Zeile mit bester Übereinstimmung
  // Score: Anzahl Zellen in der Zeile, deren normalize() in codebookNormSet enthalten ist,
  // plus Bonus für No./Anzahl/Nr.
  let best = { row: null, score: -1 };

  const start = Math.max(1, fromRow - window);
  for (let r = fromRow; r >= start; r--) {
    let score = 0, seenAny = false;
    const row = ws.getRow(r);
    for (let c = 1; c <= maxCols; c++) {
      const txt = cellToString(row.getCell(c)?.value);
      if (!txt) continue;
      seenAny = true;
      const norm = normalize(txt);
      if (codebookNormSet.has(norm)) score += 2;       // starker Treffer
      if (isNoLike(norm))            score += 1;       // No./Anzahl
    }
    if (seenAny && score > best.score) best = { row: r, score };
  }

  // Mindestscore (3) verlangen, um "Legenden"-Zeilen zu vermeiden
  return best.score >= 3 ? best.row : null;
}

function findHeaderRowAbove(ws, fromRow, maxCols = 150) {
  for (let r = fromRow; r >= 1; r--) {
    let filled = 0;
    for (let c = 1; c <= maxCols; c++) {
      const v = cellToString(ws.getRow(r).getCell(c)?.value);
      if (v) filled++;
      if (filled >= 3) return r;
    }
  }
  return null;
}

function readHeaderCells(ws, headerRowIdx, maxCols = 150) {
  const out = [];
  const row = ws.getRow(headerRowIdx);
  for (let c = 1; c <= maxCols; c++) {
    const raw = cellToString(row.getCell(c)?.value);
    if (!raw) continue;
    out.push({ col: c, raw, norm: normalize(raw) });
  }
  return out;
}

function findNoColumn(headerCells) {
  for (const h of headerCells) {
    if (isNoLike(h.norm)) return h.col;
  }
  return null;
}
function isNoLike(normText) {
  return normText === 'no' ||
         normText.startsWith('no_') ||
         normText.includes('anzahl') ||
         normText === 'nr' ||
         normText.startsWith('nr_');
}

function findFirstFreeRow(ws, maxCols = 150, consecutiveEmpty = 1) {
  const last = ws.lastRow ? ws.lastRow.number : 1;
  const scanTo = last + 200;
  for (let r = 1; r <= scanTo; r++) {
    if (isRowEmpty(ws, r, maxCols)) {
      let ok = true;
      for (let k = 1; k < consecutiveEmpty; k++) {
        if (!isRowEmpty(ws, r + k, maxCols)) { ok = false; break; }
      }
      if (ok) return r;
    }
  }
  return last + 1;
}
function isRowEmpty(ws, r, maxCols = 150) {
  const row = ws.getRow(r);
  for (let c = 1; c <= maxCols; c++) {
    const v = cellToString(row.getCell(c)?.value);
    if (v) return false;
  }
  return true;
}
function cellToString(v) {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.text) return String(v.text).trim();
    if (v.richText && Array.isArray(v.richText)) {
      return v.richText.map(rt => rt.text).join('').trim();
    }
    if (v.result != null) return String(v.result).trim();
    return String(v).trim();
  }
  return String(v).trim();
}

/* --------------------------------- Handler ------------------------------------ */

export default async function handler(req, res) {
  try {
    if (req.method !== 'POST') {
      return res.status(405).json({ error: 'Method not allowed' });
    }

    const {
      templateBase64 = '',
      rows = [],
      qa = {},
      prefix = '',
      fallzahl = 0,
      codebook = null,
    } = req.body || {};

    if (!templateBase64) return res.status(400).json({ error: 'templateBase64 fehlt' });
    if (!Array.isArray(rows) || !rows.length) return res.status(400).json({ error: 'rows fehlt/leer' });
    if (!codebook) return res.status(400).json({ error: 'codebook fehlt' });

    // 1) Mapping
    const mapping = buildColumnMapping(codebook, rows[0]);

    // 2) mappedRows (Keys = Codebook-raw)
    const mappedRows = rows.map(r => buildTemplateRow(r, mapping));

    // 3) Preview
    if (String(req.query?.preview ?? "") === "1") {
      return res.status(200).json({
        preview: {
          columns: [...mapping.keys()],
          firstRow: mappedRows[0] || null,
          count: mappedRows.length
        }
      });
    }

    // 4) Schreiben
    const fileBase64 = await buildXlsxFromTemplate(
      templateBase64,
      mappedRows,
      qa,
      prefix,
      fallzahl
    );

    return res.status(200).json({ file: fileBase64 });
  } catch (err) {
    console.error("export-xlsx error:", err);
    return res.status(500).json({ error: 'Server error', details: String(err?.message || err) });
  }
}
