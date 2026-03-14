// api/export-xlsx.js
// Robust: Mapping beliebiger Template-Header → Persona-Variablen + Schreiben in die Vorlage.
// Features:
// - Heuristisches Spalten-Mapping (Q2/Q3/Q7/Q9/Q10/Q11)
// - Q10-Duplikate → 1.=Q10_Marke, 2.=Q10_Marke_Andere
// - Sonderregel: wenn Q10_Marke !== "10" → Q10_Marke_Andere = ""
// - Preview (?preview=1) zeigt columns/firstRow/count (ohne XLSX-Schreiben)
// - Writer liest die bestehende Kopfzeile (keine neue Header-Zeile), findet erste leere Datenzeile
// - "No."/ "Anzahl"/ "Nr" wird erkannt und 1..N durchnummeriert

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
  if (personaKeys.has(rawName)) return rawName;
  const byNorm = [...personaKeys].find(k => normalize(k) === col);
  if (byNorm) return byNorm;

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
  // Spalten aus dem Codebook: bevorzugt column_order_raw, sonst variables[].raw/var
  const columns =
    Array.isArray(codebook?.column_order_raw) && codebook.column_order_raw.length
      ? [...codebook.column_order_raw]
      : Array.isArray(codebook?.variables)
        ? codebook.variables.map(v => v.raw || v.var).filter(Boolean)
        : [];

  const personaKeys = new Set(Object.keys(samplePersonaRow || {}));

  // 1) Basismapping
  const mapping = new Map(); // raw -> varKey|null
  for (const raw of columns) {
    mapping.set(raw, guessVarForColumn(raw, personaKeys));
  }

  // 2) Duplikate nach "Basis" clustern
  const groups = new Map(); // base -> raw[]
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

  // Q10: genau zwei Spalten -> 1.=Marke, 2.=Marke_Andere (falls noch nicht gesetzt)
  if (groups.has("Q10")) {
    const q10cols = groups.get("Q10");
    if (q10cols.length === 2) {
      const [c1, c2] = q10cols;
      mapping.set(c1, mapping.get(c1) ?? "Q10_Marke");
      mapping.set(c2, mapping.get(c2) ?? "Q10_Marke_Andere");
    }
  }

  return mapping; // Map<rawColumn, varKey|null>
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

  // 1) Vorlage laden
  const tplBuf = Buffer.from(templateBase64, 'base64');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(tplBuf);

  // 2) Erstes Worksheet (kein fixer Name nötig)
  const ws = workbook.worksheets[0] || workbook.addWorksheet('Daten');

  // 3) Erste vollständig leere Zeile = Datenstart (unterhalb Fragen/Antwortblock)
  const startRow = findFirstFreeRow(ws, 300, 1);
  // Falls du bewusst 1 Leerzeile Abstand wünschst:  const startRow = findFirstFreeRow(ws, 300, 1) + 1;

  // 4) Kopfzeile ermitteln: letzte "inhaltstarke" Zeile darüber
  const headerRowIdx = findHeaderRowAbove(ws, startRow - 1, 300);
  if (!headerRowIdx) throw new Error('Kopfzeile im Template nicht gefunden – bitte prüfen.');

  const header = readHeaderCells(ws, headerRowIdx, 300); // [{col, raw, norm}, ...]
  if (!header.length) throw new Error('Leere Kopfzeile erkannt – bitte Template prüfen.');

  // 5) "No."-Spalte erkennen (oder Default Spalte 1)
  const noCol = findNoColumn(header) || 1;

  // 6) Rohspaltennamen → Spaltenindex abbilden
  const indexByRaw = new Map();
  for (const h of header) {
    if (!h.raw) continue;
    indexByRaw.set(h.raw, h.col);
  }

  // 7) Daten schreiben (keine neue Kopfzeile!)
  let r = startRow;
  for (let i = 0; i < rows.length; i++) {
    const obj = rows[i]; // <- bereits "mappedRows"
    const row = ws.getRow(r);

    // 7a) Laufende Nummer
    row.getCell(noCol).value = i + 1;

    // 7b) Spalten exakt per vorhandener Header-Beschriftung befüllen
    for (const h of header) {
      if (!h.raw) continue;
      if (isNoLike(h.norm)) continue; // "No." nicht überschreiben

      const v = obj[h.raw];
      if (v !== undefined && v !== null) {
        row.getCell(h.col).value = String(v);
      } else {
        // bewusst leer lassen → kein Shift/Einschub
      }
    }

    row.commit();
    r++;
  }

  // 8) Optional: QA in 2. Sheet
  if (qa && typeof qa === 'object' && Object.keys(qa).length) {
    const qaWs = workbook.addWorksheet('QA');
    let i = 1;
    for (const [k, v] of Object.entries(qa)) {
      qaWs.getCell(i, 1).value = k;
      qaWs.getCell(i, 2).value = (typeof v === 'object') ? JSON.stringify(v) : String(v ?? '');
      i++;
    }
  }

  // 9) → Base64 zurückgeben
  const out = await workbook.xlsx.writeBuffer();
  return Buffer.from(out).toString('base64');
}

/* ----------------------------- Writer-Helfer ---------------------------------- */

function findHeaderRowAbove(ws, fromRow, maxCols = 150) {
  // Nimm die letzte Zeile vor startRow, die "deutlich" Inhalt hat (>= 3 befüllte Zellen)
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
  const scanTo = last + 200; // etwas Puffer
  for (let r = 1; r <= scanTo; r++) {
    if (isRowEmpty(ws, r, maxCols)) {
      // Optional: mehrere leere Zeilen am Stück verlangen
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
    if (v.text) return String(v.text).trim();                // Plain/RichText
    if (v.richText && Array.isArray(v.richText)) {
      return v.richText.map(rt => rt.text).join('').trim();
    }
    if (v.result != null) return String(v.result).trim();    // Formelergebnis
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

    // 1) Mapping einmalig aufbauen
    const mapping = buildColumnMapping(codebook, rows[0]);

    // 2) Persona-Rows → Template-Rows (Keys = raw-Spaltennamen aus Template)
    const mappedRows = rows.map(r => buildTemplateRow(r, mapping));

    // 3) Preview-Modus: nur Mapping prüfen
    if (String(req.query?.preview ?? "") === "1") {
      return res.status(200).json({
        preview: {
          columns: [...mapping.keys()],
          firstRow: mappedRows[0] || null,
          count: mappedRows.length
        }
      });
    }

    // 4) XLSX schreiben
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
