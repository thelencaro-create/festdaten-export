// api/export-xlsx.js
// Robust: Mapping beliebiger Template-Header → Persona-Variablen + XLSX-Schreiben mit exceljs.
// Features:
// - Heuristisches Spalten-Mapping (Q2/Q3/Q7/Q9/Q10/Q11 etc.)
// - Duplikate (z. B. zwei Q10-Spalten) → deterministisch: 1.=Q10_Marke, 2.=Q10_Marke_Andere
// - Sonderregel: Wenn Q10_Marke !== "10" → Q10_Marke_Andere = ""
// - Preview (?preview=1) zeigt columns/firstRow/count ohne XLSX-Erzeugung
// - Writer findet automatisch die erste vollständig leere Zeile unterhalb deines Templates

function normalize(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

// Heuristiken für Spalten → Variablen-Key
function guessVarForColumn(rawName, personaKeys) {
  const col = normalize(rawName);

  // 1) exakte/normalisierte Übereinstimmung mit vorhandenen Persona-Keys
  if (personaKeys.has(rawName)) return rawName;
  const byNorm = [...personaKeys].find(k => normalize(k) === col);
  if (byNorm) return byNorm;

  // 2) häufige Variablenfamilien / Synonyme
  if (col.includes("q10") && (col.includes("andere") || col.includes("andern") || col.includes("other"))) {
    return "Q10_Marke_Andere";
  }
  if (col.includes("q10")) {
    return "Q10_Marke";
  }
  if (col.includes("fett")) return "Q11_Fettgehalt";
  if (col.includes("ablehn") || col.includes("keine")) return "Q9_Ablehnung";

  if (col.includes("frisch")) return "Q7_Essverhalten_Frischkaese";
  if (col.includes("gouda")) return "Q7_Essverhalten_Gouda";
  if (col.includes("butter")) return "Q7_Essverhalten_Butterkaese";
  if (col.includes("camembert")) return "Q7_Essverhalten_Camembert";

  if (col.includes("gender") || col.includes("geschlecht")) return "Q2_Gender";
  if (col.includes("age") || col.includes("alter") || col.includes("altersgruppe")) return "Q3_Age";

  if (col === "no" || col.includes("nummer")) return "No";

  // 3) Fallback
  return null;
}

// Spalten-Mapping aus Codebook + Heuristik bauen
function buildColumnMapping(codebook, samplePersonaRow) {
  // Template-Spalten: bevorzugt column_order_raw; Fallback: variables[].raw/var
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

  // 2) Duplikate pro „Basis“
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

  // Speziell Q10: genau zwei Spalten → 1.=Marke, 2.=Marke_Andere (wenn noch nicht gesetzt)
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

// Persona → Template-Row (Keys = raw-Spalten); inkl. Sonderregel Q10_Marke_Andere
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

// ====== Writer-Teil mit exceljs ======
async function buildXlsxFromTemplate(templateBase64, rows, qa, prefix, fallzahl) {
  const ExcelJS = (await import('exceljs')).default;

  // 1) Vorlage laden
  const tplBuf = Buffer.from(templateBase64, 'base64');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(tplBuf);

  // 2) Erstes Worksheet (kein fixer Name nötig)
  const ws = workbook.worksheets[0] || workbook.addWorksheet('Daten');

  // 3) Spaltenreihenfolge aus rows ableiten (stabil)
  const columnsOrder = Array.from(
    rows.reduce((set, r) => {
      Object.keys(r).forEach(k => set.add(k));
      return set;
    }, new Set())
  );

  // 4) Erste vollständig leere Zeile unterhalb der Vorlage erkennen
  const startRow = findFirstFreeRow(ws, Math.max(columnsOrder.length, 200), 1);
  // Falls du bewusst eine Leerzeile Abstand willst:  const startRow = findFirstFreeRow(ws, Math.max(columnsOrder.length, 200), 1) + 1;

  // 5) Kopfzeile schreiben
  const headerRow = ws.getRow(startRow);
  columnsOrder.forEach((key, idx) => {
    headerRow.getCell(idx + 1).value = key;
  });
  headerRow.font = { bold: true };
  headerRow.commit();

  // 6) Datenzeilen
  let r = startRow + 1;
  for (const obj of rows) {
    const row = ws.getRow(r);
    columnsOrder.forEach((key, idx) => {
      const v = obj[key];
      row.getCell(idx + 1).value = (v === null || v === undefined) ? '' : String(v);
    });
    row.commit();
    r++;
  }

  // 7) Optional: QA in 2. Sheet
  if (qa && typeof qa === 'object' && Object.keys(qa).length) {
    const qaWs = workbook.addWorksheet('QA');
    let i = 1;
    for (const [k, v] of Object.entries(qa)) {
      qaWs.getCell(i, 1).value = k;
      qaWs.getCell(i, 2).value = (typeof v === 'object') ? JSON.stringify(v) : String(v ?? '');
      i++;
    }
  }

  // 8) → Base64 zurückgeben
  const out = await workbook.xlsx.writeBuffer();
  return Buffer.from(out).toString('base64');
}

// ——— Hilfsfunktionen: erste freie Zeile finden ———
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

function isRowEmpty(ws, r, maxCols) {
  const row = ws.getRow(r);
  for (let c = 1; c <= maxCols; c++) {
    const cell = row.getCell(c);
    const val = cell?.value;
    if (val !== null && val !== undefined && String(val).trim() !== '') {
      return false;
    }
  }
  return true;
}

// ——— HTTP-Handler ———
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

    // 2) Persona-Rows → Template-Rows (Keys = raw-Spalten)
    const mappedRows = rows.map(r => buildTemplateRow(r, mapping));

    // 3) Preview-Modus (nur Mapping prüfen, keine Datei schreiben)
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
