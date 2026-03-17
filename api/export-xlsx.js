// api/export-xlsx.js
// Robust: Mapping beliebiger Template-Header → Persona-Variablen + Schreiben in die Vorlage.

// ---------------- Normalisierung & Heuristik ----------------

function normalize(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

function guessVarForColumn(rawName, personaKeys) {
  const col = normalize(rawName);

  // 1) exakter/normalisierter Treffer
  if (personaKeys?.has?.(rawName)) return rawName;
  if (personaKeys) {
    const byNorm = [...personaKeys].find(k => normalize(k) === col);
    if (byNorm) return byNorm;
  }

  // 2) Heuristiken
  if (col.includes("q10") && (col.includes("andere") || col.includes("andern") || col.includes("other")))
    return "Q10_Marke_Andere";
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
  for (const raw of columns) mapping.set(raw, guessVarForColumn(raw, personaKeys));

  // Duplikat-Gruppen
  const groups = new Map();
  for (const raw of columns) {
    const nk = normalize(raw);
    const base =
      nk.includes("q10") ? "Q10" :
      nk.includes("q11") ? "Q11" :
      nk.includes("q9")  ? "Q9"  :
      nk.includes("q7")  ? "Q7"  : nk;
    if (!groups.has(base)) groups.set(base, []);
    groups.get(base).push(raw);
  }

  // Q10: 2 Spalten → 1.=Marke, 2.=Andere
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
  if (String(row["Q10_Marke"] ?? "") !== "10") row["Q10_Marke_Andere"] = ""; // Sonderregel

  const out = {};
  for (const [raw, vKey] of mapping.entries()) {
    if (!vKey) { out[raw] = ""; continue; }
    const val = row[vKey];
    out[raw] = (val === null || val === undefined) ? "" : String(val);
  }
  return out;
}

// ---------------- Robustes Decoding & XLSX-Writer ----------------

// Entfernt Daten-URL-Präfixe und prüft ZIP-"PK"
function toXlsxBuffer(b64) {
  const s = String(b64 || '').trim();
  const clean = s.includes('base64,') ? s.split('base64,').pop() : s;
  const buf = Buffer.from(clean, 'base64');
  if (buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4B) {
    throw new Error(`TemplateBase64 ist kein XLSX/ZIP (len=${buf.length}, head=${buf.slice(0,4).toString('hex')}).`);
  }
  return buf;
}

function cellToString(v) {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.text) return String(v.text).trim();
    if (v.richText && Array.isArray(v.richText)) return v.richText.map(rt => rt.text).join('').trim();
    if (v.result != null) return String(v.result).trim();
    return String(v).trim();
  }
  return String(v).trim();
}

function isNoLike(normText) {
  return normText === 'no' ||
         normText.startsWith('no_') ||
         normText.includes('anzahl') ||
         normText === 'nr' ||
         normText.startsWith('nr_');
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

function findFirstFreeRow(ws, maxCols = 150, consecutiveEmpty = 1) {
  const last = ws.lastRow ? ws.lastRow.number : 1;
  const scanTo = last + 200;
  for (let r = 1; r <= scanTo; r++) {
    const row = ws.getRow(r);
    let any = false;
    for (let c = 1; c <= maxCols; c++) {
      const val = cellToString(row.getCell(c)?.value);
      if (val) { any = true; break; }
    }
    if (!any) {
      let ok = true;
      for (let k = 1; k < consecutiveEmpty; k++) {
        const next = ws.getRow(r + k);
        let any2 = false;
        for (let c = 1; c <= maxCols; c++) {
          const val2 = cellToString(next.getCell(c)?.value);
          if (val2) { any2 = true; break; }
        }
        if (any2) { ok = false; break; }
      }
      if (ok) return r;
    }
  }
  return last + 1;
}

// Score-basierte Suche der echten Kopfzeile (unterhalb von Legenden/Skalen)
function findBestHeaderRow(ws, fromRow, codebookNormSet, maxCols = 200, window = 60) {
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
      if (codebookNormSet.has(norm)) score += 2;
      if (isNoLike(norm))            score += 1;
    }
    if (seenAny && score > best.score) best = { row: r, score };
  }
  return best.score >= 3 ? best.row : null;
}

// Smarte Wertefindung (exakt → heuristisch → normalisiert)
function getValueForHeader(obj, headerCell, personaKeys) {
  const raw = headerCell.raw;
  const norm = headerCell.norm;

  if (obj.hasOwnProperty(raw) && obj[raw] != null && obj[raw] !== '') return obj[raw];
  const varKey = guessVarForColumn(raw, personaKeys);
  if (varKey && obj.hasOwnProperty(varKey) && obj[varKey] != null && obj[varKey] !== '') return obj[varKey];
  const match = Object.keys(obj).find(k => normalize(k) === norm);
  if (match && obj[match] != null && obj[match] !== '') return obj[match];
  return '';
}

async function buildXlsxFromTemplate(templateBase64, rows, qa, prefix, fallzahl) {
  const ExcelJS = (await import('exceljs')).default;

  const tplBuf = toXlsxBuffer(templateBase64);
  console.log('tplLen=%d rows=%d magic=%s', tplBuf.length, rows.length, tplBuf.slice(0,2).toString());

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(tplBuf);

  const ws = workbook.worksheets[0] || workbook.addWorksheet('Daten');

  // Start der Daten (erste vollständig leere Zeile)
  const startRow = findFirstFreeRow(ws, 300, 1);

  // Header anhand Codebook-Normalformen scoren
  const codebookNorm = new Set(
    rows.length ? Object.keys(rows[0]).map(k => normalize(k)) : []
  );
  const headerRowIdx = findBestHeaderRow(ws, startRow - 1, codebookNorm, 400);

  if (headerRowIdx) {
    const header = readHeaderCells(ws, headerRowIdx, 400);
    const noCol = (function findNo(headerCells){
      for (const h of headerCells) if (isNoLike(h.norm)) return h.col;
      return 1;
    })(header);

    let r = startRow;
    for (let i = 0; i < rows.length; i++) {
      const obj = rows[i];
      const personaKeys = new Set(Object.keys(obj));
      const row = ws.getRow(r);

      // No. fortlaufend
      row.getCell(noCol).value = i + 1;

      // Werte spaltenexakt
      for (const h of header) {
        if (!h.raw || isNoLike(h.norm)) continue;
        const v = getValueForHeader(obj, h, personaKeys);
        if (v !== '') row.getCell(h.col).value = String(v);
      }
      row.commit();
      r++;
    }
  } else {
    // Fallback: Ohne verlässlichen Header → No. in A, Spalten ab B in Codebook-Reihenfolge
    const cols = Array.from(
      rows.reduce((set, r) => { Object.keys(r).forEach(k => set.add(k)); return set; }, new Set())
    );
    let r = startRow;
    for (let i = 0; i < rows.length; i++) {
      const obj = rows[i];
      const row = ws.getRow(r);
      row.getCell(1).value = i + 1; // No. in A
      cols.forEach((key, idx) => {
        const v = obj[key];
        if (v !== undefined && v !== null && v !== '') row.getCell(2 + idx).value = String(v);
      });
      row.commit();
      r++;
    }
  }

  // QA-Sheet (optional)
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

// ---------------- HTTP-Handler ----------------

export default async function handler(req, res) {
  try {
    if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

    const {
      templateBase64 = '',
      rows = [],
      qa = {},
      prefix = '',
      fallzahl = 0,
      codebook = null
    } = req.body || {};

    if (!templateBase64) return res.status(400).json({ error: 'templateBase64 fehlt' });
    if (!Array.isArray(rows) || !rows.length) return res.status(400).json({ error: 'rows fehlt/leer' });
    if (!codebook) return res.status(400).json({ error: 'codebook fehlt' });

    // Mapping aufbauen → Template-Rows (Keys = raw-Spaltennamen)
    const mapping = buildColumnMapping(codebook, rows[0]);
    const mappedRows = rows.map(r => buildTemplateRow(r, mapping));

    // Preview-Modus
    if (String(req.query?.preview ?? "") === "1") {
      return res.status(200).json({
        preview: {
          columns: [...mapping.keys()],
          firstRow: mappedRows[0] || null,
          count: mappedRows.length
        }
      });
    }

    // XLSX schreiben
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
