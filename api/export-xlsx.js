import ExcelJS from "exceljs";

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "15mb"
    }
  }
};

/**
 * Hilfen zur robusten Header-Zuordnung (ohne das Template zu verändern):
 * - Wir lesen die echten Spaltenüberschriften aus Zeile 1 des Templates.
 * - Wir bauen ein Mapping von "Template-Header" -> "Key in rows".
 * - Exakte Übereinstimmung hat Priorität, sonst Normalisierung / Heuristiken.
 * - Startzeile = 3 (Row 1 = Header, Row 2 = Legenden).
 */
function normalize(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

// Bekannte Alias-Muster von Variablennamen (Template-Header -> rows-Key)
function guessRowKeyForHeader(header, rowKeysNormMap) {
  const h = normalize(header);

  // 1) Direkte Normalisierung: exakter Norm-Match auf vorhandene row-Keys
  if (rowKeysNormMap.has(h)) return rowKeysNormMap.get(h);

  // 2) Spezifische Heuristiken (Q-Codes & deutschsprachige Header)
  if (h.includes("q2") || h.includes("gender") || h.includes("geschlecht")) {
    for (const k of rowKeysNormMap.keys()) if (k.startsWith("q2")) return rowKeysNormMap.get(k);
  }
  if (h.includes("q3") || h.includes("alter")) {
    for (const k of rowKeysNormMap.keys()) if (k.startsWith("q3")) return rowKeysNormMap.get(k);
  }
  if (h.includes("q7") && h.includes("frisch")) {
    for (const k of rowKeysNormMap.keys())
      if (k.includes("q7") && (k.includes("frisch") || k.includes("frischkaese")))
        return rowKeysNormMap.get(k);
  }
  if (h.includes("q7") && h.includes("gouda")) {
    for (const k of rowKeysNormMap.keys())
      if (k.includes("q7") && k.includes("gouda"))
        return rowKeysNormMap.get(k);
  }
  if (h.includes("q7") && (h.includes("butter") || h.includes("butterkaese"))) {
    for (const k of rowKeysNormMap.keys())
      if (k.includes("q7") && (k.includes("butter") || k.includes("butterkaese")))
        return rowKeysNormMap.get(k);
  }
  if (h.includes("q7") && h.includes("camembert")) {
    for (const k of rowKeysNormMap.keys())
      if (k.includes("q7") && k.includes("camembert"))
        return rowKeysNormMap.get(k);
  }
  if (h.includes("q9") && (h.includes("ablehn") || h.includes("keine"))) {
    for (const k of rowKeysNormMap.keys())
      if (k.includes("q9"))
        return rowKeysNormMap.get(k);
  }
  if (h.includes("q10") && h.includes("andere")) {
    // z. B. "Q10: Welches Produkt ...", zweite Spalte für "Andere"
    for (const k of rowKeysNormMap.keys())
      if (k.includes("q10") && (k.includes("andere") || k.includes("andern") || k.includes("other")))
        return rowKeysNormMap.get(k);
    // Fallback: "_Andere"
    for (const k of rowKeysNormMap.keys())
      if (k.endsWith("_andere"))
        return rowKeysNormMap.get(k);
  }
  if (h.includes("q10")) {
    for (const k of rowKeysNormMap.keys())
      if (k.includes("q10") && !k.endsWith("_andere"))
        return rowKeysNormMap.get(k);
  }
  if (h.includes("q11") || h.includes("fett")) {
    for (const k of rowKeysNormMap.keys())
      if (k.includes("q11") || k.includes("fett"))
        return rowKeysNormMap.get(k);
  }

  // 3) Kein Treffer -> leer lassen (Zelle bleibt leer)
  return null;
}

// Erste leere Datenzeile (ab rowStart) finden – über die Template-Spalten
function findFirstEmptyRow(sheet, rowStart, colCount) {
  let r = rowStart;
  while (true) {
    const row = sheet.getRow(r);
    let hasAny = false;
    for (let c = 1; c <= colCount; c++) {
      const v = row.getCell(c).value;
      if (v !== null && v !== undefined && String(v) !== "") { hasAny = true; break; }
    }
    if (!hasAny) return r;
    r++;
    if (r > (sheet.lastRow?.number || rowStart) + 10000) return r; // Sicherheitsabbruch
  }
}

export default async function handler(req, res) {
  try {
    if (req.method !== "POST")
      return res.status(405).json({ error: "Method not allowed" });

    const {
      templateBase64,
      rows = [],
      qa = {},
      prefix = "SITE"
    } = req.body || {};

    if (!templateBase64)
      return res.status(400).json({ error: "templateBase64 fehlt" });
    if (!Array.isArray(rows) || rows.length === 0)
      return res.status(400).json({ error: "rows fehlt/leer" });

    // 1) Template laden (Template bleibt 1:1 erhalten)
    const wb = new ExcelJS.Workbook();
    const buf = Buffer.from(
      String(templateBase64).includes("base64,")
        ? String(templateBase64).split("base64,").pop()
        : String(templateBase64),
      "base64"
    );
    await wb.xlsx.load(buf);

    const sheet = wb.worksheets[0]; // erstes Sheet
    const headerRow = sheet.getRow(1);     // Row 1 = Überschriften
    const legendRow = sheet.getRow(2);     // Row 2 = Legenden (bleibt unberührt)

    // 2) Template-Header exakt lesen (Spaltennamen aus Row 1)
    const headerNames = [];
    for (let c = 1; c <= headerRow.cellCount; c++) {
      const val = headerRow.getCell(c).value;
      if (!val) break;
      headerNames.push(String(val)); // z. B. "Q2.Gender", "Q3.Altersgruppe …"
    }
    const colCount = headerNames.length;

    // 3) Mapping "Template-Header" -> "rows-Key" bauen
    //    (Direktmatch, Normalisierung, Heuristiken)
    const sample = rows[0] || {};
    const rowKeys = Object.keys(sample);
    const rowKeysNormMap = new Map(rowKeys.map(k => [normalize(k), k]));

    const headerToRowKey = new Map();
    for (const h of headerNames) {
      // 3a) Exakter Key im Row-Objekt vorhanden?
      if (Object.prototype.hasOwnProperty.call(sample, h)) {
        headerToRowKey.set(h, h);
        continue;
      }
      // 3b) Normalisierte Übereinstimmung?
      const hNorm = normalize(h);
      if (rowKeysNormMap.has(hNorm)) {
        headerToRowKey.set(h, rowKeysNormMap.get(hNorm));
        continue;
      }
      // 3c) Heuristiken (Q2/Q3/Q7/Q9/Q10/Q11 + deutschsprachige Header)
      const guessed = guessRowKeyForHeader(h, rowKeysNormMap);
      headerToRowKey.set(h, guessed);
    }

    // Debug-Preview aktivieren (optional): ?preview=1
    if (String(req.query?.preview ?? "") === "1") {
      const mappingPreview = {};
      for (const h of headerNames) mappingPreview[h] = headerToRowKey.get(h) || null;
      return res.status(200).json({
        preview: {
          headers: headerNames,
          mapping: mappingPreview,
          firstDataRow: 3,
          rowsKeys: rowKeys
        }
      });
    }

    // 4) Erste leere Zeile ab Row 3 finden und Daten einfügen
    const START_ROW = findFirstEmptyRow(sheet, 3, colCount);

    let r = START_ROW;
    for (const persona of rows) {
      const xRow = sheet.getRow(r);

      for (let c = 1; c <= colCount; c++) {
        const header = headerNames[c - 1];
        const rowKey = headerToRowKey.get(header); // kann null sein
        const value = rowKey ? (persona[rowKey] ?? "") : "";
        xRow.getCell(c).value = value;
      }

      xRow.commit();
      r++;
    }

    // 5) QA auf separates Sheet (optional)
    if (qa && typeof qa === "object" && Object.keys(qa).length) {
      const qaSheet = wb.addWorksheet("QA");
      let i = 1;
      for (const [k, v] of Object.entries(qa)) {
        qaSheet.getCell(i, 1).value = k;
        qaSheet.getCell(i, 2).value =
          typeof v === "object" ? JSON.stringify(v) : String(v ?? "");
        i++;
      }
    }

    // 6) Datei zurückgeben (Template bleibt unverändert, nur Daten gefüllt)
    const out = await wb.xlsx.writeBuffer();
    const outB64 = Buffer.from(out).toString("base64");
    return res.status(200).json({
      file: outB64,
      fileName: `${prefix}_Festdaten.xlsx`,
      success: true
    });

  } catch (err) {
    console.error("EXPORT XLSX ERROR", err);
    return res.status(500).json({ error: err.message });
  }
}
