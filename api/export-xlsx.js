import ExcelJS from "exceljs";

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "15mb"
    }
  }
};

// -------------------------
// Helper: Normalisierung
// -------------------------
function normalize(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

// Erkenne Zähl-/Index-Spalten (z. B. "Anzahl", "No.", "Nr")
function isCountHeader(h) {
  const n = normalize(h);
  return n === "anzahl" || n === "no" || n === "nr" || n.startsWith("nr_");
}

// Feste Overrides: Template-Header → rows-Key
// (inkl. Disambiguierung für doppelte Q10-Header)
function headerOverrides() {
  return new Map([
    ["Q2.Gender", "Q2_Gender"],
    ["Q3.Altersgruppe: Welcher Altersgruppe gehören Sie an?", "Q3_Age"],
    ["Q7_Essverhalten: Wie oft verwenden Sie Frischkäse?", "Q7_Essverhalten_Frischkaese"],
    ["Q7_Essverhalten: Wie oft verwenden Sie Gouda?", "Q7_Essverhalten_Gouda"],
    ["Q7_Essverhalten: Wie oft verwenden Sie Butterkäse?", "Q7_Essverhalten_Butterkaese"],
    ["Q7_Essverhalten: Wie oft verwenden Sie Camembert?", "Q7_Essverhalten_Camembert"],
    ["Q9_Geschmacksrichtungen: Lehne Sie Geschmacksrichtungen grundsätzlich ab?", "Q9_Ablehnung"],
    // Q10 kommt im Template zweimal mit gleichem Headertext:
    ["Q10: Welches Produkt verwenden Sie hauptsächlich?", "Q10_Marke"],                // 1. Auftreten
    ["Q10: Welches Produkt verwenden Sie hauptsächlich?__SECOND__", "Q10_Marke_Andere"], // 2. Auftreten
    ["Q11_Fettgehalt", "Q11_Fettgehalt"]
  ]);
}

// Fallback-Heuristik, wenn Overrides / Direct-Match nicht greifen
function guessRowKeyForHeader(header, rowKeysNormMap) {
  const h = normalize(header);

  // 1) Normalized direct hit
  if (rowKeysNormMap.has(h)) return rowKeysNormMap.get(h);

  // 2) Heuristiken für Q-Codes
  const tryStarts = (p) => {
    for (const k of rowKeysNormMap.keys()) if (k.startsWith(p)) return rowKeysNormMap.get(k);
    return null;
  };
  if (h.includes("q2") || h.includes("gender") || h.includes("geschlecht")) {
    const r = tryStarts("q2"); if (r) return r;
  }
  if (h.includes("q3") || h.includes("alter")) {
    const r = tryStarts("q3"); if (r) return r;
  }
  if (h.includes("q7") && (h.includes("frisch") || h.includes("frischkaese"))) {
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
    for (const k of rowKeysNormMap.keys())
      if (k.includes("q10") && (k.includes("andere") || k.includes("andern") || k.includes("other") || k.endsWith("_andere")))
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

  return null;
}

// Erste komplett leere Zeile ab rowStart (für den Start der Festdaten)
function findFirstEmptyRow(sheet, rowStart, colCount) {
  let r = rowStart;
  while (true) {
    const row = sheet.getRow(r);
    let any = false;
    for (let c = 1; c <= colCount; c++) {
      const v = row.getCell(c).value;
      if (v !== null && v !== undefined && String(v) !== "") { any = true; break; }
    }
    if (!any) return r;
    r++;
    if (r > (sheet.lastRow?.number || rowStart) + 20000) return r; // Sicherheitsabbruch
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

    // 1) Template laden
    const wb = new ExcelJS.Workbook();
    const buf = Buffer.from(
      String(templateBase64).includes("base64,")
        ? String(templateBase64).split("base64,").pop()
        : String(templateBase64),
      "base64"
    );
    await wb.xlsx.load(buf);

    const sheet = wb.worksheets[0];

    // 2) Header aus Zeile 1 lesen (reine Template-Spaltenbezeichnungen)
    const headerRow = sheet.getRow(1);
    const headerNames = [];
    for (let c = 1; c <= headerRow.cellCount; c++) {
      const v = headerRow.getCell(c).value;
      if (!v) break;
      headerNames.push(String(v));
    }
    const colCount = headerNames.length;

    // 3) Mapping "Template-Header" -> "rows-Key" aufbauen
    const sample = rows[0] || {};
    const rowKeys = Object.keys(sample);
    const rowKeysNormMap = new Map(rowKeys.map(k => [normalize(k), k]));

    const overrides = headerOverrides();
    const headerToRowKey = new Map();

    // Doppelte Header disambiguieren (z. B. Q10-Header 1/2)
    const seenCount = Object.create(null);

    for (const h of headerNames) {
      // Zählspalten nicht mappen; werden später automatisch gefüllt
      if (isCountHeader(h)) {
        headerToRowKey.set(h, "__AUTO_NUMBER__");
        continue;
      }

      seenCount[h] = (seenCount[h] || 0) + 1;
      const keyForOverride = (seenCount[h] === 2) ? `${h}__SECOND__` : h;

      // 1) Fester Override
      if (overrides.has(keyForOverride)) {
        headerToRowKey.set(h, overrides.get(keyForOverride));
        continue;
      }

      // 2) Exakter Key vorhanden?
      if (Object.prototype.hasOwnProperty.call(sample, h)) {
        headerToRowKey.set(h, h);
        continue;
      }

      // 3) Normalisierung / Heuristik
      const guessed = guessRowKeyForHeader(h, rowKeysNormMap);
      headerToRowKey.set(h, guessed); // kann null sein -> Zelle bleibt leer
    }

    // Optionaler Preview (nur URL: ?preview=1; Body unverändert)
    if (String(req.query?.preview ?? "") === "1") {
      const mappingPreview = {};
      for (const h of headerNames) mappingPreview[h] = headerToRowKey.get(h) || null;
      return res.status(200).json({
        preview: {
          headers: headerNames,
          mapping: mappingPreview,
          firstDataRow: 3,
          autoNumberedHeaders: headerNames.filter(isCountHeader)
        }
      });
    }

    // 4) Erste freie Zeile ab 3 finden & schreiben
    const START_ROW = findFirstEmptyRow(sheet, 3, colCount);

    let r = START_ROW;
    for (let i = 0; i < rows.length; i++) {
      const persona = rows[i];
      const xRow = sheet.getRow(r);

      for (let c = 1; c <= colCount; c++) {
        const header = headerNames[c - 1];
        const rowKey = headerToRowKey.get(header);

        if (rowKey === "__AUTO_NUMBER__") {
          xRow.getCell(c).value = i + 1; // 1..N
          continue;
        }

        const value = rowKey ? (persona[rowKey] ?? "") : "";
        xRow.getCell(c).value = value;
      }

      xRow.commit();
      r++;
    }

    // 5) QA optional auf eigenes Sheet
    if (qa && typeof qa === "object" && Object.keys(qa).length) {
      const qaSheet = wb.addWorksheet("QA");
      let rr = 1;
      for (const [k, v] of Object.entries(qa)) {
        qaSheet.getCell(rr, 1).value = k;
        qaSheet.getCell(rr, 2).value =
          typeof v === "object" ? JSON.stringify(v) : String(v ?? "");
        rr++;
      }
    }

    // 6) Datei zurückgeben
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
