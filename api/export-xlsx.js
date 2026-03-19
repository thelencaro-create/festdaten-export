import ExcelJS from "exceljs";

export const config = {
  api: {
    bodyParser: { sizeLimit: "15mb" } // große Templates/Rows
  }
};

// -------------------------
// Utilities / Heuristiken
// -------------------------
function normalize(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

function isCountHeader(h) {
  const n = normalize(h);
  // deckt "Anzahl", "No.", "Nr" u.ä. ab
  return n === "anzahl" || n === "no" || n === "nr" || n.startsWith("nr_");
}

// Header → rows-Key (falls rows-Keys andere Schreibweise haben, z. B. Unterstriche)
function buildRowKeyIndex(rows) {
  const sample = Array.isArray(rows) && rows.length ? rows[0] : {};
  const keys = Object.keys(sample);
  const byNorm = new Map(keys.map(k => [normalize(k), k]));
  return { sample, byNorm };
}

// Disambiguierung für Q10:
//  - Header mit "keine/sondern/ander" ⇒ Q10_Marke_Andere (Freitext)
//  - Sonst ⇒ Q10_Marke (Codes)
function pickQ10Key(headerNorm, byNorm) {
  const has = (p) => headerNorm.includes(p);
  if (has("q10")) {
    const freitext = has("keine") || has("sondern") || has("ander");
    if (freitext) {
      // direkter Treffer bevorzugen
      if (byNorm.has("q10_marke_andere")) return byNorm.get("q10_marke_andere");
      // Fallback: häufige Varianten
      for (const k of ["q10_marke_andere", "q10_andere", "marke_andere"]) {
        if (byNorm.has(k)) return byNorm.get(k);
      }
      return "Q10_Marke_Andere"; // letzter Fallback
    }
    // Normaler Q10 Marken‑Code
    if (byNorm.has("q10_marke")) return byNorm.get("q10_marke");
    for (const k of ["q10", "marke", "q10_code"]) {
      if (byNorm.has(k)) return byNorm.get(k);
    }
    return "Q10_Marke";
  }
  return null;
}

function bestRowKeyForHeader(headerText, byNorm) {
  const h = headerText || "";
  const hn = normalize(h);

  // Zählspalte → spezielle Markierung
  if (isCountHeader(h)) return "__AUTO_NUMBER__";

  // Spezifische Regeln Q10
  const q10 = pickQ10Key(hn, byNorm);
  if (q10) return q10;

  // Direkte/Normalisierte Übereinstimmung
  if (byNorm.has(hn)) return byNorm.get(hn);

  // Einige typische Zuordnungen (Punkte ↔ Unterstriche, Kurztexte)
  const candidates = [
    hn,
    hn.replace(/\./g, "_"),
    hn.replace(/_/g, ".")
  ];
  for (const c of candidates) {
    if (byNorm.has(c)) return byNorm.get(c);
  }
  return null;
}

// -------------------------
// Handler
// -------------------------
export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    // Body entgegennehmen
    const {
      templateBase64,
      rows = [],
      columns,      // optional: falls du eine explizite Spaltenliste mitsendest
      qa = null,    // optional: QA, falls du später ein eigenes Sheet magst
      prefix = "SITE"
    } = req.body || {};

    if (!templateBase64) {
      return res.status(400).json({ error: "templateBase64 fehlt" });
    }
    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ error: "rows fehlt/leer" });
    }

    // Template laden (Formate bleiben erhalten, solange wir nur .value setzen)
    const wb = new ExcelJS.Workbook();
    const buf = Buffer.from(
      String(templateBase64).includes("base64,")
        ? String(templateBase64).split("base64,").pop()
        : String(templateBase64),
      "base64"
    );
    await wb.xlsx.load(buf);

    const sheet = wb.worksheets[0];
    if (!sheet) {
      return res.status(400).json({ error: "Kein Worksheet im Template gefunden." });
    }

    // Header (Zeile 1) vollständig lesen – NUR Namen, keine Stile anfassen
    const headerRow = sheet.getRow(1);
    const headerNames = [];
    for (let c = 1; c <= headerRow.cellCount; c++) {
      const v = headerRow.getCell(c).value;
      if (!v) break;
      headerNames.push(String(v));
    }
    const colCount = headerNames.length;
    if (!colCount) {
      return res.status(400).json({ error: "Headerzeile leer oder nicht gefunden (Row 1)." });
    }

    // rows-Keys indexieren (normalisiert)
    const { sample, byNorm } = buildRowKeyIndex(rows);

    // Mapping: Template-Header → rows-Key (oder "__AUTO_NUMBER__")
    const headerToKey = new Map();
    const seenCounts = Object.create(null);
    for (const h of headerNames) {
      seenCounts[h] = (seenCounts[h] || 0) + 1;

      // Falls columns explizit gesetzt wurden und exakt übereinstimmen
      if (Array.isArray(columns) && columns.includes(h)) {
        headerToKey.set(h, h);
        continue;
      }

      // Allgemeine Heuristik
      const key = bestRowKeyForHeader(h, byNorm);
      headerToKey.set(h, key); // kann null sein → Zelle bleibt leer
    }

    // Preview‑Modus (optional, wenn du ?preview=1 an die URL hängst)
    if (String(req.query?.preview ?? "") === "1") {
      const mapping = {};
      for (const h of headerNames) mapping[h] = headerToKey.get(h) || null;
      return res.status(200).json({
        preview: {
          headers: headerNames,
          mapping,
          firstDataRow: 3,
          rowsSampleKeys: Object.keys(sample)
        }
      });
    }

    // Daten ab Zeile 3 schreiben (Row 2 = Legenden bleibt unberührt)
    let r = 3;
    for (let i = 0; i < rows.length; i++) {
      const persona = rows[i] || {};
      const xRow = sheet.getRow(r);

      for (let c = 1; c <= colCount; c++) {
        const header = headerNames[c - 1];
        const key = headerToKey.get(header);

        if (key === "__AUTO_NUMBER__") {
          // 1..N automatisch
          xRow.getCell(c).value = i + 1;
          continue;
        }

        const val = key ? persona[key] : "";
        // WICHTIG: Nur .value setzen → Formate/Styles bleiben erhalten
        xRow.getCell(c).value = (val == null ? "" : val);
      }

      xRow.commit();
      r++;
    }

    // Optional: QA als eigenes Sheet anlegen
    if (qa && typeof qa === "object" && Object.keys(qa).length) {
      const qaSheet = wb.addWorksheet("QA");
      let row = 1;
      for (const [k, v] of Object.entries(qa)) {
        qaSheet.getCell(row, 1).value = k;
        qaSheet.getCell(row, 2).value =
          typeof v === "object" ? JSON.stringify(v) : String(v ?? "");
        row++;
      }
    }

    // Datei zurückgeben
    const out = await wb.xlsx.writeBuffer();
    const outB64 = Buffer.from(out).toString("base64");

    return res.status(200).json({
      file: outB64,
      fileName: `${prefix}_Festdaten.xlsx`,
      filename: `${prefix}_Festdaten.xlsx`,   // für bestehende n8n-Flows kompatibel
      success: true
    });

  } catch (err) {
    console.error("EXPORT XLSX ERROR", err);
    return res.status(500).json({ error: err.message });
  }
}
