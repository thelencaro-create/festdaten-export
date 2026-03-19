import ExcelJS from "exceljs";

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "15mb"
    }
  }
};

// -------------------------
// Utilities
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
  return n === "anzahl" || n === "no" || n === "nr" || n.startsWith("nr_");
}
function buildRowKeyIndex(rowKeys) {
  const byNorm = new Map();
  for (const k of rowKeys) byNorm.set(normalize(k), k);
  return byNorm;
}

// Kandidatenliste für Header → rows-Key (ohne feste Reihenfolge im Template)
function candidateKeysFor(headerNorm, rowKeys) {
  // Basis: alle rows-Keys, normalisiert
  const rk = rowKeys.map(k => ({ raw: k, norm: normalize(k) }));

  // Hilfsfilter
  const has = p => rk.filter(x => x.norm.includes(p));
  const starts = p => rk.filter(x => x.norm.startsWith(p));

  // Q10-Speziallogik (Marke / Andere)
  if (headerNorm.includes("q10")) {
    // „Andere“ bevorzugen, wenn Header Begriffe enthält; sonst „Marke“
    if (/(andere|andern|other)/.test(headerNorm)) {
      const andere = rk.filter(x => /(q10.*andere$|_andere$)/.test(x.norm));
      const rest = has("q10").filter(x => !/_andere$/.test(x.norm));
      return [...andere, ...rest];
    }
    // Generisch: erst Marke, dann Andere, dann sonstige q10
    const q10Marke = rk.filter(x => /^q10(_marke)?$/.test(x.norm));
    const q10Andere = rk.filter(x => /(q10.*andere$|_andere$)/.test(x.norm));
    const q10Rest = has("q10").filter(x => !q10Marke.includes(x) && !q10Andere.includes(x));
    return [...q10Marke, ...q10Andere, ...q10Rest];
  }

  // Q11 Fett
  if (headerNorm.includes("q11") || headerNorm.includes("fett")) {
    const q11 = has("q11");
    const fett = rk.filter(x => x.norm.includes("fett"));
    return [...q11, ...fett];
  }

  // Q9 Ablehnung
  if (headerNorm.includes("q9") || headerNorm.includes("ablehn") || headerNorm.includes("keine")) {
    return has("q9");
  }

  // Q7 Verbrauch
  if (headerNorm.includes("q7")) {
    // feiner nach Frischkäse / Gouda / Butterkäse / Camembert
    const prio = ["frisch", "frischkaese", "gouda", "butter", "butterkaese", "camembert"];
    const out = [];
    for (const p of prio) out.push(...rk.filter(x => x.norm.includes("q7") && x.norm.includes(p)));
    const rest = rk.filter(x => x.norm.includes("q7") && !out.includes(x));
    return [...out, ...rest];
  }

  // Q2/Q3 (Gender/Age)
  if (headerNorm.includes("q2") || headerNorm.includes("gender") || headerNorm.includes("geschlecht")) {
    return starts("q2");
  }
  if (headerNorm.includes("q3") || headerNorm.includes("alter")) {
    return starts("q3");
  }

  // Fallback: exakte Norm-Übereinstimmung
  return [];
}

// Mapping je Auftreten (Duplikate) robust berechnen
function buildHeaderToRowKeyMapping(headerNames, sampleRow) {
  const rowKeys = Object.keys(sampleRow || {});
  const byNorm = buildRowKeyIndex(rowKeys);
  const taken = new Set();                  // bereits vergebene rows-Keys
  const counter = Object.create(null);      // Auftretenszähler je Header-Text
  const map = new Map();

  for (const h of headerNames) {
    // Zählspalten automatisch füllen
    if (isCountHeader(h)) { map.set(h, "__AUTO_NUMBER__"); continue; }

    counter[h] = (counter[h] || 0) + 1;
    const norm = normalize(h);

    // 1) Exakter Key vorhanden?
    if (Object.prototype.hasOwnProperty.call(sampleRow, h) && !taken.has(h)) {
      map.set(h, h); taken.add(h); continue;
    }

    // 2) Normalisierte Übereinstimmung?
    const direct = byNorm.get(norm);
    if (direct && !taken.has(direct)) { map.set(h, direct); taken.add(direct); continue; }

    // 3) Kandidatenliste (prio-gestützt je Header)
    const cand = candidateKeysFor(norm, rowKeys)
      .map(x => x.raw)
      .filter(k => !taken.has(k));

    // Bei Duplikaten (z. B. zweites „Q10: …“) den noch nicht vergebenen Kandidaten wählen
    const pick = cand[0] || null;
    map.set(h, pick);
    if (pick) taken.add(pick);
  }

  return map;
}

// Erste komplett leere Zeile ab rowStart (bleibt am Template unkritisch)
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
    if (r > (sheet.lastRow?.number || rowStart) + 20000) return r;
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

    const sheet = wb.worksheets[0];
    const headerRow = sheet.getRow(1);
    const legendRow = sheet.getRow(2); // wird nicht verändert (siehe Template) [1](https://ftteststudio-my.sharepoint.com/personal/thelen_carolin_ftstudio_de/_layouts/15/Doc.aspx?sourcedoc=%7BA9CD4E10-78BA-4992-8E44-E0F101970CD3%7D&file=Screener%20Daten%20Cream%20Cheese%20DE%20HH.xlsx&action=default&mobileredirect=true)

    // 2) Template-Header aus Zeile 1 lesen
    const headerNames = [];
    for (let c = 1; c <= headerRow.cellCount; c++) {
      const v = headerRow.getCell(c).value;
      if (!v) break;
      headerNames.push(String(v));
    }
    const colCount = headerNames.length;

    // 3) Mapping Header -> rows-Key (Duplikate robust behandeln)
    const sample = rows[0] || {};
    const headerToRowKey = buildHeaderToRowKeyMapping(headerNames, sample);

    // Preview-Modus (Body bleibt unverändert; nützlich zum Testen)
    if (String(req.query?.preview ?? "") === "1") {
      const mapping = {};
      for (const h of headerNames) mapping[h] = headerToRowKey.get(h) || null;
      return res.status(200).json({
        preview: {
          headers: headerNames,
          mapping,
          firstDataRow: 3,
          autoNumberedHeaders: headerNames.filter(isCountHeader)
        }
      });
    }

    // 4) Erste leere Zeile ab 3 suchen, Daten eintragen
    const START_ROW = findFirstEmptyRow(sheet, 3, colCount); // Zeile 1+2 sind Header/Legenden [1](https://ftteststudio-my.sharepoint.com/personal/thelen_carolin_ftstudio_de/_layouts/15/Doc.aspx?sourcedoc=%7BA9CD4E10-78BA-4992-8E44-E0F101970CD3%7D&file=Screener%20Daten%20Cream%20Cheese%20DE%20HH.xlsx&action=default&mobileredirect=true)

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

    // 5) QA optional als separates Sheet
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
