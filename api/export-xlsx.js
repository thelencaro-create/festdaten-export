import ExcelJS from "exceljs";

export const config = {
  api: {
    bodyParser: { sizeLimit: "15mb" } // große Templates
  }
};

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    // --- Body robust einlesen (RAW oder bereits geparst) ---
    const raw = req.body;
    const body = (typeof raw === "string")
      ? JSON.parse(raw)
      : (raw || {});

    const templateBase64 = body.templateBase64;
    const rows = Array.isArray(body.rows) ? body.rows : [];
    const prefix = body.prefix || "SITE";
    const qa = (body.qa && typeof body.qa === "object" && !Array.isArray(body.qa))
      ? body.qa
      : null;

    if (!templateBase64) return res.status(400).json({ error: "templateBase64 fehlt" });
    if (!rows.length)     return res.status(400).json({ error: "rows fehlt/leer" });

    // --- Template laden ---
    const wb = new ExcelJS.Workbook();
    const tplBuf = Buffer.from(
      String(templateBase64).includes("base64,")
        ? String(templateBase64).split("base64,").pop()
        : String(templateBase64),
      "base64"
    );
    await wb.xlsx.load(tplBuf);

    const sheet = wb.worksheets[0];
    if (!sheet) return res.status(400).json({ error: "Kein Worksheet im Template gefunden." });

    // --- Header aus Zeile 1 lesen (Spaltennamen exakt) ---
    const headerRow = sheet.getRow(1);
    const headers = [];
    for (let c = 1; c <= headerRow.cellCount; c++) {
      const v = headerRow.getCell(c).value;
      if (!v) break;
      headers.push(String(v));
    }
    const colCount = headers.length;
    if (!colCount) return res.status(400).json({ error: "Headerzeile (Row 1) leer." });

    // --- Daten ab Zeile 3 schreiben (Formatierung bleibt erhalten) ---
    let r = 3;
    for (let i = 0; i < rows.length; i++) {
      const persona = rows[i] || {};
      const xRow = sheet.getRow(r);

      for (let c = 1; c <= colCount; c++) {
        const header = headers[c - 1];

        // Auto-Nummerierung (falls vorhanden)
        if (header === "Anzahl" || header === "No." || header === "Nr" || header === "Nr.") {
          xRow.getCell(c).value = i + 1;
          continue;
        }

        const val = persona.hasOwnProperty(header) ? persona[header] : "";
        xRow.getCell(c).value = (val == null ? "" : val);
      }
      xRow.commit();
      r++;
    }

    // --- QA-Sheet nur erzeugen, wenn qa ein Objekt ist ---
    if (qa) {
      const qaSheet = wb.addWorksheet("QA");
      let row = 1;

      // rekursiv flatten
      function writeQA(prefixKey, obj) {
        for (const [k, v] of Object.entries(obj)) {
          const key = prefixKey ? `${prefixKey}.${k}` : k;
          if (v && typeof v === "object" && !Array.isArray(v)) {
            writeQA(key, v);
          } else {
            qaSheet.getCell(row, 1).value = key;
            qaSheet.getCell(row, 2).value =
              (typeof v === "string" || typeof v === "number")
                ? v
                : JSON.stringify(v);
            row++;
          }
        }
      }
      writeQA("", qa);
    }

    const outBuf = await wb.xlsx.writeBuffer();
    const outB64 = Buffer.from(outBuf).toString("base64");
    return res.status(200).json({
      file: outB64,
      fileName: `${prefix}_Festdaten.xlsx`,
      success: true
    });

  } catch (err) {
    console.error("EXPORT XLSX ERROR", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
