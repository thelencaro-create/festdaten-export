import ExcelJS from "exceljs";

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "15mb"
    }
  }
};

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

    // 1) Template einlesen
    const workbook = new ExcelJS.Workbook();
    const templateBuffer = Buffer.from(templateBase64, "base64");
    await workbook.xlsx.load(templateBuffer);

    const sheet = workbook.worksheets[0];

    // ---------------------------
    // 2) Header-Zeile (Row 1)
    // ---------------------------
    const headerRow = sheet.getRow(1);
    const headerValues = [];

    for (let c = 1; c <= headerRow.cellCount; c++) {
      const val = headerRow.getCell(c).value;
      if (!val) break;
      headerValues.push(String(val));  // exakte Spaltennamen aus Template
    }

    const colCount = headerValues.length;

    // ---------------------------
    // 3) Festdaten ab Zeile 3 einsetzen
    // ---------------------------
    let writeRowIndex = 3;

    for (let i = 0; i < rows.length; i++) {
      const persona = rows[i];
      const excelRow = sheet.getRow(writeRowIndex);

      for (let c = 1; c <= colCount; c++) {
        const colName = headerValues[c - 1];

        // im persona-Objekt exakt nach gleichem Key suchen
        const cellValue = persona[colName] ?? "";
        excelRow.getCell(c).value = cellValue;
      }

      excelRow.commit();
      writeRowIndex++;
    }

    // ---------------------------
    // 4) QA-Tab hinzufügen (optional)
    // ---------------------------
    if (qa && typeof qa === "object" && Object.keys(qa).length > 0) {
      const qaSheet = workbook.addWorksheet("QA");
      let r = 1;
      for (const [key, val] of Object.entries(qa)) {
        qaSheet.getCell(r, 1).value = key;
        qaSheet.getCell(r, 2).value =
          typeof val === "object" ? JSON.stringify(val) : String(val);
        r++;
      }
    }

    // ---------------------------
    // 5) Datei zurückgeben
    // ---------------------------
    const out = await workbook.xlsx.writeBuffer();
    const outBase64 = Buffer.from(out).toString("base64");

    return res.status(200).json({
      file: outBase64,
      fileName: `${prefix}_Festdaten.xlsx`
    });

  } catch (err) {
    console.error("EXPORT ERROR", err);
    return res.status(500).json({
      error: err.message,
      stack: err.stack
    });
  }
}
