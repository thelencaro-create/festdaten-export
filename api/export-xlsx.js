import ExcelJS from "exceljs";

export const config = {
  api: {
    bodyParser: { sizeLimit: "15mb" }
  }
};

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    const {
      templateBase64,
      rows = [],
      prefix = "SITE"
    } = req.body || {};

    if (!templateBase64) {
      return res.status(400).json({ error: "templateBase64 fehlt" });
    }

    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ error: "rows fehlt/leer" });
    }

    // --- 1) Template laden ---
    const workbook = new ExcelJS.Workbook();
    const templateBuffer = Buffer.from(
      templateBase64.includes("base64,")
        ? templateBase64.split("base64,").pop()
        : templateBase64,
      "base64"
    );
    await workbook.xlsx.load(templateBuffer);

    const sheet = workbook.worksheets[0];
    if (!sheet) {
      return res.status(400).json({ error: "Template enthält kein Worksheet" });
    }

    // --- 2) Header aus Zeile 1 auslesen (genaue Template-Header) ---
    const headerRow = sheet.getRow(1);
    const headerNames = [];
    for (let col = 1; col <= headerRow.cellCount; col++) {
      const head = headerRow.getCell(col).value;
      if (!head) break;
      headerNames.push(String(head));
    }

    const colCount = headerNames.length;

    // --- 3) Daten ab Zeile 3 schreiben (Zeile 1 & 2 werden 1:1 aus Template gelassen!) ---
    let rowIndex = 3;

    for (const persona of rows) {
      const excelRow = sheet.getRow(rowIndex);

      for (let c = 1; c <= colCount; c++) {
        const header = headerNames[c - 1];

        // rows müssen EXAKT dieselben Keys haben wie die headerNames
        const value = persona[header] ?? "";

        // WICHTIG:
        // nur .value setzen → ALLE Formatierungen bleiben erhalten.
        excelRow.getCell(c).value = value;
      }

      excelRow.commit();
      rowIndex++;
    }
// 4) ===== OPTIONAL: QA-SHEET ERZEUGEN =====
if (qa && typeof qa === "object" && Object.keys(qa).length) {
  const qaSheet = wb.addWorksheet("QA");

  let row = 1;

  // Funktion zum rekursiven Schreiben (unterstützt verschachtelte Objekte)
  function writeQA(prefix, obj) {
    for (const [key, value] of Object.entries(obj)) {
      if (value !== null && typeof value === "object" && !Array.isArray(value)) {
        writeQA(`${prefix}${key}.`, value);
      } else {
        qaSheet.getCell(row, 1).value = `${prefix}${key}`;
        qaSheet.getCell(row, 2).value =
          typeof value === "string" || typeof value === "number"
            ? value
            : JSON.stringify(value);
        row++;
      }
    }
  }

  writeQA("", qa);
}
    // --- 5) Datei zurückgeben ---
    const outputBuffer = await workbook.xlsx.writeBuffer();
    const base64 = Buffer.from(outputBuffer).toString("base64");

    return res.status(200).json({
      file: base64,
      fileName: `${prefix}_Festdaten.xlsx`,
      success: true
    });

  } catch (err) {
    console.error("EXPORT XLSX ERROR", err);
    return res.status(500).json({ error: err.message });
  }
}
