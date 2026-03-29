import ExcelJS from "exceljs";

export const config = {
  api: { bodyParser: { sizeLimit: "50mb" } },
};

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    const raw = req.body;
    const body = typeof raw === "string" ? JSON.parse(raw) : raw;

    const {
      templateBase64 = null,
      rows = [],
      column_order = [],
      prefix = "SITE",
    } = body;

    if (!templateBase64) {
      return res.status(400).json({ error: "templateBase64 missing" });
    }
    if (!Array.isArray(rows)) {
      return res.status(400).json({ error: "rows must be array" });
    }
    if (!Array.isArray(column_order) || !column_order.length) {
      return res.status(400).json({ error: "column_order missing" });
    }

    /***********************************************
     * TEMPLATE LADEN
     ***********************************************/
    const buffer = Buffer.from(templateBase64, "base64");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    const ws = wb.worksheets[0];

    /***********************************************
     * HEADER-SPALTEN AUS TEMPLATE AUSLESEN
     ***********************************************/
    const headerRow = ws.getRow(1);
    const colMap = {};  // key → Spaltenindex

    headerRow.eachCell((cell, colNumber) => {
      const header = String(cell.value).trim();
      colMap[header] = colNumber;
    });

    /***********************************************
     * STARTROW AUTOMATISCH FINDEN (erste leere Zeile nach Header)
     ***********************************************/
    let startRow = 2;

    /***********************************************
     * DATEN EINTRAGEN (FORMATIERUNG BLEIBT!)
     ***********************************************/
    rows.forEach((rowData, i) => {
      const excelRow = ws.getRow(startRow + i);

      // "Anzahl" — sofern im Template
      if (colMap["Anzahl"]) {
        excelRow.getCell(colMap["Anzahl"]).value = String(i + 1);
      }

      // Jede Variable in die vom Template vorgesehenen Spalten
      column_order.forEach((key) => {
        const col = colMap[key];
        if (!col) return; // Template hat die Spalte nicht → überspringen
        excelRow.getCell(col).value = rowData[key] ?? "";
      });

      // "No." falls im Template
      if (colMap["No."]) {
        excelRow.getCell(colMap["No."]).value = String(i + 1);
      }
    });

    /***********************************************
     * EXPORT
     ***********************************************/
    const out = await wb.xlsx.writeBuffer();

    return res.status(200).json({
      file: Buffer.from(out).toString("base64"),
      fileName: `${prefix}_Festdaten.xlsx`,
      success: true,
      version: "template-perfect",
    });

  } catch (err) {
    console.error("EXPORT ERROR", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
