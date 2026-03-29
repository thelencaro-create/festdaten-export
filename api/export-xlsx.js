import ExcelJS from "exceljs";

export const config = {
  api: { bodyParser: { sizeLimit: "50mb" } },
};

export default async function handler(req, res) {
  try {
    const {
      templateBase64,
      rows,
      column_order,
      prefix = "SITE",
    } = typeof req.body === "string" ? JSON.parse(req.body) : req.body;

    if (!templateBase64) throw new Error("templateBase64 missing");
    if (!rows || !Array.isArray(rows)) throw new Error("rows missing");
    if (!column_order || !column_order.length) throw new Error("column_order missing");

    // Workbook laden
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(Buffer.from(templateBase64, "base64"));
    const ws = wb.worksheets[0];

    /****************************************************
     * 1) Erste komplett leere Zeile finden
     *    (unterhalb des vertikalen Block A)
     ****************************************************/
    let insertHeaderAt = null;

    ws.eachRow((row, rowNr) => {
      const filled = row.values.filter(v => v !== null && v !== "").length;
      if (filled === 0 && !insertHeaderAt) {
        insertHeaderAt = rowNr;
      }
    });

    if (!insertHeaderAt) {
      insertHeaderAt = ws.rowCount + 1;
    }

    /****************************************************
     * 2) Horizontale Headerzeile erzeugen
     ****************************************************/
    const headerRow = ws.getRow(insertHeaderAt);

    const headerValues = [
      "Anzahl",
      ...column_order,
      "No."
    ];

    headerValues.forEach((val, idx) => {
      headerRow.getCell(idx + 1).value = val;
    });

    headerRow.commit();

    /****************************************************
     * 3) Daten einfügen
     ****************************************************/
    rows.forEach((src, i) => {
      const excelRow = ws.getRow(insertHeaderAt + 1 + i);

      excelRow.getCell(1).value = i + 1; // Anzahl

      column_order.forEach((key, idx) => {
        excelRow.getCell(idx + 2).value = src[key] ?? "";
      });

      excelRow.getCell(column_order.length + 2).value = i + 1; // No.
    });

    const out = await wb.xlsx.writeBuffer();

    return res.status(200).json({
      file: Buffer.from(out).toString("base64"),
      fileName: `${prefix}_Festdaten.xlsx`,
      success: true,
    });

  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
