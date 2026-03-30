import ExcelJS from "exceljs";

export const config = {
  api: { bodyParser: { sizeLimit: "50mb" } },
};

export default async function handler(req, res) {
  try {
    const body =
      typeof req.body === "string" ? JSON.parse(req.body) : req.body;

    const {
      templateBase64,
      headerOrder,
      rows,
      prefix = "SITE",
    } = body;

    if (!templateBase64) throw new Error("templateBase64 missing");
    if (!Array.isArray(headerOrder) || !headerOrder.length)
      throw new Error("headerOrder missing");
    if (!Array.isArray(rows))
      throw new Error("rows missing");

    // Workbook laden
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(Buffer.from(templateBase64, "base64"));
    const ws = wb.worksheets[0];

    /****************************************************
     * 1) Erste komplett leere Zeile nach vertikalem Block
     ****************************************************/
    let insertRow = null;

    ws.eachRow((row, rowNr) => {
      const filled = row.values.filter(
        (v) => v !== null && v !== ""
      ).length;
      if (filled === 0 && insertRow === null) {
        insertRow = rowNr;
      }
    });

    if (!insertRow) {
      insertRow = ws.rowCount + 1;
    }

    /****************************************************
     * 2) Headerzeile erzeugen
     ****************************************************/
    const header = ["Anzahl", ...headerOrder, "No."];
    const headerRow = ws.getRow(insertRow);

    header.forEach((val, idx) => {
      headerRow.getCell(idx + 1).value = val;
    });

    headerRow.commit();

    /****************************************************
     * 3) Daten einfügen
     ****************************************************/
    rows.forEach((src, i) => {
      const excelRow = ws.getRow(insertRow + 1 + i);

      // Anzahl
      excelRow.getCell(1).value = i + 1;

      // Daten-Spalten
      headerOrder.forEach((key, idx) => {
        excelRow.getCell(idx + 2).value = src[key] ?? "";
      });

      // No.-Spalte
      excelRow.getCell(headerOrder.length + 2).value = i + 1;
    });

    const out = await wb.xlsx.writeBuffer();

    return res.status(200).json({
      file: Buffer.from(out).toString("base64"),
      fileName: `${prefix}_Festdaten.xlsx`,
      success: true,
    });
  } catch (err) {
    return res.status(500).json({
      error: err.message,
    });
  }
}
