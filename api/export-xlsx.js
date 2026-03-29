import ExcelJS from "exceljs";

export const config = {
  api: { bodyParser: { sizeLimit: "50mb" } },
};

export default async function handler(req, res) {
  try {
    const raw = req.body;
    const body = typeof raw === "string" ? JSON.parse(raw) : raw;

    const {
      templateBase64,
      rows,
      column_order,
      prefix = "SITE",
    } = body;

    if (!templateBase64) throw new Error("templateBase64 missing");
    if (!Array.isArray(rows)) throw new Error("rows must be array");
    if (!Array.isArray(column_order) || !column_order.length)
      throw new Error("column_order missing");

    // TEMPLATE LADEN
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(Buffer.from(templateBase64, "base64"));
    const ws = wb.worksheets[0];

    /*****************************************************
     * 1) HEADER-ZEILE AUTOMATISCH FINDEN
     *    erste Zeile, die mindestens 50% der column_order
     *    horizontal enthält
     *****************************************************/
    let headerRowNr = null;

    ws.eachRow((row, rowNr) => {
      const values = row.values
        .map(v => (typeof v === "string" ? v.trim() : v))
        .filter(Boolean);

      const hits = column_order.filter(c => values.includes(c)).length;

      if (hits >= Math.ceil(column_order.length * 0.5)) {
        headerRowNr = rowNr;
      }
    });

    if (!headerRowNr) throw new Error("Horizontale Headerrow nicht gefunden");

    const headerRow = ws.getRow(headerRowNr);

    /*****************************************************
     * 2) SPALTEN-MAPPING BILDEN
     *****************************************************/
    const colMap = {};
    headerRow.eachCell((cell, colNr) => {
      const name = String(cell.value ?? "").trim();
      if (column_order.includes(name)) {
        colMap[name] = colNr;
      }
    });

    /*****************************************************
     * 3) ERSTE DATEN-ZEILE FINDEN
     *    erste Zeile nach Header, in der SPALTE A (col=1)
     *    eine Zahl steht.
     *****************************************************/
    let dataStart = null;

    for (let r = headerRowNr + 1; r < headerRowNr + 500; r++) {
      const v = ws.getRow(r).getCell(1).value;
      if (v !== null && !isNaN(Number(v))) {
        dataStart = r;
        break;
      }
    }

    if (!dataStart) throw new Error("Start der Datenzeilen nicht gefunden");

    /*****************************************************
     * 4) HORIZONTALE DATEN EINTRAGEN (FORMAT SAFE!)
     *****************************************************/
    rows.forEach((src, i) => {
      const targetRow = ws.getRow(dataStart + i);

      column_order.forEach(key => {
        const col = colMap[key];
        if (!col) return;
        targetRow.getCell(col).value = src[key] ?? "";
      });
    });

    // EXPORT
    const out = await wb.xlsx.writeBuffer();

    return res.status(200).json({
      file: Buffer.from(out).toString("base64"),
      fileName: `${prefix}_Festdaten.xlsx`,
      success: true,
      version: "universal-template-1.0",
    });

  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
