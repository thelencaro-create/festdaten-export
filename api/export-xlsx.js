import ExcelJS from "exceljs";

export const config = {
  api: { bodyParser: { sizeLimit: "50mb" } },
};

export default async function handler(req, res) {
  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;

    const { templateBase64, headerOrder, rows, prefix = "SITE" } = body;

    if (!templateBase64) throw new Error("templateBase64 missing");
    if (!Array.isArray(headerOrder)) throw new Error("headerOrder missing");
    if (!Array.isArray(rows)) throw new Error("rows missing");

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(Buffer.from(templateBase64, "base64"));
    const ws = wb.worksheets[0];

    /***********************************************
     * 1) MATRIX-START ERKENNEN (universal)
     ***********************************************/
    let headerRowIndex = null;
    let dataStartIndex = null;

    ws.eachRow((row, rowIdx) => {
      const values = row.values;

      const isHeaderCandidate =
        values.filter((v) => typeof v === "string" && v.trim() !== "").length >= 3;

      const isDataCandidate =
        typeof values[1] === "number" ||
        (typeof values[1] === "string" && /^\d+$/.test(values[1]));

      if (isHeaderCandidate && !headerRowIndex) {
        headerRowIndex = rowIdx;
      }

      if (isDataCandidate && !dataStartIndex) {
        dataStartIndex = rowIdx;
      }
    });

    const matrixIsPresent = headerRowIndex && dataStartIndex;

    /***********************************************
     * 2) MATRIX-HEADER SETZEN (falls nötig)
     ***********************************************/
    let writeHeaderAt = null;

    if (matrixIsPresent) {
      writeHeaderAt = headerRowIndex;
    } else {
      // Typ A Templates (isi, Mischfett)
      // wir suchen erste komplett "technisch leere" Zeile (keine Persona)
      let firstEmpty = null;
      ws.eachRow((row, idx) => {
        const filled = row.values.filter((v) => v !== null && v !== "").length;
        if (filled === 0 && !firstEmpty) firstEmpty = idx;
      });

      writeHeaderAt = firstEmpty || ws.rowCount + 1;

      // Header schreiben
      const headerRow = ws.getRow(writeHeaderAt);
      const header = ["Anzahl", ...headerOrder, "No."];

      header.forEach((v, i) => {
        headerRow.getCell(i + 1).value = v;
      });

      // style
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF005BBB" },
        };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });

      headerRow.commit();

      dataStartIndex = writeHeaderAt + 1;
    }

    /***********************************************
     * 3) ROWS schreiben (universell)
     ***********************************************/
    rows.forEach((src, i) => {
      const excelRow = ws.getRow(dataStartIndex + i);
      excelRow.getCell(1).value = i + 1;

      headerOrder.forEach((key, idx) => {
        excelRow.getCell(idx + 2).value = src[key] ?? "";
      });

      excelRow.getCell(headerOrder.length + 2).value = i + 1;
      excelRow.commit();
    });

    /***********************************************
     * 4) Datei zurückgeben
     ***********************************************/
    const out = await wb.xlsx.writeBuffer();

    return res.status(200).json({
      file: Buffer.from(out).toString("base64"),
      fileName: `${prefix}_Festdaten.xlsx`,
      success: true,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
