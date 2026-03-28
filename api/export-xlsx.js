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
      fallzahl = null,
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

    /************************************************************
     * 1️⃣ TEMPLATE EINLESEN
     ************************************************************/
    const buffer = Buffer.from(templateBase64, "base64");

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    // Sheet 1 ist das Template
    const ws = wb.worksheets[0];

    /************************************************************
     * 2️⃣ DATEN INS TEMPLATE SCHREIBEN
     *    (KEINE Formatierung, KEIN Styling, KEIN Reset)
     ************************************************************/
    // Wir gehen davon aus, dass DEIN TEMPLATE bereits
    // die Header stehen hat und Row-Start klar ist.

    let startRow = 2; // Beispiel: nach Header
    rows.forEach((row, idx) => {
      const excelRow = ws.getRow(startRow + idx);

      column_order.forEach((key, colIndex) => {
        excelRow.getCell(colIndex + 1).value = row[key] ?? "";
      });

      excelRow.commit();
    });

    /************************************************************
     * 3️⃣ KEIN QA-SHEET! — das macht n8n ab jetzt.
     ************************************************************/

    /************************************************************
     * 4️⃣ RETURN ALS BASE64
     ************************************************************/
    const out = await wb.xlsx.writeBuffer();

    return res.status(200).json({
      file: Buffer.from(out).toString("base64"),
      fileName: `${prefix}_Festdaten.xlsx`,
      success: true,
      version: "template-mode-clean",
    });

  } catch (err) {
    console.error("EXPORT ERROR", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
