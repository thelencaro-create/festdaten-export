// /api/export-xlsx.js  (Next.js API route)
// oder bei App Router: export default async function POST(req) { … }
import ExcelJS from "exceljs";

export const config = {
  api: { bodyParser: { sizeLimit: "50mb" } },
};

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    // ---------------------
    // 0) Eingabekörper lesen
    // ---------------------
    const raw = req.body;
    const body =
      typeof raw === "string" ? JSON.parse(raw) :
      raw && typeof raw === "object" ? raw :
      {};

    const {
      templateBase64,
      rows,
      qa,
      sheets,
      prefix = "SITE",
      headerOrder,
      dataStartRow = 4
    } = body;

    if (!templateBase64) {
      return res.status(400).json({ error: "templateBase64 fehlt" });
    }

    // ---------------------
    // 1) Template laden
    // ---------------------
    const wb = new ExcelJS.Workbook();
    const tplBuf = Buffer.from(
      templateBase64.includes("base64,")
        ? templateBase64.split("base64,").pop()
        : templateBase64,
      "base64"
    );
    await wb.xlsx.load(tplBuf);

    const tplSheet = wb.worksheets[0];
    if (!tplSheet) {
      return res.status(400).json({ error: "Kein Worksheet im Template" });
    }

    // ---------------------
    // 2) Header aus Row 1 lesen
    // ---------------------
    function readHeaderTexts(ws) {
      const headerRow = ws.getRow(1);
      const texts = [];
      for (let c = 1; c <= headerRow.cellCount; c++) {
        const v = headerRow.getCell(c).value;
        if (!v) break;
        texts.push(String(v));
      }
      return texts;
    }

    const templateHeaders = headerOrder && headerOrder.length
      ? headerOrder
      : readHeaderTexts(tplSheet);

    if (!Array.isArray(templateHeaders) || templateHeaders.length === 0) {
      return res.status(400).json({ error: "Keine Header im Template" });
    }

    // ---------------------
    // 3) Header→Key Mapping
    // ---------------------
    function makeHeaderToKeyMap(headers) {
      const map = new Map();
      const norm = (s) => String(s).toLowerCase();

      for (let i = 0; i < headers.length; i++) {
        const h = headers[i], H = norm(h);

        // Q2 Gender
        if (H.includes("q2") && (H.includes("gender") || H.includes("geschlecht")))
          { map.set(i+1, "Q2_Gender"); continue; }

        // Q3 Age
        if (H.includes("q3") && (H.includes("age") || H.includes("alter")))
          { map.set(i+1, "Q3_Age"); continue; }

        // Q7 Essverhalten
        if (H.includes("q7") && H.includes("frischk"))
          { map.set(i+1, "Q7_Essverhalten_Frischkaese"); continue; }
        if (H.includes("q7") && H.includes("gouda"))
          { map.set(i+1, "Q7_Essverhalten_Gouda"); continue; }
        if (H.includes("q7") && H.includes("butter"))
          { map.set(i+1, "Q7_Essverhalten_Butterkaese"); continue; }
        if (H.includes("q7") && H.includes("camembert"))
          { map.set(i+1, "Q7_Essverhalten_Camembert"); continue; }

        // Q9
        if (H.includes("q9"))
          { map.set(i+1, "Q9_Ablehnung"); continue; }

        // Q10 Marke & Freitext
        if (H.includes("q10") && (H.includes("produkt") || H.includes("verwenden")))
          { map.set(i+1, "Q10_Marke"); continue; }
        if (H.includes("q10") && (H.includes("keine") || H.includes("markenname") || H.includes("andere")))
          { map.set(i+1, "Q10_Marke_Andere"); continue; }

        // Q11 Fett
        if (H.includes("q11") && H.includes("fett"))
          { map.set(i+1, "Q11_Fettgehalt"); continue; }

        // Fallback: Header-namen direkt verwenden
        const fallback = h.replace(/\./g, "_").replace(/\s+/g, "_");
        map.set(i+1, fallback);
      }
      return map;
    }

    const h2k = makeHeaderToKeyMap(templateHeaders);

    // ---------------------
    // 4) Datenzeilen schreiben
    // ---------------------
    function writeDataRowsByHeader(ws, dataRows) {
      if (!Array.isArray(dataRows) || dataRows.length === 0) return;

      // genügend Zeilen erzeugen
      const need = Math.max(0, dataRows.length - (ws.rowCount - (dataStartRow - 1)));
      if (need > 0) ws.duplicateRow(dataStartRow, need, true);

      for (let i = 0; i < dataRows.length; i++) {
        const rowNum = dataStartRow + i;
        const targetRow = ws.getRow(rowNum);
        const src = dataRows[i];

        for (let c = 1; c <= templateHeaders.length; c++) {
          const key = h2k.get(c);
          const val = key && Object.prototype.hasOwnProperty.call(src, key)
            ? src[key]
            : "";
          targetRow.getCell(c).value = val == null ? "" : val;
        }
        targetRow.commit();
      }
    }

    // ---------------------
    // 5) Multi-Sheet Handling
    // ---------------------
    if (Array.isArray(sheets) && sheets.length > 0) {

      // Erstes Sheet → in Template schreiben
      writeDataRowsByHeader(tplSheet, sheets[0].rows || []);
      if (sheets[0].qa) buildQASheet(wb, sheets[0].qa);

      // Weitere Sheets erzeugen
      for (let i = 1; i < sheets.length; i++) {
        const sh = sheets[i];

        const ws = wb.addWorksheet(sh.name || `Sheet_${i+1}`);

        // Header kopieren
        for (let r = 1; r < dataStartRow; r++) {
          const srcRow = tplSheet.getRow(r);
          const dstRow = ws.getRow(r);
          for (let c = 1; c <= srcRow.cellCount; c++) {
            dstRow.getCell(c).value = srcRow.getCell(c).value;
            dstRow.getCell(c).style = { ...srcRow.getCell(c).style };
          }
          dstRow.commit();
        }

        // Spaltenbreiten übernehmen
        ws.columns = tplSheet.columns.map(col => ({ width: col.width || 10 }));

        // Daten
        writeDataRowsByHeader(ws, sh.rows || []);

        // QA pro Sheet (optional)
        if (sh.qa) buildQASheet(wb, sh.qa, `QA_${sh.name || i+1}`);
      }
    }

    // ---------------------
    // 6) Single-Sheet
    // ---------------------
    else if (Array.isArray(rows)) {
      writeDataRowsByHeader(tplSheet, rows);
      if (qa) buildQASheet(wb, qa);
    }

    // ---------------------
    // 7) QA-Sheet Builder
    // ---------------------
    function buildQASheet(workbook, qaObj, name = "QA") {
      if (!qaObj) return;
      const ws = workbook.addWorksheet(name);
      let r = 1;

      const bar = (x, max=100, len=20)=>{
        const v = Math.max(-max, Math.min(max, Number(x)||0));
        const n = Math.round(Math.abs(v)/max*len);
        return (v>=0? "█".repeat(n) : "-"+"█".repeat(n));
      };

      const fill = (cell, diff)=>{
        let color="FF92D050";
        if (Math.abs(diff)<=1) color="FFFFFF00";
        if (diff< -1) color="FFFF0000";
        cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:color}};
      };

      ws.getCell(r,1).value="QA Übersicht"; ws.getCell(r,1).font={bold:true,size:14}; r+=2;

      ws.getCell(r,1).value="Stichprobe"; ws.getCell(r,2).value=qaObj.stichprobe||0; r+=2;

      // … Rest wie gehabt (gekürzt für Klarheit)
    }

    // ---------------------
    // 8) Datei zurückgeben
    // ---------------------
    const buf = await wb.xlsx.writeBuffer();
    const outB64 = Buffer.from(buf).toString("base64");
    const fileName =
      Array.isArray(sheets) && sheets.length > 1
        ? `${prefix}_Festdaten_multi.xlsx`
        : `${prefix}_Festdaten.xlsx`;

    return res.status(200).json({ file: outB64, fileName });

  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
