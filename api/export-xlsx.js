// /api/export-xlsx.js
// v4.2 — Multi‑Sheet SUPPORT, robustes JSON‑Parsing, voller Logging-Support,
//        kein .values mehr, KEIN duplicateRow() (schreibt Zeilen direkt)
import ExcelJS from "exceljs";

export const config = {
  api: { bodyParser: { sizeLimit: "50mb" } },
};

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    // ----------------------------------------------------------
    // 0) BODY LOADING + LOGGING
    // ----------------------------------------------------------
    const raw = req.body;

    console.log("RAW BODY RECEIVED TYPE:", typeof raw);
    if (typeof raw === "string") {
      console.log("RAW BODY STRING (first 500 chars):", raw.slice(0, 500));
    } else {
      console.log("RAW BODY OBJECT KEYS:", raw ? Object.keys(raw) : null);
    }

    let body;
    try {
      body =
        typeof raw === "string"
          ? JSON.parse(raw)
          : raw && typeof raw === "object"
          ? raw
          : {};
    } catch (e) {
      console.error("JSON.parse FAILED:", e?.message);
      return res.status(400).json({
        error: "Invalid JSON body",
        detail: String(e?.message ?? e),
      });
    }

    console.log("PARSED BODY:", JSON.stringify(body, null, 2));

    // ----------------------------------------------------------
    // 1) BODY ENTPACKEN
    // ----------------------------------------------------------
    const {
      templateBase64,
      rows, // optional (single sheet)
      qa,   // optional (single QA)
      sheets, // MULTI-SHEET [{ name, rows, qa }]
      prefix = "SITE",
      headerOrder = [],
      dataStartRow = 4,
    } = body;

    if (!templateBase64) {
      return res.status(400).json({ error: "templateBase64 fehlt" });
    }

    // ----------------------------------------------------------
    // 2) TEMPLATE LADEN
    // ----------------------------------------------------------
    const wb = new ExcelJS.Workbook();

    const tplBuf = Buffer.from(
      String(templateBase64).includes("base64,")
        ? String(templateBase64).split("base64,").pop()
        : String(templateBase64),
      "base64"
    );

    await wb.xlsx.load(tplBuf);

    const tplSheet = wb.worksheets[0];
    if (!tplSheet) {
      return res.status(400).json({ error: "Kein Worksheet im Template" });
    }

    // ----------------------------------------------------------
    // 3) VISIBLE HEADERS (Spalten der ersten Zeile)
    // ----------------------------------------------------------
    function readVisibleHeaders(ws) {
      const headerRow = ws.getRow(1);
      const out = [];
      for (let c = 1; c <= headerRow.cellCount; c++) {
        const v = headerRow.getCell(c).value;
        if (!v) break;
        out.push(String(v));
      }
      return out;
    }

    const visibleHeaders = readVisibleHeaders(tplSheet);

    // ----------------------------------------------------------
    // 4) HEADER → KEY MAPPING
    // ----------------------------------------------------------
    const lc = (s) => String(s || "").toLowerCase().trim();
    const isNum = (h) =>
      ["no.", "nr.", "no", "nr", "anzahl"].includes(lc(h));
    const isQ10Andere = (h) =>
      lc(h).includes("andere") ||
      lc(h).includes("keine") ||
      lc(h).includes("markenname");

    function keyForHeaderText(txt) {
      const h = lc(txt);

      if (isNum(txt)) return "__NUM__";

      if (h.includes("q2") && (h.includes("gender") || h.includes("geschlecht")))
        return "Q2_Gender";

      if (h.includes("q3") && (h.includes("age") || h.includes("alter")))
        return "Q3_Age";

      if (h.includes("q7") && h.includes("frisch"))
        return "Q7_Essverhalten_Frischkaese";
      if (h.includes("q7") && h.includes("gouda"))
        return "Q7_Essverhalten_Gouda";
      if (h.includes("q7") && (h.includes("butter") || h.includes("butterkäse")))
        return "Q7_Essverhalten_Butterkaese";
      if (h.includes("q7") && h.includes("camembert"))
        return "Q7_Essverhalten_Camembert";

      if (h.startsWith("q9") || h.includes("geschmack") || h.includes("lehne"))
        return "Q9_Ablehnung";

      if (h.includes("q10")) {
        if (isQ10Andere(txt)) return "Q10_Marke_Andere";
        if (
          h.includes("welches") ||
          h.includes("haupt") ||
          h.includes("marke") ||
          h.includes("produkt")
        )
          return "Q10_Marke";
      }

      if (h.includes("q11") && h.includes("fett")) return "Q11_Fettgehalt";

      // Fallback: exakter Treffer aus headerOrder
      if (Array.isArray(headerOrder)) {
        const direct = headerOrder.find((k) => lc(k) === h);
        if (direct) return direct;
      }

      return null;
    }

    const columnPlan = visibleHeaders.map(keyForHeaderText);

    // ----------------------------------------------------------
    // 5) HEADER + FORMAT KOPIEREN (Zeilen < dataStartRow)
    // ----------------------------------------------------------
    function cloneHeaderAndWidths(src, dst) {
      for (let r = 1; r < dataStartRow; r++) {
        const sRow = src.getRow(r);
        const dRow = dst.getRow(r);

        for (let c = 1; c <= sRow.cellCount; c++) {
          const sCell = sRow.getCell(c);
          const dCell = dRow.getCell(c);

          dCell.value = sCell.value;
          if (sCell.style) dCell.style = { ...sCell.style };
          if (sCell.font) dCell.font = { ...sCell.font };
          if (sCell.alignment) dCell.alignment = { ...sCell.alignment };
          if (sCell.border) dCell.border = { ...sCell.border };
          if (sCell.fill) dCell.fill = { ...sCell.fill };
          if (sCell.numFmt) dCell.numFmt = sCell.numFmt;
        }

        dRow.commit();
      }

      dst.columns = src.columns.map((col) => ({
        width: col.width || 10,
      }));
    }

    // ----------------------------------------------------------
    // 6) DATA WRITER — OHNE duplicateRow()
    //    (legt Zeilen direkt an; kopiert optional Stil aus Musterzeile,
    //     falls dataStartRow im Template existiert)
    // ----------------------------------------------------------
    function writeDataRowsByPlan(ws, dataRows, plan) {
      if (!Array.isArray(dataRows) || !dataRows.length) return;

      // Optionale Musterzeile (falls im Template vorhanden)
      const sampleExists = Boolean(ws._rows?.[dataStartRow - 1]);
      const sampleRow = sampleExists ? ws.getRow(dataStartRow) : null;

      for (let i = 0; i < dataRows.length; i++) {
        const rowNum = dataStartRow + i;
        const dst = ws.getRow(rowNum);
        const src = dataRows[i] || {};

        for (let c = 1; c <= plan.length; c++) {
          const key = plan[c - 1];
          const dCell = dst.getCell(c);

          // Minimalen Stil aus der Musterzeile kopieren (sofern vorhanden)
          if (sampleRow) {
            const sCell = sampleRow.getCell(c);
            if (sCell.style) dCell.style = { ...sCell.style };
            if (sCell.font) dCell.font = { ...sCell.font };
            if (sCell.alignment) dCell.alignment = { ...sCell.alignment };
            if (sCell.border) dCell.border = { ...sCell.border };
            if (sCell.fill) dCell.fill = { ...sCell.fill };
            if (sCell.numFmt) dCell.numFmt = sCell.numFmt;
          }

          if (key === "__NUM__") {
            dCell.value = i + 1;
          } else if (!key) {
            dCell.value = "";
          } else {
            dCell.value = src[key] ?? "";
          }
        }

        dst.commit();
      }
    }

    // ----------------------------------------------------------
    // 7) QA SHEET (OPTIONAL)
    // ----------------------------------------------------------
    function buildQASheet(workbook, qaObj, name = "QA") {
      if (!qaObj || typeof qaObj !== "object") return;

      const ws = workbook.addWorksheet(name);
      let r = 1;

      ws.getCell(r, 1).value = "QA Übersicht";
      ws.getCell(r, 1).font = { bold: true, size: 14 };
      r += 2;

      for (const [k, v] of Object.entries(qaObj)) {
        ws.getCell(r, 1).value = k;
        ws.getCell(r, 2).value =
          typeof v === "object" ? JSON.stringify(v) : String(v);
        r++;
      }
    }

    // ----------------------------------------------------------
    // 8) MULTI-SHEET HANDLING
    // ----------------------------------------------------------
    cloneHeaderAndWidths(tplSheet, tplSheet);

    if (Array.isArray(sheets) && sheets.length > 0) {
      // Erstes Sheet ins Template
      writeDataRowsByPlan(tplSheet, sheets[0]?.rows ?? [], columnPlan);
      if (sheets[0]?.qa) buildQASheet(wb, sheets[0].qa, "QA");

      // Weitere Sheets erzeugen
      for (let i = 1; i < sheets.length; i++) {
        const el = sheets[i];
        const ws = wb.addWorksheet(String(el?.name || `Sheet_${i + 1}`));

        cloneHeaderAndWidths(tplSheet, ws);
        writeDataRowsByPlan(ws, el?.rows ?? [], columnPlan);

        if (el?.qa) buildQASheet(wb, el.qa, `QA_${el?.name || i + 1}`);
      }
    }
    // ----------------------------------------------------------
    // 8b) SINGLE SHEET FALLBACK
    // ----------------------------------------------------------
    else if (Array.isArray(rows)) {
      writeDataRowsByPlan(tplSheet, rows, columnPlan);
      if (qa) buildQASheet(wb, qa, "QA");
    }

    else {
      return res.status(400).json({ error: "Weder rows noch sheets übergeben." });
    }

    // ----------------------------------------------------------
    // 9) RETURN FILE
    // ----------------------------------------------------------
    const buf = await wb.xlsx.writeBuffer();
    const outB64 = Buffer.from(buf).toString("base64");

    const fileName =
      Array.isArray(sheets) && sheets.length > 1
        ? `${prefix}_Festdaten_multi.xlsx`
        : `${prefix}_Festdaten.xlsx`;

    res.setHeader("X-Exporter-Version", "4.2");
    return res.status(200).json({
      file: outB64,
      fileName,
      success: true,
      version: "4.2",
    });

  } catch (err) {
    console.error("EXPORT XLSX ERROR", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
