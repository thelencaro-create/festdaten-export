// /api/export-xlsx.js  (Next.js API Route)
// v3.2 — Q10-Paar wird anhand sichtbarer Header (Row 1) korrekt zugeordnet
import ExcelJS from "exceljs";

export const config = {
  api: { bodyParser: { sizeLimit: "50mb" } },
};

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

    // ----- 0) Body einlesen -----
    const raw = req.body;
    const body =
      typeof raw === "string" ? JSON.parse(raw) :
      (raw && typeof raw === "object") ? raw : {};

    const {
      templateBase64,
      rows,                 // Single-sheet
      qa,                   // Single-sheet QA
      sheets,               // Multi-sheet: [{ name, rows, qa }]
      prefix = "SITE",
      headerOrder,          // = codebook.column_order  (Pflicht!)
      dataStartRow = 4
    } = body;

    if (!templateBase64) return res.status(400).json({ error: "templateBase64 fehlt" });
    if (!Array.isArray(headerOrder) || !headerOrder.length) {
      return res.status(400).json({ error: "headerOrder (column_order) fehlt/leer." });
    }

    // ----- 1) Template laden -----
    const wb = new ExcelJS.Workbook();
    const tplBuf = Buffer.from(
      String(templateBase64).includes("base64,")
        ? String(templateBase64).split("base64,").pop()
        : String(templateBase64),
      "base64"
    );
    await wb.xlsx.load(tplBuf);

    const tplSheet = wb.worksheets[0];
    if (!tplSheet) return res.status(400).json({ error: "Kein Worksheet im Template" });

    // ----- 2) Sichtbare Headertexte (Row 1) lesen -----
    function readTemplateHeaderTexts(ws) {
      const headerRow = ws.getRow(1);
      const out = [];
      for (let c = 1; c <= headerRow.cellCount; c++) {
        const v = headerRow.getCell(c).value;
        if (!v) break;
        out.push(String(v));
      }
      return out;
    }

    const visibleHeaderTexts = readTemplateHeaderTexts(tplSheet);

    // ----- 3) Q10-Paar anhand sichtbarer Header justieren -----
    function isAndereHeader(txt) {
      const H = String(txt || "").toLowerCase();
      return H.includes("keine") || H.includes("andere") || H.includes("markenname");
    }

    function adjustQ10OrderWithVisible(order, visible) {
      if (!Array.isArray(order) || !order.length) return order;
      const eff = order.slice();

      const iMain = eff.findIndex(k => String(k).toLowerCase() === "q10_marke");
      const iAlt  = eff.findIndex(k => String(k).toLowerCase() === "q10_marke_andere");
      if (iMain < 0 || iAlt < 0) return eff;

      const visMain = visible[iMain] || "";
      const visAlt  = visible[iAlt]  || "";

      const mainLooksAndere = isAndereHeader(visMain);
      const altLooksAndere  = isAndereHeader(visAlt);

      // Wenn die "Hauptmarke"-Spalte sichtbar als "Andere..." beschriftet ist → tauschen.
      // Oder wenn die "Andere"-Spalte sichtbar NICHT nach "andere/keine/markenname" aussieht → tauschen.
      if ((mainLooksAndere && !altLooksAndere) || (!mainLooksAndere && !altLooksAndere && visMain && visAlt)) {
        [eff[iMain], eff[iAlt]] = [eff[iAlt], eff[iMain]];
      }
      return eff;
    }

    const effectiveHeaderOrder = adjustQ10OrderWithVisible(headerOrder, visibleHeaderTexts);

    // ----- 4) Kopf+Breiten kopieren (für neue Sheets) -----
    function cloneHeaderAndWidths(srcWs, dstWs) {
      for (let r = 1; r <= Math.max(1, dataStartRow - 1); r++) {
        const sRow = srcWs.getRow(r);
        const dRow = dstWs.getRow(r);
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
      dstWs.columns = srcWs.columns.map(col => ({ width: col.width || 10 }));
    }

    // ----- 5) Daten schreiben (streng nach effectiveHeaderOrder) -----
    function isNumberingHeader(visible) {
      const h = String(visible || "").toLowerCase().trim();
      return ["anzahl", "no.", "no", "nr.", "nr"].includes(h);
    }

    function writeDataRows(ws, dataRows, effectiveOrder, visible) {
      if (!Array.isArray(dataRows) || !dataRows.length) return;

      const need = Math.max(0, dataRows.length - (ws.rowCount - (dataStartRow - 1)));
      if (need > 0) ws.duplicateRow(dataStartRow, need, true);

      // Spalten, die Nummerierung zeigen sollen
      const numberingByCol = [];
      for (let c = 1; c <= effectiveOrder.length; c++) {
        const keyLower = String(effectiveOrder[c-1] || "").toLowerCase();
        const vis = visible[c-1] || null;
        const looksNumberKey = ["anzahl","no.","no","nr.","nr"].includes(keyLower);
        const looksNumberHdr = isNumberingHeader(vis);
        numberingByCol[c] = looksNumberKey || looksNumberHdr;
      }

      for (let i = 0; i < dataRows.length; i++) {
        const rowNum = dataStartRow + i;
        const dst = ws.getRow(rowNum);
        const src = dataRows[i] || {};

        for (let c = 1; c <= effectiveOrder.length; c++) {
          const key = effectiveOrder[c-1];
          if (numberingByCol[c]) { dst.getCell(c).value = i + 1; continue; }
          dst.getCell(c).value = (src[key] == null ? "" : src[key]);
        }
        dst.commit();
      }
    }

    // ----- 6) QA-Sheet (optional) -----
    function buildQASheet(workbook, qaObj, name = "QA") {
      if (!qaObj || typeof qaObj !== "object") return;
      const ws = workbook.addWorksheet(name);
      let r = 1;

      ws.getCell(r,1).value = "QA Übersicht";
      ws.getCell(r,1).font = { bold: true, size: 14 };
      r += 2;

      ws.getCell(r,1).value = "Stichprobe";
      ws.getCell(r,2).value = qaObj.stichprobe || 0;
      r += 2;

      const blocks = [
        ["Brands (Soll/Ist/Diff)", qaObj.brands],
        ["Gender", qaObj.gender],
        ["Age", qaObj.age],
        ["Brand Counts", qaObj.brand],
        ["Fett", qaObj.fett],
        ["Q7", qaObj.q7]
      ];

      for (const [title, obj] of blocks) {
        if (!obj || typeof obj !== "object") continue;
        ws.getCell(r,1).value = title;
        ws.getCell(r,1).font = { bold: true };
        r++;
        const putKV = (k,v)=>{ ws.getCell(r,1).value = String(k); ws.getCell(r,2).value = (typeof v === "object") ? JSON.stringify(v) : v; r++; };
        if (title === "Brands (Soll/Ist/Diff)") {
          putKV("soll", obj.soll);
          putKV("soll_scaled", obj.soll_scaled);
          putKV("ist", obj.ist);
          putKV("diff", obj.diff);
          putKV("pools", obj.pools);
          r++;
        } else if (title === "Q7") {
          for (const [k,v] of Object.entries(obj)) putKV(k,v);
          r++;
        } else {
          for (const [k,v] of Object.entries(obj)) putKV(k,v);
          r++;
        }
      }
    }

    // ----- 7) Multi vs. Single -----
    if (Array.isArray(sheets) && sheets.length > 0) {
      // Erstes Sheet ins Template
      writeDataRows(tplSheet, sheets[0]?.rows || [], effectiveHeaderOrder, visibleHeaderTexts);
      if (sheets[0]?.qa) buildQASheet(wb, sheets[0].qa, "QA");

      // Weitere Sheets
      for (let i = 1; i < sheets.length; i++) {
        const el = sheets[i];
        const ws = wb.addWorksheet(String(el?.name || `Sheet_${i+1}`));
        cloneHeaderAndWidths(tplSheet, ws);
        writeDataRows(ws, el?.rows || [], effectiveHeaderOrder, visibleHeaderTexts);
        if (el?.qa) buildQASheet(wb, el.qa, `QA_${el?.name || i+1}`);
      }
    } else if (Array.isArray(rows)) {
      writeDataRows(tplSheet, rows, effectiveHeaderOrder, visibleHeaderTexts);
      if (qa) buildQASheet(wb, qa, "QA");
    } else {
      return res.status(400).json({ error: "Weder rows noch sheets übergeben." });
    }

    // ----- 8) Datei zurückgeben -----
    const buf = await wb.xlsx.writeBuffer();
    const outB64 = Buffer.from(buf).toString("base64");
    const fileName =
      (Array.isArray(sheets) && sheets.length > 1)
        ? `${prefix}_Festdaten_multi.xlsx`
        : `${prefix}_Festdaten.xlsx`;

    return res.status(200).json({ file: outB64, fileName, success: true });

  } catch (err) {
    console.error("EXPORT XLSX ERROR", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
