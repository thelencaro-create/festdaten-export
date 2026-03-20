// /api/export-xlsx.js
// v3.3 — Sichtbare-Header-first Mapping für ALLE Spalten (+ Q10-Paar sicher), Nummerierung & QA inklusive
import ExcelJS from "exceljs";

export const config = {
  api: { bodyParser: { sizeLimit: "50mb" } },
};

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

    // ---- 0) Body lesen ----
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
      headerOrder,          // optional Sicherheitsnetz (column_order)
      dataStartRow = 4
    } = body;

    if (!templateBase64) return res.status(400).json({ error: "templateBase64 fehlt" });

    // ---- 1) Template laden ----
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

    // ---- 2) Sichtbare Headertexte (Row 1) lesen ----
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
    const visible = readVisibleHeaders(tplSheet);

    // ---- 3) Heuristiken ----
    const lc = (s)=>String(s||"").toLowerCase();
    const isNumHdr = (t) => {
      const h = lc(t).trim();
      return ["anzahl","no.","no","nr.","nr"].includes(h);
    };
    const isQ10AndereHdr = (t) => {
      const h = lc(t);
      return h.includes("keine") || h.includes("andere") || h.includes("markenname");
    };

    // ---- 4) Sichtbare-Header-first: Spaltenplan aufstellen (pro Spalte 1 Key oder "__NUM__") ----
    function keyForHeaderText(txt) {
      const h = lc(txt);

      // Nummerierung
      if (isNumHdr(txt)) return "__NUM__";

      // Q2 Gender
      if (h.includes("q2") && (h.includes("gender") || h.includes("geschlecht"))) return "Q2_Gender";

      // Q3 Age
      if (h.includes("q3") && (h.includes("alter") || h.includes("age") || h.includes("altersgruppe"))) return "Q3_Age";

      // Q7
      if (h.includes("q7") && h.includes("frisch"))     return "Q7_Essverhalten_Frischkaese";
      if (h.includes("q7") && h.includes("gouda"))      return "Q7_Essverhalten_Gouda";
      if (h.includes("q7") && (h.includes("butter") || h.includes("butterkäse"))) return "Q7_Essverhalten_Butterkaese";
      if (h.includes("q7") && h.includes("camembert"))  return "Q7_Essverhalten_Camembert";

      // Q9
      if (h.startsWith("q9") || h.includes("geschmacksrichtungen") || h.includes("lehne")) return "Q9_Ablehnung";

      // Q10 Marke / Andere
      if (h.includes("q10")) {
        if (isQ10AndereHdr(txt)) return "Q10_Marke_Andere";
        // Hauptmarke erkennen – möglichst großzügig:
        if (h.includes("welches") || h.includes("produkt") || h.includes("verwenden") || h.includes("haupts") || h.includes("marke")) {
          return "Q10_Marke";
        }
      }

      // Q11 Fett
      if (h.includes("q11") && h.includes("fett")) return "Q11_Fettgehalt";

      // Fallback: falls headerOrder gegeben ist, probiere exakten Treffer (z. B. bereits „Q10_Marke“ als Klartext)
      if (Array.isArray(headerOrder)) {
        const direct = headerOrder.find(k => lc(k) === h);
        if (direct) return direct;
      }

      // Unbekannt → nichts schreiben
      return null;
    }

    // Spaltenplan bauen (eine Position pro sichtbarer Spalte)
    const columnPlan = visible.map(keyForHeaderText);

    // ---- 5) Kopf und Breiten klonen (für neue Sheets) ----
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

    // ---- 6) Schreiben nach columnPlan ----
    function writeDataRowsByPlan(ws, dataRows, plan) {
      if (!Array.isArray(dataRows) || !dataRows.length) return;

      const need = Math.max(0, dataRows.length - (ws.rowCount - (dataStartRow - 1)));
      if (need > 0) ws.duplicateRow(dataStartRow, need, true);

      for (let i = 0; i < dataRows.length; i++) {
        const rowNum = dataStartRow + i;
        const dst = ws.getRow(rowNum);
        const src = dataRows[i] || {};

        for (let c = 1; c <= plan.length; c++) {
          const key = plan[c-1];
          if (key === "__NUM__") { dst.getCell(c).value = i + 1; continue; }
          if (!key)             { dst.getCell(c).value = "";   continue; }
          dst.getCell(c).value  = (src[key] == null ? "" : src[key]);
        }
        dst.commit();
      }
    }

    // ---- 7) QA-Sheet (optional) ----
    function buildQASheet(workbook, qaObj, name = "QA") {
      if (!qaObj || typeof qaObj !== "object") return;
      const ws = workbook.addWorksheet(name);
      let r = 1;

      ws.getCell(r,1).value = "QA Übersicht";
      ws.getCell(r,1).font  = { bold: true, size: 14 };
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

      const putKV = (ws,r,k,v)=>{ ws.getCell(r,1).value=String(k); ws.getCell(r,2).value=(typeof v==="object")?JSON.stringify(v):v; return r+1; };

      for (const [title, obj] of blocks) {
        if (!obj || typeof obj !== "object") continue;
        ws.getCell(r,1).value = title;
        ws.getCell(r,1).font  = { bold: true };
        r++;
        if (title === "Brands (Soll/Ist/Diff)") {
          r = putKV(ws,r,"soll",obj.soll);
          r = putKV(ws,r,"soll_scaled",obj.soll_scaled);
          r = putKV(ws,r,"ist",obj.ist);
          r = putKV(ws,r,"diff",obj.diff);
          r = putKV(ws,r,"pools",obj.pools);
          r++;
        } else if (title === "Q7") {
          for (const [k,v] of Object.entries(obj)) r = putKV(ws,r,k,v);
          r++;
        } else {
          for (const [k,v] of Object.entries(obj)) r = putKV(ws,r,k,v);
          r++;
        }
      }
    }

    // ---- 8) Multi vs. Single ----
    if (Array.isArray(sheets) && sheets.length > 0) {
      // Erstes Sheet ins Template
      writeDataRowsByPlan(tplSheet, sheets[0]?.rows || [], columnPlan);
      if (sheets[0]?.qa) buildQASheet(wb, sheets[0].qa, "QA");

      // Weitere Sheets
      for (let i = 1; i < sheets.length; i++) {
        const el = sheets[i];
        const ws = wb.addWorksheet(String(el?.name || `Sheet_${i+1}`));
        cloneHeaderAndWidths(tplSheet, ws);
        writeDataRowsByPlan(ws, el?.rows || [], columnPlan);
        if (el?.qa) buildQASheet(wb, el.qa, `QA_${el?.name || i+1}`);
      }
    } else if (Array.isArray(rows)) {
      writeDataRowsByPlan(tplSheet, rows, columnPlan);
      if (qa) buildQASheet(wb, qa, "QA");
    } else {
      return res.status(400).json({ error: "Weder rows noch sheets übergeben." });
    }

    // ---- 9) Datei zurückgeben + Version ----
    const buf = await wb.xlsx.writeBuffer();
    const outB64 = Buffer.from(buf).toString("base64");
    const fileName =
      (Array.isArray(sheets) && sheets.length > 1)
        ? `${prefix}_Festdaten_multi.xlsx`
        : `${prefix}_Festdaten.xlsx`;

    // Version sichtbar machen (Option A)
    res.setHeader("X-Exporter-Version", "3.3");
    return res.status(200).json({ file: outB64, fileName, success: true, version: "3.3" });

  } catch (err) {
    console.error("EXPORT XLSX ERROR", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
