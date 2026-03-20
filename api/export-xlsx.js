// /api/export-xlsx.js
import ExcelJS from "exceljs";

export const config = {
  api: { bodyParser: { sizeLimit: "50mb" } },
};

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

    // 0) Body robust einlesen
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
      headerOrder,          // = codebook.column_order  (wird zwingend erwartet)
      dataStartRow = 4      // bei dir fix 4
    } = body;

    if (!templateBase64) return res.status(400).json({ error: "templateBase64 fehlt" });

    // 1) Template laden
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

    // 2) column_order (= headerOrder) ist die einzige Wahrheit zur Spaltenreihenfolge
    if (!Array.isArray(headerOrder) || headerOrder.length === 0) {
      return res.status(400).json({ error: "headerOrder (column_order) fehlt/leer – Export abgebrochen." });
    }

    // 3) Hilfsfunktionen

    // prüft ob ein Header-Text (aus Template-Zeile 1) eine Laufnummer-Spalte meint
    function isNumberingHeader(txt) {
      const H = String(txt || "").trim().toLowerCase();
      return ["anzahl", "no.", "no", "nr.", "nr"].includes(H);
    }

    // liest die sichtbaren Template-Header-Texte aus Row 1 (nur für Nummerierungs-Erkennung)
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

    // kopiert Kopf (Zeilen 1..dataStartRow-1) & Spaltenbreiten ins neue Worksheet
    function cloneHeaderAndWidths(srcWs, dstWs) {
      // Kopfzeilen 1..dataStartRow-1
      for (let r = 1; r <= Math.max(1, dataStartRow - 1); r++) {
        const sRow = srcWs.getRow(r);
        const dRow = dstWs.getRow(r);
        for (let c = 1; c <= sRow.cellCount; c++) {
          const sCell = sRow.getCell(c);
          const dCell = dRow.getCell(c);
          dCell.value = sCell.value;
          // Styles kopieren
          if (sCell.style) dCell.style = { ...sCell.style };
          if (sCell.font) dCell.font = { ...sCell.font };
          if (sCell.alignment) dCell.alignment = { ...sCell.alignment };
          if (sCell.border) dCell.border = { ...sCell.border };
          if (sCell.fill) dCell.fill = { ...sCell.fill };
          if (sCell.numFmt) dCell.numFmt = sCell.numFmt;
        }
        dRow.commit();
      }
      // Spaltenbreiten
      dstWs.columns = srcWs.columns.map(col => ({ width: col.width || 10 }));
    }

    // schreibt dataRows exakt gemäß headerOrder in Spalte 1..N
    // + setzt Nummerierung, wenn Spalte "Anzahl/No./Nr." erkannt wird (entweder in headerOrder oder im sichtbaren Header)
    function writeDataRows(ws, dataRows, headerOrder, visibleHeaderTexts) {
      if (!Array.isArray(dataRows) || dataRows.length === 0) return;

      const need = Math.max(0, dataRows.length - (ws.rowCount - (dataStartRow - 1)));
      if (need > 0) ws.duplicateRow(dataStartRow, need, true);

      // vorbereiten: welche Spalten sind Nummerierungs-Spalten?
      // a) per Key-Name in headerOrder
      const numberingKeys = new Set(["anzahl", "no.", "no", "nr.", "nr"]);
      // b) per sichtbarem Template-Header
      const numberColsByVisible = [];
      for (let c = 1; c <= headerOrder.length; c++) {
        const headKey = String(headerOrder[c-1] || "").toLowerCase();
        const visible = visibleHeaderTexts[c-1] || null;
        const isKeyNumbering = numberingKeys.has(headKey);
        const isVisibleNumbering = isNumberingHeader(visible);
        numberColsByVisible[c] = (isKeyNumbering || isVisibleNumbering);
      }

      for (let i = 0; i < dataRows.length; i++) {
        const rowNum = dataStartRow + i;
        const dstRow = ws.getRow(rowNum);
        const src    = dataRows[i] || {};

        for (let c = 1; c <= headerOrder.length; c++) {
          const key = headerOrder[c - 1];  // exakter Row-Key laut column_order
          const isNumbering = numberColsByVisible[c];

          if (isNumbering) {
            // Laufende Nummer (1..N) schreiben; Template-NumFmt bleibt erhalten
            dstRow.getCell(c).value = i + 1;
            continue;
          }
          dstRow.getCell(c).value = (src[key] == null ? "" : src[key]);
        }
        dstRow.commit();
      }
    }

    // QA-Sheet (einfach, stabil – du kannst es bei Bedarf erweitern)
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

      // Generische Flach-Ausgabe der Blöcke (brands/gender/age/fett/q7)
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

        const writeKV = (k, v) => {
          ws.getCell(r,1).value = String(k);
          ws.getCell(r,2).value = (typeof v === "object") ? JSON.stringify(v) : v;
          r++;
        };

        if (title === "Brands (Soll/Ist/Diff)") {
          writeKV("soll", obj.soll);
          writeKV("soll_scaled", obj.soll_scaled);
          writeKV("ist", obj.ist);
          writeKV("diff", obj.diff);
          writeKV("pools", obj.pools);
          r++;
        } else if (title === "Q7") {
          for (const [k, v] of Object.entries(obj)) writeKV(k, v);
          r++;
        } else {
          for (const [k, v] of Object.entries(obj)) writeKV(k, v);
          r++;
        }
      }
    }

    // 4) Multi vs Single
    const visibleHeaderTexts = readTemplateHeaderTexts(tplSheet);

    if (Array.isArray(sheets) && sheets.length > 0) {
      // Erstes Dataset ins Templatesheet
      writeDataRows(tplSheet, sheets[0]?.rows || [], headerOrder, visibleHeaderTexts);
      if (sheets[0]?.qa) buildQASheet(wb, sheets[0].qa, "QA");

      // Weitere Sheets
      for (let i = 1; i < sheets.length; i++) {
        const el = sheets[i];
        const ws = wb.addWorksheet(String(el?.name || `Sheet_${i+1}`));
        cloneHeaderAndWidths(tplSheet, ws);
        writeDataRows(ws, el?.rows || [], headerOrder, visibleHeaderTexts);
        if (el?.qa) buildQASheet(wb, el.qa, `QA_${el?.name || i+1}`);
      }
    } else if (Array.isArray(rows)) {
      writeDataRows(tplSheet, rows, headerOrder, visibleHeaderTexts);
      if (qa) buildQASheet(wb, qa, "QA");
    } else {
      return res.status(400).json({ error: "Weder rows noch sheets übergeben." });
    }

    // 5) Datei zurückgeben
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
