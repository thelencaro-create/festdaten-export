// /api/export-xlsx.js
import ExcelJS from "exceljs";

export const config = { api: { bodyParser: { sizeLimit: "25mb" } } };

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

    // --- Body robust einlesen ---
    const raw = req.body;
    const body = (typeof raw === "string") ? JSON.parse(raw) : (raw || {});
    const { templateBase64, rows, qa, prefix, sheets, headerOrder, dataStartRow } = body;

    if (!templateBase64) return res.status(400).json({ error: "templateBase64 fehlt" });

    const wb = new ExcelJS.Workbook();
    const tplBuf = Buffer.from(
      String(templateBase64).includes("base64,")
        ? String(templateBase64).split("base64,").pop()
        : String(templateBase64),
      "base64"
    );
    await wb.xlsx.load(tplBuf);

    // === Helpers ===========================================================
    const DATA_START = Number.isFinite(Number(dataStartRow)) ? Number(dataStartRow) : 3;

    const toStr = (v) => (v == null ? "" : String(v));
    const isObj = (x) => x && typeof x === "object" && !Array.isArray(x);

    // Kopieren von Merges vom Template ins Ziel (nur die Bereiche, die bisher existieren)
    function copyMerges(src, dst) {
      // ExcelJS speichert Merges intern; wir lesen die Ranges über model (private API fallback)
      const merges = src?._merges ? Array.from(src._merges) : [];
      for (const m of merges) {
        try { dst.mergeCells(m); } catch { /* ignore overlapping merges */ }
      }
    }

    // Werte + Styles Zelle für Zelle kopieren
    function copyCell(srcCell, dstCell) {
      dstCell.value = srcCell.value;
      if (srcCell.style) dstCell.style = { ...srcCell.style };
      if (srcCell.font) dstCell.font = { ...srcCell.font };
      if (srcCell.alignment) dstCell.alignment = { ...srcCell.alignment };
      if (srcCell.border) dstCell.border = { ...srcCell.border };
      if (srcCell.fill) dstCell.fill = { ...srcCell.fill };
      if (srcCell.numFmt) dstCell.numFmt = srcCell.numFmt;
      if (srcCell.protection) dstCell.protection = { ...srcCell.protection };
    }

    // gesamten Kopfbereich (1..DATA_START-1) Zeile für Zeile kopieren
    function cloneHeader(srcWs, dstWs, headerRowsEnd = DATA_START - 1) {
      // Spaltenbreiten kopieren
      dstWs.columns = srcWs.columns.map(c => ({ width: c.width || 10 }));
      // Kopfzeilen (1..headerRowsEnd) kopieren
      for (let r = 1; r <= Math.max(1, headerRowsEnd); r++) {
        const srcRow = srcWs.getRow(r);
        const dstRow = dstWs.getRow(r);
        for (let c = 1; c <= srcRow.cellCount; c++) {
          copyCell(srcRow.getCell(c), dstRow.getCell(c));
        }
        dstRow.commit();
      }
      // Prototyp‑Datenzeile (DATA_START) – nur Styles übernehmen
      const protoSrc = srcWs.getRow(DATA_START);
      const protoDst = dstWs.getRow(DATA_START);
      for (let c = 1; c <= protoSrc.cellCount; c++) {
        copyCell(protoSrc.getCell(c), protoDst.getCell(c));
        // Werte der Prototypzeile Leer lassen
        protoDst.getCell(c).value = null;
      }
      protoDst.commit();

      // Merges übernehmen (Header + ggf. Prototyp)
      copyMerges(srcWs, dstWs);
    }

    // Daten schreiben: ENTWEDER nach headerOrder (empfohlen) ODER heuristisch über Row 1
    function writeDataRows(ws, dataRows, headerOrderOpt) {
      if (!Array.isArray(dataRows) || dataRows.length === 0) return;

      let headers = null;

      if (Array.isArray(headerOrderOpt) && headerOrderOpt.length) {
        // feste Reihenfolge aus Template/Codebook
        headers = headerOrderOpt.slice();
      } else {
        // Heuristik (Fallback): nimm die erste nicht-leere Zeile als Header
        const headerRow = ws.getRow(1);
        const h = [];
        for (let c = 1; c <= headerRow.cellCount; c++) {
          const v = headerRow.getCell(c).value;
          if (v == null || v === "") break;
          h.push(String(v));
        }
        headers = h;
      }

      if (!headers || headers.length === 0) throw new Error("Export: keine Header bestimmt (headerOrder übergeben!).");

      // ggf. ausreichende Anzahl Datenzeilen (als Style‑Kopie der Prototyp‑Zeile) einfügen
      const need = Math.max(0, dataRows.length - (ws.rowCount - (DATA_START - 1)));
      if (need > 0) ws.duplicateRow(DATA_START, need, true);

      // Zellen füllen – strikt nach Spaltenindex 1..headers.length
      for (let i = 0; i < dataRows.length; i++) {
        const src = dataRows[i] || {};
        const rowNum = DATA_START + i;
        const xRow = ws.getRow(rowNum);

        for (let c = 1; c <= headers.length; c++) {
          const header = headers[c - 1];
          const cell = xRow.getCell(c);
          // Wenn keys exakt dem Header entsprechen: direkt nehmen
          // Falls deine Rows intern kürzere Keys benutzen, mappst du sie upstream.
          const val = Object.prototype.hasOwnProperty.call(src, header) ? src[header] : "";
          cell.value = (val == null ? "" : val);
        }
        xRow.commit();
      }
    }

    // QA‑Sheet (wie zuvor)
    function buildQASheet(workbook, qaObj, name = "QA") {
      if (!qaObj || typeof qaObj !== "object") return;
      const ws = workbook.addWorksheet(name);
      let r = 1;
      const bar = (x, maxAbs = 100, len = 20) => {
        const v = Math.max(-maxAbs, Math.min(maxAbs, Number(x) || 0));
        const n = Math.round(Math.abs(v) / maxAbs * len);
        const blocks = "█".repeat(n);
        return v >= 0 ? blocks : `-${blocks}`;
      };
      const setTrafficFill = (cell, diff) => {
        let color = "FF92D050"; // grün
        if (Math.abs(diff) <= 1) color = "FFFFFF00"; // gelb
        if (diff < -1) color = "FFFF0000";          // rot
        cell.fill = { type: 'pattern', pattern:'solid', fgColor:{ argb: color } };
      };

      ws.getCell(r,1).value = "QA Übersicht"; ws.getCell(r,1).font = { bold:true, size:14 }; r+=2;

      ws.getCell(r,1).value = "Stichprobe"; ws.getCell(r,2).value = qaObj.stichprobe || 0; r+=2;

      if (qaObj.brands) {
        ws.getCell(r,1).value = "Marken Soll (scaled)"; ws.getCell(r,1).font = { bold:true }; r++;
        ws.getCell(r,1).value = "Kerrygold"; ws.getCell(r,2).value = qaObj.brands.soll_scaled?.Kerrygold ?? null; r++;
        ws.getCell(r,1).value = "Andere";    ws.getCell(r,2).value = qaObj.brands.soll_scaled?.Andere ?? null; r+=2;

        ws.getCell(r,1).value = "Marken Ist"; ws.getCell(r,1).font = { bold:true }; r++;
        const kgIst = qaObj.brands.ist?.Kerrygold ?? 0;
        const anIst = qaObj.brands.ist?.Andere ?? 0;
        ws.getCell(r,1).value = "Kerrygold"; ws.getCell(r,2).value = kgIst; r++;
        ws.getCell(r,1).value = "Andere";    ws.getCell(r,2).value = anIst; r+=2;

        ws.getCell(r,1).value = "Differenzen"; ws.getCell(r,1).font = { bold:true }; r++;
        const kgDiff = qaObj.brands.diff?.Kerrygold ?? 0;
        const anDiff = qaObj.brands.diff?.Andere ?? 0;
        ws.getCell(r,1).value = "Kerrygold"; ws.getCell(r,2).value = kgDiff; setTrafficFill(ws.getCell(r,2), kgDiff);
        ws.getCell(r,3).value = bar(kgDiff, Math.max(qaObj.stichprobe||1, 100)); r++;
        ws.getCell(r,1).value = "Andere";    ws.getCell(r,2).value = anDiff; setTrafficFill(ws.getCell(r,2), anDiff);
        ws.getCell(r,3).value = bar(anDiff, Math.max(qaObj.stichprobe||1, 100)); r+=2;

        if (qaObj.brands.pools) {
          ws.getCell(r,1).value = "Free-Text Code / Anteil"; ws.getCell(r,1).font = { bold:true }; r++;
          ws.getCell(r,1).value = String(qaObj.brands.pools.free_text_code || "");
          ws.getCell(r,2).value = qaObj.brands.pools.free_text_share || 0; ws.getCell(r,2).numFmt = "0.00%";
          ws.getCell(r,3).value = qaObj.brands.pools.free_text_rate || 0;  ws.getCell(r,3).numFmt = "0.00%"; r+=2;
        }
      }

      const blocks = [["Gender", qaObj.gender], ["Age", qaObj.age], ["Brand", qaObj.brand], ["Fett", qaObj.fett]];
      for (const [title, map] of blocks) {
        if (!map) continue;
        ws.getCell(r,1).value = title; ws.getCell(r,1).font = { italic:true }; r++;
        for (const [k,v] of Object.entries(map)) { ws.getCell(r,1).value = k; ws.getCell(r,2).value = v; r++; }
        r++;
      }

      // simpler Auto-Fit
      ws.columns.forEach((col) => {
        let max = 10;
        for (let rr = 1; rr <= ws.rowCount; rr++) {
          const v = toStr(ws.getRow(rr).getCell(col.number).value);
          if (v.length > max) max = v.length;
        }
        col.width = Math.min(Math.max(10, Math.ceil(max * 1.1)), 60);
      });
    }

    // === Single vs. Multi ================================================
    const tplSheet = wb.worksheets[0];
    if (!tplSheet) return res.status(400).json({ error: "Kein Worksheet im Template gefunden." });

    if (Array.isArray(sheets) && sheets.length > 0) {
      // Erstes Dataset schreibt ins Template-Sheet
      writeDataRows(tplSheet, sheets[0].rows || [], headerOrder);
      if (sheets[0].qa) buildQASheet(wb, sheets[0].qa, "QA");

      // Weitere Datasets in neue Sheets (Header komplett klonen)
      for (let i = 1; i < sheets.length; i++) {
        const el = sheets[i];
        const ws = wb.addWorksheet(String(el.name || `Sheet_${i+1}`), {
          properties: { ...tplSheet.properties },
          pageSetup:  { ...tplSheet.pageSetup  },
          views:      tplSheet.views ? JSON.parse(JSON.stringify(tplSheet.views)) : undefined
        });
        cloneHeader(tplSheet, ws, DATA_START - 1);
        writeDataRows(ws, el.rows || [], headerOrder);
        if (el.qa) buildQASheet(wb, el.qa, `QA_${el.name || i+1}`);
      }
    } else {
      // Single Sheet
      writeDataRows(tplSheet, rows || [], headerOrder);
      if (qa) buildQASheet(wb, qa, "QA");
    }

    const outBuf = await wb.xlsx.writeBuffer();
    const outB64 = Buffer.from(outBuf).toString("base64");
    const fileName = (Array.isArray(sheets) && sheets.length > 1)
      ? `${String(prefix || "SITE")}_Festdaten_multi.xlsx`
      : `${String(prefix || "SITE")}_Festdaten.xlsx`;

    return res.status(200).json({ file: outB64, fileName, success: true });

  } catch (err) {
    console.error("EXPORT XLSX ERROR", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
