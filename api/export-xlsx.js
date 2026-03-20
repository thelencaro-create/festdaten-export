// /api/export-xlsx.js
import ExcelJS from "exceljs";

export const config = {
  api: { bodyParser: { sizeLimit: "25mb" } } // falls große Templates
};

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    // --- Body robust einlesen (RAW oder bereits geparst) ---
    const raw = req.body;
    const body = (typeof raw === "string") ? JSON.parse(raw) : (raw || {});
    const { templateBase64, rows, qa, prefix, sheets } = body;

    if (!templateBase64) return res.status(400).json({ error: "templateBase64 fehlt" });

    // --- Template laden ---
    const wb = new ExcelJS.Workbook();
    const tplBuf = Buffer.from(
      String(templateBase64).includes("base64,")
        ? String(templateBase64).split("base64,").pop()
        : String(templateBase64),
      "base64"
    );
    await wb.xlsx.load(tplBuf);

    // ### Hilfsfunktionen ####################################################

    // Auto-Spaltenbreite (simpler Heuristik-basiert)
    function autoFitColumns(worksheet, fromRow = 1, toRow = null) {
      const lastRow = toRow || worksheet.rowCount;
      worksheet.columns.forEach((col) => {
        let max = 8; // Mindestbreite
        for (let r = fromRow; r <= lastRow; r++) {
          const v = String(worksheet.getRow(r).getCell(col.number).value ?? "");
          if (v.length > max) max = v.length;
        }
        col.width = Math.min(Math.max(10, Math.ceil(max * 1.1)), 60);
      });
    }

    // kopiert Style-Eigenschaften von Quellzelle auf Zielzelle
    function copyCellStyle(src, dst) {
      if (!src || !dst) return;
      dst.style = { ...src.style }; // ExcelJS kopiert style-Objekt
      // Einzelkomponenten defensiv setzen (falls style-Objekt unvollständig ist)
      if (src.font)       dst.font = { ...src.font };
      if (src.alignment)  dst.alignment = { ...src.alignment };
      if (src.border)     dst.border = { ...src.border };
      if (src.fill)       dst.fill = { ...src.fill };
      if (src.numFmt)     dst.numFmt = src.numFmt;
      if (src.protection) dst.protection = { ...src.protection };
    }

    // dupliziert eine ganze Zeile (Werte optional ignorieren) als "Style-Schablone"
    function duplicateStyleRow(worksheet, templateRowNumber, countToInsert) {
      if (countToInsert <= 0) return;
      worksheet.duplicateRow(templateRowNumber, countToInsert, true); // insert = true
      // Werte der frisch eingefügten Zeilen leeren (nur falls Templates Werte tragen)
      for (let i = 0; i < countToInsert; i++) {
        const r = worksheet.getRow(templateRowNumber + i);
        r.eachCell((cell) => { if (cell.type !== ExcelJS.ValueType.Merge) cell.value = null; });
        r.commit();
      }
    }

    // schreibt rows[] in ein Worksheet ab dataStartRow, ohne Header zu berühren
    function writeDataRowsPreservingStyles(worksheet, rowsArray, dataStartRow = 3) {
      if (!Array.isArray(rowsArray) || rowsArray.length === 0) {
        throw new Error("rows fehlt/leer");
      }
      // Header aus Zeile 1 (dein Template) lesen → Reihenfolge bleibt exakt
      const headerRow = worksheet.getRow(1);
      const headers = [];
      for (let c = 1; c <= headerRow.cellCount; c++) {
        const v = headerRow.getCell(c).value;
        if (v == null || v === "") break; // bis zur ersten leeren Zelle
        headers.push(String(v));
      }
      if (!headers.length) throw new Error("Headerzeile (Row 1) leer.");

      // sicherstellen, dass genug Datenzeilen mit Style existieren
      // wir nehmen Zeile dataStartRow als "Prototyp" für Styles/Merges
      const need = Math.max(0, rowsArray.length - (worksheet.rowCount - (dataStartRow - 1)));
      if (need > 0) duplicateStyleRow(worksheet, dataStartRow, need);

      // Zellen beschreiben (KEIN row.values = …), nur Werte setzen
      for (let i = 0; i < rowsArray.length; i++) {
        const src = rowsArray[i] || {};
        const rowNum = dataStartRow + i;
        const xRow = worksheet.getRow(rowNum);

        for (let c = 1; c <= headers.length; c++) {
          const header = headers[c - 1];
          const cell = xRow.getCell(c);
          const val = Object.prototype.hasOwnProperty.call(src, header) ? src[header] : "";
          cell.value = (val == null ? "" : val);
        }
        xRow.commit();
      }
    }

    // QA-Sheet erzeugen: mit Ampel (grün/gelb/rot) & "Delta‑Bars" (Blockgrafik)
    function buildQASheet(workbook, qa, name = "QA") {
      if (!qa || typeof qa !== "object") return;
      const ws = workbook.addWorksheet(name);
      let r = 1;

      // Helper für Delta-Balken (ASCII)
      const bar = (x, maxAbs = 100, len = 20) => {
        const v = Math.max(-maxAbs, Math.min(maxAbs, Number(x) || 0));
        const n = Math.round(Math.abs(v) / maxAbs * len);
        const blocks = "█".repeat(n);
        return v >= 0 ? blocks : `-${blocks}`;
      };

      // Helper für Ampel: Färbe Zelle anhand diff
      const setTrafficFill = (cell, diff) => {
        let color = "FF92D050"; // grün
        if (Math.abs(diff) <= 1) color = "FFFFFF00"; // gelb (nahe 0)
        if (diff < -1) color = "FFFF0000"; // rot (negativ)
        cell.fill = { type: 'pattern', pattern:'solid', fgColor:{ argb: color } };
      };

      ws.getCell(r,1).value = "QA Übersicht"; ws.getCell(r,1).font = { bold:true, size:14 }; r+=2;

      // Stichprobe
      ws.getCell(r,1).value = "Stichprobe"; ws.getCell(r,2).value = qa.stichprobe || 0; r+=2;

      // Brands (falls vorhanden)
      if (qa.brands) {
        ws.getCell(r,1).value = "Marken Soll (scaled)"; ws.getCell(r,1).font = { bold:true }; r++;
        ws.getCell(r,1).value = "Kerrygold"; ws.getCell(r,2).value = qa.brands.soll_scaled?.Kerrygold ?? null; r++;
        ws.getCell(r,1).value = "Andere";    ws.getCell(r,2).value = qa.brands.soll_scaled?.Andere ?? null; r+=2;

        ws.getCell(r,1).value = "Marken Ist"; ws.getCell(r,1).font = { bold:true }; r++;
        const kgIst = qa.brands.ist?.Kerrygold ?? 0;
        const anIst = qa.brands.ist?.Andere ?? 0;
        ws.getCell(r,1).value = "Kerrygold"; ws.getCell(r,2).value = kgIst; r++;
        ws.getCell(r,1).value = "Andere";    ws.getCell(r,2).value = anIst; r+=2;

        ws.getCell(r,1).value = "Differenzen"; ws.getCell(r,1).font = { bold:true }; r++;
        const kgDiff = qa.brands.diff?.Kerrygold ?? 0;
        const anDiff = qa.brands.diff?.Andere ?? 0;
        ws.getCell(r,1).value = "Kerrygold"; ws.getCell(r,2).value = kgDiff; setTrafficFill(ws.getCell(r,2), kgDiff);
        ws.getCell(r,3).value = bar(kgDiff, Math.max(qa.stichprobe||1, 100)); r++;
        ws.getCell(r,1).value = "Andere";    ws.getCell(r,2).value = anDiff; setTrafficFill(ws.getCell(r,2), anDiff);
        ws.getCell(r,3).value = bar(anDiff, Math.max(qa.stichprobe||1, 100)); r+=2;

        // Free-Text Anteil, wenn vorhanden
        if (qa.brands.pools) {
          ws.getCell(r,1).value = "Free-Text Code / Anteil"; ws.getCell(r,1).font = { bold:true }; r++;
          ws.getCell(r,1).value = String(qa.brands.pools.free_text_code || "");
          ws.getCell(r,2).value = qa.brands.pools.free_text_share || 0; ws.getCell(r,2).numFmt = "0.00%";
          ws.getCell(r,3).value = qa.brands.pools.free_text_rate || 0; ws.getCell(r,3).numFmt = "0.00%"; r+=2;
        }
      }

      // Beliebige weitere Blöcke flach ausschreiben
      ws.getCell(r,1).value = "Verteilungen (Gender/Age/Brand/Fett)"; ws.getCell(r,1).font = { bold:true }; r++;
      for (const [title, map] of [["Gender", qa.gender], ["Age", qa.age], ["Brand", qa.brand], ["Fett", qa.fett]]) {
        if (!map) continue;
        ws.getCell(r,1).value = title; ws.getCell(r,1).font = { italic:true }; r++;
        for (const [k,v] of Object.entries(map)) { ws.getCell(r,1).value = k; ws.getCell(r,2).value = v; r++; }
        r++;
      }

      autoFitColumns(ws, 1, ws.rowCount);
    }

    // Multi-Sheet Support: sheets:[{name, rows, qa}]
    async function writeSingleSheetFromTemplate(ws, dataRows, optQA) {
      writeDataRowsPreservingStyles(ws, dataRows, 3);
      if (optQA) buildQASheet(wb, optQA, "QA");
    }

    async function writeMultiSheetsFromTemplate(tplSheet, list) {
      // Erstes Element befüllt das bestehende Template-Sheet
      const first = list[0];
      await writeSingleSheetFromTemplate(tplSheet, first.rows || [], first.qa || null);
      // weitere Elemente: neue Sheets anlegen und Header/Styles reproduzieren
      for (let i = 1; i < list.length; i++) {
        const el = list[i];
        const ws = wb.addWorksheet(String(el.name || `Sheet_${i+1}`));

        // Kopfzeilen-/Vorlagenbereich (z. B. Zeilen 1..2) aus dem Template übernehmen
        // Kopiere Werte + Styles Zelle für Zelle
        const headerMaxRow = 2; // Falls dein Template 2 Kopfzeilen hat; ggf. anpassen
        for (let r = 1; r <= headerMaxRow; r++) {
          const srcRow = tplSheet.getRow(r);
          const dstRow = ws.getRow(r);
          for (let c = 1; c <= srcRow.cellCount; c++) {
            const srcCell = srcRow.getCell(c);
            const dstCell = dstRow.getCell(c);
            dstCell.value = srcCell.value; // Text/Formula übernehmen
            copyCellStyle(srcCell, dstCell); // Styles kopieren
          }
          dstRow.commit();
        }
        // Spaltenbreiten vom Template übernehmen
        ws.columns = tplSheet.columns.map(col => ({ width: col.width || 10 }));

        // Datenzeile 3 als Prototyp im Zielsheet anlegen
        // (Kopiere Styles der Zeile 3 aus dem Templatesheet)
        const srcProto = tplSheet.getRow(3);
        const dstProto = ws.getRow(3);
        for (let c = 1; c <= srcProto.cellCount; c++) {
          const srcCell = srcProto.getCell(c);
          const dstCell = dstProto.getCell(c);
          copyCellStyle(srcCell, dstCell);
        }
        dstProto.commit();

        // Jetzt Daten schreiben (inkl. Duplicate-Row bei Bedarf)
        writeDataRowsPreservingStyles(ws, el.rows || [], 3);

        // QA pro Sheet? (optional)
        if (el.qa) buildQASheet(wb, el.qa, `QA_${el.name || i+1}`);
      }
    }

    // ### Single vs Multi Sheets #############################################
    const tplSheet = wb.worksheets[0];
    if (!tplSheet) return res.status(400).json({ error: "Kein Worksheet im Template gefunden." });

    if (Array.isArray(sheets) && sheets.length > 0) {
      await writeMultiSheetsFromTemplate(tplSheet, sheets);
    } else {
      await writeSingleSheetFromTemplate(tplSheet, rows || [], qa || null);
    }

    // --- Ausgabe als Base64 zurück ---
    const outBuf = await wb.xlsx.writeBuffer();
    const outB64 = Buffer.from(outBuf).toString("base64");
    const fileName = Array.isArray(sheets) && sheets.length > 0
      ? `${String(prefix || "SITE")}_Festdaten_multi.xlsx`
      : `${String(prefix || "SITE")}_Festdaten.xlsx`;

    return res.status(200).json({ file: outB64, fileName, success: true });

  } catch (err) {
    console.error("EXPORT XLSX ERROR", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
