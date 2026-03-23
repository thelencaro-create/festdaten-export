/*********************************************************************
 * EXPORT XLSX v5.0 — KOMPLETTER END-TO-END EXPORTER
 *
 *  ✔ Sheet 1  = Hamburg (oder anderer sheetName)
 *               → Kompakt wie QA, keine Template-Texte
 *               → Nur Datenzeilen
 *
 *  ✔ Sheet 2  = QA
 *               → Soll / Ist / Diff
 *               → sauber formatiert
 *
 *  ✔ Alle Werte STRING -> keine leeren Zellen
 *  ✔ Kein Template-Duplikat, keine "1 = männlich"-Rows usw.
*********************************************************************/

import ExcelJS from "exceljs";

export const config = {
  api: { bodyParser: { sizeLimit: "50mb" } },
};

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    // Body sicher laden
    const raw = req.body;
    const body =
      typeof raw === "string"
        ? JSON.parse(raw)
        : raw && typeof raw === "object"
        ? raw
        : {};

    const {
      rows = [],              // → 125 Datenzeilen (Personas)
      column_order = [],      // → Reihenfolge wie Q2_Gender usw.
      sheetName = "Hamburg",  // → Hauptdatenblatt
      prefix = "SITE",
      targetsSummary = null   // → QA-Daten aus deinem Screener-Parser
    } = body;

    if (!Array.isArray(rows)) {
      return res.status(400).json({ error: "rows must be array" });
    }
    if (!Array.isArray(column_order) || !column_order.length) {
      return res.status(400).json({ error: "column_order missing" });
    }

    /******************************************************************
     * WORKBOOK
     ******************************************************************/
    const wb = new ExcelJS.Workbook();

    /******************************************************************
     * SHEET 1 — DATENBLATT (kompakt wie QA)
     ******************************************************************/
    const wsData = wb.addWorksheet(sheetName);

    // Kopfzeile: Anzahl + alle Variablen + No.
    const header = ["Anzahl", ...column_order, "No."];
    const hdr = wsData.addRow(header);

    hdr.eachCell((c) => {
      c.font = { bold: true };
      c.alignment = { vertical: "middle", horizontal: "center" };
      c.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
        bottom: { style: "thin" },
      };
    });

    // Datenzeilen
    rows.forEach((row, i) => {
      const line = [
        String(i + 1),
        ...column_order.map((key) => String(row?.[key] ?? "")),
        String(i + 1),
      ];

      const xlRow = wsData.addRow(line);
      xlRow.eachCell((c) => {
        c.alignment = { vertical: "middle", horizontal: "center" };
        c.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" },
          bottom: { style: "thin" },
        };
      });
    });

    wsData.columns.forEach((col) => (col.width = 16));

    /******************************************************************
     * SHEET 2 — QA (Soll / Ist / Diff)
     ******************************************************************/
    const wsQA = wb.addWorksheet("QA");

    // Kopf „QA sheet“
    const qaHeader = ["Variable", "Code", "Soll", "Ist", "Diff", "Notes"];
    const qhRow = wsQA.addRow(qaHeader);

    qhRow.eachCell((c) => {
      c.font = { bold: true };
      c.alignment = { vertical: "middle", horizontal: "center" };
      c.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
        bottom: { style: "thin" },
      };
    });

    if (targetsSummary && typeof targetsSummary === "object") {
      const blocks = Object.entries(targetsSummary);

      for (const [section, info] of blocks) {
        if (section === "fallzahl") continue; // reine Meta-Info

        if (!info || typeof info !== "object") continue;
        const soll = info.values || {};
        const ist  = info.ist    || {};
        const diff = {};

        Object.keys({ ...soll, ...ist }).forEach((code) => {
          diff[code] = (Number(ist[code] ?? 0) - Number(soll[code] ?? 0));
        });

        // Abschnittsname
        const sec = wsQA.addRow([section.toUpperCase()]);
        sec.font = { bold: true };
        sec.eachCell((c) => {
          c.alignment = { vertical: "middle", horizontal: "left" };
        });

        // Wertezeilen
        Object.keys({ ...soll, ...ist }).forEach((code) => {
          const row = wsQA.addRow([
            section,
            code,
            String(soll[code] ?? 0),
            String(ist[code] ?? 0),
            String(diff[code] ?? 0),
            Array.isArray(info.notes) ? info.notes.join("; ") : ""
          ]);

          row.eachCell((c) => {
            c.alignment = { vertical: "middle", horizontal: "center" };
            c.border = {
              top: { style: "thin" },
              left: { style: "thin" },
              right: { style: "thin" },
              bottom: { style: "thin" },
            };
          });
        });

        wsQA.addRow([]); // Leerzeile
      }
    }

    wsQA.columns.forEach((col) => (col.width = 18));

    /******************************************************************
     * EXPORT
     ******************************************************************/
    const buf = await wb.xlsx.writeBuffer();

    return res.status(200).json({
      file: Buffer.from(buf).toString("base64"),
      fileName: `${prefix}_Festdaten.xlsx`,
      success: true,
      version: "5.0",
    });

  } catch (err) {
    console.error("EXPORT ERROR", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
