// api/export-xlsx.js
// Robustes Template-Mapping für beliebige Screener/Excel-Layouts.
// Erwartet im Body: templateBase64, rows, qa, prefix, fallzahl, codebook
// Gibt standardmäßig { file: "<base64>" } zurück (TODO: buildXlsx… an deine Writer-Funktion anbinden).
// Mit ?preview=1 gibt der Endpoint eine Mapping-Vorschau zurück (kein XLSX), um das Matching zu testen.

function normalize(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

// Heuristiken: mappe Template-Spalten (raw-Label) → Variablen-Key in Persona-Row
function guessVarForColumn(rawName, personaKeys) {
  const col = normalize(rawName);

  // 1) Direkter Treffer, falls die Spalte bereits exakt ein Persona-Key ist
  //    (z. B. "Q2_Gender", häufig nach Normalisierung vorhanden)
  if (personaKeys.has(rawName)) return rawName; // exakter raw
  const byNorm = [...personaKeys].find(k => normalize(k) === col);
  if (byNorm) return byNorm;

  // 2) Heuristiken für häufige Variablen
  if (col.includes("q10") && (col.includes("andere") || col.includes("andern") || col.includes("other"))) {
    return "Q10_Marke_Andere";
  }
  if (col.includes("q10")) {
    return "Q10_Marke";
  }

  if (col.includes("fett")) return "Q11_Fettgehalt";
  if (col.includes("ablehn") || col.includes("keine")) return "Q9_Ablehnung";

  if (col.includes("frisch")) return "Q7_Essverhalten_Frischkaese";
  if (col.includes("gouda")) return "Q7_Essverhalten_Gouda";
  if (col.includes("butter")) return "Q7_Essverhalten_Butterkaese";
  if (col.includes("camembert")) return "Q7_Essverhalten_Camembert";

  if (col.includes("gender") || col.includes("geschlecht")) return "Q2_Gender";
  if (col.includes("age") || col.includes("alter") || col.includes("altersgruppe")) return "Q3_Age";

  if (col === "no" || col.includes("nummer")) return "No";

  // Fallback: kein Mapping (wird später als "" befüllt)
  return null;
}

// Erzeuge eine Spalten-Mapping-Tabelle (raw -> varKey) inkl. Duplikat-Regeln (Q10*)
function buildColumnMapping(codebook, samplePersonaRow) {
  // Spalten aus Codebook ableiten: bevorzugt column_order_raw; Fallback auf variables[].raw
  const columns = Array.isArray(codebook?.column_order_raw) && codebook.column_order_raw.length
    ? [...codebook.column_order_raw]
    : Array.isArray(codebook?.variables)
      ? codebook.variables.map(v => v.raw || v.var).filter(Boolean)
      : [];

  // Persona-Keys sammeln (für Direkt-Treffer und Heuristiken)
  const personaKeys = new Set(Object.keys(samplePersonaRow || {}));

  // Schritt 1: Basis-Mapping per Heuristik
  const mapping = new Map(); // raw -> varKey|null
  for (const raw of columns) {
    mapping.set(raw, guessVarForColumn(raw, personaKeys));
  }

  // Schritt 2: Duplikate behandeln, v. a. für Q10 (Marke / Marke_Andere)
  // Finde Gruppen gleicher "Basis" (z. B. alle Spalten, deren normalize() "q10…" enthält)
  const groups = new Map(); // baseKey -> array of raw names
  for (const raw of columns) {
    const nk = normalize(raw);
    const base =
      nk.includes("q10") ? "Q10" :
      nk.includes("q11") ? "Q11" :
      nk.includes("q9")  ? "Q9"  :
      nk.includes("q7")  ? "Q7"  :
      nk; // allgemeine Basis
    if (!groups.has(base)) groups.set(base, []);
    groups.get(base).push(raw);
  }

  // Spezielle Regel: Wenn es genau ZWEI Q10-Spalten gibt und eine davon ist noch nicht als "..._Andere" gemappt:
  if (groups.has("Q10")) {
    const q10cols = groups.get("Q10");
    if (q10cols.length === 2) {
      // Sortiere stabil (wie im Template), weise 1. → Marke, 2. → Marke_Andere
      const [c1, c2] = q10cols;
      mapping.set(c1, mapping.get(c1) ?? "Q10_Marke");
      mapping.set(c2, mapping.get(c2) ?? "Q10_Marke_Andere");
    }
  }

  return mapping; // Map<rawColumn, varKey|null>
}

// Wende die Sonderregel & Typisierung an, bilde dann eine template-kompatible Row (Keys = raw-Spalten)
function buildTemplateRow(personaRow, mapping) {
  const row = { ...personaRow };

  // Sonderregel: Marke_Andere nur, wenn Marke == "10"
  if (String(row["Q10_Marke"] ?? "") !== "10") {
    row["Q10_Marke_Andere"] = "";
  }

  const out = {};
  for (const [raw, vKey] of mapping.entries()) {
    if (!vKey) { out[raw] = ""; continue; }
    const val = row[vKey];
    out[raw] = (val === null || val === undefined) ? "" : String(val);
  }
  return out;
}

// TODO: Hier deine echte XLSX-Erzeugung aufrufen (exceljs / xlsx / eigener Writer).
// Aktuell: Platzhalter, der einfach die Vorlage zurückgibt.
async function buildXlsxFromTemplate(templateBase64, rows, qa, prefix, fallzahl) {
  // IMPLEMENTIERUNG:
  // - Vorlage (templateBase64) als Workbook öffnen
  // - Sheet lokalisieren
  // - rows (Array<Object mit raw-Spalten>) Zeile für Zeile schreiben
  // - qa ggf. in ein 2. Sheet schreiben
  // - Workbook als Base64 serialisieren → return
  return templateBase64;
}

export default async function handler(req, res) {
  try {
    if (req.method !== 'POST') {
      return res.status(405).json({ error: 'Method not allowed' });
    }

    const {
      templateBase64 = '',
      rows = [],
      qa = {},
      prefix = '',
      fallzahl = 0,
      codebook = null,
    } = req.body || {};

    if (!templateBase64) {
      return res.status(400).json({ error: 'templateBase64 fehlt' });
    }
    if (!Array.isArray(rows) || !rows.length) {
      return res.status(400).json({ error: 'rows fehlt/leer' });
    }
    if (!codebook) {
      return res.status(400).json({ error: 'codebook fehlt' });
    }

    // 1) Mapping pro Template aufbauen
    const mapping = buildColumnMapping(codebook, rows[0]);

    // 2) Persona-Rows → Template-Rows (Keys = raw-Spalten)
    const mappedRows = rows.map(r => buildTemplateRow(r, mapping));

    // 3) Optional: Preview für Debugging (?preview=1)
    if (String(req.query?.preview ?? "") === "1") {
      return res.status(200).json({
        preview: {
          columns: [...mapping.keys()],
          firstRow: mappedRows[0] || null,
          count: mappedRows.length
        }
      });
    }

    // 4) XLSX bauen (hier deinen Writer einhängen)
    const fileBase64 = await buildXlsxFromTemplate(
      templateBase64,
      mappedRows,
      qa,
      prefix,
      fallzahl
    );

    return res.status(200).json({ file: fileBase64 });
  } catch (err) {
    console.error("export-xlsx error:", err);
    return res.status(500).json({
      error: 'Server error',
      details: String(err?.message || err)
    });
  }
}
