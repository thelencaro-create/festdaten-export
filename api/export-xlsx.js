// api/export-xlsx.js
export default async function handler(req, res) {
  try {
    if (req.method !== 'POST') {
      return res.status(405).json({ error: 'Method not allowed' });
    }

    // In Vercel-Node-Funktionen ist req.body bei Content-Type: application/json bereits geparst.
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

    // TODO: Hier XLSX generieren. Für den Smoke-Test geben wir es pass-through zurück.
    const fileBase64 = templateBase64;

    return res.status(200).json({ file: fileBase64 });
  } catch (err) {
    return res.status(500).json({
      error: 'Server error',
      details: String(err?.message || err),
    });
  }
}
