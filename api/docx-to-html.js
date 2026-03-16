// api/docx-to-html.js
// Vercel Serverless Function: DOCX (multipart oder base64) -> { html }
// Requires: npm i mammoth
import mammoth from 'mammoth';

export const config = {
  runtime: 'nodejs18.x', // oder höher
  regions: ['fra1'],     // optional: nahe EU
  memory: 512,           // optional
  maxDuration: 30,       // optional
};

function fromBase64(b64) {
  const s = String(b64 ?? '').trim();
  const clean = s.includes('base64,') ? s.split('base64,').pop() : s;
  return Buffer.from(clean, 'base64');
}

export default async function handler(req, res) {
  try {
    if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

    let buffer = null;

    // 1) Multipart-Upload via n8n HTTP Request (Send Binary Data)
    if (req.files?.file?.[0]?.buffer) {
      buffer = req.files.file[0].buffer;
    }
    // 2) JSON { fileBase64: "..." }
    else if (req.body?.fileBase64) {
      buffer = fromBase64(req.body.fileBase64);
    }

    if (!buffer?.length) return res.status(400).json({ error: 'No DOCX provided' });

    const result = await mammoth.convertToHtml({ buffer });
    const html = String(result.value ?? '').trim();

    return res.status(200).json({ html, messages: result.messages ?? [] });
  } catch (err) {
    console.error('docx-to-html error:', err);
    return res.status(500).json({ error: String(err?.message ?? err) });
  }
}
