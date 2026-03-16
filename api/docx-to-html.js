// api/docx-to-html.js
import formidable from "formidable";
import fs from "fs";
import mammoth from "mammoth";

// JSON-Body einlesen (für die JSON-Variante)
async function readRawBody(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  return Buffer.concat(chunks).toString("utf8");
}

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    const ctype = String(req.headers["content-type"] || "").toLowerCase();
    let buffer = null;

    // A) multipart/form-data (n8n "n8n Binary File")
    if (ctype.includes("multipart/form-data")) {
      const form = formidable({ multiples: false });
      const { files } = await new Promise((resolve, reject) => {
        form.parse(req, (err, fields, files) => (err ? reject(err) : resolve({ fields, files })));
      });

      const f = files?.file;
      const fileObj = Array.isArray(f) ? f[0] : f;
      if (!fileObj?.filepath) {
        return res.status(400).json({ error: "No DOCX file received (multipart)" });
      }
      buffer = fs.readFileSync(fileObj.filepath);
    }
    // B) application/json ({"fileBase64": "...", optional data:URL})
    else if (ctype.includes("application/json")) {
      const raw = await readRawBody(req);
      let body = {};
      try { body = JSON.parse(raw || "{}"); } catch {}
      const fileBase64 = String(body.fileBase64 || "").trim();
      if (!fileBase64) {
        return res.status(400).json({ error: "No DOCX file received (json)" });
      }
      const clean = fileBase64.includes("base64,") ? fileBase64.split("base64,").pop() : fileBase64;
      buffer = Buffer.from(clean, "base64");
    }

    if (!buffer) {
      return res.status(415).json({ error: "Unsupported media type. Use multipart/form-data or JSON with fileBase64." });
    }

    // DOCX → HTML
    const result = await mammoth.convertToHtml({ buffer });
    return res.status(200).json({ html: result.value, messages: result.messages ?? [] });
  } catch (err) {
    console.error("docx-to-html error:", err);
    return res.status(500).json({ error: "Upload or conversion failed", details: String(err?.message ?? err) });
  }
}
