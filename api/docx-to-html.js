// api/docx-to-html.js

export const config = {
  runtime: "nodejs18.x"   // <<< WICHTIG: verhindert EDGE-Runtime!
};

import formidable from "formidable";
import fs from "fs";
import mammoth from "mammoth";

// Raw JSON Body Reader (für JSON Uploads)
async function readRawBody(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  return Buffer.concat(chunks).toString("utf8");
}

export const api = {
  bodyParser: false,
};

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    const contentType = req.headers["content-type"] || "";
    let buffer = null;

    // ========== VARIANTE A: multipart/form-data (n8n Binary File) ==========
    if (contentType.includes("multipart/form-data")) {
      const form = formidable({ multiples: false });

      const { files } = await new Promise((resolve, reject) => {
        form.parse(req, (err, fields, files) => {
          if (err) reject(err);
          else resolve({ fields, files });
        });
      });

      const uploaded = files.file;
      const fileObj = Array.isArray(uploaded) ? uploaded[0] : uploaded;

      if (!fileObj?.filepath) {
        return res.status(400).json({ error: "No DOCX file received (multipart)" });
      }

      buffer = fs.readFileSync(fileObj.filepath);
    }

    // ========== VARIANTE B: JSON Upload (fileBase64) ==========
    else if (contentType.includes("application/json")) {
      const raw = await readRawBody(req);
      const body = JSON.parse(raw || "{}");

      if (!body.fileBase64) {
        return res.status(400).json({ error: "No DOCX file received (json)" });
      }

      const base64 = body.fileBase64.replace(/^data:.*;base64,/, "");
      buffer = Buffer.from(base64, "base64");
    }

    if (!buffer) {
      return res.status(400).json({
        error: "Unsupported content type (send multipart/form-data or JSON with fileBase64)."
      });
    }

    // DOCX → HTML
    const result = await mammoth.convertToHtml({ buffer });

    return res.status(200).json({
      html: result.value,
      messages: result.messages ?? [],
    });

  } catch (err) {
    console.error("docx-to-html error:", err);
    return res.status(500).json({
      error: "Server crashed",
      details: err.message,
    });
  }
}
