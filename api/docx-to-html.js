// api/docx-to-html.js
import mammoth from "mammoth";

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    const { fileBase64 } = req.body;

    if (!fileBase64) {
      return res.status(400).json({ error: "No DOCX file provided" });
    }

    const clean = fileBase64.replace(/^data:.*;base64,/, "");
    const buffer = Buffer.from(clean, "base64");

    const result = await mammoth.convertToHtml({ buffer });

    res.status(200).json({ html: result.value });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
