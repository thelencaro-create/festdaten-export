// api/docx-to-html.js
import formidable from "formidable";
import fs from "fs";
import mammoth from "mammoth";

export const config = {
  api: {
    bodyParser: false, // WICHTIG: multipart selbst parsen
  },
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  // multipart/form-data Parser
  const form = formidable({
    multiples: false,
    keepExtensions: false,
  });

  try {
    const { files } = await new Promise((resolve, reject) => {
      form.parse(req, (err, fields, files) => {
        if (err) reject(err);
        else resolve({ fields, files });
      });
    });

    // n8n sendet das File im Field 'file'
    const uploadedFile = files.file;

    if (!uploadedFile) {
      return res.status(400).json({ error: "No DOCX file provided" });
    }

    // formidable liefert file.filepath = temp Pfad
    const buffer = fs.readFileSync(uploadedFile.filepath);

    // DOCX → HTML
    const result = await mammoth.convertToHtml({ buffer });

    return res.status(200).json({
      html: result.value,
      messages: result.messages ?? [],
    });
  } catch (err) {
    return res.status(500).json({
      error: "Upload or conversion failed",
      details: err.message,
    });
  }
}
