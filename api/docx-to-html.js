// api/docx-to-html.js

import { IncomingForm } from "formidable";
import fs from "fs";
import mammoth from "mammoth";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }
  
  const form = new IncomingForm({ multiples: false });

  form.parse(req, async (err, fields, files) => {
    try {
      if (err) {
        res.status(500).json({ error: "Failed to parse" });
        return;
      }

      const uploaded = files.file;
      if (!uploaded) {
        res.status(400).json({ error: "No DOCX file provided" });
        return;
      }

      const buffer = fs.readFileSync(uploaded.filepath);
      const result = await mammoth.convertToHtml({ buffer });
      res.status(200).json({ html: result.value });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
};
