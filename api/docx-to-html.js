import mammoth from "mammoth";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  let buffer = null;

  // 1) Multipart (Send Binary Data aus n8n)
  if (req.files?.file?.[0]?.buffer) {
    buffer = req.files.file[0].buffer;
  }

  // 2) JSON body mit Base64
  if (!buffer && req.body?.fileBase64) {
    buffer = Buffer.from(req.body.fileBase64, "base64");
  }

  if (!buffer) {
    return res.status(400).json({ error: "No DOCX file provided" });
  }

  const result = await mammoth.convertToHtml({ buffer });
  return res.status(200).json({ html: result.value });
}
