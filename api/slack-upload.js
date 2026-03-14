export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();

  const { url } = req.query;
  if (!url) return res.status(400).json({ error: "No upload URL" });

  const chunks = [];
  req.on("data", chunk => chunks.push(chunk));
  await new Promise(r => req.on("end", r));
  const body = Buffer.concat(chunks);

  const response = await fetch(decodeURIComponent(url), {
    method: "POST",
    headers: { "Content-Type": req.headers["content-type"] || "text/html" },
    body
  });

  res.status(response.ok ? 200 : 500).json({ ok: response.ok });
}
