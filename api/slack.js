export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, x-slack-token");
  if (req.method === "OPTIONS") return res.status(200).end();

  const token = process.env.VITE_SLACK_TOKEN || req.headers["x-slack-token"];
  if (!token) return res.status(401).json({ ok: false, error: "No Slack token configured" });

  const { endpoint, ...params } = req.query;
  if (!endpoint) return res.status(400).json({ ok: false, error: "No endpoint specified" });

  try {
    let response;
    if (req.method === "POST") {
      const contentType = req.headers["content-type"] || "";
      if (contentType.includes("application/json")) {
        response = await fetch(`https://slack.com/api/${endpoint}`, {
          method: "POST",
          headers: {
            "Authorization": `Bearer ${token}`,
            "Content-Type": "application/json"
          },
          body: JSON.stringify(req.body)
        });
      } else {
        const qs = new URLSearchParams(params).toString();
        const chunks = [];
        req.on("data", chunk => chunks.push(chunk));
        await new Promise(r => req.on("end", r));
        const rawBody = Buffer.concat(chunks);
        response = await fetch(`https://slack.com/api/${endpoint}${qs ? "?" + qs : ""}`, {
          method: "POST",
          headers: {
            "Authorization": `Bearer ${token}`,
            "Content-Type": contentType
          },
          body: rawBody
        });
      }
    } else {
      const qs = new URLSearchParams(params).toString();
      response = await fetch(`https://slack.com/api/${endpoint}${qs ? "?" + qs : ""}`, {
        headers: { "Authorization": `Bearer ${token}` }
      });
    }
    const data = await response.json();
    res.status(200).json(data);
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
}
