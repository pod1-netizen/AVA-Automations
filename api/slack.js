export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, x-slack-token");
  if (req.method === "OPTIONS") return res.status(200).end();

  const token = process.env.VITE_SLACK_TOKEN || req.headers["x-slack-token"];
  if (!token) return res.status(401).json({ ok: false, error: "No Slack token configured" });

  const { endpoint, ...params } = req.query;
  if (!endpoint) return res.status(400).json({ ok: false, error: "No endpoint specified" });

  const qs = new URLSearchParams(params).toString();
  const url = `https://slack.com/api/${endpoint}${qs ? "?" + qs : ""}`;

  try {
    const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    const data = await response.json();
    res.status(200).json(data);
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
}
