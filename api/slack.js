export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  const { endpoint, ...params } = req.query;
  const token = req.headers["x-slack-token"];
  if (!token) return res.status(401).json({ error: "No token" });

  const qs = new URLSearchParams(params).toString();
  const url = `https://slack.com/api/${endpoint}${qs ? "?" + qs : ""}`;
  const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const data = await response.json();
  res.status(200).json(data);
}
