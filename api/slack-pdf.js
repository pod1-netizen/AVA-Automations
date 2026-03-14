import chromium from "@sparticuz/chromium";
import puppeteer from "puppeteer-core";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();

  const chunks = [];
  req.on("data", chunk => chunks.push(chunk));
  await new Promise(r => req.on("end", r));
  const html = Buffer.concat(chunks).toString("utf-8");

  const browser = await puppeteer.launch({
    args: chromium.args,
    defaultViewport: chromium.defaultViewport,
    executablePath: await chromium.executablePath(),
    headless: chromium.headless,
  });

  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: "networkidle0" });
  const pdf = await page.pdf({ format: "A4", printBackground: true });
  await browser.close();

  res.setHeader("Content-Type", "application/pdf");
  res.status(200).send(pdf);
}
