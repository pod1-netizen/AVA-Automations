import { useState, useEffect, useCallback, useMemo, useRef } from "react";

const NAVY = "#07111f";
const NAVY2 = "#0c1d35";
const GOLD = "#f0b429";
const TEAL = "#0d9488";
const GREEN = "#22c55e";

const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxbxOCptDyXw0x5dezcYq_a7u3e5rC-gnB2TOKB1QCBD0crE0f3WZx4Kx_b5MT69Od4/exec";

// ── Static fallback data (used instantly while live data loads) ───────────────
const FALLBACK_SHEET_DATA = {
  "Bryan Cruz": {
    tasks: [
      { name: "Reel for Bryan & Rey for 146 Westlawn", date: "2026-02-02", hours: 1.0, va: "DJ", cat: "Content & Marketing" },
      { name: "Reel Revision for Split-level Doelger Home", date: "2026-02-03", hours: 0.75, va: "DJ", cat: "Content & Marketing" },
      { name: "Revised Westlake Property Tour Video", date: "2026-02-05", hours: 0.92, va: "DJ", cat: "Content & Marketing" },
      { name: "173 Morton Virtual Tour", date: "2026-02-07", hours: 1.0, va: "Echo", cat: "Content & Marketing" },
      { name: "Revise 137 Morton Reel", date: "2026-02-10", hours: 0.5, va: "Echo", cat: "Content & Marketing" },
      { name: "Made a Property Reel", date: "2026-02-12", hours: 1.0, va: "Echo", cat: "Content & Marketing" },
      { name: "Edited Final Drafts of Doelger Homes Educational Videos", date: "2026-02-17", hours: 0.75, va: "DJ", cat: "Content & Marketing" },
      { name: "Researched TikTok Content (Educational, POV, Trend, Lifestyle)", date: "2026-02-20", hours: 0.5, va: "DJ", cat: "Administrative" },
      { name: "Sync Meeting", date: "2026-02-24", hours: 0.25, va: "DJ", cat: "Management" },
      { name: "Created AVA's KPI Report for Bryan", date: "2026-02-26", hours: 0.5, va: "DJ", cat: "Administrative" },
      { name: "New listing photos edit - 220 Elm St", date: "2026-03-02", hours: 1.25, va: "DJ", cat: "Content & Marketing" },
      { name: "Social media caption batch for March", date: "2026-03-04", hours: 0.75, va: "Echo", cat: "Content & Marketing" },
      { name: "Reel revision - 220 Elm St walkthrough", date: "2026-03-06", hours: 0.5, va: "DJ", cat: "Content & Marketing" },
    ],
    wins: []
  },
  "Leo": {
    tasks: [
      { name: "Morning Recurring Task FUB + Google Calendar", date: "2026-02-01", hours: 0.5, va: "Clyde", cat: "Administrative" },
      { name: "Notion KPI dashboard setup", date: "2026-02-03", hours: 1.5, va: "Clyde", cat: "Administrative" },
      { name: "Requested disclosures for 461 Southgate Ave", date: "2026-02-04", hours: 1.0, va: "Clyde", cat: "Administrative" },
      { name: "AVA's Ad on Social Media Channels", date: "2026-02-06", hours: 0.67, va: "DJ", cat: "Content & Marketing" },
      { name: "Ava Ad - 4 samples", date: "2026-02-07", hours: 1.0, va: "Cly", cat: "Design & Editing" },
      { name: "Ava PubMat", date: "2026-02-10", hours: 0.5, va: "Cly", cat: "Design & Editing" },
      { name: "755 Mountain View Dr Property Post", date: "2026-02-17", hours: 1.0, va: "DJ", cat: "Content & Marketing" },
      { name: "1754 Technology Dr Property Post", date: "2026-02-19", hours: 0.42, va: "DJ", cat: "Administrative" },
      { name: "Called SAMCAR for assistant pricing", date: "2026-02-21", hours: 1.0, va: "Allen", cat: "Administrative" },
      { name: "FUB tasks & lead reminder", date: "2026-02-24", hours: 0.42, va: "Ellaine", cat: "Administrative" },
      { name: "FUB tasks & lead follow-up batch", date: "2026-02-25", hours: 1.0, va: "Tashia", cat: "Administrative" },
      { name: "Open house prep - 755 Mountain View Dr", date: "2026-03-01", hours: 0.5, va: "Clyde", cat: "Administrative" },
      { name: "Posted Just Sold graphic - 461 Southgate", date: "2026-03-03", hours: 0.75, va: "DJ", cat: "Content & Marketing" },
      { name: "CRM lead tagging cleanup", date: "2026-03-05", hours: 1.0, va: "Tashia", cat: "Administrative" },
    ],
    wins: [
      { va: "DJ", note: "Siiiicccckkkk! - Leo", date: "2026-02-17" },
      { va: "Cly", note: "Wow this is really clean! I like it - Leo", date: "2026-02-07" },
    ]
  },
  "Nick": {
    tasks: [
      { name: "Notion KPI dashboard setup", date: "2026-02-02", hours: 1.5, va: "Clyde", cat: "Administrative" },
      { name: "Reels: 238 Warwick St, 25 Ramsell St", date: "2026-02-04", hours: 0.59, va: "Tashia", cat: "Content & Marketing" },
      { name: "Open House Flyer Templates", date: "2026-02-05", hours: 1.0, va: "Tashia", cat: "Design & Editing" },
      { name: "Open House Flyer - 183 Victoria St. SF", date: "2026-02-07", hours: 0.67, va: "Tashia", cat: "Design & Editing" },
      { name: "Open House Flyer - 15-17 Harriet St", date: "2026-02-10", hours: 0.5, va: "Tashia", cat: "Design & Editing" },
      { name: "Market Report (272 new listings, 124 sold, 13 coming soon)", date: "2026-02-16", hours: 2.0, va: "Echo", cat: "Document Prep" },
      { name: "Email to Kimmie - OH flyer for printing", date: "2026-02-18", hours: 0.17, va: "Tashia", cat: "Administrative" },
      { name: "Email Newsletter for Feb 25", date: "2026-02-20", hours: 1.5, va: "Clyde", cat: "Administrative" },
      { name: "QR code linktree & Curb hero inquiry", date: "2026-02-24", hours: 0.5, va: "Ellaine", cat: "Administrative" },
      { name: "Ran CMA for 29 Cove Ln", date: "2026-02-26", hours: 0.5, va: "Echo", cat: "Administrative" },
      { name: "Open House Flyer - 340 Lakeview", date: "2026-03-02", hours: 0.75, va: "Tashia", cat: "Design & Editing" },
      { name: "Instagram story batch for Nick listings", date: "2026-03-05", hours: 1.0, va: "Echo", cat: "Content & Marketing" },
    ],
    wins: []
  },
  "Joji": {
    tasks: [
      { name: "Joji's Real Estate Journey: Reel (initial)", date: "2026-02-01", hours: 8.0, va: "Echo", cat: "Content & Marketing" },
      { name: "Real Estate Journey Reel Revision x1", date: "2026-02-05", hours: 6.0, va: "Echo", cat: "Content & Marketing" },
      { name: "Real Estate Journey Reel Revision x2", date: "2026-02-08", hours: 6.0, va: "Echo", cat: "Content & Marketing" },
      { name: "Requested disclosures & rent rolls for 8 properties", date: "2026-02-11", hours: 3.5, va: "Ellaine", cat: "Administrative" },
      { name: "Organized disclosure sheet", date: "2026-02-13", hours: 1.0, va: "Ellaine", cat: "Document Prep" },
      { name: "Ran CMA for 3 properties", date: "2026-02-17", hours: 3.0, va: "Ellaine", cat: "Administrative" },
      { name: "Followed up on disclosures - 1279 Florida St", date: "2026-02-20", hours: 0.5, va: "Ellaine", cat: "Administrative" },
      { name: "Real Estate Journey Reel Revision x3", date: "2026-02-22", hours: 6.0, va: "Echo", cat: "Content & Marketing" },
      { name: "Joji Real Estate Journey (Pending Review)", date: "2026-03-01", hours: 2.5, va: "Echo", cat: "Content & Marketing" },
      { name: "Property comp research - 1455 Mission", date: "2026-03-04", hours: 1.5, va: "Ellaine", cat: "Administrative" },
    ],
    wins: []
  },
  "Ray Guardado & Fiona Santos": {
    tasks: [
      { name: "Property Reel 4184 Callan Blvd DC", date: "2026-02-02", hours: 1.0, va: "Joseph", cat: "Content & Marketing" },
      { name: "Posted 1314 Danberry Ln DC Property Reel", date: "2026-02-04", hours: 0.75, va: "Joseph", cat: "Content & Marketing" },
      { name: "Posted Property Reel - 2401 Ardee", date: "2026-02-06", hours: 0.75, va: "Aldwin", cat: "Content & Marketing" },
      { name: "Property Reel for 137 Morton Dr DC", date: "2026-02-10", hours: 1.25, va: "Joseph", cat: "Content & Marketing" },
      { name: "Posting - Property Reel - 783 Solana", date: "2026-02-12", hours: 0.5, va: "Aldwin", cat: "Platform Management" },
      { name: "Client Testimonial Carousel: Abbas Testimonial", date: "2026-02-17", hours: 1.5, va: "Joseph", cat: "Content & Marketing" },
      { name: "408 Cyrpress Property reel/thumbnail/caption", date: "2026-02-19", hours: 1.5, va: "Joseph", cat: "Content & Marketing" },
      { name: "2411 Catalpa property reel", date: "2026-02-22", hours: 1.25, va: "Aldwin", cat: "Content & Marketing" },
      { name: "Posting: Abbas Testimonial (FB, TT, IG, LI)", date: "2026-02-25", hours: 0.5, va: "Aldwin", cat: "Platform Management" },
      { name: "New listing reel - 520 Harbor View", date: "2026-03-03", hours: 1.0, va: "Joseph", cat: "Content & Marketing" },
      { name: "Posting schedule for March week 1", date: "2026-03-05", hours: 0.5, va: "Aldwin", cat: "Platform Management" },
    ],
    wins: [{ va: "Joseph", note: "408 Cyrpress reel approved by client", date: "2026-02-19" }]
  },
  "Jacky": {
    tasks: [
      { name: "Carousel + Did You Know - Published IG, FB, TT", date: "2026-02-03", hours: 0.5, va: "Aldwin", cat: "Platform Management" },
      { name: "Posting: How to Prepare for Homeownership in the Bay Area", date: "2026-02-10", hours: 0.33, va: "Aldwin", cat: "Platform Management" },
      { name: "Posting - Carousel: The Rise of Micro-Units in San Francisco", date: "2026-02-18", hours: 0.25, va: "Aldwin", cat: "Administrative" },
      { name: "Bay Area market tips carousel", date: "2026-03-02", hours: 0.5, va: "Aldwin", cat: "Platform Management" },
    ],
    wins: []
  },
  "NSP": {
    tasks: [
      { name: "Alzheimer's Disease & Dementia Care Seminar content", date: "2026-02-02", hours: 3.0, va: "Hannah", cat: "Administrative" },
      { name: "Revised caption & creatives, Updated Social Pilot", date: "2026-02-04", hours: 2.0, va: "Hannah", cat: "Administrative" },
      { name: "Uploaded content to Social Pilot for approval", date: "2026-02-06", hours: 1.0, va: "Hannah", cat: "Administrative" },
      { name: "Content calendar and script for March", date: "2026-02-10", hours: 2.0, va: "Hannah", cat: "Content & Marketing" },
      { name: "KPI Report prep", date: "2026-02-12", hours: 1.0, va: "Echo", cat: "Administrative" },
      { name: "Reel: Client Home Visit Revision", date: "2026-02-14", hours: 3.0, va: "Echo", cat: "Content & Marketing" },
      { name: "Sit Rep Weekly Report", date: "2026-02-17", hours: 0.92, va: "Aldwin", cat: "Administrative" },
      { name: "Outreach calls (34), Contracts sent (2), Referrals (3)", date: "2026-02-19", hours: 4.0, va: "Alex Castillo", cat: "Client Relations" },
      { name: "CRM Outreach (25 contacts), Contracts Sent (3)", date: "2026-02-21", hours: 4.0, va: "Alex Castillo", cat: "Client Relations" },
      { name: "Reel: Success story testimonial", date: "2026-02-24", hours: 2.0, va: "Echo", cat: "Content & Marketing" },
      { name: "Social Pilot scheduling for March wk1", date: "2026-03-01", hours: 1.5, va: "Hannah", cat: "Administrative" },
      { name: "CRM Outreach (30 contacts), follow-ups", date: "2026-03-03", hours: 3.5, va: "Alex Castillo", cat: "Client Relations" },
      { name: "Design & Editing: flyer revision for seminar", date: "2026-03-06", hours: 0.42, va: "Aldwin", cat: "Design & Editing" },
    ],
    wins: [{ va: "Alex Castillo", note: "Smooth outreach, fully compliant with NSP conditions", date: "2026-02-21" }]
  },
  "Tien Le": {
    tasks: [
      { name: "Followed up for revisions/approvals, sent revised recap video", date: "2026-02-16", hours: 1.0, va: "Ellaine", cat: "Design & Editing" },
      { name: "Created AVA's KPI Report for Tien Le", date: "2026-02-24", hours: 1.0, va: "DJ", cat: "Administrative" },
      { name: "Social post scheduling for March", date: "2026-03-02", hours: 0.75, va: "DJ", cat: "Administrative" },
    ],
    wins: []
  }
};

// ── Static client metadata (Slack channels, groups) ───────────────────────────
const CLIENT_META = [
  { display: "Bryan Cruz",                  slack: "#kinetic-bryan-cruz",                       group: "Kinetic" },
  { display: "Ed Barreto",                  slack: "#kinetic-ed-barreto",                       group: "Kinetic" },
  { display: "Ray Guardado & Fiona Santos", slack: "#kinetic-ray-guardado-and-fiona-santos",    group: "Kinetic" },
  { display: "Guillean Arradaza",           slack: "#kinetic-guillean-arradaza",                group: "Kinetic" },
  { display: "Tien Le",                     slack: "#kinetic-tien-le-private",                  group: "Kinetic", priv: true },
  { display: "Leo",                         slack: "#leo-team-jacky-nick-joji",                 group: "Leo Team" },
  { display: "Jacky",                       slack: "#leo-team-jacky-nick-joji",                 group: "Leo Team" },
  { display: "Nick",                        slack: "#leo-team-jacky-nick-joji",                 group: "Leo Team" },
  { display: "Joji",                        slack: "#leo-team-jacky-nick-joji",                 group: "Leo Team" },
  { display: "Chris Yanguas",               slack: "#client-chris-yanguas",                     group: "Other" },
  { display: "Joey Boy Colo",               slack: "#client-joey-boy-colo",                     group: "Other" },
  { display: "NSP",                         slack: "#nsp-general",                              group: "Other" },
  { display: "7Edu",                        slack: "#7edu-ava-general",                         group: "Other" },
];

// ── Parse live Apps Script JSON into SHEET_DATA shape ────────────────────────
function parseLiveData(rows) {
  const result = {};
  rows.forEach(r => {
    const clientName = r.clientName?.trim();
    if (!clientName) return;
    if (!result[clientName]) result[clientName] = { tasks: [], wins: [] };

    // Parse date — Apps Script returns "MM-dd-yyyy HH:mm" or column H formatted date
    let dateStr = "";
    const raw = r.date || r.timestamp || "";
    // Try to extract YYYY-MM-DD from various formats
    const isoMatch = raw.match(/(\d{4})-(\d{2})-(\d{2})/);
    const mdyMatch = raw.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (isoMatch) {
      dateStr = isoMatch[0];
    } else if (mdyMatch) {
      const [, m, d, y] = mdyMatch;
      const year = y.length === 2 ? "20" + y : y;
      dateStr = `${year}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;
    }

    if (r.taskName) {
      result[clientName].tasks.push({
        name: r.taskName,
        date: dateStr,
        hours: parseFloat(r.hours) || 0,
        va: r.vaName || "",
        cat: r.category || "Administrative",
      });
    }
    if (r.winNote && r.winNote.trim()) {
      result[clientName].wins.push({
        va: r.vaName || "",
        note: r.winNote,
        date: dateStr,
      });
    }
  });
  return result;
}


// ── CSV Import parser ─────────────────────────────────────────────────────────
function parseCSVData(csvText) {
  const lines = csvText.trim().split(/\r?\n/);
  if (lines.length < 2) return null;

  const header = lines[0].split(",").map(h => h.trim().toLowerCase().replace(/[^a-z0-9]/g, ""));
  const idx = (candidates) => {
    for (const c of candidates) {
      const i = header.findIndex(h => h.includes(c));
      if (i !== -1) return i;
    }
    return -1;
  };

  const iClient = idx(["clientname","client"]);
  const iVA     = idx(["vaname","va"]);
  const iHours  = idx(["totalhours","hours"]);
  const iCat    = idx(["taskcategory","category","cat"]);
  const iTask   = idx(["taskscompleted","taskname","task"]);
  const iWin    = idx(["specialwins","winnote","win","note"]);
  const iDate   = idx(["formatteddate","date","timestamp"]);

  function parseRow(line) {
    const cols = []; let cur = "", inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') { inQ = !inQ; }
      else if (ch === "," && !inQ) { cols.push(cur.trim()); cur = ""; }
      else { cur += ch; }
    }
    cols.push(cur.trim()); return cols;
  }

  const result = {};
  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const cols = parseRow(lines[i]);
    const clientName = iClient >= 0 ? cols[iClient]?.replace(/^"|"$/g,"").trim() : "";
    if (!clientName) continue;
    if (!result[clientName]) result[clientName] = { tasks: [], wins: [] };

    const rawDate = iDate >= 0 ? cols[iDate]?.replace(/^"|"$/g,"").trim() : "";
    let dateStr = "";
    const isoMatch = rawDate.match(/(\d{4})-(\d{2})-(\d{2})/);
    const mdyMatch = rawDate.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (isoMatch) { dateStr = isoMatch[0]; }
    else if (mdyMatch) {
      const [,m,d,y] = mdyMatch;
      dateStr = `${y.length===2?"20"+y:y}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;
    }

    const taskName = iTask  >= 0 ? cols[iTask]?.replace(/^"|"$/g,"").trim()  : "";
    const vaName   = iVA    >= 0 ? cols[iVA]?.replace(/^"|"$/g,"").trim()    : "";
    const hours    = iHours >= 0 ? parseFloat(cols[iHours]) || 0             : 0;
    const cat      = iCat   >= 0 ? cols[iCat]?.replace(/^"|"$/g,"").trim()   : "Administrative";
    const winNote  = iWin   >= 0 ? cols[iWin]?.replace(/^"|"$/g,"").trim()   : "";

    if (taskName) result[clientName].tasks.push({ name: taskName, date: dateStr, hours, va: vaName, cat });
    if (winNote)  result[clientName].wins.push({ va: vaName, note: winNote, date: dateStr });
  }
  return Object.keys(result).length > 0 ? result : null;
}

// ── Build CLIENTS list from sheetData + CLIENT_META ───────────────────────────
function buildClients(sheetData) {
  const liveKeys = new Set(Object.keys(sheetData));
  return CLIENT_META.map(m => ({
    ...m,
    dataKey: liveKeys.has(m.display) ? m.display : null,
  }));
}

function getMonthPeriods() {
  const periods = [];
  const now = new Date();
  for (let m = 3; m >= 0; m--) {
    const d = new Date(now.getFullYear(), now.getMonth() - m, 1);
    const year = d.getFullYear(), month = d.getMonth();
    const monthName = d.toLocaleDateString("en-US", { month: "long", year: "numeric" });
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const fmt = dt => dt.toLocaleDateString("en-US", { month: "short", day: "numeric" });
    const p1s = new Date(year, month, 1), p1e = new Date(year, month, 14);
    const p2s = new Date(year, month, 15), p2e = new Date(year, month, daysInMonth);
    const fullS = new Date(year, month, 1), fullE = new Date(year, month, daysInMonth);
    periods.push({
      month: monthName,
      slots: [
        { label: `${monthName.split(" ")[0]} 1-14, ${year}`, start: p1s.toISOString().split("T")[0], end: p1e.toISOString().split("T")[0] },
        { label: `${monthName.split(" ")[0]} 15-${daysInMonth}, ${year}`, start: p2s.toISOString().split("T")[0], end: p2e.toISOString().split("T")[0] },
        { label: `${monthName.split(" ")[0]} 1-${daysInMonth}, ${year}`, start: fullS.toISOString().split("T")[0], end: fullE.toISOString().split("T")[0] },
      ]
    });
  }
  return periods;
}

function groupClients(clients) {
  const g = {};
  clients.forEach(c => { if (!g[c.group]) g[c.group] = []; g[c.group].push(c); });
  return g;
}

function filterData(rawData, period) {
  if (!rawData || !period) return null;
  const filtered = rawData.tasks.filter(t => t.date >= period.start && t.date <= period.end);
  if (!filtered.length) return null;
  const totalHours = Math.round(filtered.reduce((s, t) => s + t.hours, 0) * 100) / 100;
  const vas = {}, cats = {}, byCategory = {}, vaDetails = {};
  filtered.forEach(t => {
    vas[t.va] = Math.round(((vas[t.va] || 0) + t.hours) * 100) / 100;
    cats[t.cat] = Math.round(((cats[t.cat] || 0) + t.hours) * 100) / 100;
    if (!byCategory[t.cat]) byCategory[t.cat] = [];
    byCategory[t.cat].push(t.name);
    // Per-VA detail
    if (!vaDetails[t.va]) vaDetails[t.va] = { hours: 0, tasks: [], cats: {} };
    vaDetails[t.va].hours = Math.round((vaDetails[t.va].hours + t.hours) * 100) / 100;
    vaDetails[t.va].tasks.push({ name: t.name, hours: t.hours, cat: t.cat, date: t.date });
    vaDetails[t.va].cats[t.cat] = Math.round(((vaDetails[t.va].cats[t.cat] || 0) + t.hours) * 100) / 100;
  });
  const wins = (rawData.wins || []).filter(w => w.date >= period.start && w.date <= period.end);
  return { totalHours, taskCount: filtered.length, vas, cats, byCategory, wins, vaDetails };
}

// ── Prompt — matches Bryan Cruz PDF structure exactly ─────────────────────────
function buildPrompt(client, period, data) {
  const vaList = data ? Object.entries(data.vas).map(([v, h]) => `${v} (${h}h)`).join(", ") : "N/A";
  const catList = data ? Object.entries(data.cats).map(([c, h]) => `${c}: ${h}h`).join(", ") : "N/A";
  const tasksByCat = data ? JSON.stringify(data.byCategory, null, 2) : "{}";
  const winList = data?.wins?.length ? data.wins.map(w => `${w.va}: ${w.note}`).join("; ") : "none";
  const totalH = data?.totalHours || 0;
  const selfH = Math.round(totalH * 2.5 * 100) / 100;
  const savedH = Math.round((selfH - totalH) * 100) / 100;
  const oppLow = (selfH * 100).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const oppHigh = (selfH * 150).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const costAvoid = (selfH * 75).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  return `You are generating a structured KPI report JSON for Ava Virtual Agents Inc.

Client: ${client.display}
Period: ${period.label}
Slack: ${client.slack}
Total hours logged: ${totalH}h across ${data?.taskCount || 0} tasks
VAs active: ${vaList}
Hours by category: ${catList}
Tasks by category (use these EXACTLY): ${tasksByCat}
Client wins/feedback: ${winList}

Return ONLY a raw JSON object. No markdown. No explanation. No code fences.

Match this EXACT schema used by the Bryan Cruz KPI report:

{
  "client": "${client.display}",
  "period": "${period.label}",
  "slack": "${client.slack}",
  "totals": {
    "submitted": ${data?.taskCount || 0},
    "completed": ${data?.taskCount || 0},
    "completion_rate": "91%",
    "avg_turnaround": "< 24 Hours",
    "in_progress": 0,
    "overdue": 0,
    "total_hours": ${totalH}
  },
  "metric_statuses": {
    "submitted": "-",
    "completed": "High",
    "completion_rate": "On Track",
    "avg_turnaround": "Excellent",
    "in_progress": "Minimal",
    "overdue": "Monitor",
    "total_hours": "Comprehensive Service"
  },
  "tasks_data": {
    "categories": { "<use exact cat names from input>": <hours as number> },
    "by_category": { "<cat>": ["exact task names from input"] }
  },
  "quality": {
    "first_time_rate": "89%",
    "revision_rate": "11%",
    "approval_rate": "100%"
  },
  "communication": {
    "avg_response": "< 15 Mins",
    "daily_updates": "100%",
    "slack_engagement": "Active"
  },
  "time_saved": {
    "total_ava_hours": ${totalH},
    "total_self_hours": ${selfH},
    "total_saved_hours": ${savedH},
    "breakdown": [
      { "category": "<cat>", "ava_hours": <number>, "self_hours": <ava*2.5>, "hours_saved": <self-ava> }
    ]
  },
  "value": {
    "opp_low": "${oppLow}",
    "opp_high": "${oppHigh}",
    "cost_avoidance": "${costAvoid}",
    "total_low": "${oppLow}",
    "total_high": "${oppHigh}",
    "rate_low": 100,
    "rate_high": 150
  },
  "recap": {
    "updated": "${new Date().toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" })}",
    "overall_performance": "[2-3 sentences referencing specific tasks completed, VAs by name, and any notable outcomes. Energetic, professional tone like the Bryan Cruz report.]",
    "roi_impact": "[1-2 sentences about hours reclaimed, pipeline impact, and financial value vs. service cost.]",
    "market_positioning": "[1-2 sentences about brand/digital strategy built this sprint and scalability.]"
  },
  "recognitions": [
    { "number": 1, "name": "<VA name>", "note": "<specific achievement, concise>" },
    { "number": 2, "name": "<VA name>", "note": "<specific achievement>" }
  ],
  "monthly_balance": {
    "cap": 300,
    "used": ${totalH},
    "remaining": ${300 - totalH}
  }
}

RULES:
- Use EXACT task names from input. Do not invent tasks.
- self_hours = ava_hours * 2.5 for every row
- hours_saved = self_hours - ava_hours
- Include one recognition entry per active VA with real task references
- overall_performance, roi_impact, market_positioning must be specific to this client and their tasks
- Do not add categories that don't exist in the input data`;
}

// ── PDF HTML template — matches Bryan Cruz report exactly ─────────────────────
function generatePDFHTML(d, client, period) {
  const now = new Date().toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });
  const T = d.totals || {};
  const MS = d.metric_statuses || {};
  const TD = d.tasks_data || {};
  const Q = d.quality || {};
  const C = d.communication || {};
  const TS = d.time_saved || {};
  const V = d.value || {};
  const R = d.recap || {};
  const recs = d.recognitions || [];
  const MB = d.monthly_balance || {};

  // Status badge color
  const statusColor = s => {
    const map = { "-": "#94a3b8", "High": "#22c55e", "On Track": "#22c55e", "Excellent": "#22c55e", "Minimal": "#22c55e", "Monitor": "#f59e0b", "Comprehensive Service": "#0d9488" };
    return map[s] || "#94a3b8";
  };

  const metricsRows = [
    ["Tasks Submitted", T.submitted, MS.submitted],
    ["Tasks Completed", T.completed, MS.completed],
    ["Task Completion Rate", T.completion_rate, MS.completion_rate],
    ["Average Task Turnaround Time", T.avg_turnaround, MS.avg_turnaround],
    ["Tasks in Progress", T.in_progress, MS.in_progress],
    ["Overdue Tasks", T.overdue, MS.overdue],
    ["Total Work Investment\n(Including Management)", `${T.total_hours} Hours`, MS.total_hours],
  ].map(([metric, value, status]) => `
    <tr>
      <td class="metric-cell">${metric.replace("\n", "<br>")}</td>
      <td class="value-cell">${value}</td>
      <td class="status-cell"><span class="status-badge" style="color:${statusColor(status)}">${status}</span></td>
    </tr>`).join("");

  // Work investment breakdown bullet
  const catBreakdown = Object.entries(TD.categories || {})
    .map(([c, h]) => `${h} hrs ${c}`)
    .join(" + ");

  // Task breakdown by category
  const taskBreakdown = Object.entries(TD.by_category || {}).map(([cat, tasks]) => `
    <div class="cat-section">
      <div class="cat-heading">${cat}</div>
      ${tasks.map(t => `<div class="task-line">&#10004; ${t}</div>`).join("")}
    </div>`).join("");

  // Time saved rows
  const tsRows = (TS.breakdown || []).map(r => `
    <tr>
      <td>${r.category}</td>
      <td>${r.ava_hours} hrs.</td>
      <td>${r.self_hours} hrs.</td>
      <td>${r.hours_saved} hrs.</td>
    </tr>`).join("");

  // Recognitions — numbered list
  const recList = recs.map(r =>
    `<div class="rec-item">${r.number}) <strong>${r.name}:</strong> ${r.note}</div>`
  ).join("");

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>KPI Report - ${client.display} - ${period.label}</title>
<style>
/* ── Reset & Base ─────────────────── */
* { margin: 0; padding: 0; box-sizing: border-box; }
html, body {
  font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
  font-size: 9.5pt;
  color: #1a1a2e;
  background: white;
  width: 816px;
  margin: 0 auto;
}
@page { margin: 0; size: letter; }
@media print {
  html, body { width: auto; margin: 0; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .no-break { page-break-inside: avoid; }
}

/* ── Header ───────────────────────── */
.header {
  background: #07111f;
  position: relative;
  overflow: hidden;
  padding: 20px 36px 16px;
  min-height: 90px;
}
.header-circle-1 {
  position: absolute; top: -50px; right: -50px;
  width: 200px; height: 200px; border-radius: 50%;
  background: rgba(13,148,136,0.15);
  border: 40px solid rgba(13,148,136,0.08);
}
.header-circle-2 {
  position: absolute; top: 10px; right: 80px;
  width: 100px; height: 100px; border-radius: 50%;
  background: rgba(13,148,136,0.08);
}
.header-inner {
  display: flex; justify-content: space-between; align-items: center;
  position: relative; z-index: 1;
}
.logo-group { display: flex; align-items: center; gap: 12px; }
.logo-mark {
  width: 36px; height: 36px; border-radius: 8px;
  background: #0d9488;
  display: flex; align-items: center; justify-content: center;
  color: white; font-weight: 900; font-size: 18px; flex-shrink: 0;
}
.brand-name { color: white; font-size: 14pt; font-weight: 700; letter-spacing: 0.01em; }
.brand-tagline { color: #64748b; font-size: 7.5pt; margin-top: 2px; }
.report-title { color: #f0b429; font-size: 24pt; font-weight: 900; letter-spacing: 2px; }

/* ── Teal accent bar ──────────────── */
.accent-bar { background: #0d9488; height: 3px; }

/* ── Sub-header meta ──────────────── */
.meta-bar {
  background: #0c1d35;
  padding: 9px 36px;
  display: flex; gap: 0;
  border-bottom: 1px solid rgba(255,255,255,0.05);
}
.meta-item { flex: 1; }
.meta-item + .meta-item { border-left: 1px solid rgba(255,255,255,0.06); padding-left: 20px; }
.meta-label { color: #475569; font-size: 6.5pt; text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 2px; }
.meta-value { color: white; font-size: 9.5pt; font-weight: 700; }

/* ── Page content ─────────────────── */
.page-content { padding: 22px 36px 56px; }

/* ── Section heading ──────────────── */
.section-title {
  font-size: 10.5pt; font-weight: 800; color: #07111f;
  border-bottom: 2.5px solid #0d9488;
  padding-bottom: 4px; margin-bottom: 12px; margin-top: 20px;
  text-transform: uppercase; letter-spacing: 0.06em;
}
.section-title:first-child { margin-top: 0; }

/* ── Metrics table ────────────────── */
.metrics-table { width: 100%; border-collapse: collapse; font-size: 9pt; }
.metrics-table thead tr { background: #07111f; }
.metrics-table thead th {
  color: white; font-weight: 700; font-size: 8.5pt;
  padding: 8px 12px; text-align: left;
  border-right: 1px solid rgba(255,255,255,0.1);
}
.metrics-table thead th:last-child { border-right: none; }
.metrics-table tbody tr:nth-child(odd) td { background: #f8fafc; }
.metrics-table tbody tr:nth-child(even) td { background: white; }
.metrics-table tbody tr:last-child td { background: #e8f4f8; font-weight: 700; }
.metric-cell { padding: 8px 12px; border: 0.5px solid #e2e8f0; color: #1e293b; width: 55%; line-height: 1.4; }
.value-cell { padding: 8px 12px; border: 0.5px solid #e2e8f0; color: #1e293b; font-weight: 700; width: 20%; text-align: center; }
.status-cell { padding: 8px 12px; border: 0.5px solid #e2e8f0; width: 25%; }
.status-badge { font-weight: 700; font-size: 8.5pt; }

/* ── Breakdown bullet ─────────────── */
.breakdown-line {
  margin: 10px 0 6px; font-size: 8.5pt; color: #374151; line-height: 1.7;
  padding-left: 14px; position: relative;
}
.breakdown-line::before { content: "•"; position: absolute; left: 2px; }

/* ── Task breakdown ───────────────── */
.cat-section { margin-bottom: 12px; }
.cat-heading {
  font-weight: 700; font-size: 9pt; color: #07111f;
  text-decoration: underline; margin-bottom: 4px;
}
.task-line { font-size: 8.5pt; color: #374151; padding: 1.5px 0 1.5px 10px; line-height: 1.45; }

/* ── QC + Communication 2-col ─────── */
.two-col { display: flex; gap: 0; margin-top: 14px; }
.col-half { flex: 1; padding: 12px 16px; background: #f8fafc; border: 0.5px solid #e2e8f0; }
.col-half + .col-half { border-left: none; }
.col-heading { font-weight: 700; font-size: 9pt; color: #07111f; margin-bottom: 7px; }
.col-item { font-size: 8.5pt; color: #374151; padding: 2px 0; }
.col-item::before { content: "• "; }

/* ── Time saved table ─────────────── */
.ts-note { font-size: 7.5pt; color: #94a3b8; margin: 7px 0 10px; line-height: 1.55; font-style: italic; }
.data-table { width: 100%; border-collapse: collapse; font-size: 9pt; }
.data-table thead tr { background: #07111f; }
.data-table thead th {
  color: white; padding: 8px 12px; font-weight: 700; font-size: 8.5pt;
  border-right: 1px solid rgba(255,255,255,0.1);
  text-align: left;
}
.data-table thead th:not(:first-child) { text-align: center; }
.data-table thead th:last-child { border-right: none; }
.data-table tbody tr:nth-child(odd) td { background: white; }
.data-table tbody tr:nth-child(even) td { background: #f8fafc; }
.data-table tbody td { padding: 7px 12px; border: 0.5px solid #e2e8f0; color: #374151; }
.data-table tbody td:not(:first-child) { text-align: center; }
.data-table tfoot td {
  padding: 8px 12px; font-weight: 800; color: #07111f;
  background: #e0f2fe; border: 0.5px solid #bae6fd;
  text-align: center;
}
.data-table tfoot td:first-child { text-align: left; }

/* ── Value Analysis ───────────────── */
.value-block { margin: 4px 0; font-size: 9pt; color: #374151; line-height: 1.7; padding-left: 16px; }

/* ── KPI Recap ────────────────────── */
.recap-meta { font-size: 8pt; color: #64748b; margin-bottom: 14px; }
.recap-section { margin-bottom: 12px; }
.recap-label { font-weight: 800; font-size: 9.5pt; color: #07111f; margin-bottom: 3px; }
.recap-text { font-size: 8.5pt; color: #374151; line-height: 1.7; }

/* ── Special Recognition ──────────── */
.rec-intro { font-size: 8.5pt; color: #374151; margin-bottom: 7px; font-style: italic; }
.rec-item { font-size: 8.5pt; color: #374151; padding: 3px 0; line-height: 1.5; }

/* ── Monthly Balance ──────────────── */
.balance-list { margin-top: 6px; }
.balance-item { font-size: 8.5pt; color: #374151; padding: 2px 0; }
.balance-item::before { content: "• "; }

/* ── Footer ───────────────────────── */
.footer {
  position: fixed; bottom: 0; left: 0; right: 0;
  background: #07111f; border-top: 2px solid #0d9488;
  padding: 9px 36px; display: flex; justify-content: space-between; align-items: center;
}
.footer-left { color: #475569; font-size: 7pt; }
.footer-right { color: #475569; font-size: 7pt; }
.footer-logo { color: #0d9488; font-weight: 800; font-size: 8pt; }
</style>
</head>
<body>

<!-- HEADER -->
<div class="header">
  <div class="header-circle-1"></div>
  <div class="header-circle-2"></div>
  <div class="header-inner">
    <div class="logo-group">
      <div class="logo-mark">A</div>
      <div>
        <div class="brand-name">Ava</div>
        <div class="brand-tagline">Astro Virtual Agents Inc.</div>
      </div>
    </div>
    <div class="report-title">KPI REPORT</div>
  </div>
</div>
<div class="accent-bar"></div>
<div class="meta-bar">
  <div class="meta-item"><div class="meta-label">Prepared by</div><div class="meta-value">Ava</div></div>
  <div class="meta-item"><div class="meta-label">Prepared For</div><div class="meta-value">${client.display}</div></div>
  <div class="meta-item"><div class="meta-label">Date Period</div><div class="meta-value">${period.label}</div></div>
</div>

<div class="page-content">

  <!-- TASK & PRODUCTIVITY METRICS -->
  <div class="section-title no-break">Task &amp; Productivity Metrics</div>
  <table class="metrics-table no-break">
    <thead>
      <tr><th>Metric</th><th style="text-align:center">Value</th><th>Status</th></tr>
    </thead>
    <tbody>${metricsRows}</tbody>
  </table>

  <div class="breakdown-line">
    <strong>Work Investment Breakdown (${T.total_hours} Hours):</strong><br>
    ${catBreakdown}
  </div>

  <!-- TASK BREAKDOWN -->
  <div class="section-title" style="margin-top:18px">Task Breakdown</div>
  ${taskBreakdown}

  <!-- QUALITY + COMMUNICATION -->
  <div class="two-col no-break">
    <div class="col-half">
      <div class="col-heading">Quality Metrics</div>
      <div class="col-item">First-Time Completion Rate: ${Q.first_time_rate || "--"}</div>
      <div class="col-item">Revision Rate: ${Q.revision_rate || "--"}</div>
      <div class="col-item">Client Approval Rate: ${Q.approval_rate || "--"}</div>
    </div>
    <div class="col-half">
      <div class="col-heading">Communication</div>
      <div class="col-item">Avg Response Time: ${C.avg_response || "--"}</div>
      <div class="col-item">Daily Progress Updates: ${C.daily_updates || "--"}</div>
      <div class="col-item">Slack Engagement: ${C.slack_engagement || "--"}</div>
    </div>
  </div>

  <!-- TIME SAVED -->
  <div class="section-title" style="margin-top:20px">Time Saved for Client</div>
  <p style="font-size:8.5pt;font-weight:700;margin-bottom:4px">Total Client Time Savings: ${TS.total_saved_hours} Hours</p>
  <p class="ts-note">* Based on 2.5x multiplier methodology (Ava hours &times; 2.5 = Client self-work hours)</p>
  <table class="data-table no-break">
    <thead>
      <tr><th>Category</th><th>AVA Hours</th><th>Self Hours</th><th>Hours Saved</th></tr>
    </thead>
    <tbody>${tsRows}</tbody>
    <tfoot>
      <tr>
        <td>TOTAL</td>
        <td>${TS.total_ava_hours} hrs.</td>
        <td>${TS.total_self_hours} hrs.</td>
        <td>${TS.total_saved_hours} hrs.</td>
      </tr>
    </tfoot>
  </table>

  <!-- VALUE ANALYSIS -->
  <div class="section-title" style="margin-top:20px">Value Analysis</div>
  <div class="value-block">Opportunity Value: <strong>$${V.opp_low} &ndash; $${V.opp_high}</strong> (Based on $${V.rate_low}&ndash;$${V.rate_high}/hr saved)</div>
  <div class="value-block">Cost Avoidance: <strong>$${V.cost_avoidance}</strong> (Estimated savings vs. multi-agency freelancers)</div>
  <div class="value-block">Total Value Delivered: <strong>$${V.total_low} &ndash; $${V.total_high}</strong></div>

  <!-- KPI RECAP -->
  <div class="section-title" style="margin-top:20px">KPI Recap</div>
  <div class="recap-meta">Dashboard Last Updated: ${R.updated || now} | Status: <strong>Active &amp; Productive</strong></div>

  <div class="recap-section no-break">
    <div class="recap-label">Overall Performance:</div>
    <div class="recap-text">${R.overall_performance || "--"}</div>
  </div>
  <div class="recap-section no-break">
    <div class="recap-label">ROI Impact:</div>
    <div class="recap-text">${R.roi_impact || "--"}</div>
  </div>
  <div class="recap-section no-break">
    <div class="recap-label">Market Positioning Summary:</div>
    <div class="recap-text">${R.market_positioning || "--"}</div>
  </div>

  <!-- SPECIAL RECOGNITION -->
  <div class="section-title" style="margin-top:20px">Special Recognition</div>
  <div class="rec-intro">Standout achievements this period:</div>
  ${recList}

  <!-- MONTHLY BALANCE -->
  <div class="section-title" style="margin-top:20px">Monthly Balance</div>
  <div class="balance-list">
    <div class="balance-item">Monthly Cap: ${MB.cap} Hours</div>
    <div class="balance-item">Total Used: ${MB.used} Hours</div>
    <div class="balance-item">Remaining: ${MB.remaining} Hours</div>
  </div>

</div>

<!-- FOOTER -->
<div class="footer">
  <div class="footer-left"><span class="footer-logo">Ava</span> &nbsp;Astro Virtual Agents Inc. &nbsp;&#183;&nbsp; Confidential KPI Report</div>
  <div class="footer-right">${now}</div>
</div>

</body>
</html>`;
}

// ── BarRow ─────────────────────────────────────────────────────────────────────
function BarRow({ label, value, max, color }) {
  return (
    <div style={{ marginBottom: 8 }}>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 3 }}>
        <span style={{ color: "#94a3b8", maxWidth: "72%", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{label}</span>
        <span style={{ color: "#f1f5f9", fontWeight: 700 }}>{value}h</span>
      </div>
      <div style={{ height: 3, background: "rgba(255,255,255,0.05)", borderRadius: 2 }}>
        <div style={{ height: 3, borderRadius: 2, background: color, width: `${Math.min(100, (value / max) * 100)}%`, transition: "width 0.4s" }} />
      </div>
    </div>
  );
}



// ── VA Dashboard helpers ──────────────────────────────────────────────────────
const VA_MASTER_LIST = [
  "Aldwin", "Alex Castillo", "Allen", "Charm", "Clyde",
  "DJ", "Echo", "Ellaine", "Hannah",
  "Joseph", "Portia", "Tashia"
];

const VA_NAME_MAP = {
  "charm": "Charm", "Charn": "Charm",
  "Cly": "Clyde",
  "Echi": "Echo",
  "Alex Castilllo": "Alex Castillo",
  "ellaine": "Ellaine",
};

function normalizeVAName(name) {
  return VA_NAME_MAP[name] || name;
}

function getAllVAs(sheetData) {
  const fromSheet = new Set();
  Object.values(sheetData).forEach(cd => cd.tasks.forEach(t => { if (t.va) fromSheet.add(normalizeVAName(t.va)); }));
  const combined = new Set([...VA_MASTER_LIST, ...fromSheet]);
  ["Leo Morales", "Fiona"].forEach(c => combined.delete(c));
  return Array.from(combined).sort();
}

function getVAData(vaName, period, sheetData) {
  if (!vaName || !period) return null;
  const tasks = [], clientHours = {}, cats = {}, wins = [];
  Object.entries(sheetData).forEach(([clientName, cd]) => {
    const filtered = cd.tasks.filter(t => normalizeVAName(t.va) === vaName && t.date >= period.start && t.date <= period.end);
    filtered.forEach(t => {
      tasks.push({ ...t, client: clientName });
      clientHours[clientName] = Math.round(((clientHours[clientName] || 0) + t.hours) * 100) / 100;
      cats[t.cat] = Math.round(((cats[t.cat] || 0) + t.hours) * 100) / 100;
    });
    (cd.wins || []).filter(w => w.va === vaName && w.date >= period.start && w.date <= period.end)
      .forEach(w => wins.push({ ...w, client: clientName }));
  });
  if (!tasks.length) return null;
  const totalHours = Math.round(tasks.reduce((s, t) => s + t.hours, 0) * 100) / 100;
  return { totalHours, taskCount: tasks.length, tasks, clientHours, cats, wins };
}

const CAT_COMPLEXITY = {
  "Content & Marketing": 3.0, "Design & Editing": 3.5, "Document Prep": 2.5,
  "Client Relations": 2.0, "Administrative": 1.5, "Platform Management": 1.5, "Management": 1.0,
};
const CAT_HIDDEN_NOTE = {
  "Content & Marketing": "scripting, research, multiple revision rounds",
  "Design & Editing": "rendering, feedback loops, asset sourcing",
  "Document Prep": "data gathering, formatting, accuracy checks",
  "Client Relations": "follow-ups, CRM logging, prep",
  "Administrative": "cross-checking, coordination, tracking",
  "Platform Management": "scheduling, formatting, platform nuances",
  "Management": "direct effort",
};

function getProductivityRating(data) {
  const outputRate = data.taskCount / data.totalHours;
  const breadth = Object.keys(data.cats).length;
  const hiddenHours = Object.entries(data.cats).reduce((s, [cat, h]) => s + h * ((CAT_COMPLEXITY[cat] || 1.5) - 1), 0);
  if (outputRate >= 3 && breadth >= 3) return "High Output · Lean Hours";
  if (hiddenHours > data.totalHours * 1.5) return "High Complexity Output";
  if (outputRate >= 2) return "Strong Efficiency";
  if (breadth >= 4) return "Wide Coverage";
  return "Solid Contributor";
}

function VADashboard({ sheetData }) {
  const allVAs = getAllVAs(sheetData);
  const months = getMonthPeriods();
  const [selVA, setSelVA] = useState(null);
  const [selPeriod, setSelPeriod] = useState(null);
  const [openMonth, setOpenMonth] = useState(months[months.length - 1].month);
  const data = useMemo(() => getVAData(selVA, selPeriod, sheetData), [selVA, selPeriod]);
  const outputRate = data ? Math.round((data.taskCount / data.totalHours) * 10) / 10 : 0;
  const breadth = data ? Object.keys(data.cats).length : 0;
  const hiddenHours = data ? Math.round(Object.entries(data.cats).reduce((s, [cat, h]) => s + h * ((CAT_COMPLEXITY[cat] || 1.5) - 1), 0) * 10) / 10 : 0;
  const rating = data ? getProductivityRating(data) : "";

  return (
    <div style={{ display: "grid", gridTemplateColumns: "220px 1fr", minHeight: "100vh" }}>
      <div style={{ background: "#060e1b", borderRight: "1px solid rgba(255,255,255,0.06)", overflowY: "auto", position: "sticky", top: 0, height: "100vh" }}>
        <div style={{ padding: "14px 12px 8px" }}>
          <div style={{ fontSize: 9, color: "#475569", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 8, display: "flex", alignItems: "center", gap: 5 }}>
            <span style={{ background: "#8b5cf6", color: "#fff", width: 15, height: 15, borderRadius: "50%", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 8, fontWeight: 700 }}>1</span>
            Select VA
          </div>
          {allVAs.map(va => (
            <button key={va} onClick={() => setSelVA(va)}
              style={{ width: "100%", textAlign: "left", padding: "7px 9px", borderRadius: 6, marginBottom: 2, background: selVA === va ? "rgba(139,92,246,0.15)" : "transparent", border: selVA === va ? "1px solid #8b5cf6" : "1px solid transparent", color: selVA === va ? "#fff" : "#94a3b8", cursor: "pointer", fontSize: 12, fontWeight: selVA === va ? 600 : 400, display: "flex", alignItems: "center", gap: 8 }}>
              <div style={{ width: 26, height: 26, borderRadius: "50%", background: selVA === va ? "#8b5cf6" : "rgba(255,255,255,0.06)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 800, color: selVA === va ? "#fff" : "#475569", flexShrink: 0 }}>{va.charAt(0).toUpperCase()}</div>
              {va}
            </button>
          ))}
        </div>
        <div style={{ height: 1, background: "rgba(255,255,255,0.04)", margin: "4px 12px" }} />
        <div style={{ padding: "10px 12px 16px" }}>
          <div style={{ fontSize: 9, color: "#475569", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 9, display: "flex", alignItems: "center", gap: 5 }}>
            <span style={{ background: GOLD, color: NAVY, width: 15, height: 15, borderRadius: "50%", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 8, fontWeight: 700 }}>2</span>
            Period
          </div>
          {months.map(m => (
            <div key={m.month} style={{ marginBottom: 4 }}>
              <button onClick={() => setOpenMonth(openMonth === m.month ? null : m.month)}
                style={{ width: "100%", textAlign: "left", padding: "5px 9px", borderRadius: 6, background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.05)", color: "#64748b", cursor: "pointer", fontSize: 11, fontWeight: 600, display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 2 }}>
                <span>{m.month}</span><span style={{ fontSize: 9, color: "#334155" }}>{openMonth === m.month ? "▲" : "▼"}</span>
              </button>
              {openMonth === m.month && m.slots.map(p => {
                const active = selPeriod?.label === p.label;
                return (
                  <button key={p.label} onClick={() => setSelPeriod(p)}
                    style={{ width: "100%", textAlign: "left", padding: "6px 12px", borderRadius: 6, marginBottom: 2, background: active ? "rgba(240,180,41,0.1)" : "rgba(255,255,255,0.02)", border: active ? `1px solid ${GOLD}` : "1px solid transparent", color: active ? "#fff" : "#94a3b8", cursor: "pointer", fontSize: 11 }}>
                    {p.label}
                  </button>
                );
              })}
            </div>
          ))}
        </div>
      </div>

      <div style={{ padding: "28px 36px", overflowY: "auto" }}>
        {!selVA && <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "70vh", opacity: 0.4, gap: 12 }}><div style={{ fontSize: 48 }}>👤</div><div style={{ fontSize: 16, fontWeight: 700, color: "#94a3b8" }}>Select a VA + period</div></div>}
        {selVA && !selPeriod && <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "70vh", opacity: 0.5, gap: 10 }}><div style={{ fontSize: 36 }}>📅</div><div style={{ fontSize: 14, fontWeight: 700, color: "#94a3b8" }}>Now select a reporting period</div></div>}
        {selVA && selPeriod && !data && <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "70vh", opacity: 0.5, gap: 10 }}><div style={{ fontSize: 36 }}>📭</div><div style={{ fontSize: 14, fontWeight: 700, color: "#94a3b8" }}>No tasks found for {selVA} in this period</div></div>}
        {selVA && selPeriod && data && (
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            {/* Header */}
            <div style={{ background: `linear-gradient(135deg,${NAVY2},#1a2744)`, borderRadius: 14, padding: "20px 26px", borderBottom: "3px solid #8b5cf6" }}>
              <div style={{ fontSize: 9, color: "#8b5cf6", textTransform: "uppercase", letterSpacing: "0.14em", marginBottom: 4 }}>VA Performance Dashboard</div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                  <div style={{ width: 52, height: 52, borderRadius: "50%", background: "linear-gradient(135deg,#8b5cf6,#6d28d9)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, fontWeight: 900, color: "#fff" }}>{selVA.charAt(0).toUpperCase()}</div>
                  <div>
                    <div style={{ fontSize: 22, fontWeight: 800, color: "#fff" }}>{selVA}</div>
                    <div style={{ fontSize: 10, color: "#475569", marginTop: 2 }}>Virtual Assistant · {selPeriod.label}</div>
                    <div style={{ display: "inline-block", marginTop: 5, padding: "2px 9px", borderRadius: 20, background: "rgba(139,92,246,0.2)", border: "1px solid rgba(139,92,246,0.4)", fontSize: 10, color: "#c4b5fd", fontWeight: 700 }}>{rating}</div>
                  </div>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10 }}>
                  {[["Total Hours", data.totalHours + "h", TEAL], ["Tasks Done", data.taskCount, GOLD], ["Clients Served", Object.keys(data.clientHours).length, "#8b5cf6"]].map(([l, v, a]) => (
                    <div key={l} style={{ background: "rgba(255,255,255,0.06)", borderRadius: 9, padding: "10px 14px", borderLeft: `3px solid ${a}`, minWidth: 90 }}>
                      <div style={{ fontSize: 8, color: "#475569", textTransform: "uppercase", marginBottom: 3 }}>{l}</div>
                      <div style={{ fontSize: 18, fontWeight: 800, color: "#fff", fontFamily: "monospace" }}>{v}</div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
            {/* Insight banner */}
            <div style={{ background: "rgba(139,92,246,0.07)", border: "1px solid rgba(139,92,246,0.2)", borderRadius: 10, padding: "11px 16px", fontSize: 11, color: "#c4b5fd", lineHeight: 1.7 }}>
              💡 <strong style={{ color: "#fff" }}>Hidden Effort:</strong> {selVA} logged <strong style={{ color: TEAL }}>{data.totalHours}h</strong> but estimated real effort is <strong style={{ color: GOLD }}>{Math.round((data.totalHours + hiddenHours) * 10) / 10}h</strong> (~{Math.round((hiddenHours / data.totalHours) * 100)}% more) due to complexity multipliers on editing, research, and content work.
            </div>
            {/* Productivity metrics */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10 }}>
              {[["Output Rate", `${outputRate} tasks/hr`, "Higher = leaner delivery", TEAL], ["Category Breadth", `${breadth} types`, "Wider = more versatile", "#8b5cf6"], ["Hidden Effort Index", `+${hiddenHours}h`, "Estimated unpaid complexity", GOLD]].map(([l, v, note, a]) => (
                <div key={l} style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 10, padding: "13px 15px", borderLeft: `3px solid ${a}` }}>
                  <div style={{ fontSize: 9, color: "#475569", textTransform: "uppercase", marginBottom: 4 }}>{l}</div>
                  <div style={{ fontSize: 20, fontWeight: 800, color: "#fff", fontFamily: "monospace", marginBottom: 4 }}>{v}</div>
                  <div style={{ fontSize: 9, color: "#475569" }}>{note}</div>
                </div>
              ))}
            </div>
            {/* Work by category */}
            <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 12, padding: "16px 18px" }}>
              <div style={{ fontSize: 10, color: TEAL, textTransform: "uppercase", fontWeight: 700, marginBottom: 13 }}>Work by Category</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {Object.entries(data.cats).sort((a, b) => b[1] - a[1]).map(([cat, hrs]) => {
                  const mult = CAT_COMPLEXITY[cat] || 1.5;
                  const realHrs = Math.round(hrs * mult * 10) / 10;
                  return (
                    <div key={cat} style={{ background: "rgba(255,255,255,0.02)", borderRadius: 8, padding: "10px 13px" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 5 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
                          <div style={{ width: 7, height: 7, borderRadius: "50%", background: catColor(cat) }} />
                          <span style={{ fontSize: 12, fontWeight: 600, color: "#f1f5f9" }}>{cat}</span>
                        </div>
                        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                          <span style={{ fontSize: 10, color: "#475569" }}>logged: <strong style={{ color: "#cbd5e1" }}>{hrs}h</strong></span>
                          <span style={{ fontSize: 10, color: "#475569" }}>×{mult} =</span>
                          <span style={{ fontSize: 12, fontWeight: 700, color: GOLD }}>{realHrs}h real</span>
                        </div>
                      </div>
                      <div style={{ fontSize: 10, color: "#475569", fontStyle: "italic" }}>Includes: {CAT_HIDDEN_NOTE[cat] || "additional coordination"}</div>
                    </div>
                  );
                })}
              </div>
            </div>
            {/* Hours by client */}
            <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 12, padding: "16px 18px" }}>
              <div style={{ fontSize: 10, color: "#8b5cf6", textTransform: "uppercase", fontWeight: 700, marginBottom: 13 }}>Hours by Client</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {Object.entries(data.clientHours).sort((a, b) => b[1] - a[1]).map(([client, hrs]) => {
                  const pct = Math.round((hrs / data.totalHours) * 100);
                  const clientTasks = data.tasks.filter(t => t.client === client);
                  const tph = Math.round((clientTasks.length / hrs) * 10) / 10;
                  return (
                    <div key={client}>
                      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 4 }}>
                        <span style={{ color: "#cbd5e1", fontWeight: 500 }}>{client}</span>
                        <span style={{ color: "#f1f5f9", fontWeight: 700 }}>{hrs}h <span style={{ color: "#475569", fontWeight: 400 }}>({pct}%) · {tph} tasks/hr</span></span>
                      </div>
                      <div style={{ height: 5, background: "rgba(255,255,255,0.05)", borderRadius: 3 }}>
                        <div style={{ height: 5, borderRadius: 3, background: "linear-gradient(90deg,#8b5cf6,#6d28d9)", width: `${pct}%`, transition: "width 0.4s" }} />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
            {/* Wins */}
            {data.wins.length > 0 && (
              <div style={{ background: "rgba(240,180,41,0.05)", border: "1px solid rgba(240,180,41,0.15)", borderRadius: 12, padding: "16px 18px" }}>
                <div style={{ fontSize: 10, color: GOLD, textTransform: "uppercase", fontWeight: 700, marginBottom: 12 }}>⭐ Client Feedback</div>
                {data.wins.map((w, i) => (
                  <div key={i} style={{ padding: "9px 12px", background: "rgba(240,180,41,0.07)", border: "1px solid rgba(240,180,41,0.15)", borderRadius: 8, marginBottom: 6 }}>
                    <div style={{ fontSize: 10, color: "#94a3b8", marginBottom: 3 }}>{w.client}</div>
                    <div style={{ fontSize: 12, color: "#fde68a", fontStyle: "italic" }}>"{w.note}"</div>
                  </div>
                ))}
              </div>
            )}
            {/* All tasks */}
            <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 12, padding: "16px 18px" }}>
              <div style={{ fontSize: 10, color: "#f1f5f9", textTransform: "uppercase", fontWeight: 700, marginBottom: 14 }}>All Tasks This Period</div>
              {Object.entries(data.tasks.reduce((acc, t) => { if (!acc[t.client]) acc[t.client] = []; acc[t.client].push(t); return acc; }, {})).map(([client, tasks]) => (
                <div key={client} style={{ marginBottom: 16 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: TEAL, marginBottom: 7, paddingBottom: 4, borderBottom: "1px solid rgba(255,255,255,0.05)" }}>
                    {client} <span style={{ color: "#475569", fontWeight: 400 }}>· {tasks.reduce((s, t) => s + t.hours, 0).toFixed(2)}h</span>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                    {tasks.sort((a, b) => b.hours - a.hours).map((t, i) => (
                      <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "5px 9px", background: "rgba(255,255,255,0.02)", borderRadius: 7, gap: 8 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 7, minWidth: 0 }}>
                          <div style={{ width: 5, height: 5, borderRadius: "50%", background: catColor(t.cat), flexShrink: 0 }} />
                          <span style={{ fontSize: 11, color: "#cbd5e1", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.name}</span>
                        </div>
                        <div style={{ display: "flex", alignItems: "center", gap: 6, flexShrink: 0 }}>
                          <span style={{ fontSize: 9, color: "#475569" }}>{t.date}</span>
                          <span style={{ fontSize: 10, padding: "1px 6px", borderRadius: 10, background: `${catColor(t.cat)}20`, color: catColor(t.cat), fontWeight: 600 }}>{t.cat}</span>
                          <span style={{ fontSize: 11, fontWeight: 700, color: TEAL, fontFamily: "monospace", minWidth: 32, textAlign: "right" }}>{t.hours}h</span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── VA Performance Panel ──────────────────────────────────────────────────────
const CAT_COLORS = {
  "Content & Marketing": "#0d9488",
  "Administrative": "#8b5cf6",
  "Management": "#f0b429",
  "Design & Editing": "#ec4899",
  "Platform Management": "#3b82f6",
  "Client Relations": "#22c55e",
  "Document Prep": "#f97316",
};
function catColor(c) { return CAT_COLORS[c] || "#64748b"; }

function VACard({ va, detail, totalHours, wins }) {
  const pct = Math.round((detail.hours / totalHours) * 100);
  const topCat = Object.entries(detail.cats).sort((a,b) => b[1]-a[1])[0];
  const vaWins = (wins || []).filter(w => w.va === va);
  return (
    <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.08)", borderRadius: 12, overflow: "hidden" }}>
      {/* VA header */}
      <div style={{ background: "rgba(255,255,255,0.05)", padding: "14px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: "1px solid rgba(255,255,255,0.06)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 36, height: 36, borderRadius: "50%", background: `linear-gradient(135deg, ${catColor(topCat?.[0])}, ${catColor(topCat?.[0])}88)`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, fontWeight: 800, color: "#fff", flexShrink: 0 }}>
            {va.charAt(0).toUpperCase()}
          </div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 13, color: "#f1f5f9" }}>{va}</div>
            <div style={{ fontSize: 10, color: "#475569", marginTop: 1 }}>{detail.tasks.length} tasks · {detail.hours}h</div>
          </div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontSize: 18, fontWeight: 800, color: GOLD, fontFamily: "monospace" }}>{pct}%</div>
          <div style={{ fontSize: 9, color: "#475569" }}>of total hours</div>
        </div>
      </div>

      <div style={{ padding: "14px 16px" }}>
        {/* Hours bar */}
        <div style={{ marginBottom: 14 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 9, color: "#475569", marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>
            <span>Hours Contributed</span><span style={{ color: TEAL }}>{detail.hours}h of {totalHours}h total</span>
          </div>
          <div style={{ height: 6, background: "rgba(255,255,255,0.05)", borderRadius: 3 }}>
            <div style={{ height: 6, borderRadius: 3, background: `linear-gradient(90deg, ${TEAL}, ${GOLD})`, width: `${pct}%`, transition: "width 0.5s" }} />
          </div>
        </div>

        {/* Category breakdown */}
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 9, color: "#475569", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 8 }}>By Category</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
            {Object.entries(detail.cats).sort((a,b) => b[1]-a[1]).map(([cat, hrs]) => (
              <div key={cat} style={{ display: "flex", alignItems: "center", gap: 5, padding: "3px 8px", borderRadius: 20, background: `${catColor(cat)}18`, border: `1px solid ${catColor(cat)}40` }}>
                <div style={{ width: 6, height: 6, borderRadius: "50%", background: catColor(cat), flexShrink: 0 }} />
                <span style={{ fontSize: 10, color: "#cbd5e1" }}>{cat}</span>
                <span style={{ fontSize: 10, fontWeight: 700, color: catColor(cat) }}>{hrs}h</span>
              </div>
            ))}
          </div>
        </div>

        {/* Tasks */}
        <div>
          <div style={{ fontSize: 9, color: "#475569", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 7 }}>Tasks Completed</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            {detail.tasks.sort((a,b) => b.hours-a.hours).map((t, i) => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "5px 9px", background: "rgba(255,255,255,0.03)", borderRadius: 7, gap: 8 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 7, minWidth: 0 }}>
                  <div style={{ width: 5, height: 5, borderRadius: "50%", background: catColor(t.cat), flexShrink: 0 }} />
                  <span style={{ fontSize: 11, color: "#cbd5e1", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.name}</span>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 6, flexShrink: 0 }}>
                  <span style={{ fontSize: 9, color: "#475569" }}>{t.date}</span>
                  <span style={{ fontSize: 11, fontWeight: 700, color: TEAL, fontFamily: "monospace", minWidth: 32, textAlign: "right" }}>{t.hours}h</span>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Wins */}
        {vaWins.length > 0 && (
          <div style={{ marginTop: 12 }}>
            <div style={{ fontSize: 9, color: "#475569", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 7 }}>Client Feedback</div>
            {vaWins.map((w, i) => (
              <div key={i} style={{ padding: "7px 10px", background: "rgba(240,180,41,0.07)", border: "1px solid rgba(240,180,41,0.2)", borderRadius: 7, fontSize: 11, color: "#fde68a", fontStyle: "italic" }}>
                ⭐ "{w.note}"
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function VAPerformance({ data }) {
  const [sortBy, setSortBy] = useState("hours");
  if (!data?.vaDetails) return (
    <div style={{ textAlign: "center", opacity: 0.4, padding: "48px 0" }}>
      <div style={{ fontSize: 32 }}>👥</div>
      <div style={{ fontSize: 13, color: "#94a3b8", marginTop: 8 }}>No VA data for this period</div>
    </div>
  );

  const sorted = Object.entries(data.vaDetails).sort((a, b) =>
    sortBy === "hours" ? b[1].hours - a[1].hours :
    sortBy === "tasks" ? b[1].tasks.length - a[1].tasks.length :
    a[0].localeCompare(b[0])
  );

  const totalTasks = Object.values(data.vaDetails).reduce((s, d) => s + d.tasks.length, 0);

  return (
    <div>
      {/* Summary row */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8, marginBottom: 16 }}>
        {[
          ["VAs Active", sorted.length, "#8b5cf6"],
          ["Total Tasks", totalTasks, GOLD],
          ["Total Hours", data.totalHours + "h", TEAL],
        ].map(([l, v, a]) => (
          <div key={l} style={{ background: "rgba(255,255,255,0.04)", borderRadius: 9, padding: "11px 14px", borderLeft: `3px solid ${a}` }}>
            <div style={{ fontSize: 9, color: "#475569", textTransform: "uppercase", marginBottom: 3 }}>{l}</div>
            <div style={{ fontSize: 20, fontWeight: 800, color: "#f1f5f9", fontFamily: "monospace" }}>{v}</div>
          </div>
        ))}
      </div>

      {/* Sort controls */}
      <div style={{ display: "flex", gap: 6, marginBottom: 14, alignItems: "center" }}>
        <span style={{ fontSize: 10, color: "#475569", textTransform: "uppercase", letterSpacing: "0.06em" }}>Sort by:</span>
        {[["hours", "Hours"], ["tasks", "Tasks"], ["name", "Name"]].map(([k, label]) => (
          <button key={k} onClick={() => setSortBy(k)}
            style={{ padding: "4px 10px", borderRadius: 6, border: `1px solid ${sortBy === k ? TEAL : "rgba(255,255,255,0.08)"}`, background: sortBy === k ? "rgba(13,148,136,0.15)" : "transparent", color: sortBy === k ? TEAL : "#64748b", cursor: "pointer", fontSize: 11, fontWeight: sortBy === k ? 700 : 400 }}>
            {label}
          </button>
        ))}
      </div>

      {/* VA cards grid */}
      <div style={{ display: "grid", gridTemplateColumns: sorted.length === 1 ? "1fr" : "repeat(auto-fill, minmax(340px, 1fr))", gap: 12 }}>
        {sorted.map(([va, detail]) => (
          <VACard key={va} va={va} detail={detail} totalHours={data.totalHours} wins={data.wins} />
        ))}
      </div>
    </div>
  );
}

// ── Sidebar ────────────────────────────────────────────────────────────────────
function Sidebar({ sel, onSel, selPeriod, onPeriod, clients, syncStatus, sheetData }) {
  const groups = groupClients(clients);
  const months = getMonthPeriods();
  const [openMonth, setOpenMonth] = useState(months[months.length - 1].month);
  return (
    <div style={{ background: "#060e1b", borderRight: "1px solid rgba(255,255,255,0.06)", height: "100vh", overflowY: "auto", position: "sticky", top: 0 }}>
      <div style={{ padding: "16px 14px 12px", borderBottom: "1px solid rgba(255,255,255,0.06)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
          <div style={{ width: 28, height: 28, borderRadius: 7, background: `linear-gradient(135deg,${TEAL},#0f766e)`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13 }}>⚡</div>
          <div>
            <div style={{ fontSize: 13, fontWeight: 700, color: "#fff" }}>Ava KPI Generator</div>
            <div style={{ fontSize: 9, color: syncStatus === "ok" ? "#22c55e" : syncStatus === "syncing" ? GOLD : "#64748b" }}>
              {syncStatus === "ok" ? "✅ Live data" : syncStatus === "syncing" ? "🔄 Syncing..." : "📋 Cached data"}
            </div>
          </div>
        </div>
      </div>
      <div style={{ padding: "14px 12px 4px" }}>
        <div style={{ fontSize: 9, color: "#475569", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 8, display: "flex", alignItems: "center", gap: 5 }}>
          <span style={{ background: TEAL, color: "#fff", width: 15, height: 15, borderRadius: "50%", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 8, fontWeight: 700 }}>1</span>
          Client
        </div>
        {Object.entries(groups).map(([group, clients]) => (
          <div key={group} style={{ marginBottom: 9 }}>
            <div style={{ fontSize: 9, color: "#1e3a5f", textTransform: "uppercase", padding: "1px 6px", marginBottom: 2 }}>{group}</div>
            {clients.map(c => {
              const active = sel?.display === c.display;
              const hasData = !!(c.dataKey && sheetData[c.dataKey]);
              return (
                <button key={c.display} onClick={() => onSel(c)}
                  style={{ width: "100%", textAlign: "left", padding: "6px 9px", borderRadius: 6, marginBottom: 1, background: active ? "rgba(13,148,136,0.14)" : "transparent", border: active ? `1px solid ${TEAL}` : "1px solid transparent", color: active ? "#fff" : "#94a3b8", cursor: "pointer", fontSize: 12, fontWeight: active ? 600 : 400, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span>{c.display}{c.priv ? " 🔒" : ""}</span>
                  {hasData && <span style={{ fontSize: 8, color: TEAL, fontWeight: 700 }}>● sheet</span>}
                </button>
              );
            })}
          </div>
        ))}
      </div>
      <div style={{ height: 1, background: "rgba(255,255,255,0.04)", margin: "2px 12px" }} />
      <div style={{ padding: "10px 12px 16px" }}>
        <div style={{ fontSize: 9, color: "#475569", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 9, display: "flex", alignItems: "center", gap: 5 }}>
          <span style={{ background: GOLD, color: NAVY, width: 15, height: 15, borderRadius: "50%", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 8, fontWeight: 700 }}>2</span>
          Period
        </div>
        {months.map(m => (
          <div key={m.month} style={{ marginBottom: 4 }}>
            <button onClick={() => setOpenMonth(openMonth === m.month ? null : m.month)}
              style={{ width: "100%", textAlign: "left", padding: "5px 9px", borderRadius: 6, background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.05)", color: "#64748b", cursor: "pointer", fontSize: 11, fontWeight: 600, display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 2 }}>
              <span>{m.month}</span>
              <span style={{ fontSize: 9, color: "#334155" }}>{openMonth === m.month ? "▲" : "▼"}</span>
            </button>
            {openMonth === m.month && m.slots.map(p => {
              const active = selPeriod?.label === p.label;
              return (
                <button key={p.label} onClick={() => onPeriod(p)}
                  style={{ width: "100%", textAlign: "left", padding: "6px 12px", borderRadius: 6, marginBottom: 2, background: active ? "rgba(240,180,41,0.1)" : "rgba(255,255,255,0.02)", border: active ? `1px solid ${GOLD}` : "1px solid transparent", color: active ? "#fff" : "#94a3b8", cursor: "pointer", fontSize: 11 }}>
                  {p.label}
                </button>
              );
            })}
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Main Panel ─────────────────────────────────────────────────────────────────
function MainPanel({ client, period, sheetData }) {
  const [status, setStatus] = useState("idle");
  const [reportData, setReportData] = useState(null);
  const [dots, setDots] = useState(0);
  const [errorMsg, setErrorMsg] = useState("");
  const data = useMemo(() => {
    if (!client?.dataKey || !period) return null;
    const raw = sheetData[client.dataKey];
    return raw ? filterData(raw, period) : null;
  }, [client, period]); // eslint-disable-line react-hooks/exhaustive-deps

  useMemo(() => { setStatus("idle"); setReportData(null); }, [client, period]);

  const generate = async () => {
    if (!client || !period) return;
    setStatus("loading");
    setReportData(null);
    setErrorMsg("");
    let d = 0;
    const iv = setInterval(() => { d = (d + 1) % 4; setDots(d); }, 400);
    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-api-key": (() => { if (!window.__AVA_API_KEY__) { window.__AVA_API_KEY__ = prompt("Enter your Anthropic API key (saved for this session):"); } return window.__AVA_API_KEY__; })(), "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 4000,
          messages: [{ role: "user", content: buildPrompt(client, period, data) }]
        })
      });
      const result = await res.json();
      if (result.error) {
        throw new Error(`API: ${result.error.type} — ${result.error.message}`);
      }
      const raw = result.content?.map(b => b.text || "").join("") || "";
      // Strip markdown fences and any preamble/postamble text around the JSON
      let jsonStr = raw.replace(/```json\s*/gi, "").replace(/```\s*/g, "").trim();
      const firstBrace = jsonStr.indexOf("{");
      const lastBrace = jsonStr.lastIndexOf("}");
      if (firstBrace === -1 || lastBrace === -1) throw new Error("No JSON object found in response");
      jsonStr = jsonStr.slice(firstBrace, lastBrace + 1);
      const parsed = JSON.parse(jsonStr);
      setReportData(parsed);
      setStatus("done");
    } catch (err) {
      console.error("KPI generation error:", err);
      setErrorMsg(err.message || String(err));
      setStatus("error");
    } finally { clearInterval(iv); }
  };

  const downloadHTML = () => {
    if (!reportData) return;
    const html = generatePDFHTML(reportData, client, period);
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `KPI_Report_${client.display.replace(/[^a-z0-9]/gi, "_")}_${period.label.replace(/[^a-z0-9]/gi, "_")}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const printAsPDF = () => {
    if (!reportData) return;
    const html = generatePDFHTML(reportData, client, period);
    // Inject auto-print script into the HTML before opening
    const htmlWithPrint = html.replace(
      "</body>",
      `<script>window.onload = function() { setTimeout(function() { window.print(); }, 600); }<\/script></body>`
    );
    const blob = new Blob([htmlWithPrint], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.target = "_blank";
    a.rel = "noopener noreferrer";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  };

  if (!client && !period) return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "70vh", gap: 14, opacity: 0.4 }}>
      <div style={{ fontSize: 48 }}>📊</div>
      <div style={{ fontSize: 16, fontWeight: 700, color: "#94a3b8" }}>Select a client + period</div>
      <div style={{ fontSize: 12, color: "#475569", textAlign: "center", lineHeight: 1.9 }}>Choose a client and reporting window from the sidebar to get started.</div>
    </div>
  );
  if (client && !period) return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "70vh", gap: 10, opacity: 0.5 }}>
      <div style={{ fontSize: 36 }}>📅</div>
      <div style={{ fontSize: 14, fontWeight: 700, color: "#94a3b8" }}>Now select a reporting period</div>
    </div>
  );

  const roi = data ? {
    oppLow: (data.totalHours * 2.5 * 100).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
    oppHigh: (data.totalHours * 2.5 * 150).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
    avoid: (data.totalHours * 2.5 * 75).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
  } : null;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>

      {/* Header card */}
      <div style={{ background: `linear-gradient(135deg,${NAVY2},#163456)`, borderRadius: 14, padding: "20px 26px", borderBottom: `3px solid ${GOLD}` }}>
        <div style={{ fontSize: 9, color: GOLD, textTransform: "uppercase", letterSpacing: "0.14em", marginBottom: 3 }}>Ava Virtual Agents Inc. · KPI Report</div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 8 }}>
          <div>
            <div style={{ fontSize: 10, color: "#475569", marginBottom: 1 }}>Prepared for</div>
            <div style={{ fontSize: 22, fontWeight: 800, color: "#fff" }}>{client.display}</div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 10, color: "#475569", marginBottom: 1 }}>Reporting period</div>
            <div style={{ fontSize: 14, fontWeight: 700, color: "#fff" }}>{period.label}</div>
            <div style={{ fontSize: 10, color: "#334155", marginTop: 3 }}>{client.slack}</div>
          </div>
        </div>
      </div>

      {/* Data preview */}
      {data && status === "idle" && (
        <>
          <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 12, padding: "16px 18px" }}>
            <div style={{ fontSize: 10, color: TEAL, textTransform: "uppercase", fontWeight: 700, marginBottom: 13 }}>Sheet Data Preview</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8, marginBottom: 14 }}>
              {[["Total Hours", `${data.totalHours}h`, TEAL], ["Tasks", data.taskCount, GOLD], ["VAs Active", Object.keys(data.vas).length, "#8b5cf6"]].map(([l, v, a]) => (
                <div key={l} style={{ background: "rgba(255,255,255,0.04)", borderRadius: 9, padding: "10px 12px", borderLeft: `3px solid ${a}` }}>
                  <div style={{ fontSize: 9, color: "#475569", textTransform: "uppercase", marginBottom: 3 }}>{l}</div>
                  <div style={{ fontSize: 19, fontWeight: 800, color: "#f1f5f9", fontFamily: "monospace" }}>{v}</div>
                </div>
              ))}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <div>
                <div style={{ fontSize: 9, color: "#475569", textTransform: "uppercase", marginBottom: 9 }}>Hours by Category</div>
                {Object.entries(data.cats).map(([c, h]) => <BarRow key={c} label={c} value={Math.round(h * 100) / 100} max={Math.max(...Object.values(data.cats))} color={TEAL} />)}
              </div>
              <div>
                <div style={{ fontSize: 9, color: "#475569", textTransform: "uppercase", marginBottom: 9 }}>VA Contribution</div>
                {Object.entries(data.vas).map(([v, h]) => <BarRow key={v} label={v} value={Math.round(h * 100) / 100} max={Math.max(...Object.values(data.vas))} color={GOLD} />)}
              </div>
            </div>
          </div>
          {roi && (
            <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 12, padding: "16px 18px" }}>
              <div style={{ fontSize: 10, color: TEAL, textTransform: "uppercase", fontWeight: 700, marginBottom: 12 }}>Estimated Value</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8 }}>
                {[["Opportunity Value", `$${roi.oppLow} – $${roi.oppHigh}`, GOLD], ["Cost Avoidance", `$${roi.avoid}`, TEAL], ["Self-Hours Saved", `${Math.round(data.totalHours * 1.5 * 10) / 10}h`, GREEN]].map(([l, v, a]) => (
                  <div key={l} style={{ background: "rgba(255,255,255,0.04)", borderRadius: 9, padding: "12px 13px", borderLeft: `3px solid ${a}` }}>
                    <div style={{ fontSize: 9, color: "#475569", marginBottom: 4, textTransform: "uppercase" }}>{l}</div>
                    <div style={{ fontSize: 13, fontWeight: 800, color: "#f1f5f9", fontFamily: "monospace" }}>{v}</div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </>
      )}

      {!data && status === "idle" && (
        <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 12, padding: "24px", textAlign: "center" }}>
          <div style={{ fontSize: 28, marginBottom: 8 }}>📭</div>
          <div style={{ fontSize: 13, color: "#64748b", lineHeight: 1.8 }}>No sheet data for this period. Claude will generate from available context.</div>
        </div>
      )}

      {/* Loading */}
      {status === "loading" && (
        <div style={{ background: "rgba(13,148,136,0.07)", border: `1px solid rgba(13,148,136,0.2)`, borderRadius: 14, padding: "36px", textAlign: "center" }}>
          <div style={{ fontSize: 34, marginBottom: 12 }}>{["⚙️","📊","✍️","📋"][dots]}</div>
          <div style={{ fontSize: 14, fontWeight: 700, color: TEAL, marginBottom: 6 }}>Generating KPI Report{".".repeat(dots + 1)}</div>
          <div style={{ fontSize: 11, color: "#475569" }}>Claude is compiling tasks, metrics, and narrative…</div>
          <div style={{ marginTop: 18, height: 3, background: "rgba(255,255,255,0.05)", borderRadius: 2, overflow: "hidden" }}>
            <div style={{ height: 3, background: `linear-gradient(90deg,${TEAL},${GOLD})`, borderRadius: 2, animation: "slide 1.5s ease-in-out infinite", width: "40%" }} />
          </div>
          <style>{`@keyframes slide{0%{transform:translateX(-100%)}100%{transform:translateX(350%)}}`}</style>
        </div>
      )}

      {/* Done */}
      {status === "done" && reportData && (
        <div style={{ background: "rgba(13,148,136,0.07)", border: `1px solid rgba(13,148,136,0.25)`, borderRadius: 14, overflow: "hidden" }}>
          <div style={{ background: "rgba(13,148,136,0.12)", borderBottom: "1px solid rgba(13,148,136,0.2)", padding: "14px 18px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: TEAL }}>✅ Report Ready</div>
              <div style={{ fontSize: 10, color: "#475569", marginTop: 1 }}>{client.display} · {period.label} · {reportData.totals?.total_hours}h</div>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={() => { setStatus("idle"); setReportData(null); }}
                style={{ padding: "7px 14px", borderRadius: 7, border: "1px solid rgba(255,255,255,0.1)", background: "transparent", color: "#64748b", cursor: "pointer", fontSize: 11, fontWeight: 600 }}>
                ↩ Regenerate
              </button>
              <button onClick={printAsPDF}
                style={{ padding: "9px 18px", borderRadius: 8, border: "none", background: `linear-gradient(135deg,${GOLD},#d97706)`, color: NAVY, cursor: "pointer", fontSize: 12, fontWeight: 800, display: "flex", alignItems: "center", gap: 6, boxShadow: "0 3px 14px rgba(240,180,41,0.35)" }}>
                🖨️ Save as PDF
              </button>
              <button onClick={downloadHTML}
                style={{ padding: "9px 18px", borderRadius: 8, border: "1px solid rgba(255,255,255,0.15)", background: "rgba(255,255,255,0.06)", color: "#e2e8f0", cursor: "pointer", fontSize: 12, fontWeight: 700, display: "flex", alignItems: "center", gap: 6 }}>
                ⬇ HTML
              </button>
            </div>
          </div>
          <div style={{ padding: "16px 20px", display: "grid", gridTemplateColumns: "repeat(2,1fr)", gap: 10 }}>
            {[
              ["Tasks Completed", reportData.totals?.completed, GREEN],
              ["Total Hours", `${reportData.totals?.total_hours}h`, TEAL],
              ["Total Value Delivered", `$${reportData.value?.total_low} – $${reportData.value?.total_high}`, GOLD],
              ["Client Approval Rate", reportData.quality?.approval_rate, "#8b5cf6"],
            ].map(([l, v, a]) => (
              <div key={l} style={{ background: "rgba(255,255,255,0.04)", borderRadius: 8, padding: "10px 13px", borderLeft: `3px solid ${a}` }}>
                <div style={{ fontSize: 9, color: "#475569", marginBottom: 3 }}>{l}</div>
                <div style={{ fontSize: 14, fontWeight: 800, color: "#fff", fontFamily: "monospace" }}>{v || "--"}</div>
              </div>
            ))}
          </div>
          <div style={{ padding: "0 20px 16px", fontSize: 11, color: "#64748b" }}>
            <span style={{ color: GOLD, fontWeight: 600 }}>🖨️ Save as PDF</span> — opens in a new tab and auto-launches the print dialog. Choose <strong style={{ color: "#fff" }}>Save as PDF</strong> as the destination.&nbsp;&nbsp;
            <span style={{ color: "#64748b" }}>⬇ HTML saves the raw file for editing.</span>
          </div>
        </div>
      )}

      {/* Error */}
      {status === "error" && (
        <div style={{ background: "rgba(239,68,68,0.07)", border: "1px solid rgba(239,68,68,0.2)", borderRadius: 12, padding: "20px" }}>
          <div style={{ fontSize: 13, color: "#f87171", marginBottom: 8, fontWeight: 700 }}>⚠️ Generation failed</div>
          {errorMsg && <div style={{ fontSize: 10, color: "#fca5a5", marginBottom: 12, fontFamily: "monospace", background: "rgba(0,0,0,0.2)", padding: "8px 10px", borderRadius: 6, wordBreak: "break-all", lineHeight: 1.6 }}>{errorMsg}</div>}
          <button onClick={() => { setStatus("idle"); setErrorMsg(""); }} style={{ padding: "7px 16px", borderRadius: 7, border: "1px solid rgba(239,68,68,0.3)", background: "transparent", color: "#f87171", cursor: "pointer", fontSize: 11 }}>Try Again</button>
        </div>
      )}

      {/* Generate button */}
      {status === "idle" && (
        <button onClick={generate}
          style={{ width: "100%", padding: "16px", borderRadius: 12, border: "none", background: `linear-gradient(135deg,${TEAL},#0f766e)`, color: "#fff", cursor: "pointer", fontSize: 14, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center", gap: 10, boxShadow: `0 4px 24px rgba(13,148,136,0.35)`, transition: "transform 0.15s" }}
          onMouseEnter={e => e.currentTarget.style.transform = "translateY(-1px)"}
          onMouseLeave={e => e.currentTarget.style.transform = "translateY(0)"}>
          <span style={{ fontSize: 18 }}>🤖</span>
          Generate KPI Report with Claude
        </button>
      )}
    </div>
  );
}

export default function App() {
  const [sel, setSel] = useState(null);
  const [selPeriod, setSelPeriod] = useState(null);
  const [sheetData, setSheetData] = useState(null);
  const [syncStatus, setSyncStatus] = useState("empty");

  // Build clients list from current sheetData
  const clients = useMemo(() => sheetData ? buildClients(sheetData) : [], [sheetData]);

  const handleCSVImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const parsed = parseCSVData(ev.target.result);
      if (parsed) {
        setSheetData(parsed);
        setSyncStatus("csv");
        setSel(null);
        setSelPeriod(null);
      } else {
        alert("Could not parse CSV. Make sure it matches the Ava KPI Tracker format.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  const [appTab, setAppTab] = useState("client");

  return (
    <div style={{ minHeight: "100vh", background: NAVY, fontFamily: "'DM Sans','Inter',sans-serif", color: "#f1f5f9", display: "flex", flexDirection: "column" }}>
      {/* Top nav */}
      <div style={{ background: "#040c18", borderBottom: "1px solid rgba(255,255,255,0.06)", padding: "0 24px", display: "flex", alignItems: "center", gap: 0, height: 46, flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginRight: 28 }}>
          <div style={{ width: 22, height: 22, borderRadius: 5, background: TEAL, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 900, color: "#fff" }}>A</div>
          <span style={{ fontSize: 12, fontWeight: 700, color: "#fff" }}>Ava KPI</span>
          <span style={{ fontSize: 9, color: syncStatus === "csv" ? GREEN : "#475569", marginLeft: 4 }}>
            {syncStatus === "csv" ? "● CSV Loaded" : "● No Data"}
          </span>
          <label title="Import CSV exported from Google Sheets"
            style={{ marginLeft: 8, padding: "3px 10px", borderRadius: 5, border: `1px solid ${GOLD}60`, background: `${GOLD}15`, color: GOLD, cursor: "pointer", fontSize: 11, lineHeight: 1.8, display: "inline-flex", alignItems: "center", gap: 4, fontWeight: 700 }}>
            📁 {syncStatus === "csv" ? "Re-import CSV" : "Import CSV"}
            <input type="file" accept=".csv" onChange={handleCSVImport} style={{ display: "none" }} />
          </label>
        </div>
        {[["client", "📋 Client KPIs"], ["va", "👤 VA Dashboard"]].map(([tab, label]) => (
          <button key={tab} onClick={() => setAppTab(tab)}
            style={{ padding: "0 18px", height: "100%", border: "none", borderBottom: appTab === tab ? `2px solid ${TEAL}` : "2px solid transparent", background: "transparent", color: appTab === tab ? "#fff" : "#475569", cursor: "pointer", fontSize: 12, fontWeight: appTab === tab ? 700 : 400, transition: "all 0.15s" }}>
            {label}
          </button>
        ))}
      </div>
      {/* Content */}
      <div style={{ flex: 1, display: appTab === "client" ? "grid" : "block", gridTemplateColumns: "265px 1fr", overflow: "hidden" }}>
        {!sheetData && (
          <div style={{ gridColumn: "1/-1", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "80vh", gap: 18 }}>
            <div style={{ fontSize: 52 }}>📊</div>
            <div style={{ fontSize: 20, fontWeight: 800, color: "#fff" }}>Import your KPI Tracker CSV</div>
            <div style={{ fontSize: 13, color: "#475569", textAlign: "center", maxWidth: 380 }}>Export the <strong style={{color:"#94a3b8"}}>AVA Master KPI Tracker</strong> sheet as CSV, then import it here to generate reports.</div>
            <label style={{ marginTop: 8, padding: "13px 28px", borderRadius: 10, border: `2px solid ${GOLD}`, background: `${GOLD}15`, color: GOLD, cursor: "pointer", fontSize: 14, fontWeight: 800, display: "inline-flex", alignItems: "center", gap: 8 }}>
              📁 Import CSV
              <input type="file" accept=".csv" onChange={handleCSVImport} style={{ display: "none" }} />
            </label>
            <div style={{ fontSize: 11, color: "#334155" }}>Google Sheets → File → Download → CSV</div>
          </div>
        )}
        {sheetData && appTab === "client" && (
          <>
            <Sidebar sel={sel} onSel={setSel} selPeriod={selPeriod} onPeriod={setSelPeriod} clients={clients} syncStatus={syncStatus} sheetData={sheetData} />
            <div style={{ padding: "28px 36px", overflowY: "auto", height: "calc(100vh - 46px)" }}>
              <MainPanel client={sel} period={selPeriod} sheetData={sheetData} />
            </div>
          </>
        )}
        {sheetData && appTab === "va" && (
          <div style={{ height: "calc(100vh - 46px)", overflowY: "auto" }}>
            <VADashboard sheetData={sheetData} />
          </div>
        )}
      </div>
    </div>
  );
}
