// nostage-agent.js — "No Lead Stage" chaser.
//
// Finds contacts that have NO lead stage, DO have an owner, and whose Create Date
// OR Latest Traffic Source Date falls in the chosen window (default: 7-14 days ago,
// i.e. about a week old and still unstaged). Groups them by consultant, emails each
// consultant their own list asking them to contact the lead and set the lead stage,
// then emails Ali a full report and writes an Excel file.
//
// SAFE MODE: DRY_RUN=true prints everything and sends nothing.
// Secrets: HUBSPOT_TOKEN, RESEND_KEY.

const XLSX = require("xlsx");

const HUBSPOT_TOKEN = process.env.HUBSPOT_TOKEN;
const RESEND_KEY = process.env.RESEND_KEY;

const CFG = {
  DRY_RUN: (process.env.DRY_RUN_INPUT || "true") === "true",

  // Window in days. Leads are included when Create Date OR Latest Traffic Source
  // Date is between MAX_AGE_DAYS and MIN_AGE_DAYS ago.
  // "7 / 14"  = turned a week old and still unstaged  (recommended)
  // MAX_AGE_DAYS = 0 means "no lower bound" (the whole backlog — very large).
  MIN_AGE_DAYS: parseInt(process.env.MIN_AGE_DAYS || "7", 10),
  MAX_AGE_DAYS: parseInt(process.env.MAX_AGE_DAYS || "14", 10),

  // Safety: never email one consultant more than this many leads in one go.
  MAX_PER_CONSULTANT: parseInt(process.env.MAX_PER_CONSULTANT || "150", 10),

  ALI_EMAIL: "razaali@hofmigration.com",

  // HOW EMAIL IS SENT:
  //  "gmail"  = sent from Ali's own Gmail (razaali@hofmigration.com). Works today,
  //             no DNS setup. Lands in Ali's Sent folder, replies come back to Ali.
  //             Needs the GMAIL_APP_PASSWORD secret.
  //  "resend" = Resend API. Consultant emails only work AFTER hofmigration.com is
  //             verified in Resend; until then only Ali's own address can receive.
  SEND_VIA: (process.env.SEND_VIA || "gmail").toLowerCase(),
  GMAIL_USER: process.env.GMAIL_USER || "razaali@hofmigration.com",
  FROM_EMAIL: process.env.FROM_EMAIL || "onboarding@resend.dev",   // resend mode only
  PORTAL_ID: "23735726",
  OUT_FILE: "no-lead-stage-report.xlsx",
};

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function hub(method, path, body) {
  const url = `https://api.hubapi.com${path}`;
  for (let a = 0; a < 6; a++) {
    const res = await fetch(url, {
      method,
      headers: { Authorization: `Bearer ${HUBSPOT_TOKEN}`, "Content-Type": "application/json" },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (res.status === 429) { await sleep(2000 * (a + 1)); continue; }
    if (!res.ok) { const t = await res.text(); throw new Error(`${method} ${path} -> ${res.status}: ${t.slice(0, 250)}`); }
    return res.status === 204 ? null : res.json();
  }
  throw new Error(`rate-limited: ${method} ${path}`);
}

// name + email for every owner in the portal
async function owners() {
  const map = {}; let after;
  for (let i = 0; i < 30; i++) {
    const d = await hub("GET", `/crm/v3/owners/?limit=100${after ? `&after=${after}` : ""}`);
    for (const o of d.results || []) {
      map[String(o.id)] = {
        name: [o.firstName, o.lastName].filter(Boolean).join(" ").trim() || o.email || String(o.id),
        email: o.email || null,
        active: o.archived !== true,
      };
    }
    after = d.paging?.next?.after;
    if (!after) break;
  }
  return map;
}

function windowBounds() {
  const now = Date.now();
  const newest = now - CFG.MIN_AGE_DAYS * 86400000;                       // e.g. 7 days ago
  const oldest = CFG.MAX_AGE_DAYS > 0 ? now - CFG.MAX_AGE_DAYS * 86400000 : 0; // 0 = no lower bound
  return { oldest, newest };
}

// Two filter groups = OR. Each group also carries the common conditions (AND within a group).
async function fetchUnstaged() {
  const { oldest, newest } = windowBounds();
  const common = [
    { propertyName: "lead_stage", operator: "NOT_HAS_PROPERTY" },
    { propertyName: "hubspot_owner_id", operator: "HAS_PROPERTY" },
  ];
  const dateFilter = (prop) =>
    oldest > 0
      ? { propertyName: prop, operator: "BETWEEN", value: String(oldest), highValue: String(newest) }
      : { propertyName: prop, operator: "LT", value: String(newest) };

  const filterGroups = [
    { filters: [...common, dateFilter("createdate")] },
    { filters: [...common, dateFilter("hs_latest_source_timestamp")] },
  ];

  const seen = new Set(), out = []; let after;
  for (let page = 0; page < 200; page++) {
    const d = await hub("POST", "/crm/v3/objects/contacts/search", {
      filterGroups,
      sorts: [{ propertyName: "createdate", direction: "DESCENDING" }],
      properties: ["firstname", "lastname", "email", "phone", "hubspot_owner_id", "createdate", "hs_latest_source", "hs_latest_source_timestamp"],
      limit: 100, after,
    });
    for (const c of d.results || []) { if (!seen.has(c.id)) { seen.add(c.id); out.push(c); } }
    after = d.paging?.next?.after;
    if (!after) break;
  }
  return out;
}

const link = (id) => `https://app.hubspot.com/contacts/${CFG.PORTAL_ID}/record/0-1/${id}`;
const day = (iso) => (iso ? String(iso).slice(0, 10) : "");

// ---- sending ----------------------------------------------------------------
let mailer = null;
function gmailTransport() {
  if (mailer) return mailer;
  const pass = process.env.GMAIL_APP_PASSWORD;
  if (!pass) throw new Error("Missing GMAIL_APP_PASSWORD secret (16-character Google app password, no spaces)");
  const nodemailer = require("nodemailer");
  mailer = nodemailer.createTransport({
    host: "smtp.gmail.com", port: 465, secure: true,
    auth: { user: CFG.GMAIL_USER, pass },
  });
  return mailer;
}

async function sendEmail(to, subject, html, attachFile) {
  try {
    if (CFG.SEND_VIA === "gmail") {
      const msg = { from: CFG.GMAIL_USER, to, subject, html };
      if (attachFile && require("fs").existsSync(attachFile))
        msg.attachments = [{ filename: attachFile, path: attachFile }];
      await gmailTransport().sendMail(msg);
      return true;
    }
    if (!RESEND_KEY) { console.log(`  (no RESEND_KEY — would have emailed ${to})`); return false; }
    const body = { from: CFG.FROM_EMAIL, to: [to], subject, html };
    if (attachFile && require("fs").existsSync(attachFile))
      body.attachments = [{ filename: attachFile, content: require("fs").readFileSync(attachFile).toString("base64") }];
    const res = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { Authorization: `Bearer ${RESEND_KEY}`, "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const t = (await res.text()).slice(0, 200);
      console.log(`  ! email to ${to} failed: ${res.status} ${t}`);
      if (res.status === 403) console.log(`    (Resend only sends to your own address until hofmigration.com is verified — switch the sender dropdown to Gmail)`);
      return false;
    }
    return true;
  } catch (e) {
    console.log(`  ! email to ${to} failed: ${e.message}`);
    return false;
  }
}

function consultantEmailHtml(firstName, leads) {
  const rows = leads.map((l) => `
    <tr>
      <td style="padding:6px 10px;border-bottom:1px solid #eee;"><a href="${link(l.id)}">${l.name}</a></td>
      <td style="padding:6px 10px;border-bottom:1px solid #eee;">${l.phone || ""}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #eee;">${l.created}</td>
    </tr>`).join("");
  return `
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#33475b;">
    <p style="background:#f5f8fa;border-left:3px solid #7c98b6;padding:8px 12px;margin:0 0 16px 0;font-size:13px;color:#516f90;">
      This is an automated compliance email. Please do not reply to this message. For any question, contact Ali Raza directly.
    </p>
    <p>Hi ${firstName},</p>
    <p>Hope you are well. The leads below are assigned to you and still have <strong>no lead stage</strong> selected.</p>
    <p>Kindly contact these clients and select the correct lead stage on each record.</p>
    <table style="border-collapse:collapse;font-size:13px;">
      <tr>
        <th align="left" style="padding:6px 10px;border-bottom:2px solid #33475b;">Lead</th>
        <th align="left" style="padding:6px 10px;border-bottom:2px solid #33475b;">Phone</th>
        <th align="left" style="padding:6px 10px;border-bottom:2px solid #33475b;">Created</th>
      </tr>
      ${rows}
    </table>
    <p style="margin-top:16px;">Total: <strong>${leads.length}</strong> leads.</p>
    <p>Thank you.</p>
    <p style="color:#7c98b6;font-size:12px;margin-bottom:2px;">Ali Raza &middot; Compliance &middot; HOF Migration</p>
    <p style="color:#a0b4c6;font-size:11px;margin-top:0;">Sent automatically by the CRM compliance system. Generated ${new Date().toISOString().slice(0, 10)}.</p>
  </div>`;
}

(async () => {
  if (!HUBSPOT_TOKEN) throw new Error("Missing HUBSPOT_TOKEN");
  const { oldest, newest } = windowBounds();
  console.log(`=== No Lead Stage agent — ${new Date().toISOString()} ===`);
  console.log(`DRY_RUN: ${CFG.DRY_RUN}`);
  console.log(`Window:  created / last source between ${oldest ? day(new Date(oldest).toISOString()) : "any time"} and ${day(new Date(newest).toISOString())}`);
  console.log(`Cap:     ${CFG.MAX_PER_CONSULTANT} leads per consultant`);
  console.log(`Sending: ${CFG.SEND_VIA === "gmail" ? `Gmail as ${CFG.GMAIL_USER}` : `Resend as ${CFG.FROM_EMAIL}`}`);
  if (!CFG.DRY_RUN && CFG.SEND_VIA === "gmail" && !process.env.GMAIL_APP_PASSWORD)
    console.log(`!! GMAIL_APP_PASSWORD secret is missing — no email can be sent.`);
  if (!CFG.DRY_RUN && CFG.SEND_VIA === "resend" && CFG.FROM_EMAIL.endsWith("resend.dev"))
    console.log(`!! Sender is onboarding@resend.dev — consultants will NOT receive anything (403). Use Gmail instead.`);
  console.log("");

  const ownerMap = await owners();
  const contacts = await fetchUnstaged();
  console.log(`Contacts with NO lead stage and a known owner in this window: ${contacts.length}`);

  // group by consultant
  const byOwner = {};
  for (const c of contacts) {
    const oid = c.properties.hubspot_owner_id;
    (byOwner[oid] ||= []).push({
      id: c.id,
      name: [c.properties.firstname, c.properties.lastname].filter(Boolean).join(" ").trim() || `Contact ${c.id}`,
      phone: c.properties.phone || "",
      email: c.properties.email || "",
      created: day(c.properties.createdate),
      source: c.properties.hs_latest_source || "",
      sourceDate: day(c.properties.hs_latest_source_timestamp),
    });
  }

  const groups = Object.entries(byOwner)
    .map(([oid, leads]) => ({
      ownerId: oid,
      name: ownerMap[oid]?.name || `Owner ${oid}`,
      email: ownerMap[oid]?.email || null,
      active: ownerMap[oid]?.active !== false,
      leads: leads.sort((a, b) => (a.created < b.created ? 1 : -1)),
    }))
    .sort((a, b) => b.leads.length - a.leads.length);

  // ---- report to the log ----
  console.log(`Consultants affected: ${groups.length}\n`);
  console.log(`LEADS PER CONSULTANT`);
  console.log(`${"count".padStart(6)}  consultant${" ".repeat(20)}email`);
  for (const g of groups)
    console.log(`${String(g.leads.length).padStart(6)}  ${g.name.padEnd(28).slice(0, 28)}  ${g.email || "NO EMAIL IN HUBSPOT"}${g.active ? "" : "  (deactivated)"}`);

  console.log(`\nLEAD NAMES PER CONSULTANT`);
  for (const g of groups) {
    console.log(`\n${g.name} (${g.leads.length}):`);
    g.leads.forEach((l) => console.log(`   - ${l.name}${l.phone ? `  ${l.phone}` : ""}  created ${l.created}  ${link(l.id)}`));
  }

  // ---- Excel report ----
  const allRows = [];
  for (const g of groups)
    for (const l of g.leads)
      allRows.push({
        Consultant: g.name, "Consultant Email": g.email || "", Lead: l.name,
        Phone: l.phone, "Lead Email": l.email, "Create Date": l.created,
        "Latest Source": l.source, "Latest Source Date": l.sourceDate, Link: link(l.id),
      });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
    groups.map((g) => ({ Consultant: g.name, Email: g.email || "NO EMAIL", "Leads with no lead stage": g.leads.length }))
  ), "Per Consultant");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(allRows), "All Leads");
  XLSX.writeFile(wb, CFG.OUT_FILE);
  console.log(`\nWrote ${CFG.OUT_FILE} (download from the run's Artifacts section).`);

  // ---- email each consultant (live runs only) ----
  let sent = 0, skipped = 0, noEmail = 0;
  if (CFG.DRY_RUN) {
    console.log(`\nDRY RUN: no consultant emails sent. They WOULD have gone to:`);
    for (const g of groups) console.log(`   ${g.email ? g.email : "NO EMAIL IN HUBSPOT"}  (${g.name}, ${Math.min(g.leads.length, CFG.MAX_PER_CONSULTANT)} leads)`);
    noEmail = groups.filter((g) => !g.email).length;
  } else {
    console.log(`\nSending consultant emails...`);
    for (const g of groups) {
      if (!g.email) { console.log(`  skip ${g.name}: no email in HubSpot`); noEmail++; skipped++; continue; }
      const leads = g.leads.slice(0, CFG.MAX_PER_CONSULTANT);
      const ok = await sendEmail(g.email, `[Automated] ${leads.length} of your leads have no lead stage selected`,
        consultantEmailHtml(g.name.split(" ")[0], leads));
      if (ok) { console.log(`  sent ${g.name} (${leads.length} leads)`); sent++; } else skipped++;
      await sleep(600);
    }
  }

  // ---- report to Ali: sent on EVERY run, including dry runs ----
  const summary = groups.map((g) =>
    `<tr><td style="padding:4px 10px;border-bottom:1px solid #eee;">${g.name}</td>
         <td style="padding:4px 10px;border-bottom:1px solid #eee;">${g.leads.length}</td>
         <td style="padding:4px 10px;border-bottom:1px solid #eee;">${g.email || '<span style="color:#c0392b;">NO EMAIL</span>'}</td></tr>`).join("");

  const banner = CFG.DRY_RUN
    ? `<p style="background:#fff4e5;border-left:3px solid #f5a623;padding:8px 12px;font-size:13px;color:#8a6d3b;">
         <strong>DRY RUN / PREVIEW.</strong> No consultant received an email. This is what would be sent on a live run.
       </p>`
    : `<p style="background:#eaf6ec;border-left:3px solid #45a163;padding:8px 12px;font-size:13px;color:#2c6b40;">
         <strong>LIVE RUN.</strong> Consultant emails have been sent.
       </p>`;

  const reportHtml = `
    <div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#33475b;">
      ${banner}
      <p>Leads with <strong>no lead stage</strong> selected, owner known, in the window
         ${oldest ? day(new Date(oldest).toISOString()) : "any time"} to ${day(new Date(newest).toISOString())}.</p>
      <p><strong>${contacts.length}</strong> leads across <strong>${groups.length}</strong> consultants.</p>
      <table style="border-collapse:collapse;font-size:13px;">
        <tr><th align="left" style="padding:4px 10px;border-bottom:2px solid #33475b;">Consultant</th>
            <th align="left" style="padding:4px 10px;border-bottom:2px solid #33475b;">Leads</th>
            <th align="left" style="padding:4px 10px;border-bottom:2px solid #33475b;">Email</th></tr>
        ${summary}
      </table>
      <p style="margin-top:14px;">${CFG.DRY_RUN
        ? `Would email: ${groups.length - noEmail} consultants &middot; no email on file: ${noEmail}`
        : `Emails sent: ${sent} &middot; skipped: ${skipped} (no email on file: ${noEmail})`}</p>
      <p>The full per-lead list is attached as a spreadsheet.</p>
      <p style="color:#a0b4c6;font-size:11px;">Sent automatically by the CRM compliance system.</p>
    </div>`;

  const okReport = await sendEmail(
    CFG.ALI_EMAIL,
    `${CFG.DRY_RUN ? "[PREVIEW] " : ""}No lead stage report — ${contacts.length} leads, ${groups.length} consultants`,
    reportHtml,
    CFG.OUT_FILE
  );
  console.log(`\nReport email to ${CFG.ALI_EMAIL}: ${okReport ? "SENT (with spreadsheet attached)" : "FAILED — see the error above"}`);
  if (!CFG.DRY_RUN) console.log(`Consultant emails sent: ${sent}, skipped: ${skipped}.`);
})().catch((e) => { console.error("FATAL:", e); process.exit(1); });
