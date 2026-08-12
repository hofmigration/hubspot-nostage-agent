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
const T = require("./email-template");

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

// Checks the chosen transport works BEFORE the run does any real work, so a
// missing secret is obvious immediately instead of silently sending nothing.
async function checkTransport() {
  if (CFG.SEND_VIA === "gmail") {
    if (!process.env.GMAIL_APP_PASSWORD) {
      console.log(`\n!! CANNOT SEND EMAIL — sender is set to Gmail but the GMAIL_APP_PASSWORD secret is missing.`);
      console.log(`   Fix one of these:`);
      console.log(`     a) Add the secret: Google Account > Security > 2-Step Verification > App passwords,`);
      console.log(`        create one, remove the spaces, then add it as repo secret GMAIL_APP_PASSWORD.`);
      console.log(`     b) Or re-run and pick "Resend - onboarding@resend.dev" in the "How to send" dropdown.`);
      return { ok: false, reason: "missing GMAIL_APP_PASSWORD" };
    }
    try {
      await gmailTransport().verify();
      console.log(`Transport: Gmail OK (authenticated as ${CFG.GMAIL_USER})`);
      return { ok: true };
    } catch (e) {
      console.log(`\n!! GMAIL LOGIN FAILED: ${e.message}`);
      console.log(`   Usually means the app password is wrong, has spaces in it, or app passwords`);
      console.log(`   are blocked for the Workspace account.`);
      return { ok: false, reason: `gmail auth failed: ${e.message}` };
    }
  }
  if (!RESEND_KEY) {
    console.log(`\n!! CANNOT SEND EMAIL — sender is set to Resend but the RESEND_KEY secret is missing.`);
    return { ok: false, reason: "missing RESEND_KEY" };
  }
  console.log(`Transport: Resend OK (as ${CFG.FROM_EMAIL})`);
  if (CFG.FROM_EMAIL.endsWith("resend.dev"))
    console.log(`  note: this sender only reaches ${CFG.ALI_EMAIL}. Consultant emails will fail with 403.`);
  return { ok: true };
}

// Last-resort path so Ali's report still arrives when the chosen transport is dead.
async function sendViaResendDirect(to, subject, html) {
  if (!RESEND_KEY) return false;
  try {
    const res = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { Authorization: `Bearer ${RESEND_KEY}`, "Content-Type": "application/json" },
      body: JSON.stringify({ from: "onboarding@resend.dev", to: [to], subject, html }),
    });
    if (!res.ok) { console.log(`  ! fallback email failed: ${res.status} ${(await res.text()).slice(0, 150)}`); return false; }
    return true;
  } catch (e) { console.log(`  ! fallback email failed: ${e.message}`); return false; }
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

function consultantEmailHtml(firstName, leads, windowLabel) {
  const rows = leads.map((l, i) => T.row({
    name: l.name,
    title: "",
    link: link(l.id),
    linkLabel: "Contact",
    details: [
      `created ${T.esc(l.created)}${l.phone ? ` &middot; ${T.esc(l.phone)}` : ""}${l.source ? ` &middot; source: ${T.esc(l.source)}` : ""}`,
    ],
    last: i === leads.length - 1,
  })).join("");

  return T.shell({
    title: "Leads Missing Lead Stage",
    subtitle: `${windowLabel} &middot; ${leads.length} lead${leads.length === 1 ? "" : "s"} assigned to you`,
    body:
      T.callout("This is an <strong>automated</strong> compliance email. Please do not reply to this message. For any question, contact Ali Raza directly.", "info") +
      T.paragraph(`Hi ${T.esc(firstName)}, hope you are well.`) +
      T.paragraph("The leads below are assigned to you and still have <strong>no lead stage</strong> selected. Kindly contact these clients and select the correct lead stage on each record.") +
      T.sectionTitle(`Your unstaged leads — ${leads.length} lead${leads.length === 1 ? "" : "s"}`) +
      rows +
      T.paragraph(`<strong>Total: ${leads.length}</strong>`, 13) +
      T.footer(),
  });
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

  const transport = await checkTransport();
  console.log("");

  const ownerMap = await owners();
  const contacts = await fetchUnstaged();
  console.log(`Contacts with NO lead stage and a known owner in this window: ${contacts.length}`);
  const windowLabel = `${oldest ? day(new Date(oldest).toISOString()) : "any time"} \u2192 ${day(new Date(newest).toISOString())}`;

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
  } else if (!transport.ok) {
    console.log(`\nSKIPPING consultant emails — email is not working (${transport.reason}).`);
    skipped = groups.length;
    noEmail = groups.filter((g) => !g.email).length;
  } else {
    console.log(`\nSending consultant emails...`);
    for (const g of groups) {
      if (!g.email) { console.log(`  skip ${g.name}: no email in HubSpot`); noEmail++; skipped++; continue; }
      const leads = g.leads.slice(0, CFG.MAX_PER_CONSULTANT);
      const ok = await sendEmail(g.email, `[Automated] ${leads.length} of your leads have no lead stage selected`,
        consultantEmailHtml(g.name.split(" ")[0], leads, windowLabel));
      if (ok) { console.log(`  sent ${g.name} (${leads.length} leads)`); sent++; } else skipped++;
      await sleep(600);
    }
  }

  // ---- report to Ali: sent on EVERY run, including dry runs ----
  const reportRows = groups.map((g) => [
    T.esc(g.name),
    `<strong>${g.leads.length}</strong>`,
    g.email ? T.link(`mailto:${g.email}`, g.email) : `<span style="color:#c0392b;">NO EMAIL</span>`,
  ]);

  const reportHtml = T.shell({
    title: "HOF No Lead Stage Report",
    subtitle: `${windowLabel} &middot; ${contacts.length} leads &middot; ${groups.length} consultants`,
    body:
      (CFG.DRY_RUN
        ? T.callout("<strong>DRY RUN / PREVIEW.</strong> No consultant received an email. This is what would be sent on a live run.", "warn")
        : T.callout("<strong>LIVE RUN.</strong> Consultant emails have been sent.", "ok")) +
      T.paragraph(`Contacts with <strong>no lead stage</strong> selected and a known owner, where the create date or latest traffic source date falls in this window.`) +
      T.sectionTitle(`Leads per consultant — ${groups.length} consultant(s)`) +
      T.table(["Consultant", "Leads", "Email"], reportRows) +
      T.paragraph(CFG.DRY_RUN
        ? `Would email <strong>${groups.length - noEmail}</strong> consultants &middot; no email on file: <strong>${noEmail}</strong>`
        : `Emails sent: <strong>${sent}</strong> &middot; skipped: <strong>${skipped}</strong> &middot; no email on file: <strong>${noEmail}</strong>`, 13) +
      T.paragraph("The full per-lead list is attached as a spreadsheet.", 13) +
      T.footer(`Generated ${new Date().toISOString().slice(0, 10)}.`),
  });

  const reportSubject = `${CFG.DRY_RUN ? "[PREVIEW] " : ""}No lead stage report — ${contacts.length} leads, ${groups.length} consultants`;

  let okReport = false, how = "";
  if (transport.ok) {
    okReport = await sendEmail(CFG.ALI_EMAIL, reportSubject, reportHtml, CFG.OUT_FILE);
    how = okReport ? `via ${CFG.SEND_VIA} (spreadsheet attached)` : "";
  }
  if (!okReport) {
    console.log(`Trying the fallback route for your report...`);
    okReport = await sendViaResendDirect(CFG.ALI_EMAIL, reportSubject, reportHtml);
    if (okReport) how = "via Resend fallback (no attachment — get the spreadsheet from Artifacts below)";
  }
  console.log(`\nReport email to ${CFG.ALI_EMAIL}: ${okReport ? `SENT ${how}` : "FAILED — no working email route. See the errors above."}`);
  if (!CFG.DRY_RUN) console.log(`Consultant emails sent: ${sent}, skipped: ${skipped}.`);
})().catch((e) => { console.error("FATAL:", e); process.exit(1); });
