// email-template.js — HOF email styling, matching the "HOF Daily Email QA" look.
// Table-based with inline styles so it renders correctly in Gmail and Outlook.

const C = {
  pageBg: "#f4f6fa",
  card: "#ffffff",
  cardBorder: "#e5e9f2",
  navyDark: "#1b2650",
  navyLite: "#2f3f87",
  heading: "#1f2d5c",
  text: "#33475b",
  meta: "#7c8aa5",
  divider: "#ebeef4",
  link: "#2f6ecb",
  alertBg: "#fdeaea",
  alertBar: "#d9534f",
  warnBg: "#fff4e5",
  warnBar: "#f5a623",
  okBg: "#eaf6ec",
  okBar: "#45a163",
  font: "-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif",
};

const esc = (s) => String(s == null ? "" : s)
  .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

// Whole email: navy header (title + stats line) over a white card.
function shell({ title, subtitle, body }) {
  return `<div style="background:${C.pageBg};padding:24px 12px;font-family:${C.font};">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td align="center">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="width:600px;max-width:600px;background:${C.card};border:1px solid ${C.cardBorder};border-radius:10px;">
  <tr><td style="background-color:${C.navyDark};background-image:linear-gradient(90deg,${C.navyDark},${C.navyLite});padding:20px 24px;border-radius:10px 10px 0 0;">
    <div style="color:#ffffff;font-size:19px;font-weight:700;line-height:1.3;">${title}</div>
    ${subtitle ? `<div style="color:#a9b6d8;font-size:12px;margin-top:6px;">${subtitle}</div>` : ""}
  </td></tr>
  <tr><td style="padding:20px 24px 24px 24px;">${body}</td></tr>
</table>
</td></tr></table></div>`;
}

const sectionTitle = (t) =>
  `<div style="color:${C.heading};font-size:14px;font-weight:700;margin:18px 0 10px 0;">${esc(t)}</div>`;

const paragraph = (html, size = 14) =>
  `<div style="color:${C.text};font-size:${size}px;line-height:1.55;margin:0 0 12px 0;">${html}</div>`;

// coloured callout: kind = "alert" | "warn" | "ok" | "info"
function callout(html, kind = "info") {
  const map = { alert: [C.alertBg, C.alertBar, "#a33b38"], warn: [C.warnBg, C.warnBar, "#8a6d3b"], ok: [C.okBg, C.okBar, "#2c6b40"], info: ["#f5f8fa", "#7c98b6", "#516f90"] };
  const [bg, bar, fg] = map[kind] || map.info;
  return `<div style="background:${bg};border-left:3px solid ${bar};padding:10px 14px;margin:0 0 14px 0;color:${fg};font-size:13px;line-height:1.5;">${html}</div>`;
}

// One list row: bold name, optional ×count, title text, optional link, optional detail lines.
function row({ name, count, title, link, linkLabel = "Open", details = [], last = false }) {
  const border = last ? "" : `border-bottom:1px solid ${C.divider};`;
  const badge = count ? `<span style="color:${C.meta};font-size:11px;">&times;${esc(count)}</span> ` : "";
  const anchor = link ? ` <a href="${esc(link)}" style="color:${C.link};text-decoration:none;font-size:12px;">${esc(linkLabel)} &#8599;</a>` : "";
  const sub = details.map((d) =>
    `<div style="color:${C.meta};font-size:11.5px;line-height:1.5;margin-top:3px;">&middot; ${d}</div>`).join("");
  return `<div style="padding:9px 0;${border}">
    <div style="color:${C.text};font-size:13px;line-height:1.5;">
      <strong style="color:${C.heading};">${esc(name)}</strong> ${badge}${title ? esc(title) : ""}${anchor}
    </div>${sub}
  </div>`;
}

// Simple table (header row + rows of cells; cells may contain HTML).
function table(headers, rows) {
  const th = headers.map((h, i) =>
    `<th align="${i === 0 ? "left" : "left"}" style="padding:7px 10px;border-bottom:2px solid ${C.heading};color:${C.heading};font-size:12px;">${esc(h)}</th>`).join("");
  const tr = rows.map((r) =>
    `<tr>${r.map((c) => `<td style="padding:7px 10px;border-bottom:1px solid ${C.divider};color:${C.text};font-size:13px;">${c}</td>`).join("")}</tr>`).join("");
  return `<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;width:100%;">
    <tr>${th}</tr>${tr}</table>`;
}

const footer = (extra = "") =>
  `<div style="margin-top:22px;padding-top:14px;border-top:1px solid ${C.divider};">
     <div style="color:${C.meta};font-size:12px;">Ali Raza &middot; Compliance &middot; HOF Migration</div>
     <div style="color:#a8b6c6;font-size:11px;margin-top:3px;">Sent automatically by the CRM compliance system.${extra ? ` ${esc(extra)}` : ""}</div>
   </div>`;

const link = (url, label) => `<a href="${esc(url)}" style="color:${C.link};text-decoration:none;">${esc(label)}</a>`;

module.exports = { shell, sectionTitle, paragraph, callout, row, table, footer, link, esc, C };
