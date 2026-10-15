// engagement-send — staff clicks "Send NDA + deposit" or "Send engagement letter" in Solonos.
// 15 Oct 2026: lead_nda takes email_subject / email_text written in Solonos; the letter starts the follow-up (Part 35).
//   POST { engagement_id, kind: "nda" | "letter" }   Authorization: Bearer <staff JWT>
//
//   nda:          fills NDA_TEMPLATE.docx, issues the €1,000 deposit invoice (PDF), emails both + the portal link.
//   letter:       generated + Kyriacos's signature stamped + emailed automatically once the NDA is signed, the client has
//                 confirmed their company details and the fee is set. Idempotent (a sent/signed letter is not re-sent unless force).
//                 Callable by staff or by the client's own portal session.
//   stamp:        after the client signs on the portal: draws their signature + date onto the PDF, stores signed_pdf_path,
//                 emails the signed copy to the client and info@. Callable by the client (own agreement) or staff.
//   receipt:      after an invoice is marked paid: receipt PDF, stored, emailed to the client.
//
// Deploy via the dashboard editor, name exactly: engagement-send, Verify JWT OFF (we check is_staff ourselves).
// PDF conversion: the Solonos Mail app registration needs Files.ReadWrite.All (Application) with admin consent.
// Needs: templates bucket with NDA_TEMPLATE.docx and Engagement_Letter_TEMPLATE.docx; send-email deployed;
// secrets GHL_SYNC_SECRET (used as the internal key between functions), PORTAL_URL (e.g. https://www.solonos.io/portal).

import { createClient } from "npm:@supabase/supabase-js@2";
import PizZip from "npm:pizzip@3.1.7";
import Docxtemplater from "npm:docxtemplater@3.50.0";
import { PDFDocument, StandardFonts, rgb } from "npm:pdf-lib@1.17.1";
import fontkit from "npm:@pdf-lib/fontkit@1.1.1";
import { getDocumentProxy } from "npm:unpdf@0.12.1";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json", ...CORS } });

const admin = () => createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!, { auth: { persistSession: false } });
const PORTAL = () => Deno.env.get("PORTAL_URL") ?? "https://www.solonos.io/portal";

// ------------------------------------------------------------------ helpers
const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
function dateParts(d = new Date()) {
  const day = d.getUTCDate();
  const suffix = (day % 10 === 1 && day !== 11) ? "st" : (day % 10 === 2 && day !== 12) ? "nd" : (day % 10 === 3 && day !== 13) ? "rd" : "th";
  const monthYear = `${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`;
  return { day: String(day), suffix, monthYear, plain: `${day}${suffix} ${monthYear}` };
}
const money = (n: number) => Number(n).toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const APPROACHES: Record<string, string> = {
  dcf: "Income Approach (Discounted Cash Flow DCF Analysis)",
  cca: "Market Approach (Comparable Company Analysis)",
  pt: "Market Approach (Precedent Transactions)",
  nav: "Asset-Based Approach (Net Asset Value - NAV Method)",
};
function approachLines(selected: string[]) {
  const s = new Set(selected); const out: { label: string }[] = [];
  if (s.has("dcf")) out.push({ label: APPROACHES.dcf });
  if (s.has("cca") && s.has("pt")) out.push({ label: "Market Approach (Comparable Company Analysis & Precedent Transactions)" });
  else if (s.has("cca")) out.push({ label: APPROACHES.cca });
  else if (s.has("pt")) out.push({ label: APPROACHES.pt });
  if (s.has("nav")) out.push({ label: APPROACHES.nav });
  return out;
}
async function sha256(bytes: Uint8Array) {
  const h = await crypto.subtle.digest("SHA-256", bytes as unknown as ArrayBuffer);
  return [...new Uint8Array(h)].map((b) => b.toString(16).padStart(2, "0")).join("");
}
function b64(bytes: Uint8Array) {
  let s = ""; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

// DOCX → PDF using Word's own engine via Microsoft Graph (needs Files.ReadWrite.All application permission).
// Returns null if anything fails, so the caller can fall back to sending the .docx.
async function docxToPdf(docx: Uint8Array, filename: string): Promise<Uint8Array | null> {
  try {
    const tenant = Deno.env.get("MS_TENANT_ID")!, user = Deno.env.get("MAIL_FROM")!;
    const body = new URLSearchParams({ client_id: Deno.env.get("MS_CLIENT_ID")!, client_secret: Deno.env.get("MS_CLIENT_SECRET")!, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
    const tok = await (await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body })).json();
    const H = { Authorization: `Bearer ${tok.access_token}` };
    const base = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(user)}/drive/root:/Solonos-temp/${encodeURIComponent(filename)}`;
    const up = await fetch(`${base}:/content`, { method: "PUT", headers: { ...H, "content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }, body: new Blob([docx as unknown as ArrayBuffer]) });
    if (!up.ok) { console.error("graph upload", up.status, await up.text()); return null; }
    const item = await up.json();
    const pdfRes = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(user)}/drive/items/${item.id}/content?format=pdf`, { headers: H });
    const pdf = pdfRes.ok ? new Uint8Array(await pdfRes.arrayBuffer()) : null;
    if (!pdfRes.ok) console.error("graph convert", pdfRes.status, await pdfRes.text());
    await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(user)}/drive/items/${item.id}`, { method: "DELETE", headers: H });
    return pdf;
  } catch (e) { console.error("docxToPdf", e); return null; }
}

async function fillTemplate(db: ReturnType<typeof admin>, file: string, data: Record<string, unknown>): Promise<Uint8Array> {
  const { data: blob, error } = await db.storage.from("templates").download(file);
  if (error || !blob) throw new Error(`Template ${file} not found in the templates bucket: ${error?.message ?? ""}`);
  const zip = new PizZip(new Uint8Array(await blob.arrayBuffer()));
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(data);
  return doc.getZip().generate({ type: "uint8array" }) as Uint8Array;
}

// ------------------------------------------------------------------ invoice + receipt PDFs (pdf-lib), approved design 13 Sep 2026
type InvoiceRow = {
  id: string; engagement_id: string; seq: number; amount: number; currency: string; issuing_entity: string; invoice_no: number; issued_at: string; due_at: string | null; paid_at: string | null;
  entity_name: string; client_name: string; client_email: string; company_name: string | null; company_address: string | null; company_country: string | null;
  company_tax_id: string | null; iban: string; swift: string; bank: string; account_name: string;
};
type Entity = { legal_name: string; address_lines: string[]; vat_no: string | null; registration_ref: string | null; tel: string | null; email: string | null };
const FONT_URL = "https://raw.githubusercontent.com/googlefonts/noto-fonts/main/hinted/ttf/NotoSans/";
const fontCache: Record<string, Uint8Array> = {};
async function fetchFont(name: string) {
  if (!fontCache[name]) fontCache[name] = new Uint8Array(await (await fetch(FONT_URL + name)).arrayBuffer());
  return fontCache[name];
}
const ascii = (t: string) => t.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^\x20-\x7E]/g, "?");
const C = { brand: rgb(0.5, 0, 0.125), accent: rgb(0.047, 0.278, 0.404), ink: rgb(0.165, 0.118, 0.125), grey: rgb(0.42, 0.33, 0.34), cream: rgb(0.953, 0.906, 0.871), line: rgb(0.86, 0.78, 0.75), on: rgb(0.98, 0.95, 0.91) };
const fmtDay = (d: string | Date) => new Date(d).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
const invNo = (i: { issuing_entity: string; invoice_no: number }) => `${i.issuing_entity}-${String(i.invoice_no).padStart(4, "0")}`;
const descFor = (i: InvoiceRow) => i.seq === 0
  ? [`Engagement deposit - business valuation${i.company_name ? ` of ${i.company_name}` : ""}`, "Non-refundable. Credited against the first instalment of the professional fee."]
  : i.seq === 1 ? ["Business valuation - first instalment (50% on signing the engagement letter)", "Less deposit already received."]
  : ["Business valuation - final instalment", "Payable before release of the final report."];

async function themedDoc(db: ReturnType<typeof admin>, title: string) {
  const pdf = await PDFDocument.create();
  const page = pdf.addPage([595.28, 841.89]);
  let font, bold, clean = (t: string) => t;
  try { pdf.registerFontkit(fontkit); font = await pdf.embedFont(await fetchFont("NotoSans-Regular.ttf"), { subset: true }); bold = await pdf.embedFont(await fetchFont("NotoSans-Bold.ttf"), { subset: true }); }
  catch (_) { font = await pdf.embedFont(StandardFonts.Helvetica); bold = await pdf.embedFont(StandardFonts.HelveticaBold); clean = ascii; }
  try {
    const { data } = await db.storage.from("templates").download("logo.png");
    if (data) { const img = await pdf.embedPng(new Uint8Array(await data.arrayBuffer())); const lw = 235, lh = lw * img.height / img.width; page.drawImage(img, { x: 50, y: 841.89 - 36 - lh, width: lw, height: lh }); }
  } catch (_) { page.drawText("CONSORTIA", { x: 50, y: 790, size: 22, font: bold, color: C.brand }); }
  page.drawText(title, { x: 545 - bold.widthOfTextAtSize(title, 26), y: 790, size: 26, font: bold, color: C.brand });
  page.drawLine({ start: { x: 50, y: 740 }, end: { x: 545, y: 740 }, thickness: 1.2, color: C.brand });
  page.drawRectangle({ x: 0, y: 0, width: 595.28, height: 40, color: C.brand });
  page.drawText("Consortia Advisory  \u00b7  info@consortiaadvisory.com  \u00b7  consortiaadvisory.com", { x: 50, y: 15, size: 8.5, font, color: C.on });
  page.drawText("Thriving for Excellence.", { x: 545 - bold.widthOfTextAtSize("Thriving for Excellence.", 8.5), y: 15, size: 8.5, font: bold, color: C.on });
  const T = (t: string, x: number, y: number, size: number, f = font, c = C.ink) => page.drawText(clean(t), { x, y, size, font: f, color: c });
  const R = (t: string, y: number, size: number, f = font, c = C.ink) => page.drawText(clean(t), { x: 545 - f.widthOfTextAtSize(clean(t), size), y, size, font: f, color: c });
  return { pdf, page, font, bold, T, R };
}
function twoColumns(d: Awaited<ReturnType<typeof themedDoc>>, leftTitle: string, left: string[], rightTitle: string, right: string[]) {
  let y = 715;
  d.T(leftTitle, 50, y, 8.5, d.bold, C.grey); d.T(rightTitle, 320, y, 8.5, d.bold, C.grey); y -= 14;
  let yy = y; left.forEach((t, i) => { d.T(t, 50, yy, 9.5, i ? d.font : d.bold); yy -= 13; });
  yy = y; right.forEach((t, i) => { d.T(t, 320, yy, 9.5, i ? d.font : d.bold); yy -= 13; });
}
function infoBand(d: Awaited<ReturnType<typeof themedDoc>>, cells: [string, string][]) {
  const y = 620;
  d.page.drawRectangle({ x: 50, y: y - 34, width: 495, height: 44, color: C.cream, borderColor: C.line, borderWidth: 0.5 });
  let cx = 62; for (const [k, v] of cells) { d.T(k, cx, y - 4, 7.5, d.font, C.grey); d.T(v, cx, y - 22, 10.5, d.bold); cx += 97; }
}
function lineItem(d: Awaited<ReturnType<typeof themedDoc>>, head: string, amountHead: string, title: string, sub: string, amount: string, box: string) {
  let y = 555;
  d.T(head, 50, y, 8.5, d.bold, C.grey); d.R(amountHead, y, 8.5, d.bold, C.grey); y -= 8;
  d.page.drawLine({ start: { x: 50, y }, end: { x: 545, y }, thickness: 0.6, color: C.line }); y -= 20;
  d.T(title, 50, y, 10.5); d.R(amount, y, 10.5, d.bold); y -= 14;
  if (sub) d.T(sub, 50, y, 8.5, d.font, C.grey); y -= 26;
  d.page.drawLine({ start: { x: 50, y }, end: { x: 545, y }, thickness: 0.6, color: C.line }); y -= 24;
  d.page.drawRectangle({ x: 345, y: y - 14, width: 200, height: 34, color: C.accent });
  d.T(box, 357, y - 2, 10, d.bold, C.on); d.page.drawText(ascii(amount), { x: 545 - 12 - d.bold.widthOfTextAtSize(ascii(amount), 13), y: y - 3, size: 13, font: d.bold, color: C.on });
  return y - 70;
}
const addressLines = (inv: InvoiceRow) => [inv.company_name || inv.client_name, ...(inv.company_name ? [`Attn: ${inv.client_name}`] : []), ...(inv.company_address || "").split(/\n|,\s*/).filter(Boolean), ...(inv.company_country ? [inv.company_country] : []), ...(inv.company_tax_id ? [inv.company_tax_id.match(/^\s*(tax|reg|vat|id)/i) ? inv.company_tax_id : `Tax / Reg. No. ${inv.company_tax_id}`] : [])];

async function invoicePdf(db: ReturnType<typeof admin>, inv: InvoiceRow, entity: Entity) {
  const d = await themedDoc(db, "INVOICE");
  const from = [entity.legal_name, ...entity.address_lines, ...(inv.issuing_entity === "CY" ? [`VAT No. ${entity.vat_no ?? ""}`, `TIC No. ${entity.registration_ref ?? ""}`, `Tel ${entity.tel ?? ""}`, entity.email ?? ""].filter((x) => !/ $/.test(x) && x) : [])];
  twoColumns(d, "From", from, "Invoice to", addressLines(inv));
  infoBand(d, [["Invoice no.", invNo(inv)], ["Date", fmtDay(inv.issued_at)], ["Due", inv.due_at ? fmtDay(inv.due_at) : "On receipt"], ["Currency", inv.currency], ["Reference", invNo(inv)]]);
  const [title, sub] = descFor(inv);
  let y = lineItem(d, "Description", "Amount", title, sub, `${inv.currency} ${money(inv.amount)}`, "Total due");
  d.T("Payment details", 50, y, 8.5, d.bold, C.grey); y -= 16;
  for (const [k, v] of [["Account name", inv.account_name], ["Bank", inv.bank], ["IBAN", inv.iban.replace(/(.{4})/g, "$1 ").trim()], ["SWIFT / BIC", inv.swift], ["Reference", invNo(inv)]]) { d.T(k, 50, y, 9.5, d.font, C.grey); d.T(v, 145, y, 9.5, d.bold); y -= 14.5; }
  y -= 10; d.T("Payment is due within 15 days of the invoice date unless stated otherwise. Please quote the reference with your transfer.", 50, y, 8, d.font, C.grey);
  return await d.pdf.save();
}

async function receiptPdf(db: ReturnType<typeof admin>, inv: InvoiceRow, entity: Entity) {
  const d = await themedDoc(db, "RECEIPT");
  twoColumns(d, "From", [entity.legal_name, ...entity.address_lines], "Received from", addressLines(inv).filter((l) => !/^Tax/.test(l)));
  infoBand(d, [["Receipt no.", `R-${invNo(inv)}`], ["Date received", fmtDay(inv.paid_at ?? new Date())], ["For invoice", invNo(inv)], ["Method", "Bank transfer"], ["Currency", inv.currency]]);
  const [title] = descFor(inv);
  let y = lineItem(d, "In respect of", "Amount received", title, "", `${inv.currency} ${money(inv.amount)}`, "Paid in full");
  d.T("Thank you. This receipt confirms settlement of the invoice above in full. No further action is required.", 50, y + 10, 9, d.font, C.grey);
  return await d.pdf.save();
}

// ------------------------------------------------------------------ PDF stamping (signatures, dates)
type Hit = { page: number; x: number; y: number; w: number; h: number; str: string };
async function findText(pdfBytes: Uint8Array, pred: (t: string) => boolean): Promise<Hit[]> {
  const doc = await getDocumentProxy(new Uint8Array(pdfBytes));
  const hits: Hit[] = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    for (const it of content.items as Array<{ str: string; transform: number[]; width: number; height: number }>) {
      if (it.str && pred(it.str.trim())) hits.push({ page: p, x: it.transform[4], y: it.transform[5], w: it.width, h: it.height || Math.abs(it.transform[3]) || 10, str: it.str.trim() });
    }
  }
  return hits;
}
async function stampImage(pdfBytes: Uint8Array, png: Uint8Array, hit: Hit, opts: { above?: number; height?: number; caption?: string }): Promise<Uint8Array> {
  const pdf = await PDFDocument.load(pdfBytes);
  const page = pdf.getPage(hit.page - 1);
  const img = await pdf.embedPng(png);
  const h = opts.height ?? 42, w = h * (img.width / img.height);
  page.drawImage(img, { x: hit.x, y: hit.y + (opts.above ?? 8), width: w, height: h });
  if (opts.caption) {
    const font = await pdf.embedFont(StandardFonts.Helvetica);
    page.drawText(ascii(opts.caption), { x: hit.x, y: hit.y - 22, size: 7, font, color: rgb(0.42, 0.33, 0.34) });
  }
  return new Uint8Array(await pdf.save());
}
function dataUrlToBytes(dataUrl: string): Uint8Array {
  const b = atob(dataUrl.split(",")[1] ?? dataUrl);
  const out = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) out[i] = b.charCodeAt(i); return out;
}
// Kyriacos's signature, embedded so nothing has to be uploaded. A kyriacos-signature.png in the templates bucket overrides it.
const FIRM_SIGNATURE_B64 = "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";
async function firmSignature(db: ReturnType<typeof admin>): Promise<Uint8Array> {
  try { const { data } = await db.storage.from("templates").download("kyriacos-signature.png"); if (data) return new Uint8Array(await data.arrayBuffer()); } catch (_) { /* fall through */ }
  return dataUrlToBytes(FIRM_SIGNATURE_B64);
}
/** Dotted signature lines on the "Yours truly" page: first = Consortia, second = client. */
async function letterLines(pdf: Uint8Array) {
  const dots = await findText(pdf, (t) => /^\.{8,}\s*$/.test(t));
  const yours = await findText(pdf, (t) => /^Yours truly/i.test(t));
  const pg = yours[0]?.page ?? dots[0]?.page;
  const onPage = dots.filter((d) => d.page === pg).sort((a, b) => b.y - a.y);
  return { firm: onPage[0], client: onPage[1], page: pg };
}
async function drawOnLine(pdfBytes: Uint8Array, png: Uint8Array, line: Hit, height: number): Promise<Uint8Array> {
  const pdf = await PDFDocument.load(pdfBytes);
  const page = pdf.getPage(line.page - 1);
  const img = await pdf.embedPng(png);
  page.drawImage(img, { x: line.x + 4, y: line.y - 6, width: height * (img.width / img.height), height });
  return new Uint8Array(await pdf.save());
}
async function drawCaption(pdfBytes: Uint8Array, pageNo: number, x: number, y: number, text: string): Promise<Uint8Array> {
  const pdf = await PDFDocument.load(pdfBytes);
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  pdf.getPage(pageNo - 1).drawText(ascii(text), { x, y, size: 7, font, color: rgb(0.42, 0.33, 0.34) });
  return new Uint8Array(await pdf.save());
}
/** Kyriacos signs the letter when it is generated: on the first dotted line. */
async function signLetterAsFirm(db: ReturnType<typeof admin>, pdf: Uint8Array): Promise<Uint8Array> {
  const png = await firmSignature(db);
  const { firm } = await letterLines(pdf);
  if (!firm) throw new Error("could not find the signature line on the letter");
  return await drawOnLine(pdf, png, firm, 34);
}
/** The client's drawn signature: NDA → above the right-hand "Authorized Signature"; letter → on the second dotted line. */
async function signAsClient(pdf: Uint8Array, kind: string, clientName: string, png: Uint8Array, when: Date, email: string): Promise<Uint8Array> {
  const caption = `Signed electronically by ${clientName} (${email}) on ${when.toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" })} ${when.toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" })} UTC via Solonos`;
  if (kind === "nda") {
    const all = await findText(pdf, () => true);
    const lastPage = Math.max(...all.map((h) => h.page));
    const onLast = all.filter((h) => h.page === lastPage);
    const norm = (t: string) => t.toLowerCase().replace(/\s+/g, " ").trim();
    // right-hand column: the client's name heading, or the "Authorized Signature" label (possibly split into pieces)
    const names = onLast.filter((h) => norm(h.str) === norm(clientName)).sort((a, b) => b.x - a.x);
    const labels = onLast.filter((h) => /authori[sz]ed/i.test(h.str)).sort((a, b) => b.x - a.x);
    const label = labels[0];
    const name = names.find((n) => !label || Math.abs(n.x - label.x) < 60) ?? names[0];
    if (!label && !name) throw new Error("could not find the client's signature block on the NDA");
    // place the signature above the label line, or just under the name heading when the label is missing
    const anchor: Hit = label ?? { ...name!, y: name!.y - 52 };
    const out = await stampImage(pdf, png, anchor, { above: 6, height: 40 });
    return await drawCaption(out, anchor.page, anchor.x, anchor.y - 34, caption);
  }
  const { client, page } = await letterLines(pdf);
  if (!client) return pdf;
  let out = await drawOnLine(pdf, png, client, 32);
  const dates = await findText(out, (t) => /(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}/.test(t));
  const dateHit = dates.filter((d) => d.page === page && d.y < client.y).sort((a, b) => b.y - a.y)[0];
  out = await drawCaption(out, page!, dateHit ? dateHit.x : client.x, dateHit ? dateHit.y - 14 : client.y - 40, caption);
  return out;
}

// ------------------------------------------------------------------ email bodies
/** Plain text typed in Solonos -> paragraphs (blank line = new paragraph, single newline = line break). */
function textToHtml(t: string) {
  const esc = (x: string) => x.replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c] as string));
  return t.replace(/\r/g, "").split(/\n{2,}/).map((p) => `<p>${esc(p).replace(/\n/g, "<br>")}</p>`).join("");
}
function shell(inner: string) {
  return `<div style="font-family:Inter,Arial,sans-serif;max-width:560px;margin:0 auto;padding:28px 24px;color:#2A1E20;line-height:1.55">
    <div style="font:600 20px Georgia,serif;color:#800020;margin-bottom:18px">Consortia Advisory</div>${inner}
    <p style="font-size:12px;color:#6B5457;margin-top:28px">Consortia Advisory · info@consortiaadvisory.com</p></div>`;
}
const button = (href: string, label: string) =>
  `<p style="margin:22px 0"><a href="${href}" style="background:#0C4767;color:#FAF3E9;padding:11px 18px;border-radius:8px;text-decoration:none;font-weight:600">${label}</a></p>`;

// ------------------------------------------------------------------ server
Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (req.method !== "POST") return json({ error: "POST only" }, 405);

  // staff check through the caller's JWT
  const me = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_ANON_KEY")!,
    { global: { headers: { Authorization: req.headers.get("Authorization") ?? "" } }, auth: { persistSession: false } });
  const { data: isStaff } = await me.rpc("is_staff");
  const body = await req.json().catch(() => ({}));
  const { engagement_id, kind, agreement_id, invoice_id } = body;
  if (!["invite", "nda", "letter", "stamp", "receipt", "invoice_pdf", "lead_nda"].includes(kind)) return json({ error: "kind must be invite|nda|letter|stamp|receipt|invoice_pdf|lead_nda" }, 400);
  if (!["stamp", "letter", "nda"].includes(kind) && !isStaff) return json({ error: "staff only" }, 403);
  if (["letter", "nda"].includes(kind) && !isStaff) {   // the client's portal may trigger the automatic NDA / letter for their own engagement
    const { data: mine } = await me.from("engagements").select("id").eq("id", engagement_id).maybeSingle();
    if (!mine) return json({ error: "not your engagement" }, 403);
  }

  const db = admin();
  try {
    // ---- stamp: after signing (client or staff may call; the caller must be able to read the agreement)
    if (kind === "stamp") {
      const { data: visible } = await me.from("agreements").select("id").eq("id", agreement_id).maybeSingle();
      if (!visible) return json({ error: "not your agreement" }, 403);
      const { data: a } = await db.from("agreements").select("*").eq("id", agreement_id).single();
      if (!a || a.status !== "signed") throw new Error("agreement is not signed");
      if (a.signed_pdf_path && !body.force) return json({ ok: true, signed_pdf_path: a.signed_pdf_path, already: true });
      const { data: e } = await db.from("engagements").select("*").eq("id", a.engagement_id).single();
      const { data: sig } = await db.from("signatures").select("*").eq("agreement_id", a.id).order("signed_at", { ascending: false }).limit(1).single();
      const src = a.pdf_path ?? null;
      if (!src) throw new Error("no PDF to stamp (the document was sent as .docx)");
      const [bucket, ...rest] = src.split("/");
      const { data: blob } = await db.storage.from(bucket).download(rest.join("/"));
      let pdf: Uint8Array = new Uint8Array(await blob!.arrayBuffer());
      if (sig?.signature_data) pdf = await signAsClient(pdf, a.kind, e.client_name, dataUrlToBytes(sig.signature_data), new Date(sig.signed_at), sig.signer_email);
      const outPath = rest.join("/").replace(/\.pdf$/i, "") + "-SIGNED.pdf";
      await db.storage.from("agreements").upload(outPath, pdf, { contentType: "application/pdf", upsert: true });
      await db.from("agreements").update({ signed_pdf_path: `agreements/${outPath}` }).eq("id", a.id);
      const label = a.kind === "nda" ? "NDA" : "Engagement Letter";
      const docTitle = a.kind === "nda" ? `${e.company_name || e.client_name} - Non Disclosure Agreement - signed` : `${label} - signed`;
      await db.from("documents").upsert({ engagement_id: e.id, kind: a.kind, name: docTitle, storage_path: `agreements/${outPath}`, visibility: "client", uploaded_by: "system" }, { onConflict: "storage_path", ignoreDuplicates: true });
      const atts = [{ filename: a.kind === "nda" ? `${docTitle}.pdf` : `${label} - ${e.client_name} - signed.pdf`, contentBase64: b64(pdf), contentType: "application/pdf" }];
      let invLine = "";
      if (a.kind === "engagement_letter") {
        // Invoice 1 is issued by the database when the letter is signed; render it, put it on the portal and attach it here
        const { data: inv1 } = await db.from("engagement_invoices").select("id, invoice_no, issuing_entity, amount, due_at, pdf_path").eq("engagement_id", e.id).eq("seq", 1).eq("status", "issued").maybeSingle();
        if (inv1) {
          const r = await renderInvoice(db, inv1.id, false);
          atts.push({ filename: `Invoice ${r.invoice_no}.pdf`, contentBase64: r.base64, contentType: "application/pdf" });
          invLine = `<p>Invoice ${r.invoice_no} for EUR ${money(inv1.amount)} (the first instalment${e.deposit_required ? ", less the deposit" : ""}) is attached and on your Solonos page with the bank details${inv1.due_at ? `, due ${fmtDay(inv1.due_at)}` : ""}. We start work as soon as it is settled.</p>`;
        }
      }
      await sendEmail({
        to: e.client_email, cc: Deno.env.get("MAIL_FROM") ?? undefined,
        subject: `Signed ${label} - ${e.company_name || e.client_name}`,
        html: shell(`<p>Dear ${e.client_name},</p><p>Thank you. Attached is the ${label} signed by both parties, for your records.</p>${a.kind === "nda" ? "<p>You can now upload the financial statements on your Solonos page.</p>" : invLine}<p>Kind regards,<br>Kyriacos<br>Consortia Advisory</p>`),
        attachments: atts,
      });
      return json({ ok: true, signed_pdf_path: `agreements/${outPath}` });
    }

    // ---- invoice_pdf: render (and email) an issued invoice that has no PDF yet
    if (kind === "invoice_pdf") {
      const { data: cur } = await db.from("engagement_invoices").select("id, status").eq("id", invoice_id).single();
      if (cur?.status === "scheduled") await db.rpc("issue_invoice", { p_invoice: invoice_id });   // Invoice 2: number it now
      const out = await renderInvoice(db, invoice_id, true);
      return json({ ok: true, ...out });
    }

    // ---- receipt: after an invoice is marked paid
    if (kind === "receipt") {
      const { data: inv } = await db.from("invoices_view").select("*").eq("id", invoice_id).single();
      if (!inv || inv.status !== "paid") throw new Error("invoice is not marked paid");
      const { data: ent } = await db.from("billing_entities").select("*").eq("code", inv.issuing_entity).single();
      const pdf = await receiptPdf(db, inv, ent);
      const rpath = `${inv.engagement_id}/Receipt-${inv.issuing_entity}-${String(inv.invoice_no).padStart(4, "0")}.pdf`;
      await db.storage.from("invoices").upload(rpath, pdf, { contentType: "application/pdf", upsert: true });
      await db.from("engagement_invoices").update({ receipt_path: `invoices/${rpath}` }).eq("id", inv.id);
      await db.from("documents").upsert({ engagement_id: inv.engagement_id, kind: "invoice", name: `Receipt ${inv.issuing_entity}-${String(inv.invoice_no).padStart(4, "0")}`, storage_path: `invoices/${rpath}`, visibility: "client", uploaded_by: "system" }, { onConflict: "storage_path", ignoreDuplicates: true });
      await sendEmail({
        to: inv.client_email, subject: `Receipt ${inv.issuing_entity}-${String(inv.invoice_no).padStart(4, "0")} - ${inv.company_name || inv.client_name}`,
        html: shell(`<p>Dear ${inv.client_name},</p><p>We have received your payment of EUR ${money(inv.amount)} for invoice ${inv.issuing_entity}-${String(inv.invoice_no).padStart(4, "0")}. The receipt is attached and also available on your Solonos page.</p><p>Kind regards,<br>Kyriacos<br>Consortia Advisory</p>`),
        attachments: [{ filename: `Receipt ${inv.issuing_entity}-${String(inv.invoice_no).padStart(4, "0")}.pdf`, contentBase64: b64(pdf), contentType: "application/pdf" }],
      });
      return json({ ok: true, receipt_path: `invoices/${rpath}` });
    }

    // ---- NDA from the lead, before any engagement exists (staff only). The PDF lives with the lead's documents
    //      and follows the lead into the engagement on conversion.
    if (kind === "lead_nda") {
      const { opportunity_id, party_name, signatory_name, signatory_title, email, address, country, is_individual, send, email_subject, email_text } = body;
      if (!opportunity_id) throw new Error("opportunity_id missing");
      const miss = [!party_name && "company or person name", !address && "address", !is_individual && !country && "country", send && !email && "email"].filter(Boolean);
      if (miss.length) throw new Error(`Missing: ${miss.join(", ")}`);
      const dp = dateParts();
      const merge = { client_name: party_name, client_address: address, client_country: country || "", is_individual: !!is_individual, nda_day: dp.day, nda_day_suffix: dp.suffix, nda_month_year: dp.monthYear };
      const bytes = await fillTemplate(db, "NDA_TEMPLATE.docx", merge);
      const DOCX_T = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
      const title = `${party_name} - Non Disclosure Agreement`;
      const slug = title.replace(/[^A-Za-z0-9.]+/g, "-").replace(/^-+|-+$/g, "");
      const base = `leads/${opportunity_id}/${Date.now()}-${slug}`;
      await db.storage.from("deliverables").upload(`${base}.docx`, bytes, { contentType: DOCX_T, upsert: true });
      const pdf = await docxToPdf(bytes, `NDA-lead-${Date.now()}.docx`);
      if (pdf) await db.storage.from("deliverables").upload(`${base}.pdf`, pdf, { contentType: "application/pdf", upsert: true });
      const storage_path = `deliverables/${base}.${pdf ? "pdf" : "docx"}`;
      const meta = { party_name, signatory_name: signatory_name || "", signatory_title: signatory_title || "", email: email || "", address, country: country || "", is_individual: !!is_individual, sent_at: send ? new Date().toISOString() : null, signed: false };
      const { data: doc, error: de } = await db.from("documents").insert({ opportunity_id, kind: "nda", name: title, storage_path, visibility: "internal", uploaded_by: "staff", meta }).select().single();
      if (de) throw new Error(de.message);
      if (send) {
        const first = String(signatory_name || party_name).trim().split(/\s+/)[0];
        const defaultText = `Dear ${first},\n\nThank you for our call. As discussed, attached is our Mutual Non-Disclosure Agreement so you can share your financial information with full confidentiality.\n\nPlease sign it and send it back by replying to this email. Once we have it, we will send you the next steps.\n\nKind regards,\nKyriacos\nUK-Qualified Chartered Accountant (ICAEW)\nConsortia Advisory`;
        await sendEmail({
          to: email, subject: String(email_subject || title).trim() || title,
          html: shell(textToHtml(String(email_text || "").trim() || defaultText)),
          attachments: [pdf ? { filename: `${title}.pdf`, contentBase64: b64(pdf), contentType: "application/pdf" } : { filename: `${title}.docx`, contentBase64: b64(bytes), contentType: DOCX_T }],
        });
      }
      return json({ ok: true, document: doc, pdf: !!pdf });
    }

    const { data: e, error } = await db.from("engagements").select("*").eq("id", engagement_id).single();
    if (error || !e) throw new Error("engagement not found");
    const missing = [!e.client_name && "signatory name", !e.client_email && "email", !e.company_address && "address", !e.company_country && "country"].filter(Boolean);
    if (kind === "invite" && (!e.client_name || !e.client_email)) throw new Error("Signatory name and email are required");
    const d = dateParts();
    const attachments: { filename: string; contentBase64: string; contentType: string }[] = [];
    const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    if (kind === "invite") {
      // ---- welcome email with the portal link, plus the deposit invoice when a deposit applies
      let depositLine = "";
      if (e.deposit_required && !e.company_country) {
        depositLine = `<li>the <strong>engagement deposit invoice</strong> for EUR ${money(1000)} will follow once you confirm your company details.</li>`;
      } else if (e.deposit_required) {
        await db.rpc("refresh_deposit_entity", { p_engagement: e.id });
        const { data: inv0 } = await db.from("engagement_invoices").select("id, status").eq("engagement_id", e.id).eq("seq", 0).maybeSingle();
        if (inv0 && inv0.status === "scheduled") await db.rpc("issue_invoice", { p_invoice: inv0.id });
        if (inv0) {
          const r = await renderInvoice(db, inv0.id, false);
          attachments.push({ filename: `Invoice ${r.invoice_no}.pdf`, contentBase64: r.base64, contentType: "application/pdf" });
          depositLine = `<li>the <strong>engagement deposit invoice</strong> for EUR ${money(1000)}, payable on acceptance and credited against the first instalment of the professional fee.</li>`;
        }
      }
      const html = shell(`
        <p>Dear ${e.client_name},</p>
        <p>Thank you for our call. To get started on the valuation of ${e.company_name || "your business"}, we have set up a private page for you${depositLine ? ":" : "."}</p>
        ${depositLine ? `<ol>${depositLine}</ol>` : ""}
        <p>On your page you confirm your company details, sign the Non-Disclosure Agreement, upload the financial statements${e.deposit_required ? ", and find the bank details for the deposit" : ""}. It takes a few minutes.</p>
        ${button(PORTAL(), "Open your Solonos page")}
        <p style="font-size:13px;color:#6B5457">Sign in with this email address (${e.client_email}); you'll receive a one-time code, no password needed.</p>
        <p>Kind regards,<br>Kyriacos<br>UK-Qualified Chartered Accountant (ICAEW)<br>Consortia Advisory</p>`);
      await sendEmail({ to: e.client_email, subject: `Your Solonos page - business valuation of ${e.company_name || e.client_name}`, html, attachments });
      await db.from("engagements").update({ sent_at: new Date().toISOString() }).eq("id", e.id);
      await db.from("activity_log").insert({ engagement_id: e.id, actor: "staff", action: "invite.sent", detail: { deposit: e.deposit_required }, visible_to_client: true });
      return json({ ok: true });
    }

    if (kind === "nda") {
      // ---- NDA: generated once the client has confirmed their company details (address + country known)
      const { data: existing } = await db.from("agreements").select("id, version").eq("engagement_id", e.id).eq("kind", "nda").in("status", ["sent", "signed"]).maybeSingle();
      if (existing && !body.force) return json({ ok: true, already: true, version: existing.version });
      if (!e.company_confirmed_at) throw new Error("The client has not confirmed their company details yet");
      if (missing.length) throw new Error(`Missing: ${missing.join(", ")}`);
      const { count } = await db.from("agreements").select("*", { count: "exact", head: true }).eq("engagement_id", e.id).eq("kind", "nda");
      const version = (count ?? 0) + 1;
      const merge = { client_name: e.client_name, client_address: e.company_address, client_country: e.company_country, is_individual: !!e.is_individual, nda_day: d.day, nda_day_suffix: d.suffix, nda_month_year: d.monthYear };
      const bytes = await fillTemplate(db, "NDA_TEMPLATE.docx", merge);
      const path = `${e.id}/NDA-v${version}.docx`;
      await db.storage.from("agreements").upload(path, bytes, { contentType: DOCX, upsert: true });
      const ndaPdf = await docxToPdf(bytes, `NDA-${e.id}-v${version}.docx`);
      const pdfPath = ndaPdf ? `${e.id}/NDA-v${version}.pdf` : null;
      if (ndaPdf) await db.storage.from("agreements").upload(pdfPath!, ndaPdf, { contentType: "application/pdf", upsert: true });
      await db.from("agreements").update({ status: "void" }).eq("engagement_id", e.id).eq("kind", "nda").eq("status", "sent");
      await db.from("agreements").insert({
        engagement_id: e.id, kind: "nda", version, status: "sent", template_version: "nda-2026-09",
        merge_data: merge, docx_path: `agreements/${path}`, pdf_path: pdfPath ? `agreements/${pdfPath}` : null,
        document_sha256: await sha256(ndaPdf ?? bytes), sent_at: new Date().toISOString(),
      });
      const ndaName = `${e.company_name || e.client_name} - Non Disclosure Agreement`;
      attachments.push(ndaPdf ? { filename: `${ndaName}.pdf`, contentBase64: b64(ndaPdf), contentType: "application/pdf" } : { filename: `${ndaName}.docx`, contentBase64: b64(bytes), contentType: DOCX });
      let depositNote = "";
      if (e.deposit_required) {
        await db.rpc("refresh_deposit_entity", { p_engagement: e.id });
        const { data: inv0 } = await db.from("engagement_invoices").select("id, status, pdf_path").eq("engagement_id", e.id).eq("seq", 0).maybeSingle();
        if (inv0 && inv0.status === "scheduled") {
          await db.rpc("issue_invoice", { p_invoice: inv0.id });
          const r = await renderInvoice(db, inv0.id, false);
          attachments.push({ filename: `Invoice ${r.invoice_no}.pdf`, contentBase64: r.base64, contentType: "application/pdf" });
          depositNote = ` The engagement deposit invoice for EUR ${money(1000)} is also attached; bank details are on the invoice and on your page.`;
        }
      }
      await sendEmail({
        to: e.client_email, subject: `Your NDA is ready to sign - ${e.company_name || e.client_name}`,
        html: shell(`<p>Dear ${e.client_name},</p><p>Thank you for confirming your details. Your Non-Disclosure Agreement is attached and ready to sign on your Solonos page. Once signed, you can upload the financial statements.${depositNote}</p>${button(PORTAL(), "Sign the NDA")}<p>Kind regards,<br>Kyriacos<br>Consortia Advisory</p>`),
        attachments,
      });
      await db.from("activity_log").insert({ engagement_id: e.id, actor: "system", action: "nda.sent", detail: { version }, visible_to_client: true });
      return json({ ok: true, version });
    }

    // ---- engagement letter: generated and sent automatically once NDA signed + company confirmed + fee set
    if (kind === "letter") {
      const { data: existing } = await db.from("agreements").select("id, version, status").eq("engagement_id", e.id).eq("kind", "engagement_letter").in("status", ["sent", "signed"]).maybeSingle();
      if (existing && !body.force) return json({ ok: true, already: true, version: existing.version, status: existing.status });
      if (!e.fee || !(e.approaches?.length)) throw new Error("Set the fee and the approaches first");
      if (!e.company_confirmed_at) throw new Error("The client has not confirmed their company details yet");
      const { data: ndaSigned } = await db.from("agreements").select("id").eq("engagement_id", e.id).eq("kind", "nda").eq("status", "signed").maybeSingle();
      if (!ndaSigned) throw new Error("The NDA is not signed yet");
      const { count } = await db.from("agreements").select("*", { count: "exact", head: true }).eq("engagement_id", e.id).eq("kind", "engagement_letter");
      const version = (count ?? 0) + 1;
      const merge = {
        client_name: e.client_name, client_name_upper: String(e.client_name).toUpperCase(), client_title: e.client_title || "Director",
        company_name: e.company_name || e.client_name, company_address: e.company_address || "", company_tax_id: e.company_tax_id || "",
        fee: Number(e.fee).toLocaleString("en-GB", { maximumFractionDigits: 0 }),
        letter_day: d.day, letter_day_suffix: d.suffix, letter_month_year: d.monthYear, letter_date: d.plain,
        approaches: approachLines(e.approaches),
      };
      const bytes = await fillTemplate(db, "Engagement_Letter_TEMPLATE.docx", merge);
      const path = `${e.id}/Engagement-Letter-v${version}.docx`;
      await db.storage.from("agreements").upload(path, bytes, { contentType: DOCX, upsert: true });
      let pdf: Uint8Array | null = await docxToPdf(bytes, `Letter-${e.id}-v${version}.docx`);
      if (pdf) pdf = await signLetterAsFirm(db, pdf);
      const pdfPath = pdf ? `${e.id}/Engagement-Letter-v${version}.pdf` : null;
      if (pdf) await db.storage.from("agreements").upload(pdfPath!, pdf, { contentType: "application/pdf", upsert: true });
      await db.from("agreements").update({ status: "void" }).eq("engagement_id", e.id).eq("kind", "engagement_letter").in("status", ["sent", "draft"]);
      await db.from("agreements").insert({
        engagement_id: e.id, kind: "engagement_letter", version, status: "sent", template_version: "el-2026-09",
        merge_data: merge, docx_path: `agreements/${path}`, pdf_path: pdfPath ? `agreements/${pdfPath}` : null,
        document_sha256: await sha256(pdf ?? bytes), sent_at: new Date().toISOString(),
      });
      const elName = `Engagement Letter - ${e.company_name || e.client_name}`;
      attachments.push(pdf ? { filename: `${elName}.pdf`, contentBase64: b64(pdf), contentType: "application/pdf" }
                           : { filename: `${elName}.docx`, contentBase64: b64(bytes), contentType: DOCX });
      const feeTxt = Number(e.fee).toLocaleString("en-GB", { maximumFractionDigits: 0 });
      const html = shell(`
        <p>Dear ${e.client_name},</p>
        <p>Thank you for signing the NDA and confirming your details. Attached is the engagement letter for the business valuation of ${e.company_name || "your business"}. The professional fee is EUR ${feeTxt}, payable 50% on signing and 50% before release of the final report; the deposit is credited against the first instalment.</p>
        <p>You can read and sign the letter on your Solonos page. Once signed, the first invoice is issued automatically and we start work as soon as it is settled.</p>
        ${button(PORTAL(), "Read and sign the engagement letter")}
        <p>Kind regards,<br>Kyriacos<br>UK-Qualified Chartered Accountant (ICAEW)<br>Consortia Advisory</p>`);
      await sendEmail({ to: e.client_email, subject: `Engagement letter - business valuation of ${e.company_name || e.client_name}`, html, attachments });
      if (["nda_pending", "docs_pending", "review_pending"].includes(e.stage)) await db.rpc("advance_engagement", { p_id: e.id, p_to: "proposal" });
      await db.from("activity_log").insert({ engagement_id: e.id, actor: "system", action: "letter.sent", detail: { version, fee: e.fee }, visible_to_client: true });
      // the lead moves to "Engagement Letter Sent" in GHL (best effort: the stage must exist in the pipeline)
      if (e.opportunity_id) {
        try {
          await fetch(`${Deno.env.get("SUPABASE_URL")}/functions/v1/ghl-sync`, {
            method: "POST", headers: { "content-type": "application/json", "x-sync-secret": Deno.env.get("GHL_SYNC_SECRET") ?? "" },
            body: JSON.stringify({ action: "move", opportunity_id: e.opportunity_id, stage_name: "Engagement Letter Sent" }),
          });
        } catch (err) { console.warn("GHL move after letter failed", err); }
      }
      // the follow-up starts by itself: touch 1 the next working day, then every 3 working days (Part 35)
      try { await db.rpc("followup_begin", { p_opportunity: e.opportunity_id ?? null, p_kind: "letter", p_check_back: null, p_engagement: e.id, p_who: "system" }); }
      catch (err) { console.warn("follow-up after letter failed", err); }
      return json({ ok: true, version });
    }
    return json({ error: "unhandled kind" }, 400);
  } catch (err) {
    return json({ ok: false, error: String(err instanceof Error ? err.message : err) }, 500);
  }
});

/** Render an issued invoice to PDF, store it, register it for the client; optionally email it. */
async function renderInvoice(db: ReturnType<typeof admin>, invoiceId: string, email: boolean) {
  const { data: inv } = await db.from("invoices_view").select("*").eq("id", invoiceId).single();
  if (!inv || !inv.invoice_no) throw new Error("invoice is not issued yet");
  const { data: ent } = await db.from("billing_entities").select("*").eq("code", inv.issuing_entity).single();
  const pdf = await invoicePdf(db, inv, ent);
  const ipath = `${inv.engagement_id}/Invoice-${invNo(inv)}.pdf`;
  await db.storage.from("invoices").upload(ipath, pdf, { contentType: "application/pdf", upsert: true });
  await db.from("engagement_invoices").update({ pdf_path: `invoices/${ipath}` }).eq("id", inv.id);
  await db.from("documents").upsert({ engagement_id: inv.engagement_id, kind: "invoice", name: `Invoice ${invNo(inv)}`, storage_path: `invoices/${ipath}`, visibility: "client", uploaded_by: "system" }, { onConflict: "storage_path", ignoreDuplicates: true });
  if (email) {
    const [title] = descFor(inv);
    await sendEmail({
      to: inv.client_email, subject: `Invoice ${invNo(inv)} - ${inv.company_name || inv.client_name}`,
      html: shell(`<p>Dear ${inv.client_name},</p><p>Please find attached invoice ${invNo(inv)} for EUR ${money(inv.amount)}: ${title.toLowerCase()}. Bank details are on the invoice and on your Solonos page.${inv.due_at ? ` Due ${fmtDay(inv.due_at)}.` : ""}</p><p>Kind regards,<br>Kyriacos<br>Consortia Advisory</p>`),
      attachments: [{ filename: `Invoice ${invNo(inv)}.pdf`, contentBase64: b64(pdf), contentType: "application/pdf" }],
    });
  }
  return { pdf_path: `invoices/${ipath}`, base64: b64(pdf), invoice_no: invNo(inv) };
}

async function sendEmail(m: { to: string; cc?: string; subject: string; html: string; attachments: { filename: string; contentBase64: string; contentType: string }[] }) {
  const r = await fetch(`${Deno.env.get("SUPABASE_URL")}/functions/v1/send-email`, {
    method: "POST",
    headers: { "content-type": "application/json", "x-sync-secret": Deno.env.get("GHL_SYNC_SECRET") ?? "" },
    body: JSON.stringify(m),
  });
  const out = await r.json().catch(() => ({}));
  if (!r.ok || out.ok === false) throw new Error(`email failed: ${out.error ?? r.status}`);
}
