// onedrive — links each engagement to a folder in the Consortia "Common" library, under Common / Projects / <category> / <client>,
//   and syncs files both ways when you press "Sync now" in Solonos.
//   Deploy as: onedrive   (Verify JWT OFF — staff are checked with is_staff on the caller's token)
//   Uses the existing secrets of the "Solonos Mail" app: MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MAIL_FROM
//   (Files.ReadWrite.All application permission, already granted). Needs Part 36.
//
//   POST { action: "discover" }                                   finds Common / Projects and lists its category folders
//   POST { action: "create", engagement_id, category_id, name }   creates (or reuses) the client folder + subfolders, then syncs
//   POST { action: "sync", engagement_id }                        Solonos -> OneDrive for new Solonos files, OneDrive -> Solonos for new OneDrive files
// Version 1.0 (23 Oct 2026)
import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type", "Access-Control-Allow-Methods": "POST, OPTIONS" };
const json = (b: unknown, s = 200) => new Response(JSON.stringify(b), { status: s, headers: { ...CORS, "content-type": "application/json" } });
const G = "https://graph.microsoft.com/v1.0";

// The folders made inside every client folder, and where each kind of Solonos file goes.
const SUB = { nda: "01 NDA & Letter", fin: "02 Financials", vb: "03 Valuation Builder", draft: "04 Drafts", final: "05 Final report", inv: "06 Invoices", other: "07 Other" };
const kindToSub: Record<string, string> = {
  nda: SUB.nda, engagement_letter: SUB.nda, client_upload: SUB.fin, valuation_builder: SUB.vb,
  draft_report: SUB.draft, final_report: SUB.final, invoice: SUB.inv, six_page_review: SUB.vb,
};
const subToKind: Record<string, string> = { [SUB.draft]: "draft_report", [SUB.final]: "final_report", [SUB.fin]: "client_upload", [SUB.nda]: "nda" };

let token = "";
async function graphToken() {
  if (token) return token;
  const body = new URLSearchParams({ client_id: Deno.env.get("MS_CLIENT_ID")!, client_secret: Deno.env.get("MS_CLIENT_SECRET")!, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const r = await fetch(`https://login.microsoftonline.com/${Deno.env.get("MS_TENANT_ID")}/oauth2/v2.0/token`, { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body });
  const j = await r.json();
  if (!j.access_token) throw new Error(`Microsoft sign-in failed: ${j.error_description ?? j.error ?? r.status}`);
  return (token = j.access_token);
}
async function g(path: string, init: RequestInit = {}) {
  const r = await fetch(path.startsWith("http") ? path : G + path, { ...init, headers: { Authorization: `Bearer ${await graphToken()}`, ...(init.headers ?? {}) } });
  if (!r.ok) throw new Error(`Microsoft ${r.status} on ${path.split("?")[0]}: ${(await r.text()).slice(0, 300)}`);
  return r.status === 204 ? null : r.headers.get("content-type")?.includes("json") ? r.json() : r;
}
const children = async (drive: string, item: string) => {
  const out: any[] = []; let next: string | null = `/drives/${drive}/items/${item}/children?$top=200&$select=id,name,folder,file,webUrl,size,lastModifiedDateTime,eTag`;
  while (next) { const j: any = await g(next); out.push(...(j.value ?? [])); next = j["@odata.nextLink"] ?? null; }
  return out;
};
async function ensureFolder(drive: string, parent: string, name: string) {
  const hit = (await children(drive, parent)).find((c) => c.folder && c.name.toLowerCase() === name.toLowerCase());
  if (hit) return hit;
  return g(`/drives/${drive}/items/${parent}/children`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }) });
}
const safeName = (s: string) => String(s).replace(/["*:<>?/\\|#%]/g, " ").replace(/\s+/g, " ").trim().slice(0, 120) || "Client";

/** Finds Common / Projects: first as a shortcut or folder in the Solonos user's OneDrive, then as a SharePoint site. */
async function discover() {
  const user = encodeURIComponent(Deno.env.get("MAIL_FROM")!);
  let drive = "", root = "", via = "";
  const mine: any = await g(`/users/${user}/drive/root/children?$top=200`);
  const common = (mine.value ?? []).find((c: any) => /^common/i.test(c.name));
  if (common?.remoteItem) { drive = common.remoteItem.parentReference?.driveId; root = common.remoteItem.id; via = "shortcut in OneDrive"; }
  else if (common?.folder) { drive = common.parentReference?.driveId; root = common.id; via = "folder in OneDrive"; }
  if (!drive) {
    try {
      const s: any = await g(`/sites?search=Common`);
      const site = (s.value ?? [])[0];
      if (site) {
        const d: any = await g(`/sites/${site.id}/drives`);
        const lib = (d.value ?? []).find((x: any) => /common/i.test(x.name)) ?? (d.value ?? [])[0];
        if (lib) { drive = lib.id; root = (await g(`/drives/${lib.id}/root`) as any).id; via = `SharePoint site ${site.displayName}`; }
      }
    } catch (e) { throw new Error(`Could not find the Common library (${String(e).slice(0, 200)}). Add the Sites.Read.All permission to the Solonos Mail app.`); }
  }
  if (!drive) throw new Error("Could not find the Common library");
  const projects = (await children(drive, root)).find((c) => c.folder && /^projects$/i.test(c.name));
  if (!projects) throw new Error('Found Common, but no "Projects" folder inside it');
  const cats = (await children(drive, projects.id)).filter((c) => c.folder).map((c) => ({ id: c.id, name: c.name, url: c.webUrl }))
    .sort((a, b) => a.name.localeCompare(b.name, "en", { numeric: true }));
  return { drive_id: drive, projects_id: projects.id, projects_url: projects.webUrl, via, categories: cats, found_at: new Date().toISOString() };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (req.method !== "POST") return json({ error: "POST only" }, 405);
  const me = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: req.headers.get("Authorization") ?? "" } }, auth: { persistSession: false } });
  const { data: isStaff } = await me.rpc("is_staff");
  if (!isStaff) return json({ error: "staff only" }, 403);
  const db = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!, { auth: { persistSession: false } });
  const body = await req.json().catch(() => ({}));
  try {
    if (body.action === "discover") {
      const d = await discover();
      await db.from("app_settings").upsert({ key: "onedrive", value: d, updated_at: new Date().toISOString() });
      return json({ ok: true, ...d });
    }

    const { data: e, error } = await db.from("engagements").select("*").eq("id", body.engagement_id).single();
    if (error || !e) throw new Error("engagement not found");

    if (body.action === "create") {
      const { data: s } = await db.from("app_settings").select("value").eq("key", "onedrive").maybeSingle();
      const cfg: any = s?.value ?? await discover();
      const cat = (cfg.categories ?? []).find((c: any) => c.id === body.category_id);
      if (!cat) throw new Error("Pick the type of project (category) first");
      const folder: any = await ensureFolder(cfg.drive_id, cat.id, safeName(body.name || e.company_name || e.client_name));
      for (const n of Object.values(SUB)) await ensureFolder(cfg.drive_id, folder.id, n);
      await db.from("engagements").update({ onedrive_drive_id: cfg.drive_id, onedrive_folder_id: folder.id, onedrive_url: folder.webUrl, onedrive_path: `Common / Projects / ${cat.name} / ${folder.name}`, onedrive_category: cat.name }).eq("id", e.id);
      e.onedrive_drive_id = cfg.drive_id; e.onedrive_folder_id = folder.id;
      const r = await sync(db, e);
      return json({ ok: true, url: folder.webUrl, path: `Common / Projects / ${cat.name} / ${folder.name}`, ...r });
    }

    if (body.action === "sync") {
      if (!e.onedrive_folder_id) throw new Error("Create the OneDrive folder first");
      return json({ ok: true, ...(await sync(db, e)) });
    }
    return json({ error: "action must be discover, create or sync" }, 400);
  } catch (err) {
    return json({ ok: false, error: String(err instanceof Error ? err.message : err) }, 500);
  }
});

/** Two-way, additive: new Solonos files go up, new OneDrive files come down. Nothing is ever deleted on either side. */
async function sync(db: any, e: any) {
  const drive = e.onedrive_drive_id, root = e.onedrive_folder_id;
  const { data: done } = await db.from("onedrive_files").select("storage_path, item_id").eq("engagement_id", e.id);
  const pushed = new Set((done ?? []).map((x: any) => x.storage_path)), known = new Set((done ?? []).map((x: any) => x.item_id));
  const subs = new Map<string, string>();
  for (const c of await children(drive, root)) if (c.folder) subs.set(c.name, c.id);
  const subId = async (name: string) => subs.get(name) ?? (subs.set(name, (await ensureFolder(drive, root, name)).id), subs.get(name)!);

  // ---- everything Solonos holds for this client
  const files: { path: string; name: string; sub: string }[] = [];
  const docs = await db.from("documents").select("kind, name, storage_path").or(`engagement_id.eq.${e.id}${e.opportunity_id ? `,opportunity_id.eq.${e.opportunity_id}` : ""}`);
  for (const d of docs.data ?? []) if (d.storage_path) files.push({ path: d.storage_path, name: d.name, sub: kindToSub[d.kind] ?? SUB.other });
  const ags = await db.from("agreements").select("kind, status, version, pdf_path, signed_pdf_path").eq("engagement_id", e.id).neq("status", "void");
  for (const a of ags.data ?? []) {
    const label = a.kind === "nda" ? "Non Disclosure Agreement" : "Engagement Letter";
    const who = e.company_name || e.client_name;
    if (a.signed_pdf_path) files.push({ path: a.signed_pdf_path, name: `${who} - ${label} - signed`, sub: SUB.nda });
    else if (a.pdf_path) files.push({ path: a.pdf_path, name: `${who} - ${label} v${a.version}`, sub: SUB.nda });
  }
  const inv = await db.from("engagement_invoices").select("pdf_path, receipt_path, invoice_no, issuing_entity").eq("engagement_id", e.id);
  for (const i of inv.data ?? []) {
    const no = `${i.issuing_entity ?? ""}-${String(i.invoice_no ?? "").padStart(4, "0")}`.replace(/^-/, "");
    if (i.pdf_path) files.push({ path: i.pdf_path, name: `Invoice ${no}`, sub: SUB.inv });
    if (i.receipt_path) files.push({ path: i.receipt_path, name: `Receipt ${no}`, sub: SUB.inv });
  }

  let up = 0, down = 0; const errors: string[] = [];
  for (const f of files) {
    if (pushed.has(f.path)) continue;
    try {
      const [bucket, ...rest] = f.path.split("/");
      const { data: blob, error } = await db.storage.from(bucket).download(rest.join("/"));
      if (error || !blob) throw new Error(error?.message ?? "missing in storage");
      const ext = (rest.join("/").match(/\.[A-Za-z0-9]{2,5}$/) ?? [".pdf"])[0];
      const fname = safeName(f.name.endsWith(ext) ? f.name : f.name + ext);
      const parent = await subId(f.sub);
      const item: any = await g(`/drives/${drive}/items/${parent}:/${encodeURIComponent(fname)}:/content?@microsoft.graph.conflictBehavior=rename`, { method: "PUT", body: blob });
      await db.from("onedrive_files").insert({ engagement_id: e.id, storage_path: f.path, item_id: item.id, name: item.name, direction: "up" });
      known.add(item.id); up++;
    } catch (err) { errors.push(`${f.name}: ${String(err).slice(0, 160)}`); }
  }

  // ---- files you added in OneDrive (any subfolder, one level deep)
  for (const [subName, sid] of subs) {
    for (const c of await children(drive, sid)) {
      if (!c.file || known.has(c.id)) continue;
      try {
        const r = await fetch(`${G}/drives/${drive}/items/${c.id}/content`, { headers: { Authorization: `Bearer ${await graphToken()}` } });
        if (!r.ok) throw new Error(`download ${r.status}`);
        const bytes = new Uint8Array(await r.arrayBuffer());
        const path = `engagements/${e.id}/onedrive/${Date.now()}-${safeName(c.name).replace(/\s+/g, "-")}`;
        const { error: ue } = await db.storage.from("deliverables").upload(path, bytes, { contentType: c.file.mimeType ?? "application/octet-stream", upsert: true });
        if (ue) throw new Error(ue.message);
        const kind = subToKind[subName] ?? "working_paper";
        await db.from("documents").insert({ engagement_id: e.id, kind, name: c.name.replace(/\.[A-Za-z0-9]{2,5}$/, ""), storage_path: `deliverables/${path}`, visibility: "internal", uploaded_by: "onedrive" });
        await db.from("onedrive_files").insert({ engagement_id: e.id, storage_path: `deliverables/${path}`, item_id: c.id, name: c.name, direction: "down" });
        known.add(c.id); down++;
      } catch (err) { errors.push(`${c.name}: ${String(err).slice(0, 160)}`); }
    }
  }
  await db.from("engagements").update({ onedrive_synced_at: new Date().toISOString() }).eq("id", e.id);
  return { uploaded: up, downloaded: down, errors };
}
