// onedrive — links each engagement to a folder in the Consortia "Common" library, under Common / Projects / <category> / <client>,
//   and syncs files both ways when you press "Sync now" in Solonos.
//   Deploy as: onedrive   (Verify JWT OFF — staff are checked with is_staff on the caller's token)
//   Uses the existing secrets of the "Solonos Mail" app: MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MAIL_FROM
//   (Files.ReadWrite.All application permission, already granted). Needs Part 36.
//
//   POST { action: "discover" }                                   finds Common / Projects and lists its category folders
//   POST { action: "create", engagement_id, category_id, name }   creates (or reuses) the client folder + subfolders, then syncs
//   POST { action: "sync", engagement_id }                        Solonos -> OneDrive for new Solonos files, OneDrive -> Solonos for new OneDrive files
//   POST { action: "roots" | "browse" | "search" | "upload" | "mkdir" }  the OneDrive page in Solonos (partner only)
//   POST { action: "sync_all" } with header x-push-secret         every 15 minutes from Supabase Cron: syncs every linked engagement
//   POST { action: "dataroom", engagement_id }                  makes "Data Room" + the template folders inside the project folder
//   POST { action: "share_dataroom", engagement_id, email, message }  Microsoft invites that email to the Data Room folder (can edit)
// Version 1.3 (23 Oct 2026): data room (template folders, shared with the client through Microsoft); the sync no longer copies
//   anything from the Data Room down to Solonos, and skips OneDrive files over 20 MB.
// Version 1.2 (23 Oct 2026): browse / upload / mkdir with engagement_id work for any staff, but only inside that project's folder
//   (the OneDrive tab on a project). The whole-OneDrive page stays partner only.
// Version 1.1 (23 Oct 2026): OneDrive page (live browse, search, upload, new folder) and automatic sync. 1.0: engagement folders.
import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type", "Access-Control-Allow-Methods": "POST, OPTIONS" };
const json = (b: unknown, s = 200) => new Response(JSON.stringify(b), { status: s, headers: { ...CORS, "content-type": "application/json" } });
const G = "https://graph.microsoft.com/v1.0";

// The folders made inside every client folder, and where each kind of Solonos file goes.
// Data room folders by project type (you can add, rename or delete folders in OneDrive afterwards)
const DATAROOM = "Data Room";
const ROOM_SELL = ["1 Advisory Reports", "2 Corporate Documents", "3 Due Diligence Materials", "4 Financial Documents", "5 Legal Documents", "6 Marketing Materials", "7 Operational Information", "8 Transaction Process Documents"];
const ROOM_BASIC = ["1 Financial Statements", "2 Management Accounts", "3 Budget & Forecast", "4 Corporate & Legal", "5 Contracts", "6 Other"];
const roomFor = (t: string) => (["sell_side", "ma"].includes(t) ? ROOM_SELL : ROOM_BASIC);
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
  const db = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!, { auth: { persistSession: false } });
  const body = await req.json().catch(() => ({}));
  // the 15-minute schedule
  if (body.action === "sync_all") {
    if (!Deno.env.get("PUSH_SECRET") || req.headers.get("x-push-secret") !== Deno.env.get("PUSH_SECRET")) return json({ error: "forbidden" }, 403);
    const { data: list } = await db.from("engagements").select("*").not("onedrive_folder_id", "is", null);
    const out: unknown[] = [];
    for (const e of list ?? []) { try { const r = await sync(db, e); if (r.uploaded || r.downloaded || r.errors.length) out.push({ id: e.id, ...r }); } catch (err) { out.push({ id: e.id, error: String(err).slice(0, 200) }); } }
    return json({ ok: true, engagements: (list ?? []).length, changed: out });
  }
  const me = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: req.headers.get("Authorization") ?? "" } }, auth: { persistSession: false } });
  const { data: isStaff } = await me.rpc("is_staff");
  if (!isStaff) return json({ error: "staff only" }, 403);
  try {
    // ---- the OneDrive tab on a project: any staff, but only inside that project's own folder
    if (["browse", "upload", "mkdir"].includes(body.action) && body.engagement_id) {
      const { data: en } = await db.from("engagements").select("onedrive_drive_id, onedrive_folder_id").eq("id", body.engagement_id).single();
      if (!en?.onedrive_folder_id) throw new Error("This project has no OneDrive folder yet");
      if (body.drive !== en.onedrive_drive_id) return json({ error: "Outside this project's folder" }, 403);
      if (body.item !== en.onedrive_folder_id) {
        const root: any = await g(`/drives/${en.onedrive_drive_id}/items/${en.onedrive_folder_id}?$select=name,parentReference`);
        const rootPath = decodeURIComponent(`${root.parentReference?.path ?? ""}/${root.name}`);
        const it: any = await g(`/drives/${en.onedrive_drive_id}/items/${body.item}?$select=parentReference`);
        const p = decodeURIComponent(it.parentReference?.path ?? "");
        if (!(p === rootPath || p.startsWith(rootPath + "/"))) return json({ error: "Outside this project's folder" }, 403);
      }
      return json({ ok: true, ...(await page(db, body)) });
    }
    // ---- the OneDrive page: partner only, because the Microsoft permission covers the whole company account
    if (["roots", "browse", "search", "upload", "mkdir"].includes(body.action)) {
      const { data: u } = await me.auth.getUser();
      const { data: st } = await db.from("staff").select("role").eq("user_id", u?.user?.id ?? "").maybeSingle();
      if (st?.role !== "partner") return json({ error: "The OneDrive page is for partners only" }, 403);
      return json({ ok: true, ...(await page(db, body)) });
    }
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

    if (body.action === "dataroom") {
      if (!e.onedrive_folder_id) throw new Error("Create the project's OneDrive folder first");
      const room: any = await ensureFolder(e.onedrive_drive_id, e.onedrive_folder_id, DATAROOM);
      for (const n of roomFor(e.engagement_type)) await ensureFolder(e.onedrive_drive_id, room.id, n);
      await db.from("engagements").update({ dataroom_folder_id: room.id, dataroom_url: room.webUrl }).eq("id", e.id);
      return json({ ok: true, url: room.webUrl, folders: roomFor(e.engagement_type) });
    }

    if (body.action === "share_dataroom") {
      if (!e.dataroom_folder_id) throw new Error("Create the data room first");
      const email = String(body.email || "").trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error("Enter the client's email");
      try {
        await g(`/drives/${e.onedrive_drive_id}/items/${e.dataroom_folder_id}/invite`, {
          method: "POST", headers: { "content-type": "application/json" },
          body: JSON.stringify({ recipients: [{ email }], roles: ["write"], requireSignIn: true, sendInvitation: true, message: String(body.message || "").slice(0, 2000) || undefined }),
        });
      } catch (err) {
        const m = String(err);
        if (/sharing|external|guest|policy|notAllowed|accessDenied/i.test(m)) throw new Error("Microsoft refused to share outside Consortia. Turn on external sharing for the Common library (SharePoint admin, Sharing: New and existing guests), then try again.");
        throw err;
      }
      const list = [...(Array.isArray(e.dataroom_shared) ? e.dataroom_shared : []).filter((x: any) => x?.email !== email), { email, at: new Date().toISOString() }];
      await db.from("engagements").update({ dataroom_shared: list }).eq("id", e.id);
      await db.from("activity_log").insert({ engagement_id: e.id, actor: "staff", action: "dataroom.shared", detail: { email }, visible_to_client: true });
      return json({ ok: true, shared: list });
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

const pick = (c: any, driveId: string) => ({ id: c.id, name: c.name, folder: !!c.folder, count: c.folder?.childCount ?? null, size: c.size ?? null, modified: c.lastModifiedDateTime ?? null, url: c.webUrl, drive: c.parentReference?.driveId ?? driveId, path: c.parentReference?.path ?? null });

async function page(db: any, b: any) {
  const user = encodeURIComponent(Deno.env.get("MAIL_FROM")!);
  if (b.action === "roots") {
    const mine: any = await g(`/users/${user}/drive/root?$select=id,name,webUrl,parentReference`);
    const roots: any[] = [{ id: mine.id, drive: mine.parentReference?.driveId, name: "My OneDrive (Consortia)", url: mine.webUrl, folder: true }];
    let common: any = null, note = "";
    try {
      const { data: s } = await db.from("app_settings").select("value").eq("key", "onedrive").maybeSingle();
      common = s?.value?.drive_id ? s.value : await discover();
      if (!s?.value?.drive_id) await db.from("app_settings").upsert({ key: "onedrive", value: common, updated_at: new Date().toISOString() });
      const root: any = await g(`/drives/${common.drive_id}/root?$select=id,name,webUrl`);
      roots.unshift({ id: root.id, drive: common.drive_id, name: "Common", url: root.webUrl, folder: true });
    } catch (e) { note = String(e).slice(0, 300); }
    return { roots, note };
  }
  if (b.action === "browse") {
    const items = (await children(b.drive, b.item)).map((c) => pick(c, b.drive))
      .sort((x, y) => (Number(y.folder) - Number(x.folder)) || x.name.localeCompare(y.name, "en", { numeric: true }));
    const self: any = await g(`/drives/${b.drive}/items/${b.item}?$select=id,name,webUrl`);
    return { items, folder: { id: self.id, name: self.name, url: self.webUrl } };
  }
  if (b.action === "search") {
    const q = String(b.q ?? "").replace(/'/g, "''").trim();
    if (q.length < 2) return { items: [] };
    const drives: string[] = [b.drive].filter(Boolean);
    if (!drives.length) {
      const mine: any = await g(`/users/${user}/drive?$select=id`); drives.push(mine.id);
      const { data: s } = await db.from("app_settings").select("value").eq("key", "onedrive").maybeSingle();
      if (s?.value?.drive_id) drives.unshift(s.value.drive_id);
    }
    const items: any[] = [];
    for (const d of drives) { try { const r: any = await g(`/drives/${d}/root/search(q='${encodeURIComponent(q)}')?$top=50`); items.push(...(r.value ?? []).map((c: any) => pick(c, d))); } catch (_) { /* one drive may refuse */ } }
    return { items: items.slice(0, 80) };
  }
  if (b.action === "mkdir") {
    const f: any = await g(`/drives/${b.drive}/items/${b.item}/children`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ name: safeName(b.name), folder: {}, "@microsoft.graph.conflictBehavior": "rename" }) });
    return { item: pick(f, b.drive) };
  }
  // upload: small files straight in (the page sends files up to 25 MB)
  const bin = Uint8Array.from(atob(String(b.contentBase64 ?? "")), (c) => c.charCodeAt(0));
  const f: any = await g(`/drives/${b.drive}/items/${b.item}:/${encodeURIComponent(safeName(b.name))}:/content?@microsoft.graph.conflictBehavior=rename`, { method: "PUT", body: bin });
  return { item: pick(f, b.drive) };
}

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
    if (subName === DATAROOM) continue;                 // the data room stays in OneDrive only (many, big files)
    for (const c of await children(drive, sid)) {
      if (!c.file || known.has(c.id) || (c.size ?? 0) > 20 * 1024 * 1024) continue;
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
