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
//   POST { action: "dataroom", engagement_id }                  makes an empty "Data Room" folder inside the project folder
//   POST { action: "share_dataroom", engagement_id, email, message }  Microsoft invites that email to the Data Room folder (can edit)
//   POST { action: "grant", engagement_id, email, scope: "client" }                       client: 3 Deliverables (view), 4 Uploads + 5 Virtual Data Room (edit)
//   POST { action: "grant", engagement_id, email, scope: "folder", folder_id, role, expires_at }  a buyer: one folder inside 5 Virtual Data Room
//   POST { action: "revoke", engagement_id, access_id }                                    removes that access in OneDrive
//   POST { action: "vdr_folders", engagement_id }                                         folders inside the data room (two levels), for the picker
// Version 1.5 (23 Oct 2026): project folder = 1 Engagement Letter & NDA, 2 Workings, 3 Deliverables, 4 Uploads, 5 Virtual Data Room.
//   Access per person and folder, with an end date (the 15-minute run removes expired access). Nothing is copied from OneDrive to Solonos.
// Version 1.4 (23 Oct 2026): the data room starts empty; Kyriacos makes the folders himself.
// Version 1.3 (23 Oct 2026): data room (template folders, shared with the client through Microsoft); the sync no longer copies
//   anything from the Data Room down to Solonos, and skips OneDrive files over 20 MB.
// Version 1.2 (23 Oct 2026): browse / upload / mkdir with engagement_id work for any staff, but only inside that project's folder
//   (the OneDrive tab on a project). The whole-OneDrive page stays partner only.
// Version 1.1 (23 Oct 2026): OneDrive page (live browse, search, upload, new folder) and automatic sync. 1.0: engagement folders.
import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type", "Access-Control-Allow-Methods": "POST, OPTIONS" };
const json = (b: unknown, s = 200) => new Response(JSON.stringify(b), { status: s, headers: { ...CORS, "content-type": "application/json" } });
const G = "https://graph.microsoft.com/v1.0";

// The five folders inside every project folder, and where each kind of Solonos file goes.
//   The client gets 3, 4 and 5. Staff (Kyriacos, Margarita) see everything through Solonos.
const SUB = { nda: "1. Engagement Letter & NDA", work: "2. Workings", deliv: "3. Deliverables", uploads: "4. Uploads", vdr: "5. Virtual Data Room" };
const kindToSub: Record<string, string> = {
  nda: SUB.nda, engagement_letter: SUB.nda, invoice: SUB.nda, client_upload: SUB.uploads,
  valuation_builder: SUB.work, six_page_review: SUB.work, working_paper: SUB.work, other: SUB.work,
  draft_report: SUB.deliv, final_report: SUB.deliv,
};

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
    const now = Date.now();
    for (const e of list ?? []) {
      const due = (Array.isArray(e.dataroom_shared) ? e.dataroom_shared : []).filter((a: any) => !a.revoked_at && a.expires_at && Date.parse(a.expires_at) < now);
      if (due.length) { try { await revoke(db, e, (a: any) => due.some((d: any) => d.id === a.id), "expired"); out.push({ id: e.id, expired: due.map((d: any) => d.email) }); } catch (err) { out.push({ id: e.id, error: String(err).slice(0, 200) }); } }
    }
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
      let vdr: any = null;
      for (const n of Object.values(SUB)) { const f: any = await ensureFolder(cfg.drive_id, folder.id, n); if (n === SUB.vdr) vdr = f; }
      await db.from("engagements").update({ onedrive_drive_id: cfg.drive_id, onedrive_folder_id: folder.id, onedrive_url: folder.webUrl, onedrive_path: `Common / Projects / ${cat.name} / ${folder.name}`, onedrive_category: cat.name, dataroom_folder_id: vdr?.id ?? null, dataroom_url: vdr?.webUrl ?? null }).eq("id", e.id);
      e.onedrive_drive_id = cfg.drive_id; e.onedrive_folder_id = folder.id;
      const r = await sync(db, e);
      return json({ ok: true, url: folder.webUrl, path: `Common / Projects / ${cat.name} / ${folder.name}`, ...r });
    }

    if (body.action === "dataroom") {          // older projects: add the five folders
      if (!e.onedrive_folder_id) throw new Error("Create the project's OneDrive folder first");
      let vdr: any = null;
      for (const n of Object.values(SUB)) { const f: any = await ensureFolder(e.onedrive_drive_id, e.onedrive_folder_id, n); if (n === SUB.vdr) vdr = f; }
      await db.from("engagements").update({ dataroom_folder_id: vdr.id, dataroom_url: vdr.webUrl }).eq("id", e.id);
      return json({ ok: true, url: vdr.webUrl });
    }

    if (body.action === "vdr_folders") {
      if (!e.dataroom_folder_id) return json({ ok: true, folders: [] });
      const out: any[] = [{ id: e.dataroom_folder_id, name: "Whole data room", depth: 0 }];
      for (const c of (await children(e.onedrive_drive_id, e.dataroom_folder_id)).filter((x) => x.folder).sort((x, y) => x.name.localeCompare(y.name, "en", { numeric: true }))) {
        out.push({ id: c.id, name: c.name, depth: 1 });
        for (const d of (await children(e.onedrive_drive_id, c.id)).filter((x) => x.folder).sort((x, y) => x.name.localeCompare(y.name, "en", { numeric: true }))) out.push({ id: d.id, name: `${c.name} / ${d.name}`, depth: 2 });
      }
      return json({ ok: true, folders: out });
    }

    if (body.action === "grant") {
      if (!e.onedrive_folder_id) throw new Error("Create the project's OneDrive folder first");
      const email = String(body.email || "").trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error("Enter a valid email");
      const targets: { item: string; name: string; url: string; role: "read" | "write" }[] = [];
      let label = "";
      if (body.scope === "client") {
        for (const [n, role] of [[SUB.deliv, "read"], [SUB.uploads, "write"], [SUB.vdr, "write"]] as const) {
          const f: any = await ensureFolder(e.onedrive_drive_id, e.onedrive_folder_id, n);
          targets.push({ item: f.id, name: n, url: f.webUrl, role });
        }
        label = "Client: Deliverables (view), Uploads and Virtual Data Room (edit)";
      } else {
        if (!e.dataroom_folder_id) throw new Error("This project has no Virtual Data Room folder yet");
        const it: any = await g(`/drives/${e.onedrive_drive_id}/items/${body.folder_id}?$select=id,name,webUrl,parentReference`);
        if (it.id !== e.dataroom_folder_id) {
          const room: any = await g(`/drives/${e.onedrive_drive_id}/items/${e.dataroom_folder_id}?$select=name,parentReference`);
          const roomPath = decodeURIComponent(`${room.parentReference?.path ?? ""}/${room.name}`);
          if (!decodeURIComponent(it.parentReference?.path ?? "").startsWith(roomPath)) throw new Error("Pick a folder inside 5. Virtual Data Room");
        }
        const role = body.role === "write" ? "write" : "read";
        targets.push({ item: it.id, name: it.id === e.dataroom_folder_id ? SUB.vdr : it.name, url: it.webUrl, role });
        label = `${it.id === e.dataroom_folder_id ? "Whole data room" : it.name} (${role === "write" ? "view and upload" : "view only"})`;
      }
      const items: any[] = [];
      for (const t of targets) {
        let r: any;
        try {
          r = await g(`/drives/${e.onedrive_drive_id}/items/${t.item}/invite`, {
            method: "POST", headers: { "content-type": "application/json" },
            body: JSON.stringify({ recipients: [{ email }], roles: [t.role], requireSignIn: true, sendInvitation: true, message: String(body.message || "").slice(0, 2000) || undefined }),
          });
        } catch (err) {
          const m = String(err);
          if (/sharing|external|guest|policy|notAllowed|accessDenied/i.test(m)) throw new Error(`Microsoft refused to share with ${email}: external sharing is off for this library. ${m.slice(0, 200)}`);
          throw err;
        }
        items.push({ item: t.item, name: t.name, url: t.url, role: t.role, perms: (r?.value ?? []).map((p: any) => p.id).filter(Boolean) });
      }
      const expires = body.expires_at ? new Date(`${String(body.expires_at).slice(0, 10)}T23:59:00`).toISOString() : null;
      const entry = { id: crypto.randomUUID(), email, scope: body.scope === "client" ? "client" : "folder", label, items, expires_at: expires, at: new Date().toISOString(), revoked_at: null };
      const list = [...(Array.isArray(e.dataroom_shared) ? e.dataroom_shared : []), entry];
      await db.from("engagements").update({ dataroom_shared: list }).eq("id", e.id);
      await db.from("activity_log").insert({ engagement_id: e.id, actor: "staff", action: "access.granted", detail: { email, label, expires_at: expires }, visible_to_client: false });
      return json({ ok: true, access: list });
    }

    if (body.action === "revoke") {
      const list = await revoke(db, e, (a: any) => a.id === body.access_id, "removed by staff");
      return json({ ok: true, access: list });
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
  for (const d of docs.data ?? []) if (d.storage_path) files.push({ path: d.storage_path, name: d.name, sub: kindToSub[d.kind] ?? SUB.work });
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
    if (i.pdf_path) files.push({ path: i.pdf_path, name: `Invoice ${no}`, sub: SUB.nda });
    if (i.receipt_path) files.push({ path: i.receipt_path, name: `Receipt ${no}`, sub: SUB.nda });
  }

  let up = 0; const down = 0; const errors: string[] = [];
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

  // Files added in OneDrive stay in OneDrive (often many and large): nothing is copied down to Solonos.
  await db.from("engagements").update({ onedrive_synced_at: new Date().toISOString() }).eq("id", e.id);
  return { uploaded: up, downloaded: down, errors };
}

/** Removes the matching access entries in OneDrive and marks them removed. */
async function revoke(db: any, e: any, match: (a: any) => boolean, why: string) {
  const list = (Array.isArray(e.dataroom_shared) ? e.dataroom_shared : []).map((a: any) => ({ ...a }));
  for (const a of list) {
    if (a.revoked_at || !match(a)) continue;
    for (const it of a.items ?? []) for (const p of it.perms ?? []) {
      try { await g(`/drives/${e.onedrive_drive_id}/items/${it.item}/permissions/${p}`, { method: "DELETE" }); }
      catch (err) { if (!/404/.test(String(err))) throw err; }
    }
    a.revoked_at = new Date().toISOString(); a.revoked_why = why;
    await db.from("activity_log").insert({ engagement_id: e.id, actor: why === "expired" ? "system" : "staff", action: "access.removed", detail: { email: a.email, label: a.label, why }, visible_to_client: false });
  }
  await db.from("engagements").update({ dataroom_shared: list }).eq("id", e.id);
  e.dataroom_shared = list;
  return list;
}
