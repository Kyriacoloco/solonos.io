// ghl-sync — single-file version for the Supabase Dashboard editor (Edge Functions → Deploy a new function → Via Editor).
// Version 1.1 (18 Sep 2026): also stores the lead's booked call (GHL calendar appointment) on the opportunity. Needs Part 25.
// Version 1.3 (9 Oct 2026): call times from GHL that carry no timezone are read as Europe/Athens (they were being stored as UTC, showing 3 hours late).
// Version 1.2 (2 Oct 2026): keeps a no-show as the call status (for the ghost follow-ups), "move" accepts stage_name,
//   new action "note" adds a note to the GHL contact, and the secret header may call move/note (used by engagement-send).
// Name the function exactly:  ghl-sync   and turn OFF 'Verify JWT'.
// Secrets: GHL_TOKEN, GHL_LOCATION_ID, GHL_SYNC_SECRET, GHL_PIPELINE, GHL_FROM_STAGE
// Token scopes: View Contacts, View Opportunities, View Locations, View Custom Fields

import { createClient } from "npm:@supabase/supabase-js@2";

/** Service-role client. Bypasses RLS — only ever used inside Edge Functions. */
function adminClient() {
  return createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    { auth: { persistSession: false } },
  );
}

/** Client bound to the caller's JWT so RLS + is_staff() apply. */
function userClient(req: Request) {
  return createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_ANON_KEY")!,
    { global: { headers: { Authorization: req.headers.get("Authorization") ?? "" } }, auth: { persistSession: false } },
  );
}

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type, x-sync-secret",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
};
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json", ...CORS } });

function requireSecret(req: Request, header: string, envName: string) {
  const expected = Deno.env.get(envName);
  if (!expected) throw new Error(`${envName} not set`);
  const got = req.headers.get(header) ?? new URL(req.url).searchParams.get("secret");
  if (got !== expected) return json({ error: "unauthorised" }, 401);
  return null;
}

// Minimal GoHighLevel API v2 client. Read-only. Private Integration Token auth.
// Base URL and Version header are GHL's; if GHL bumps the version header, change it here only.

const BASE = "https://services.leadconnectorhq.com";
const VERSION = "2021-07-28";

function headers() {
  return {
    Authorization: `Bearer ${Deno.env.get("GHL_TOKEN")!}`,
    Version: VERSION,
    Accept: "application/json",
  };
}

async function get<T>(path: string, params: Record<string, string> = {}): Promise<T> {
  const url = new URL(BASE + path);
  for (const [k, v] of Object.entries(params)) url.searchParams.set(k, v);
  let attempt = 0;
  while (true) {
    const res = await fetch(url, { headers: headers() });
    if (res.status === 429 && attempt < 4) {            // GHL rate limit: back off and retry
      attempt++;
      await new Promise((r) => setTimeout(r, 500 * 2 ** attempt));
      continue;
    }
    if (!res.ok) throw new Error(`GHL ${res.status} ${path}: ${await res.text()}`);
    return (await res.json()) as T;
  }
}

type GhlStage = { id: string; name: string; position: number };
type GhlPipeline = { id: string; name: string; stages: GhlStage[] };
type GhlOpportunity = {
  id: string; name: string; pipelineId: string; pipelineStageId: string; status: string;
  monetaryValue?: number; source?: string; assignedTo?: string; contactId?: string;
  createdAt?: string; updatedAt?: string; contact?: Record<string, unknown>;
  [k: string]: unknown;
};
type GhlContact = {
  id: string; firstName?: string; lastName?: string; email?: string; phone?: string;
  companyName?: string; country?: string;
  customFields?: { id: string; value?: unknown; field_value?: unknown }[];
  [k: string]: unknown;
};

const loc = () => Deno.env.get("GHL_LOCATION_ID")!;

async function pipelines(): Promise<GhlPipeline[]> {
  const r = await get<{ pipelines: GhlPipeline[] }>("/opportunities/pipelines", { locationId: loc() });
  return r.pipelines ?? [];
}

/** All opportunities with the given status (default: open). Pages until GHL stops returning rows. */
async function searchOpportunities(status = "open"): Promise<GhlOpportunity[]> {
  const out: GhlOpportunity[] = [];
  let page = 1;
  while (true) {
    // note: this endpoint takes snake_case location_id, unlike the rest of the API
    const r = await get<{ opportunities: GhlOpportunity[]; meta?: { total?: number; nextPage?: number | null } }>(
      "/opportunities/search", { location_id: loc(), status, limit: "100", page: String(page) },
    );
    const rows = r.opportunities ?? [];
    out.push(...rows);
    if (rows.length < 100 || r.meta?.nextPage == null) break;
    page = r.meta.nextPage;
  }
  return out;
}

async function moveOpportunity(id: string, stageId: string): Promise<void> {
  const res = await fetch(`${BASE}/opportunities/${id}`, { method: "PUT", headers: { ...headers(), "content-type": "application/json" }, body: JSON.stringify({ pipelineStageId: stageId }) });
  if (!res.ok) throw new Error(`GHL ${res.status} moving opportunity: ${await res.text()}`);
}

async function addContactNote(contactId: string, body: string): Promise<void> {
  const res = await fetch(`${BASE}/contacts/${contactId}/notes`, { method: "POST", headers: { ...headers(), "content-type": "application/json" }, body: JSON.stringify({ body }) });
  if (!res.ok) throw new Error(`GHL ${res.status} adding note: ${await res.text()}`);
}

async function opportunity(id: string): Promise<GhlOpportunity> {
  const r = await get<{ opportunity: GhlOpportunity }>(`/opportunities/${id}`);
  return r.opportunity;
}

type GhlCustomField = { id: string; name: string; fieldKey?: string; dataType?: string; position?: number };
async function customFields(): Promise<GhlCustomField[]> {
  const r = await get<{ customFields: GhlCustomField[] }>(`/locations/${loc()}/customFields`);
  return r.customFields ?? [];
}

async function contact(id: string): Promise<GhlContact> {
  const r = await get<{ contact: GhlContact }>(`/contacts/${id}`);
  return r.contact;
}
const ghl = { pipelines, searchOpportunities, opportunity, contact, customFields, moveOpportunity, addContactNote };


// ghl-sync — mirrors one GHL pipeline into pipeline_stages / contacts / opportunities.
//
//   GET  /ghl-sync?mode=full   header x-sync-secret     → full pull of stages + all open opportunities (+ their contacts)
//   POST /ghl-sync             header x-sync-secret     → webhook from a GHL workflow. The payload is treated as a
//                                                          doorbell only: we read the opportunity id out of it and
//                                                          re-fetch the record from the API, so payload shape never matters.
//
// Deploy with --no-verify-jwt (GHL can't send a Supabase JWT). The shared secret is the auth.


type Db = ReturnType<typeof adminClient>;

// id → name map for custom fields, refreshed at most once per invocation
let fieldNames: Record<string, string> | null = null;
async function loadFieldNames(db: Db): Promise<Record<string, string>> {
  if (fieldNames) return fieldNames;
  const defs = await ghl.customFields();
  for (const f of defs) {
    await db.from("ghl_custom_fields").upsert({ ghl_id: f.id, name: f.name, field_key: f.fieldKey ?? null, data_type: f.dataType ?? null, position: f.position ?? null, synced_at: new Date().toISOString() });
  }
  fieldNames = Object.fromEntries(defs.map((f) => [f.id, f.name]));
  return fieldNames;
}

async function upsertContact(db: Db, c: GhlContact) {
  const names = await loadFieldNames(db);
  const custom: Record<string, unknown> = {};
  for (const f of c.customFields ?? []) {
    const v = f.value ?? f.field_value;
    if (v == null || v === "") continue;
    custom[names[f.id] ?? f.id] = Array.isArray(v) ? v.join(", ") : v;
  }
  await db.from("contacts").upsert({
    ghl_id: c.id, first_name: c.firstName ?? null, last_name: c.lastName ?? null,
    email: c.email ?? null, phone: c.phone ?? null, company: c.companyName ?? null,
    country: c.country ?? null, custom, raw: c, synced_at: new Date().toISOString(),
  });
}

type GhlAppointment = { id: string; title?: string; startTime?: string; endTime?: string; appointmentStatus?: string; calendarId?: string };

const CAL_TZ = Deno.env.get("GHL_CALENDAR_TZ") ?? "Europe/Athens";
/** GHL sometimes returns "2026-10-09T11:00:00" with no offset: that is calendar-local time, not UTC. */
function withZone(t: string): string {
  if (!t || /(Z|[+-]\d{2}:?\d{2})$/i.test(t)) return t;
  const naive = new Date(t + "Z");                                                   // pretend UTC to get the parts
  const parts = Object.fromEntries(new Intl.DateTimeFormat("en-GB", { timeZone: CAL_TZ, timeZoneName: "longOffset", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" }).formatToParts(naive).map((p) => [p.type, p.value]));
  const m = /GMT([+-]\d{2}):?(\d{2})?/.exec(parts.timeZoneName ?? "GMT+00:00");
  const off = m ? `${m[1]}:${m[2] ?? "00"}` : "+00:00";
  return t.replace(/(\.\d+)?$/, "") + off;
}
let appointmentsBlocked = false;   // set when the token lacks the calendar scope, so we stop asking

/** The lead's call: the next upcoming appointment that is not cancelled, else the most recent past one. */
async function contactCall(db: Db, contactId: string): Promise<{ call_at: string | null; call_status: string | null; call_title: string | null }> {
  const none = { call_at: null, call_status: null, call_title: null };
  if (appointmentsBlocked) return none;
  let events: GhlAppointment[] = [];
  try {
    const r = await get<{ events?: GhlAppointment[] }>(`/contacts/${contactId}/appointments`);
    events = r.events ?? [];
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    if (/GHL 40[13]/.test(msg)) {
      appointmentsBlocked = true;
      await db.from("sync_log").insert({ kind: "appointments", ok: false, detail: { error: "GHL token has no calendar access; add the calendars/events.readonly scope to see call dates", raw: msg.slice(0, 300) } });
      return none;
    }
    throw e;
  }
  // cancelled/invalid are ignored; a no-show stays, so the board can flag it
  const live = events.filter((a) => a.startTime && !/cancel|invalid/i.test(a.appointmentStatus ?? "")).map((a) => ({ ...a, startTime: withZone(a.startTime!) }));
  const now = Date.now();
  const upcoming = live.filter((a) => new Date(a.startTime!).getTime() >= now && !/noshow|no_show/i.test(a.appointmentStatus ?? "")).sort((a, b) => a.startTime!.localeCompare(b.startTime!));
  const past = live.filter((a) => new Date(a.startTime!).getTime() < now).sort((a, b) => b.startTime!.localeCompare(a.startTime!));
  const pick = upcoming[0] ?? past[0];
  if (!pick) return none;
  return { call_at: pick.startTime!, call_status: pick.appointmentStatus ?? null, call_title: pick.title ?? null };
}

async function upsertOpportunity(db: Db, o: GhlOpportunity) {
  if (o.contactId) {
    // fetch the contact only if we don't have it yet or it's older than a day
    const { data } = await db.from("contacts").select("synced_at").eq("ghl_id", o.contactId).maybeSingle();
    const stale = !data || Date.now() - new Date(data.synced_at).getTime() > 86_400_000;
    if (stale) await upsertContact(db, await ghl.contact(o.contactId));
  }
  const call = o.contactId ? await contactCall(db, o.contactId) : { call_at: null, call_status: null, call_title: null };
  await db.from("opportunities").upsert({
    ghl_id: o.id, contact_id: o.contactId ?? null, name: o.name ?? "(unnamed)",
    stage_id: o.pipelineStageId ?? null, status: o.status ?? "open",
    monetary_value: o.monetaryValue ?? null, source: o.source ?? null, assigned_to: o.assignedTo ?? null,
    ghl_created_at: o.createdAt ?? null, ghl_updated_at: o.updatedAt ?? null,
    ...call,
    raw: o, synced_at: new Date().toISOString(),
  });
}

async function fullPull(db: Db) {
  const pipeline = await pickPipeline();
  for (const s of pipeline.stages) {
    await db.from("pipeline_stages").upsert({ ghl_id: s.id, name: s.name, position: s.position, synced_at: new Date().toISOString() });
  }
  const opps = (await ghl.searchOpportunities("open")).filter((o) => stageAllowed(o.pipelineStageId, pipeline.stages));
  for (const o of opps) await upsertOpportunity(db, o);
  await loadFieldNames(db);
  return { pipeline: pipeline.name, stages: pipeline.stages.length, opportunities: opps.length, custom_fields: Object.keys(fieldNames ?? {}).length, from_stage: fromStage() || null, only_stages: onlyStages() };
}

/**
 * Which leads reach the platform:
 *   GHL_PIPELINE    = "Business Valuation"   pick the pipeline whose name contains this (case-insensitive). Empty = first pipeline.
 *   GHL_FROM_STAGE  = "Free Call Booked"     mirror this stage and every stage to its right. Empty = all stages.
 *   GHL_ONLY_STAGES = "A, B"                 alternative: an explicit list of stage names. Ignored when GHL_FROM_STAGE is set.
 */
function onlyStages(): string[] {
  return (Deno.env.get("GHL_ONLY_STAGES") ?? "").split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
}
function fromStage(): string { return (Deno.env.get("GHL_FROM_STAGE") ?? "").trim().toLowerCase(); }
function stageAllowed(stageId: string | undefined, stages: GhlStage[]): boolean {
  const stage = stages.find((s) => s.id === stageId);
  const from = fromStage();
  if (from) {
    const start = stages.find((s) => s.name.toLowerCase() === from);
    if (!start) throw new Error(`GHL_FROM_STAGE "${from}" is not a stage in this pipeline: ${stages.map((s) => s.name).join(" | ")}`);
    return !!stage && stage.position >= start.position;
  }
  const only = onlyStages();
  if (only.length === 0) return true;
  return only.includes(stage?.name?.toLowerCase() ?? "");
}
async function pickPipeline(): Promise<GhlPipeline> {
  const all = await ghl.pipelines();
  const want = (Deno.env.get("GHL_PIPELINE") ?? "").trim().toLowerCase();
  const p = want ? all.find((x) => x.name.toLowerCase().includes(want)) : all[0];
  if (!p) throw new Error(`No pipeline matching "${want}". Pipelines: ${all.map((x) => x.name).join(" | ")}`);
  return p;
}

/** Pull an opportunity id out of whatever GHL's workflow webhook sent. */
function findOpportunityId(body: unknown): string | null {
  if (!body || typeof body !== "object") return null;
  const b = body as Record<string, unknown>;
  const direct = b.opportunity_id ?? b.opportunityId ?? (b.opportunity as Record<string, unknown> | undefined)?.id ?? b.id;
  if (typeof direct === "string" && direct.length > 10) return direct;
  for (const v of Object.values(b)) {
    if (v && typeof v === "object") { const r = findOpportunityId(v); if (r) return r; }
  }
  return null;
}

// Staff signed in to the app may trigger a full pull with their own login (no secret needed)
async function isStaffCaller(req: Request): Promise<boolean> {
  if (!req.headers.get("Authorization")) return false;
  const { data } = await userClient(req).rpc("is_staff");
  return data === true;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });   // browser preflight
  const url = new URL(req.url);
  const fullPullRequested = url.searchParams.get("mode") === "full";
  const denied = requireSecret(req, "x-sync-secret", "GHL_SYNC_SECRET");
  const staffCall = fullPullRequested || (req.method === "POST" && (req.headers.get("content-type") ?? "").includes("json") && req.headers.get("Authorization"));
  if (denied && !(staffCall && await isStaffCaller(req))) return denied;
  const db = adminClient();

  if (fullPullRequested) {
    try {
      const detail = await fullPull(db);
      await db.from("sync_log").insert({ kind: "full_pull", ok: true, detail });
      return json({ ok: true, ...detail });
    } catch (e) {
      await db.from("sync_log").insert({ kind: "full_pull", ok: false, detail: { error: String(e) } });
      return json({ ok: false, error: String(e) }, 500);
    }
  }

  if (req.method === "POST") {
    let body: unknown = null;
    try { body = await req.json(); } catch { /* GHL sometimes sends form-encoded; fall through */ }
    // staff drags a lead to another column: push the stage to GHL, then re-read it
    const b = body as Record<string, unknown> | null;
    const trusted = async () => (await isStaffCaller(req)) || req.headers.get("x-sync-secret") === Deno.env.get("GHL_SYNC_SECRET");
    // adds a note on the GHL contact (follow-up touches, so Jesper sees them too)
    if (b?.action === "note") {
      if (!(await trusted())) return json({ error: "staff only" }, 403);
      try {
        let contactId = b.contact_id ? String(b.contact_id) : "";
        if (!contactId && b.opportunity_id) {
          const { data } = await db.from("opportunities").select("contact_id").eq("ghl_id", String(b.opportunity_id)).maybeSingle();
          contactId = data?.contact_id ?? "";
        }
        if (!contactId) throw new Error("no contact for this lead");
        await ghl.addContactNote(contactId, String(b.body ?? ""));
        return json({ ok: true });
      } catch (e) {
        await db.from("sync_log").insert({ kind: "webhook", event: "note", ok: false, detail: { error: String(e), id: b.opportunity_id } });
        return json({ ok: false, error: String(e) }, 500);
      }
    }
    if (b?.action === "move") {
      if (!(await trusted())) return json({ error: "staff only" }, 403);
      try {
        let stageId = b.stage_id ? String(b.stage_id) : "";
        if (!stageId && b.stage_name) {   // e.g. "Engagement Letter Sent"
          const { data } = await db.from("pipeline_stages").select("ghl_id").ilike("name", String(b.stage_name)).maybeSingle();
          if (!data) throw new Error(`stage "${b.stage_name}" not found; add it in GHL and sync`);
          stageId = data.ghl_id;
        }
        await ghl.moveOpportunity(String(b.opportunity_id), stageId);
        const o = await ghl.opportunity(String(b.opportunity_id));
        await upsertOpportunity(db, o);
        await db.from("sync_log").insert({ kind: "webhook", event: "move", ok: true, detail: { id: b.opportunity_id, stage: stageId } });
        return json({ ok: true });
      } catch (e) {
        return json({ ok: false, error: String(e) }, 500);
      }
    }
    const id = findOpportunityId(body);
    const event = (body as Record<string, unknown> | null)?.type ?? (body as Record<string, unknown> | null)?.event ?? "webhook";
    if (!id) {
      await db.from("sync_log").insert({ kind: "webhook", event: String(event), ok: false, detail: { error: "no opportunity id in payload", body } });
      return json({ ok: false, error: "no opportunity id in payload" }, 400);
    }
    try {
      const o = await ghl.opportunity(id);
      const pipeline = await pickPipeline();
      if (o.pipelineId !== pipeline.id || o.status !== "open" || !stageAllowed(o.pipelineStageId, pipeline.stages)) {
        // keep the row (engagements may point at it) but take it off the board: leads_board only shows status = 'open'
        await upsertOpportunity(db, { ...o, status: o.status === "open" ? "filtered" : o.status });
        await db.from("sync_log").insert({ kind: "webhook", event: String(event), ok: true, detail: { id, dropped: true, stage: o.pipelineStageId, status: o.status } });
        return json({ ok: true, id, dropped: true });
      }
      await upsertOpportunity(db, o);
      await db.from("sync_log").insert({ kind: "webhook", event: String(event), ok: true, detail: { id, stage: o.pipelineStageId, status: o.status } });
      return json({ ok: true, id });
    } catch (e) {
      await db.from("sync_log").insert({ kind: "webhook", event: String(event), ok: false, detail: { id, error: String(e) } });
      return json({ ok: false, error: String(e) }, 500);
    }
  }

  return json({ error: "use GET ?mode=full or POST webhook" }, 405);
});
