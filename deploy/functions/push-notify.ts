// push-notify — pings your phone for tasks that are due (each task at its time) and sends the morning digest
//   (08:30 Athens on weekdays, by push and by email to info@, skipped when there is nothing to do).
//   Run every 5 minutes by Supabase Cron (see PART35_CRON.sql). POST with header x-push-secret: <PUSH_SECRET>.
//   Secrets: VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY, VAPID_SUBJECT (mailto:info@consortiaadvisory.com), PUSH_SECRET,
//            GHL_SYNC_SECRET (to call send-email), DIGEST_EMAIL (optional, default info@consortiaadvisory.com)
// Version 2.0 (15 Oct 2026): morning digest; the separate 09:00 overdue ping is gone (the digest carries the overdue count).
import { createClient } from "npm:@supabase/supabase-js@2";
import webpush from "npm:web-push@3.6.7";

const json = (b: unknown, s = 200) => new Response(JSON.stringify(b), { status: s, headers: { "content-type": "application/json" } });
const SITE = "https://www.solonos.io";

Deno.serve(async (req) => {
  if (req.method !== "POST") return json({ error: "POST only" }, 405);
  if (req.headers.get("x-push-secret") !== Deno.env.get("PUSH_SECRET")) return json({ error: "forbidden" }, 403);
  const db = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
  webpush.setVapidDetails(Deno.env.get("VAPID_SUBJECT") ?? "mailto:info@consortiaadvisory.com", Deno.env.get("VAPID_PUBLIC_KEY")!, Deno.env.get("VAPID_PRIVATE_KEY")!);
  const body = await req.json().catch(() => ({}));

  const { data: subs } = await db.from("push_subscriptions").select("*");
  const notes: { title: string; body: string; url: string; tag: string }[] = [];
  const out: Record<string, unknown> = { ok: true };

  // ---- the morning digest (once a day; the database decides whether it is time)
  if (body.digest !== false) {
    const { data: d, error } = await db.rpc("morning_digest_due");
    if (error) console.warn("digest", error.message);
    if (d) {
      const text = digestText(d);
      notes.push({ title: text.title, body: text.body, url: "/app/tasks", tag: "digest" });
      try { await sendEmail({ to: Deno.env.get("DIGEST_EMAIL") || "info@consortiaadvisory.com", subject: text.title, html: text.html }); out.digest_email = true; }
      catch (e) { console.warn("digest email", e); out.digest_email = String(e); }
      out.digest = true;
    }
  }

  // ---- each task at its own time
  const { data: tasks, error } = await db.rpc("tasks_to_push");
  if (error) return json({ ok: false, error: error.message }, 500);
  const due = (tasks ?? []).filter((t: any) => t.reason === "due");
  for (const t of due) {
    const who = String(t.who ?? "").split("|")[0].trim() || "Task";
    const url = t.opportunity_id ? `/app/leads?open=${encodeURIComponent(t.opportunity_id)}` : t.engagement_id ? `/app/engagements/${t.engagement_id}` : "/app/tasks";
    notes.push({ title: `${who} · ${label(t.kind)}`, body: t.title, url, tag: `task-${t.id}` });
  }

  if (!notes.length) return json({ ...out, sent: 0 });
  if (!subs?.length) { if (due.length) await db.rpc("task_mark_pushed", { p_ids: due.map((t: any) => t.id), p_reason: "due" }); return json({ ...out, sent: 0, note: "no devices" }); }

  let sent = 0; const dead: string[] = [];
  for (const s of subs) {
    const sub = { endpoint: s.endpoint, keys: s.keys };
    for (const n of notes) {
      try { await webpush.sendNotification(sub, JSON.stringify(n), { TTL: 3600 }); sent++; }
      catch (e: any) { if (e?.statusCode === 404 || e?.statusCode === 410) { dead.push(s.id); break; } console.warn("push failed", e?.statusCode, e?.body); }
    }
  }
  if (dead.length) await db.from("push_subscriptions").delete().in("id", dead);
  await db.from("push_subscriptions").update({ last_ok_at: new Date().toISOString() }).not("id", "in", `(${dead.map((d) => `"${d}"`).join(",") || '"00000000-0000-0000-0000-000000000000"'})`);
  if (due.length) await db.rpc("task_mark_pushed", { p_ids: due.map((t: any) => t.id), p_reason: "due" });
  return json({ ...out, sent, devices: subs.length - dead.length, removed: dead.length });
});

function label(kind: string) {
  return ({ call: "Call", whatsapp: "WhatsApp", voicenote: "Voice note", email: "Email", meeting: "Meeting", checkin: "Check-in" } as Record<string, string>)[kind] ?? "Task";
}
const plural = (n: number, w: string) => `${n} ${w}${n === 1 ? "" : "s"}`;

/** "Thu 15 Oct: 4 to contact (2 calls, 1 voice note, 1 email), 1 overdue. 09:00 Email Primož (GA Kuhinje), touch 2 of 5 …" */
function digestText(d: any) {
  const date = new Date(d.date + "T00:00:00");
  const day = date.toLocaleDateString("en-GB", { weekday: "short", day: "numeric", month: "short" });
  const order = ["call", "voicenote", "whatsapp", "email", "checkin", "meeting", "other"];
  const names: Record<string, string> = { call: "call", voicenote: "voice note", whatsapp: "WhatsApp", email: "email", checkin: "check-in", meeting: "meeting", other: "task" };
  const counts = order.filter((k) => d.counts?.[k]).map((k) => plural(d.counts[k], names[k]));
  const head = `${day}: ${d.n_today} to contact${counts.length ? ` (${counts.join(", ")})` : ""}${d.n_overdue ? `, ${d.n_overdue} overdue` : ""}.`;
  const line = (r: any) => `${r.overdue ? "OVERDUE " : ""}${r.time} ${label(r.kind)} ${r.who}${r.company && r.company !== r.who ? ` (${r.company})` : ""}${r.step > 0 ? `, touch ${r.step} of ${r.steps_total}` : r.kind === "checkin" ? ", check-in" : ""}`;
  const rows: string[] = (d.rows ?? []).map(line);
  const events: string[] = (d.events ?? []).map((e: any) => `${e.who}: ${describe(e.action, e.detail)}`);
  const title = `Solonos · ${plural(d.n_today, "touch")} today${d.n_overdue ? `, ${d.n_overdue} overdue` : ""}`;
  const body = [head, ...rows.slice(0, 3), events.length ? `Overnight: ${events.slice(0, 2).join("; ")}` : ""].filter(Boolean).join("\n");
  const esc = (s: string) => s.replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c] as string));
  const link = (r: any) => r.opportunity_id ? `${SITE}/app/leads?open=${encodeURIComponent(r.opportunity_id)}` : r.engagement_id ? `${SITE}/app/engagements/${r.engagement_id}` : `${SITE}/app/tasks`;
  const html = `<div style="font-family:Inter,Arial,sans-serif;font-size:15px;line-height:1.55;color:#2A1E20;max-width:560px">
<p style="margin:0 0 12px"><strong>${esc(head)}</strong></p>
<ul style="padding-left:18px;margin:0 0 14px">${(d.rows ?? []).map((r: any) => `<li style="margin:0 0 6px"><a href="${link(r)}" style="color:#800020">${esc(line(r))}</a></li>`).join("")}</ul>
${events.length ? `<p style="margin:0 0 12px"><strong>Overnight</strong><br>${events.map(esc).join("<br>")}</p>` : ""}
<p style="margin:0"><a href="${SITE}/app/tasks" style="color:#800020">Open Tasks</a></p></div>`;
  return { title, body, html };
}

function describe(action: string, detail: any) {
  const d = typeof detail === "string" ? detail : detail?.name || detail?.kind || "";
  return ({ "agreement.signed": "signed", "deposit.claimed": "says the deposit is paid", "invoice.claimed": "says an invoice is paid", "client.upload": "uploaded a file", "client.company": "confirmed company details", "client.login": "opened their page" } as Record<string, string>)[action] ?? action.replace("client.", "") + (d ? ` ${d}` : "");
}

async function sendEmail(m: { to: string; subject: string; html: string }) {
  const r = await fetch(`${Deno.env.get("SUPABASE_URL")}/functions/v1/send-email`, {
    method: "POST",
    headers: { "content-type": "application/json", "x-sync-secret": Deno.env.get("GHL_SYNC_SECRET") ?? "" },
    body: JSON.stringify(m),
  });
  const out = await r.json().catch(() => ({}));
  if (!r.ok || out.ok === false) throw new Error(`email failed: ${out.error ?? r.status}`);
}
