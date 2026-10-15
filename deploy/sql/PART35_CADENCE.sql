-- =====================================================================
-- Part 35 (15 Oct 2026): one follow-up rhythm for everything.
--   Touch 1 the next working day after the trigger (Call 1, valuation sent, letter sent, draft sent),
--   then every 3 working days after the last touch. 5 touches with no reply = parked on the 90-day list
--   (a paying client after a draft goes "Quiet" instead and resurfaces every 30 days).
--   Only the NEXT touch exists as a task; the plan is visible on the lead.
--   Also: Cyprus + Greek public holidays, task times + push (Parts 33/34 folded in), leads_list view
--   for the list screen, and the 08:30 morning digest.
-- Run once, after Parts 31 and 32 (Parts 33 and 34 are included here; running them first is harmless).
-- Safe to re-run.
-- =====================================================================

-- ------------------------------------------------------------ columns (Parts 33/34)
alter table lead_tasks add column if not exists followup_id uuid references followups(id) on delete cascade;
alter table lead_tasks add column if not exists step int;
alter table lead_tasks add column if not exists channel text;      -- email | voicenote | call | whatsapp | checkin
alter table lead_tasks add column if not exists subject text;
alter table lead_tasks add column if not exists message text;      -- template; the app fills {first} {company} {thing} ...
alter table lead_tasks add column if not exists outcome text;      -- no_reply | replied | not_now | lost | checkin | done
alter table lead_tasks add column if not exists due_time time not null default '09:00';
alter table lead_tasks add column if not exists notified_at timestamptz;
alter table lead_tasks add column if not exists overdue_notified_on date;
create index if not exists lead_tasks_by_followup on lead_tasks (followup_id) where done_at is null;

-- followups: more kinds, engagement follow-ups, "quiet" end state
alter table followups add column if not exists engagement_id uuid references engagements(id) on delete cascade;
alter table followups add column if not exists steps_total int;
alter table followups alter column opportunity_id drop not null;
do $$
declare c record;
begin
  for c in select conname from pg_constraint where conrelid = 'followups'::regclass and contype = 'c' loop
    execute format('alter table followups drop constraint %I', c.conname);
  end loop;
end $$;
alter table followups add constraint followups_kind_check   check (kind in ('call1','call2','valuation','letter','draft','checkin'));
alter table followups add constraint followups_status_check check (status in ('active','checkin','quiet','replied','not_now','lost','closed'));
alter table followups add constraint followups_target_check check (opportunity_id is not null or engagement_id is not null);
create index if not exists followups_by_engagement on followups (engagement_id) where status in ('active','checkin','quiet');

-- ------------------------------------------------------------ working days: weekends and CY/GR public holidays
create table if not exists holidays (day date primary key, name text not null);
alter table holidays enable row level security;
drop policy if exists staff_all on holidays;
create policy staff_all on holidays for all to authenticated using (is_staff()) with check (is_staff());
grant select, insert, update, delete on holidays to authenticated;
insert into holidays (day, name) values
  ('2026-01-01','New Year'), ('2026-01-06','Epiphany'), ('2026-02-23','Clean Monday'), ('2026-03-25','25 March'),
  ('2026-04-01','Cyprus National Day'), ('2026-04-10','Good Friday'), ('2026-04-13','Easter Monday'), ('2026-05-01','Labour Day'),
  ('2026-06-01','Whit Monday'), ('2026-08-15','Assumption'), ('2026-10-01','Cyprus Independence'), ('2026-10-28','Ochi Day'),
  ('2026-12-24','Christmas Eve'), ('2026-12-25','Christmas'), ('2026-12-26','Boxing Day'), ('2026-12-31','New Year''s Eve'),
  ('2027-01-01','New Year'), ('2027-01-06','Epiphany'), ('2027-03-15','Clean Monday'), ('2027-03-25','25 March'),
  ('2027-04-01','Cyprus National Day'), ('2027-04-30','Good Friday'), ('2027-05-03','Easter Monday'), ('2027-05-01','Labour Day'),
  ('2027-06-21','Whit Monday'), ('2027-08-15','Assumption'), ('2027-10-01','Cyprus Independence'), ('2027-10-28','Ochi Day'),
  ('2027-12-24','Christmas Eve'), ('2027-12-25','Christmas'), ('2027-12-26','Boxing Day'), ('2027-12-31','New Year''s Eve')
on conflict (day) do nothing;

create or replace function is_working_day(p date) returns boolean language sql stable as $$
  select extract(isodow from p) < 6 and not exists (select 1 from holidays h where h.day = p)
$$;
create or replace function add_working_days(p date, n int) returns date language plpgsql stable as $$
declare d date := p; k int := 0;
begin
  while k < n loop d := d + 1; if is_working_day(d) then k := k + 1; end if; end loop;
  return d;
end $$;

-- ------------------------------------------------------------ the one sequence (Settings edits this row)
create or replace function followup_seq() returns jsonb language sql stable as $$
  select value from app_settings where key = 'followups'
$$;
do $$
declare cur jsonb := coalesce((select value from app_settings where key = 'followups'), '{}'::jsonb);
        def jsonb := $json$
{
  "first_day": 1,
  "gap_days": 3,
  "checkin_days": 90,
  "quiet_days": 30,
  "times": { "call": "10:00", "whatsapp": "09:30", "voicenote": "09:30", "email": "09:00", "other": "09:00", "checkin": "09:00" },
  "things": { "call1": "the points from our call", "call2": "the points from our last call", "valuation": "the Market Example Range I sent you", "letter": "the engagement letter", "draft": "the draft valuation" },
  "chase": { "steps": [
    { "channel": "email", "title": "Email: did it land?", "subject": "{company}: quick check",
      "text": "Hi {first},\n\nDid you get a chance to look at {thing}? Happy to walk you through it; 15 minutes is enough.\n\nKyriacos" },
    { "channel": "voicenote", "title": "WhatsApp voice note",
      "text": "(voice note, 20 seconds) Hi {first}, Kyriacos here. I sent you an email about {thing}. One quick question I would like to ask you about {company}; let me know when you have a moment." },
    { "channel": "call", "title": "Call; if no answer, a short voice note",
      "text": "(if no answer, voice note) Hi {first}, tried to call you just now, nothing urgent. When would suit you for a quick word about {thing}?" },
    { "channel": "email", "title": "Email with something new", "subject": "{company}: one more point",
      "text": "Hi {first},\n\nOne more point since my last message: {point}.\n\nIf you want to go through it, I can do {slot1} or {slot2}.\n\nKyriacos" },
    { "channel": "voicenote", "title": "Voice note: should I close your file?",
      "text": "(voice note) Hi {first}, Kyriacos. I have not heard back, which usually means the timing is not right, and that is completely fine. Should I close your file for now, or shall I check back with you in a few months? A one-word reply is enough." }
  ] },
  "first_touch": {
    "call1": { "subject": "{company}: next step after our call",
      "text": "Hi {first},\n\nThank you for the call. As discussed, here is the next step for {company}.\n\n{booking}\n\nKyriacos" }
  },
  "checkin_reasons": ["Year-end accounts", "A deal in their sector", "The thing they said mattered"]
}
$json$::jsonb;
begin
  -- keep what he may have changed before (times, booking link, check-in days and reasons); everything else comes from the new default
  def := def || jsonb_strip_nulls(jsonb_build_object('times', coalesce(def->'times', '{}'::jsonb) || coalesce(cur->'times', '{}'::jsonb),
                                                     'booking', cur->'booking', 'checkin_days', cur->'checkin_days', 'checkin_reasons', cur->'checkin_reasons'));
  insert into app_settings (key, value) values ('followups', def)
  on conflict (key) do update set value = excluded.value, updated_at = now();
end $$;

create or replace function task_default_time(p_kind text) returns time language sql stable as $$
  select coalesce((followup_seq()->'times'->>p_kind)::time, '09:00'::time)
$$;
create or replace function task_kind_for(p_channel text) returns text language sql immutable as $$
  select case when p_channel like 'call%' then 'call' when p_channel = 'voicenote' then 'voicenote' when p_channel = 'whatsapp' then 'whatsapp'
              when p_channel in ('email','letter') then 'email' else 'other' end
$$;
create or replace function followup_kind_label(p_kind text) returns text language sql immutable as $$
  select case p_kind when 'call1' then 'after Call 1' when 'call2' then 'after Call 2' when 'valuation' then 'after the valuation'
                     when 'letter' then 'after the engagement letter' when 'draft' then 'after the draft report' else p_kind end
$$;

/** The task row for touch n of a follow-up (no insert). Step 1 may have a per-kind first message. */
create or replace function followup_touch_row(p_kind text, p_step int)
returns table (title text, channel text, kind text, subject text, message text) language sql stable as $$
  select coalesce(st->>'title', 'Touch ' || p_step),
         coalesce(st->>'channel', 'email'),
         task_kind_for(coalesce(st->>'channel', 'email')),
         coalesce(ft->>'subject', st->>'subject'),
         coalesce(ft->>'text', st->>'text')
  from (select followup_seq()->'chase'->'steps'->(p_step - 1) as st,
               case when p_step = 1 then followup_seq()->'first_touch'->p_kind else null end as ft) x
$$;
create or replace function followup_steps_total() returns int language sql stable as $$
  select coalesce(jsonb_array_length(followup_seq()->'chase'->'steps'), 5)
$$;

/** Internal: start a follow-up without the staff check (edge functions call this with the service role). */
create or replace function followup_begin(p_opportunity text, p_kind text, p_check_back date, p_engagement uuid, p_who text)
returns followups language plpgsql security definer set search_path = public as $$
declare f followups; r record; d date; total int := followup_steps_total(); first_day int := coalesce((followup_seq()->>'first_day')::int, 1);
begin
  if p_kind not in ('call1','call2','valuation','letter','draft','checkin') then raise exception 'kind must be call1, call2, valuation, letter, draft or checkin'; end if;
  if p_opportunity is null and p_engagement is null then raise exception 'A lead or an engagement is needed'; end if;
  if p_kind = 'draft' and p_engagement is null then raise exception 'The draft follow-up runs on an engagement'; end if;
  -- one running follow-up per lead/engagement; a running one is replaced (the lead moved on to a new trigger)
  for f in select * from followups where status in ('active','checkin','quiet')
             and ((p_opportunity is not null and opportunity_id = p_opportunity) or (p_engagement is not null and engagement_id = p_engagement)) loop
    delete from lead_tasks where followup_id = f.id and done_at is null;
    update followups set status = 'closed', ended_at = now() where id = f.id;
  end loop;
  if p_kind = 'checkin' then
    d := coalesce(p_check_back, current_date + coalesce((followup_seq()->>'checkin_days')::int, 90));
    insert into followups (opportunity_id, engagement_id, kind, status, step, next_due, steps_total, created_by)
    values (p_opportunity, p_engagement, 'checkin', 'checkin', 0, d, 0, auth.uid()) returning * into f;
    insert into lead_tasks (opportunity_id, engagement_id, followup_id, step, title, kind, channel, due_on, due_time, created_by)
    values (p_opportunity, p_engagement, f.id, 0, '90-day check-in (only with a reason)', 'other', 'checkin', d, task_default_time('checkin'), p_who);
  else
    d := add_working_days(current_date, first_day);
    insert into followups (opportunity_id, engagement_id, kind, status, step, next_due, steps_total, created_by)
    values (p_opportunity, p_engagement, p_kind, 'active', 1, d, total, auth.uid()) returning * into f;
    select * into r from followup_touch_row(p_kind, 1);
    insert into lead_tasks (opportunity_id, engagement_id, followup_id, step, title, kind, channel, subject, message, due_on, due_time, created_by)
    values (p_opportunity, p_engagement, f.id, 1, r.title, r.kind, r.channel, r.subject, r.message, d, task_default_time(r.kind), p_who);
  end if;
  insert into lead_notes (opportunity_id, engagement_id, body, author)
  values (p_opportunity, p_engagement, case when p_kind = 'checkin' then 'On the 90-day check-in list. Next check-in ' || to_char(d, 'DD Mon YYYY')
                else 'Follow-up started (' || followup_kind_label(p_kind) || '): ' || total || ' touches, every ' || coalesce(followup_seq()->>'gap_days', '3') || ' working days. First ' || to_char(d, 'DD Mon YYYY') end, p_who);
  return f;
end $$;
revoke all on function followup_begin(text, text, date, uuid, text) from public, anon, authenticated;
grant execute on function followup_begin(text, text, date, uuid, text) to service_role;

/** Start a follow-up (staff). kind: call1 | call2 | valuation | letter | draft | checkin. p_engagement for engagement follow-ups. */
drop function if exists followup_start(text, text, date);
create or replace function followup_start(p_opportunity text, p_kind text, p_check_back date default null, p_engagement uuid default null)
returns followups language plpgsql security definer set search_path = public as $$
begin
  if not is_staff() then raise exception 'Staff only'; end if;
  return followup_begin(p_opportunity, p_kind, p_check_back, p_engagement, coalesce((select email from auth.users where id = auth.uid()), 'staff'));
end $$;
revoke all on function followup_start(text, text, date, uuid) from public, anon;
grant execute on function followup_start(text, text, date, uuid) to authenticated;

/** Done on a task. Touches: no_reply | replied | not_now | lost | checkin. Your own tasks: outcome ignored. */
create or replace function task_done(p_task uuid, p_outcome text default null, p_note text default null, p_check_back date default null)
returns lead_tasks language plpgsql security definer set search_path = public as $$
declare t lead_tasks; f followups; who text; nxt date; label text; r record; gap int := coalesce((followup_seq()->>'gap_days')::int, 3);
        total int; is_eng boolean;
begin
  if not is_staff() then raise exception 'Staff only'; end if;
  select * into t from lead_tasks where id = p_task for update;
  if t.id is null then raise exception 'Task not found'; end if;
  if t.done_at is not null then raise exception 'Already done'; end if;
  who := coalesce((select email from auth.users where id = auth.uid()), 'staff');

  if t.followup_id is null then
    update lead_tasks set done_at = now(), outcome = 'done', note = coalesce(nullif(trim(p_note), ''), note) where id = t.id returning * into t;
    if t.opportunity_id is not null or t.engagement_id is not null then
      insert into lead_notes (opportunity_id, engagement_id, body, author)
      values (t.opportunity_id, t.engagement_id, 'Done: ' || t.title || case when coalesce(trim(p_note), '') <> '' then ' · ' || trim(p_note) else '' end, who);
    end if;
    return t;
  end if;

  select * into f from followups where id = t.followup_id for update;
  total := coalesce(f.steps_total, followup_steps_total());
  is_eng := f.kind = 'draft';
  if p_outcome is null or p_outcome not in ('no_reply','replied','not_now','lost','checkin') then raise exception 'Tell me what happened: no_reply, replied, not_now, lost or checkin'; end if;
  if t.channel = 'checkin' and p_outcome = 'checkin' and coalesce(trim(p_note), '') = '' then raise exception 'A check-in needs a reason'; end if;
  update lead_tasks set done_at = now(), outcome = p_outcome, note = coalesce(nullif(trim(p_note), ''), note) where id = t.id returning * into t;

  if p_outcome = 'no_reply' then
    if t.step >= 1 and t.step < total then
      nxt := add_working_days(current_date, gap);
      select * into r from followup_touch_row(f.kind, t.step + 1);
      insert into lead_tasks (opportunity_id, engagement_id, followup_id, step, title, kind, channel, subject, message, due_on, due_time, created_by)
      values (f.opportunity_id, f.engagement_id, f.id, t.step + 1, r.title, r.kind, r.channel, r.subject, r.message, nxt, task_default_time(r.kind), who);
      update followups set step = t.step + 1, next_due = nxt, last_touch_on = current_date where id = f.id;
      label := 'No reply. Touch ' || (t.step + 1) || ' of ' || total || ' on ' || to_char(nxt, 'DD Mon YYYY');
    elsif is_eng then
      nxt := current_date + coalesce((followup_seq()->>'quiet_days')::int, 30);
      update followups set status = 'quiet', next_due = nxt, step = t.step, last_touch_on = current_date where id = f.id;
      insert into lead_tasks (opportunity_id, engagement_id, followup_id, step, title, kind, channel, due_on, due_time, created_by)
      values (f.opportunity_id, f.engagement_id, f.id, 0, 'Quiet client: check in', 'other', 'checkin', nxt, task_default_time('checkin'), who);
      label := 'No reply after ' || total || ' touches. Client is Quiet; check in ' || to_char(nxt, 'DD Mon YYYY');
    else
      nxt := current_date + coalesce((followup_seq()->>'checkin_days')::int, 90);
      update followups set status = 'checkin', next_due = nxt, step = t.step, last_touch_on = current_date where id = f.id;
      insert into lead_tasks (opportunity_id, engagement_id, followup_id, step, title, kind, channel, due_on, due_time, created_by)
      values (f.opportunity_id, f.engagement_id, f.id, 0, '90-day check-in (only with a reason)', 'other', 'checkin', nxt, task_default_time('checkin'), who);
      label := 'No reply after ' || total || ' touches. On the 90-day list; next check-in ' || to_char(nxt, 'DD Mon YYYY');
    end if;
  elsif p_outcome = 'replied' then
    delete from lead_tasks where followup_id = f.id and done_at is null;
    update followups set status = 'replied', ended_at = now(), last_touch_on = current_date where id = f.id;
    label := 'Replied. Follow-up closed.';
  elsif p_outcome = 'not_now' then
    nxt := coalesce(p_check_back, current_date + coalesce((followup_seq()->>'checkin_days')::int, 90));
    delete from lead_tasks where followup_id = f.id and done_at is null;
    update followups set status = 'checkin', next_due = nxt, last_touch_on = current_date where id = f.id;
    insert into lead_tasks (opportunity_id, engagement_id, followup_id, step, title, kind, channel, due_on, due_time, created_by)
    values (f.opportunity_id, f.engagement_id, f.id, 0, 'Check back (they said not now)', 'other', 'checkin', nxt, task_default_time('checkin'), who);
    label := 'Not now. Check back ' || to_char(nxt, 'DD Mon YYYY');
  elsif p_outcome = 'lost' then
    delete from lead_tasks where followup_id = f.id and done_at is null;
    update followups set status = 'lost', ended_at = now(), last_touch_on = current_date where id = f.id;
    label := 'Said no. Follow-up closed.';
  else -- a check-in done: plan the next one (90 days for leads, 30 for quiet clients)
    nxt := coalesce(p_check_back, current_date + case when f.status = 'quiet' then coalesce((followup_seq()->>'quiet_days')::int, 30) else coalesce((followup_seq()->>'checkin_days')::int, 90) end);
    update followups set next_due = nxt, last_touch_on = current_date where id = f.id;
    insert into lead_tasks (opportunity_id, engagement_id, followup_id, step, title, kind, channel, due_on, due_time, created_by)
    values (f.opportunity_id, f.engagement_id, f.id, 0, case when f.status = 'quiet' then 'Quiet client: check in' else '90-day check-in (only with a reason)' end, 'other', 'checkin', nxt, task_default_time('checkin'), who);
    label := 'Check-in done. Next ' || to_char(nxt, 'DD Mon YYYY');
  end if;

  insert into followup_touches (followup_id, opportunity_id, step, channel, outcome, note, next_due, done_by)
  values (f.id, coalesce(t.opportunity_id, f.engagement_id::text), t.step, t.channel, p_outcome, nullif(trim(p_note), ''), nxt, who);
  insert into lead_notes (opportunity_id, engagement_id, body, author)
  values (t.opportunity_id, coalesce(t.engagement_id, f.engagement_id),
          coalesce(initcap(t.channel) || ': ', '') || label || case when coalesce(trim(p_note), '') <> '' then ' · ' || trim(p_note) else '' end, who);
  return t;
end $$;
drop function if exists task_done(uuid, text);
drop function if exists followup_log(uuid, text, text, text, date);

create or replace function followup_cancel(p_followup uuid) returns void
language plpgsql security definer set search_path = public as $$
begin
  if not is_staff() then raise exception 'Staff only'; end if;
  delete from lead_tasks where followup_id = p_followup and done_at is null;
  update followups set status = 'closed', ended_at = now() where id = p_followup and status in ('active','checkin','quiet');
end $$;
revoke all on function task_done(uuid, text, text, date), followup_cancel(uuid) from public, anon;
grant execute on function task_done(uuid, text, text, date), followup_cancel(uuid) to authenticated;

-- ------------------------------------------------------------ devices + what to push (Part 34, minus the separate 09:00 overdue ping)
create table if not exists push_subscriptions (
  id          uuid primary key default gen_random_uuid(),
  user_id     uuid not null,
  endpoint    text not null unique,
  keys        jsonb not null,
  user_agent  text,
  created_at  timestamptz not null default now(),
  last_ok_at  timestamptz
);
alter table push_subscriptions enable row level security;
drop policy if exists own_rows on push_subscriptions;
create policy own_rows on push_subscriptions for all to authenticated using (user_id = auth.uid() and is_staff()) with check (user_id = auth.uid() and is_staff());
grant select, insert, update, delete on push_subscriptions to authenticated;

-- ------------------------------------------------------------ the board: every open task with who and how to reach them
drop view if exists tasks_board;
create view tasks_board as
select t.id, t.opportunity_id, coalesce(t.engagement_id, f.engagement_id) as engagement_id, t.title, t.kind, t.channel, t.subject, t.message, t.due_on, t.due_time, t.note, t.created_at,
       t.followup_id, t.step, f.kind as followup_kind, f.status as followup_status,
       coalesce(f.steps_total, followup_steps_total()) as steps_total,
       coalesce(l.name, e.company_name, e.client_name) as who,
       coalesce(l.first_name, split_part(e.client_name, ' ', 1)) as first_name,
       coalesce(l.company, e.company_name) as company,
       coalesce(l.phone, '') as phone, coalesce(l.email, e.client_email) as email,
       l.stage_name
from lead_tasks t
left join followups f on f.id = t.followup_id
left join leads_board l on l.ghl_id = t.opportunity_id
left join engagements e on e.id = coalesce(t.engagement_id, f.engagement_id)
where t.done_at is null;
alter view tasks_board set (security_invoker = on);

create or replace function tasks_to_push()
returns table (id uuid, kind text, title text, who text, due_on date, due_time time, opportunity_id text, engagement_id uuid, reason text)
language plpgsql security definer set search_path = public as $$
declare now_local timestamp := (now() at time zone 'Europe/Athens'); today date := now_local::date; t time := now_local::time;
begin
  if extract(isodow from today) > 5 or t < '08:00' or t > '20:00' then return; end if;
  return query
    select b.id, b.kind, b.title, b.who, b.due_on, b.due_time, b.opportunity_id, b.engagement_id, 'due'::text
      from tasks_board b join lead_tasks x on x.id = b.id
     where b.due_on = today and b.due_time <= t and x.notified_at is null;
end $$;
revoke all on function tasks_to_push() from public, anon, authenticated;
grant execute on function tasks_to_push() to service_role;

create or replace function task_mark_pushed(p_ids uuid[], p_reason text) returns void
language sql security definer set search_path = public as $$
  update lead_tasks set notified_at = case when p_reason = 'due' then now() else notified_at end,
                        overdue_notified_on = case when p_reason = 'overdue' then (now() at time zone 'Europe/Athens')::date else overdue_notified_on end
   where id = any(p_ids);
$$;
revoke all on function task_mark_pushed(uuid[], text) from public, anon, authenticated;
grant execute on function task_mark_pushed(uuid[], text) to service_role;

-- ------------------------------------------------------------ suggestions: who looks like a ghost right now
drop view if exists followup_suggestions;
create view followup_suggestions as
select l.ghl_id as opportunity_id, l.name, l.first_name, l.phone, l.stage_name, l.call_at, l.call_status, l.next_step_at,
       case when l.call_status ~* 'noshow|no_show' and l.call_at < now() then 'noshow' else 'silence' end as reason,
       case when l.stage_name ~* 'letter' then 'letter' when l.stage_name ~* 'offer' then 'valuation' else 'call1' end as suggested_kind
from leads_board l
where not exists (select 1 from followups f where f.opportunity_id = l.ghl_id and f.status in ('active','checkin','quiet'))
  and not exists (select 1 from engagements e where e.opportunity_id = l.ghl_id)
  and ( (l.call_status ~* 'noshow|no_show' and l.call_at < now() and l.call_at > now() - interval '21 days')
     or (l.next_step_at is not null and add_working_days(l.next_step_at, 3) < current_date
         and coalesce(l.last_note_at, '1970-01-01'::timestamptz) < l.next_step_at::timestamptz) );
alter view followup_suggestions set (security_invoker = on);

-- ------------------------------------------------------------ the list screen: one row per open lead, sorted by what needs doing
drop view if exists leads_list;
create view leads_list as
with nt as (
  select distinct on (t.opportunity_id) t.opportunity_id, t.id as task_id, t.kind as task_kind, t.channel as task_channel, t.title as task_title,
         t.due_on, t.due_time, t.step, t.followup_id
  from lead_tasks t where t.done_at is null and t.opportunity_id is not null
  order by t.opportunity_id, t.due_on, t.due_time
), lt as (
  select opportunity_id, max(done_at) as last_touch_at from lead_tasks where done_at is not null and opportunity_id is not null group by opportunity_id
)
select l.*,
       nt.task_id, nt.task_kind, nt.task_channel, nt.task_title, nt.due_on, nt.due_time, nt.step,
       f.id as followup_id, f.kind as followup_kind, f.status as followup_status, coalesce(f.steps_total, followup_steps_total()) as steps_total,
       greatest(lt.last_touch_at, l.last_note_at) as last_touch_at,
       case when greatest(lt.last_touch_at, l.last_note_at) is null then null
            else (current_date - greatest(lt.last_touch_at, l.last_note_at)::date) end as days_since_touch,
       case when f.status in ('checkin','quiet') then 'parked'
            when nt.due_on < current_date then 'overdue'
            when nt.due_on = current_date then 'today'
            when nt.due_on is null then 'none'
            else 'future' end as bucket,
       case when f.status in ('checkin','quiet') then 4 when nt.due_on < current_date then 0 when nt.due_on = current_date then 1 when nt.due_on is null then 2 else 3 end as bucket_order
from leads_board l
left join nt on nt.opportunity_id = l.ghl_id
left join followups f on f.opportunity_id = l.ghl_id and f.status in ('active','checkin','quiet')
left join lt on lt.opportunity_id = l.ghl_id;
alter view leads_list set (security_invoker = on);

-- ------------------------------------------------------------ the morning digest (08:30 Athens, weekdays, only when there is something)
create table if not exists digest_log (sent_on date primary key, sent_at timestamptz not null default now(), body jsonb);

/** What the 08:30 message says. Returns null when there is nothing due and nothing overdue. */
create or replace function morning_digest() returns jsonb language plpgsql security definer set search_path = public as $$
declare now_local timestamp := (now() at time zone 'Europe/Athens'); today date := now_local::date;
        rows jsonb; counts jsonb; n_today int; n_over int; events jsonb; since timestamptz;
begin
  select count(*) filter (where due_on = today), count(*) filter (where due_on < today) into n_today, n_over from tasks_board;
  if coalesce(n_today, 0) + coalesce(n_over, 0) = 0 then return null; end if;
  select coalesce(jsonb_object_agg(k, c), '{}'::jsonb) into counts
    from (select coalesce(case when channel = 'checkin' then 'checkin' else kind end, 'other') as k, count(*) as c from tasks_board where due_on = today group by 1) x;
  select coalesce(jsonb_agg(jsonb_build_object('time', to_char(due_time, 'HH24:MI'), 'kind', case when channel = 'checkin' then 'checkin' else kind end,
                    'who', trim(split_part(coalesce(who, ''), '|', 1)), 'company', company, 'title', title, 'step', step, 'steps_total', steps_total,
                    'due_on', due_on, 'overdue', due_on < today, 'opportunity_id', opportunity_id, 'engagement_id', engagement_id) order by due_on, due_time), '[]'::jsonb)
    into rows from (select * from tasks_board where due_on <= today order by due_on, due_time limit 12) b;
  since := coalesce((select max(sent_at) from digest_log), now() - interval '1 day');
  select coalesce(jsonb_agg(jsonb_build_object('who', coalesce(e.company_name, e.client_name, a.actor), 'action', a.action, 'detail', a.detail, 'at', a.created_at) order by a.created_at), '[]'::jsonb)
    into events
    from activity_log a left join engagements e on e.id = a.engagement_id
   where a.created_at > since and (a.action like 'client.%' or a.action in ('agreement.signed','deposit.claimed','invoice.claimed'))
     and a.actor like '%@%' and a.actor not ilike '%consortiaadvisory.com';
  return jsonb_build_object('date', today, 'n_today', n_today, 'n_overdue', n_over, 'counts', counts, 'rows', rows, 'events', events);
end $$;
revoke all on function morning_digest() from public, anon;
grant execute on function morning_digest() to authenticated;   -- Settings shows a preview

/** Called by push-notify with the service role: the digest to send now, or null (wrong time, weekend, already sent, nothing to say). */
create or replace function morning_digest_due() returns jsonb language plpgsql security definer set search_path = public as $$
declare now_local timestamp := (now() at time zone 'Europe/Athens'); today date := now_local::date; t time := now_local::time;
        at time := coalesce((followup_seq()->>'digest_time')::time, '08:30'::time); d jsonb;
begin
  if extract(isodow from today) > 5 then return null; end if;
  if t < at or t > at + interval '2 hours' then return null; end if;
  if exists (select 1 from digest_log where sent_on = today) then return null; end if;
  d := morning_digest();
  insert into digest_log (sent_on, body) values (today, d);   -- recorded even when empty, so we do not look again today
  return d;
end $$;
revoke all on function morning_digest_due() from public, anon, authenticated;
grant execute on function morning_digest_due() to service_role;

-- ------------------------------------------------------------ existing data: running sequences from Parts 31-34 get one next touch
do $$
declare f record; nxt date; r record; who text := 'migration';
begin
  for f in select * from followups where status = 'active' loop
    if exists (select 1 from lead_tasks x where x.followup_id = f.id and x.done_at is null) then
      -- keep only the earliest open touch; the rest are planned, not tasks
      delete from lead_tasks x where x.followup_id = f.id and x.done_at is null
         and x.id <> (select id from lead_tasks y where y.followup_id = f.id and y.done_at is null order by due_on, due_time limit 1);
    else
      nxt := greatest(coalesce(f.next_due, current_date), current_date);
      select * into r from followup_touch_row(f.kind, greatest(f.step, 1));
      insert into lead_tasks (opportunity_id, engagement_id, followup_id, step, title, kind, channel, subject, message, due_on, due_time, created_by)
      values (f.opportunity_id, f.engagement_id, f.id, greatest(f.step, 1), r.title, r.kind, r.channel, r.subject, r.message, nxt, task_default_time(r.kind), who);
    end if;
    update followups set steps_total = coalesce(steps_total, followup_steps_total()) where id = f.id;
  end loop;
  for f in select * from followups where status = 'checkin' and not exists (select 1 from lead_tasks x where x.followup_id = followups.id and x.done_at is null) loop
    insert into lead_tasks (opportunity_id, engagement_id, followup_id, step, title, kind, channel, due_on, due_time, created_by)
    values (f.opportunity_id, f.engagement_id, f.id, 0, '90-day check-in (only with a reason)', 'other', 'checkin', coalesce(f.next_due, current_date + 90), task_default_time('checkin'), who);
  end loop;
end $$;
