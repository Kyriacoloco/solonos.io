-- =====================================================================
-- Part 39 (23 Oct 2026): simpler flow. Lead (CRM) -> Engagement (any type: NDA and maybe an invoice) -> Project (any time).
--   engagements.engagement_type   valuation | sell_side | ma | business_plan | feasibility | financial_modeling |
--                                 transfer_pricing | teaser | form_8802 | restructuring | cfo | other
--   staff_new_engagement(...)     an engagement without a lead; sends nothing
--   staff_add_invoice(...)        an invoice of any amount on any engagement; numbered now, emailed only when you choose
--   staff_create_project(...)     a project from any engagement at any time, even with details missing
-- Valuations keep their existing flow. Safe to re-run.
-- =====================================================================

-- ---- engagement type
alter table engagements add column if not exists engagement_type text not null default 'valuation';
alter table engagements drop constraint if exists engagements_type_check;
alter table engagements add constraint engagements_type_check check (engagement_type in
  ('valuation','sell_side','ma','business_plan','feasibility','financial_modeling','transfer_pricing','teaser','form_8802','restructuring','cfo','other'));

-- ---- project type follows the engagement type; notes on the project
alter table projects drop constraint if exists projects_service_line_check;
alter table projects add constraint projects_service_line_check check (service_line in
  ('valuation','sell_side','ma','business_plan','feasibility','financial_modeling','transfer_pricing','teaser','form_8802','restructuring','cfo','other'));
alter table projects add column if not exists notes text;

create or replace function engagement_type_label(t text) returns text language sql immutable as $$
  select case t
    when 'valuation' then 'Business Valuation' when 'sell_side' then 'Sell-side M&A' when 'ma' then 'M&A'
    when 'business_plan' then 'Business Plan' when 'feasibility' then 'Feasibility Study' when 'financial_modeling' then 'Financial Modelling'
    when 'transfer_pricing' then 'Transfer Pricing' when 'teaser' then 'Teaser' when 'form_8802' then '8802 Form'
    when 'restructuring' then 'Restructuring Report' when 'cfo' then 'CFO Services' else 'Project' end
$$;

-- ---- an engagement without a lead
create or replace function staff_new_engagement(p_type text, p_client_name text, p_client_email text,
  p_company text default null, p_country text default null, p_opportunity text default null) returns engagements
language plpgsql security definer set search_path = public as $$
declare e engagements;
begin
  if not is_staff() or is_limited_user() then raise exception 'staff only'; end if;
  if coalesce(trim(p_client_name), '') = '' then raise exception 'Client name is required'; end if;
  insert into engagements (engagement_type, client_name, client_email, company_name, company_country, opportunity_id, stage, owner_id)
  values (coalesce(nullif(p_type, ''), 'other'), trim(p_client_name), nullif(trim(coalesce(p_client_email, '')), ''),
          nullif(trim(coalesce(p_company, '')), ''), nullif(trim(coalesce(p_country, '')), ''), p_opportunity, 'deposit_pending', auth.uid())
  returning * into e;
  insert into activity_log (engagement_id, actor, action, detail, visible_to_client)
  values (e.id, coalesce(auth.uid()::text, 'staff'), 'engagement.created', jsonb_build_object('type', e.engagement_type, 'manual', true), false);
  return e;
end $$;
grant execute on function staff_new_engagement(text, text, text, text, text, text) to authenticated;

-- ---- invoices of any amount (seq 3 and up, trigger 'manual', with their own description)
alter table engagement_invoices add column if not exists description text;
alter table engagement_invoices drop constraint if exists engagement_invoices_seq_check;
alter table engagement_invoices add constraint engagement_invoices_seq_check check (seq >= 0);
alter table engagement_invoices drop constraint if exists engagement_invoices_trigger_check;
alter table engagement_invoices add constraint engagement_invoices_trigger_check check (trigger in ('deposit','engagement_signed','final_report','manual'));

create or replace view invoices_view as
select
  i.id, i.engagement_id, i.seq, i.trigger, i.amount, i.currency, i.issuing_entity, i.invoice_no,
  i.status, i.stripe_link, i.pdf_path, i.receipt_path, i.issued_at, i.due_at, i.paid_at, i.client_claimed_at, i.client_reference,
  be.legal_name as entity_name,
  e.client_name, e.company_name, e.company_address, e.company_country, e.company_tax_id, e.client_email,
  b.iban, b.swift, b.bank, b.account_name,
  i.description
from engagement_invoices i
join engagements e on e.id = i.engagement_id
join billing_entities be on be.code = i.issuing_entity
left join bank_accounts b on b.entity = i.issuing_entity and b.currency = i.currency;
alter view invoices_view set (security_invoker = on);
grant select on invoices_view to authenticated;

create or replace function staff_add_invoice(p_engagement uuid, p_amount numeric, p_description text, p_currency text default 'EUR')
returns engagement_invoices language plpgsql security definer set search_path = public as $$
declare e engagements; i engagement_invoices; n int;
begin
  if not is_staff() or is_limited_user() then raise exception 'staff only'; end if;
  select * into e from engagements where id = p_engagement;
  if e.id is null then raise exception 'Engagement not found'; end if;
  if coalesce(p_amount, 0) <= 0 then raise exception 'Enter an amount'; end if;
  if coalesce(trim(p_description), '') = '' then raise exception 'Enter what the invoice is for'; end if;
  select greatest(3, coalesce(max(seq), 0) + 1) into n from engagement_invoices where engagement_id = p_engagement;
  insert into engagement_invoices (engagement_id, seq, trigger, amount, currency, issuing_entity, status, description)
  values (p_engagement, n, 'manual', p_amount, coalesce(nullif(p_currency, ''), 'EUR'), default_billing_entity(e.company_country), 'scheduled', trim(p_description))
  returning * into i;
  return issue_invoice(i.id);     -- numbered now (CY or UK sequence by the client's country)
end $$;
grant execute on function staff_add_invoice(uuid, numeric, text, text) to authenticated;

-- ---- projects from any engagement at any time
-- the existing "Start valuation" path also uses this; it now returns the project if one already exists
create or replace function create_project_from_engagement(p_engagement uuid) returns projects
language plpgsql security definer set search_path = public as $$
declare e engagements; p projects; n int;
begin
  select * into p from projects where engagement_id = p_engagement order by started_at limit 1;
  if p.id is not null then return p; end if;
  select * into e from engagements where id = p_engagement;
  select coalesce(max(substring(code from 5)::int), 0) + 1 into n from projects;
  insert into projects (engagement_id, code, title, fee, owner_id, service_line)
  values (p_engagement, 'PRJ-' || lpad(n::text, 4, '0'),
          engagement_type_label(e.engagement_type) || ' - ' || coalesce(e.company_name, e.client_name), e.fee, e.owner_id, e.engagement_type)
  returning * into p;
  if e.engagement_type = 'valuation' then
    insert into tasks (project_id, title, stage) values
      (p.id, 'Normalise financial statements', 'draft'),
      (p.id, 'Build DCF model', 'draft'),
      (p.id, 'Comparable set & multiples', 'draft'),
      (p.id, 'Draft report', 'draft'),
      (p.id, 'Partner review', 'internal_review'),
      (p.id, 'Client walkthrough call', 'client_review'),
      (p.id, 'Issue final report', 'final');
  end if;
  insert into activity_log (engagement_id, project_id, actor, action, detail, visible_to_client)
  values (p_engagement, p.id, 'system', 'project.created', jsonb_build_object('code', p.code), false);
  return p;
end $$;
revoke all on function create_project_from_engagement(uuid) from public, anon, authenticated;

create or replace function staff_create_project(p_engagement uuid) returns projects
language plpgsql security definer set search_path = public as $$
declare e engagements; p projects;
begin
  if not is_staff() or is_limited_user() then raise exception 'staff only'; end if;
  select * into e from engagements where id = p_engagement;
  if e.id is null then raise exception 'Engagement not found'; end if;
  p := create_project_from_engagement(p_engagement);
  -- non-valuation engagements show "Project active"; valuations keep their own stages (letter, invoices)
  if e.engagement_type <> 'valuation' and e.stage not in ('completed','declined') then
    update engagements set stage = 'valuation_active' where id = p_engagement;
  end if;
  return p;
end $$;
grant execute on function staff_create_project(uuid) to authenticated;

-- ---- Margarita (role "user") works on project tasks, so the project task list is open to her
drop policy if exists no_limited_user on tasks;
