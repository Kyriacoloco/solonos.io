-- PART 45 (4 Nov 2026): leads are never deleted. Kyriacos: "Never delete any lead."
-- The GHL sync only ever adds or updates; this makes the database refuse a delete from anywhere (app, SQL, API).
create or replace function public.leads_never_deleted() returns trigger language plpgsql as $$
begin
  raise exception 'Leads are never deleted in Solonos (% %). Mark the lead Lost or Disqualified instead.', tg_table_name, coalesce(old.ghl_id, '');
end $$;
drop trigger if exists trg_no_delete_opportunities on public.opportunities;
create trigger trg_no_delete_opportunities before delete on public.opportunities for each row execute function public.leads_never_deleted();
drop trigger if exists trg_no_delete_contacts on public.contacts;
create trigger trg_no_delete_contacts before delete on public.contacts for each row execute function public.leads_never_deleted();
select count(*) as leads_in_solonos from public.opportunities;
