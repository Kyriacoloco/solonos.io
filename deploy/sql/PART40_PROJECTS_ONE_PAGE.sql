-- =====================================================================
-- Part 40 (23 Oct 2026, Solonos 3.16): one "Projects" section. Each project is one record (the engagement row)
-- with email, tasks, files, NDA, invoices and notes on the same page. The task list sits on a projects row that
-- is made the first time a task is added. That no longer changes the status; you set the status yourself.
-- Safe to re-run.
-- =====================================================================
create or replace function staff_create_project(p_engagement uuid) returns projects
language plpgsql security definer set search_path = public as $$
declare e engagements;
begin
  if not is_staff() then raise exception 'staff only'; end if;      -- Margarita (role user) adds tasks too
  select * into e from engagements where id = p_engagement;
  if e.id is null then raise exception 'Project not found'; end if;
  return create_project_from_engagement(p_engagement);
end $$;
grant execute on function staff_create_project(uuid) to authenticated;
