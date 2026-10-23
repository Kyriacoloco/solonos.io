-- Part 38 (23 Oct 2026): a third staff role, "user", who only sees Projects (engagements and their files).
-- The database itself blocks the sales side (leads, contacts, calls, follow-ups, tasks, ads, settings) for this role,
-- so hiding the menu is not the only protection. Safe to re-run.
alter table staff drop constraint if exists staff_role_check;
alter table staff add constraint staff_role_check check (role in ('partner','analyst','user'));

create or replace function is_limited_user() returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from staff where user_id = auth.uid() and role = 'user')
$$;
grant execute on function is_limited_user() to authenticated;

do $$
declare t text;
begin
  foreach t in array array['opportunities','contacts','call_recordings','lead_notes','lead_stage_events','lead_tasks',
                           'followups','followup_touches','meta_insights_daily','ghl_custom_fields',
                           'pipeline_stages','sync_log','digest_log','app_settings']
  loop
    if to_regclass('public.'||t) is not null then
      execute format('drop policy if exists no_limited_user on %I', t);
      execute format('create policy no_limited_user on %I as restrictive for all to authenticated using (not is_limited_user()) with check (not is_limited_user())', t);
    end if;
  end loop;
end $$;

-- a limited user can see her own staff row only
drop policy if exists no_limited_user on staff;
create policy no_limited_user on staff as restrictive for all to authenticated
  using (not is_limited_user() or user_id = auth.uid()) with check (not is_limited_user());
