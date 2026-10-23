-- Part 41 (23 Oct 2026, Solonos 3.18): data room = a "Data Room" folder inside the project's OneDrive folder,
-- shared with the client through Microsoft. Solonos keeps where it is and who it was shared with. Safe to re-run.
alter table engagements add column if not exists dataroom_folder_id text;
alter table engagements add column if not exists dataroom_url text;
alter table engagements add column if not exists dataroom_shared jsonb not null default '[]'::jsonb;
select 'ok' as done;
