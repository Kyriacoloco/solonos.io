-- =====================================================================
-- Part 36 (23 Oct 2026): each engagement can have a OneDrive folder under Common / Projects / <category> / <client>.
--   engagements.onedrive_*   where the folder is
--   onedrive_files           which file went where, so "Sync now" never copies the same file twice
-- Safe to re-run.
-- =====================================================================
alter table engagements add column if not exists onedrive_drive_id  text;
alter table engagements add column if not exists onedrive_folder_id text;
alter table engagements add column if not exists onedrive_url       text;
alter table engagements add column if not exists onedrive_path      text;
alter table engagements add column if not exists onedrive_category  text;
alter table engagements add column if not exists onedrive_synced_at timestamptz;

create table if not exists onedrive_files (
  id            bigserial primary key,
  engagement_id uuid not null references engagements(id) on delete cascade,
  storage_path  text not null,          -- bucket/path in Solonos
  item_id       text not null,          -- the OneDrive item
  name          text,
  direction     text not null check (direction in ('up','down')),
  created_at    timestamptz not null default now(),
  unique (engagement_id, storage_path)
);
create index if not exists onedrive_files_by_item on onedrive_files (engagement_id, item_id);
alter table onedrive_files enable row level security;
drop policy if exists staff_all on onedrive_files;
create policy staff_all on onedrive_files for all to authenticated using (is_staff()) with check (is_staff());
grant select on onedrive_files to authenticated;
