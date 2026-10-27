-- Part 42 (27 Oct 2026, Solonos 3.23): where each project is in its process (steps depend on the engagement type,
-- e.g. sell-side: Value and prepare, Go to market, Offers, Negotiate and complete). 0 = first step. Safe to re-run.
alter table engagements add column if not exists process_step int not null default 0;
select 'ok' as done;
