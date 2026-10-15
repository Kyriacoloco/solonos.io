-- =====================================================================
-- Part 35b (15 Oct 2026): run push-notify every 5 minutes. It sends each task's ping at its time and,
-- once a day at 08:30 Athens (weekdays), the morning digest. Replaces the Part 34 schedule.
-- BEFORE running: replace <PUSH_SECRET> with the PUSH_SECRET secret of the push-notify function.
-- Safe to re-run (it replaces the schedule).
-- =====================================================================
create extension if not exists pg_cron;
create extension if not exists pg_net;
select cron.unschedule(jobid) from cron.job where jobname in ('push-notify', 'push-digest');
select cron.schedule('push-notify', '*/5 * * * *', $$
  select net.http_post(
    url := 'https://uirbiqfkgnfnfugfclmu.supabase.co/functions/v1/push-notify',
    headers := '{"content-type":"application/json","x-push-secret":"<PUSH_SECRET>"}'::jsonb,
    body := '{}'::jsonb
  );
$$);
