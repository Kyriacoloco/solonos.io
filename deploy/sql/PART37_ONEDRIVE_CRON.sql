-- Part 37 (23 Oct 2026): sync every engagement's OneDrive folder every 15 minutes. Safe to re-run.
select cron.unschedule('onedrive-sync') where exists (select 1 from cron.job where jobname = 'onedrive-sync');
select cron.schedule('onedrive-sync', '*/15 * * * *', $$
  select net.http_post(
    url     := 'https://uirbiqfkgnfnfugfclmu.supabase.co/functions/v1/onedrive',
    headers := jsonb_build_object('content-type','application/json',
                 'x-push-secret', (select decrypted_secret from vault.decrypted_secrets where name = 'push_secret')),
    body    := '{"action":"sync_all"}'::jsonb,
    timeout_milliseconds := 120000);
$$);
