-- Emails the family head when their fc_families.status flips to 'active'
-- (the field checked at login — status==='pending' shows the waiting
-- screen). Approval happens via a manual status edit in Supabase Studio
-- with no application code path involved; only a best-effort Telegram
-- alert fires at signup time (notifyOwnerOfSignup), nothing on approval.
-- Same pattern as derasar-boli/DealLagi/reminder/Contract-Note-Converter —
-- shared Cloudflare Worker email relay (telegram-notify.unigoods2026.workers.dev,
-- action:"sendEmail"). fc_families itself has no email column - the head's
-- email lives on fc_users (is_head=true, family_id=families.id).

create or replace function fc_notify_approval() returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  v_email text;
begin
  if NEW.status = 'active' and OLD.status is distinct from 'active'
     and coalesce(NEW.is_demo, false) = false then
    select email into v_email from fc_users where family_id = NEW.id and is_head = true limit 1;
    if v_email is not null then
      perform net.http_post(
        url := 'https://telegram-notify.unigoods2026.workers.dev/',
        headers := '{"Content-Type":"application/json"}'::jsonb,
        body := jsonb_build_object(
          'action', 'sendEmail',
          'to', v_email,
          'subject', 'Your Family Cash account is approved',
          'html', '<p>Hi,</p>'
            || '<p>Your family account <b>' || coalesce(NEW.name, '') || '</b> has been approved on Family Cash. You can now log in and start using the app:</p>'
            || '<p><a href="https://familycash.anyapps.in">https://familycash.anyapps.in</a></p>'
            || '<p style="font-size:13px;color:#666;">Questions? Contact vkvcoder.support@gmail.com</p>'
        )
      );
    end if;
  end if;
  return NEW;
end;
$$;

drop trigger if exists fc_families_notify_approval on fc_families;
create trigger fc_families_notify_approval
after update on fc_families
for each row execute function fc_notify_approval();
