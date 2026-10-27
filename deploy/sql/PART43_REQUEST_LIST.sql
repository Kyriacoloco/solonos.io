-- =====================================================================
-- Part 43 (27 Oct 2026, Solonos 3.28): the document request list per project.
--   client_documents gets a section, an order, the files the client sent (in OneDrive) and the client's note.
--   portal_request_file / portal_request_note are what the client's page calls. Safe to re-run.
-- =====================================================================
alter table client_documents add column if not exists section text;
alter table client_documents add column if not exists sort_order int not null default 0;
alter table client_documents add column if not exists files jsonb not null default '[]'::jsonb;   -- [{name, item, url, at}]
alter table client_documents add column if not exists client_note text;
alter table client_documents add column if not exists note_at timestamptz;

-- the client sent a file for one item (the file itself is in OneDrive, 4. Uploads)
create or replace function portal_request_file(p_doc uuid, p_name text, p_item text, p_url text)
returns client_documents language plpgsql security definer set search_path = public as $$
declare d client_documents;
begin
  select * into d from client_documents where id = p_doc;
  if d.id is null or not is_my_engagement(d.engagement_id) then raise exception 'Not your document'; end if;
  update client_documents
     set files = coalesce(files, '[]'::jsonb) || jsonb_build_array(jsonb_build_object('name', p_name, 'item', p_item, 'url', p_url, 'at', now())),
         uploaded_at = coalesce(uploaded_at, now()), uploaded_by = lower(auth.jwt() ->> 'email')
   where id = p_doc returning * into d;
  update engagements set last_client_activity_at = now() where id = d.engagement_id;
  insert into activity_log (engagement_id, actor, action, detail, visible_to_client)
  values (d.engagement_id, auth.jwt() ->> 'email', 'client.uploaded', jsonb_build_object('label', d.label, 'file', p_name), true);
  return d;
end $$;
grant execute on function portal_request_file(uuid, text, text, text) to authenticated;

-- the client explains why an item is missing (or anything else about it)
create or replace function portal_request_note(p_doc uuid, p_note text)
returns client_documents language plpgsql security definer set search_path = public as $$
declare d client_documents;
begin
  select * into d from client_documents where id = p_doc;
  if d.id is null or not is_my_engagement(d.engagement_id) then raise exception 'Not your document'; end if;
  update client_documents set client_note = nullif(trim(p_note), ''), note_at = now() where id = p_doc returning * into d;
  update engagements set last_client_activity_at = now() where id = d.engagement_id;
  insert into activity_log (engagement_id, actor, action, detail, visible_to_client)
  values (d.engagement_id, auth.jwt() ->> 'email', 'client.note', jsonb_build_object('label', d.label, 'note', left(coalesce(p_note, ''), 300)), true);
  return d;
end $$;
grant execute on function portal_request_note(uuid, text) to authenticated;

-- the client may read every column of their own items (the policy already exists; this widens the grant)
grant select on client_documents to authenticated;

-- ---- LUXALAD: the Information Request List of 23 Oct 2026 (replaces the untouched default items)
do $$
declare eid uuid := '5dba111e-93f6-4752-bb39-35a2e02c72f0';
begin
  if exists (select 1 from engagements where id = eid) and not exists (select 1 from client_documents where engagement_id = eid and section is not null) then
    delete from client_documents where engagement_id = eid and uploaded_at is null;
    insert into client_documents (engagement_id, section, sort_order, label, required)
    select eid, s, n, l, true from (values
  ('A. Financial', 1, 'Annual Financial Statements accounts for the last 5 years'),
  ('A. Financial', 2, 'Monthly management accounts for the last 24 months and this year to date'),
  ('A. Financial', 3, 'Revenue by month for 36 months, split into set-up fees, royalties, packaging and other'),
  ('A. Financial', 4, 'Revenue by franchisee, by month'),
  ('A. Financial', 5, 'For each unit: date signed, date opened, set-up fee charged and received, and what it cost you to open'),
  ('A. Financial', 6, 'Royalties charged and collected per franchisee, and anything still unpaid'),
  ('A. Financial', 7, 'All costs by category for 24 months, including freelancers'),
  ('A. Financial', 8, 'Personal or one-off costs paid by the company, including your own pay'),
  ('A. Financial', 9, 'Budget and five-year plan'),
  ('A. Financial', 10, 'Loans, leasing, and any money owed between you and the company'),
  ('A. Financial', 11, 'Tax returns for the last 4 years (corporate tax and VAT) and any letters from the tax office'),
  ('A. Financial', 12, 'Revenue by country and currency, and which company invoices in each country'),
  ('B. Commercial', 13, 'List of all franchisees: location, start date, contract length, fee rate and status'),
  ('B. Commercial', 14, 'Units that closed: date, reason and any money refunded'),
  ('B. Commercial', 15, 'Any complaints or disputes with franchisees'),
  ('B. Commercial', 16, 'Sales per unit, by month'),
  ('B. Commercial', 17, 'Franchise agreement template, and any versions that differ by country'),
  ('B. Commercial', 18, 'Cost to open a new unit, and what a franchisee invests in total'),
  ('B. Commercial', 19, 'Profit and loss of two or three typical units, and how long they take to pay back'),
  ('B. Commercial', 20, 'Leads, calls and signings by month for 24 months, and marketing spend'),
  ('B. Commercial', 21, 'Pipeline: units signed but not open, and those in discussion'),
  ('B. Commercial', 22, 'Packaging: supplier, buying price, selling price and volumes'),
  ('B. Commercial', 23, 'Main suppliers and delivery platform agreements'),
  ('B. Commercial', 24, 'Growth plan and any new product or income ideas'),
  ('B. Commercial', 25, 'Existing brochure and sales material'),
  ('C. Operational', 26, 'Company history: when you started, each opening and closure, each country'),
  ('C. Operational', 27, 'Operations manual and training material'),
  ('C. Operational', 28, 'What you provide to franchisees: training, marketing, support and quality checks'),
  ('C. Operational', 29, 'Who does what, including freelancers, and your own role day to day'),
  ('C. Operational', 30, 'List of freelancers, their contracts and how each is paid'),
  ('C. Operational', 31, 'Systems and accounts you use (booking, ads, domains, social media) and whose name they are in'),
  ('C. Operational', 32, 'How franchisee sales are reported to you and how royalties are collected'),
  ('C. Operational', 33, 'Company documents, shareholders and ownership percentages, and any other group companies'),
  ('C. Operational', 34, 'What exactly is included in the sale'),
  ('C. Operational', 35, 'Trademark registrations and licences in each country'),
  ('C. Operational', 36, 'Any disputes, guarantees, or contracts that need consent if the company is sold'),
  ('D. Brand and pictures', 37, 'Logo files and brand guidelines (colours and fonts)'),
  ('D. Brand and pictures', 38, 'High-resolution photos: salads, units, team and openings (with consent from anyone shown)'),
  ('D. Brand and pictures', 39, 'Menu and store design: layout, signage and packaging'),
  ('D. Brand and pictures', 40, 'Social media accounts and follower numbers')
    ) as t(s, n, l);
  end if;
end $$;
select (select count(*) from client_documents where engagement_id = '5dba111e-93f6-4752-bb39-35a2e02c72f0') as luxalad_items;
