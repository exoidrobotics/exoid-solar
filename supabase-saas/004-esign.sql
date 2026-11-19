-- 004 — Online contract signing
-- Run once in Supabase → SQL Editor → New query → paste → Run.
-- Adds a table of signing links. Clients (not logged in) can only read or sign a contract through the two
-- functions at the bottom, and only with the long random token in their link.

create table if not exists public.contract_links (
  id           uuid primary key default gen_random_uuid(),
  token        text not null unique,
  company_id   uuid not null references public.companies on delete cascade,
  proposal_id  uuid references public.proposals on delete set null,
  created_by   uuid references auth.users on delete set null,
  created_at   timestamptz not null default now(),
  expires_at   timestamptz not null default now() + interval '60 days',
  doc          jsonb not null default '{}'::jsonb,       -- contract text and reference details, frozen when sent
  status       text not null default 'sent' check (status in ('sent','signed','void')),
  viewed_at    timestamptz,
  signed_at    timestamptz,
  signer_name  text,
  signature    text,                                     -- PNG data URL of the drawn signature
  signer_meta  jsonb
);
create index if not exists contract_links_company_idx  on public.contract_links(company_id, created_at desc);
create index if not exists contract_links_proposal_idx on public.contract_links(proposal_id);

alter table public.contract_links enable row level security;
drop policy if exists contract_links_read   on public.contract_links;
drop policy if exists contract_links_insert on public.contract_links;
drop policy if exists contract_links_update on public.contract_links;
create policy contract_links_read   on public.contract_links for select to authenticated using (company_id = my_company());
create policy contract_links_insert on public.contract_links for insert to authenticated
  with check (company_id = my_company() and created_by = auth.uid() and company_active(company_id) and length(token) >= 20);
create policy contract_links_update on public.contract_links for update to authenticated
  using (company_id = my_company() and status = 'sent') with check (company_id = my_company());   -- a signed agreement can't be voided

revoke all on public.contract_links from anon;
grant  select, insert on public.contract_links to authenticated;
grant  update (status) on public.contract_links to authenticated;   -- teams can only void a link; signing goes through sign_contract

-- Client side: read the contract behind a token
create or replace function public.get_contract(p_token text) returns json
language plpgsql security definer set search_path = public as $$
declare r contract_links; c companies;
begin
  select * into r from contract_links where token = p_token;
  if not found or r.status = 'void' then
    return json_build_object('ok', false, 'error', 'This signing link is no longer valid. Please ask the contractor for a new one.');
  end if;
  if r.status <> 'signed' and r.expires_at < now() then
    return json_build_object('ok', false, 'error', 'This signing link has expired. Please ask the contractor for a new one.');
  end if;
  if r.viewed_at is null then update contract_links set viewed_at = now() where id = r.id; end if;
  select * into c from companies where id = r.company_id;
  return json_build_object('ok', true, 'doc', r.doc, 'status', r.status, 'signed_at', r.signed_at,
    'signer_name', r.signer_name, 'signature', r.signature,
    'company', json_build_object('name', c.name, 'logo', c.logo, 'phone', c.phone, 'email', c.email, 'address', c.address));
end $$;

-- Client side: sign it (once)
create or replace function public.sign_contract(p_token text, p_name text, p_signature text) returns json
language plpgsql security definer set search_path = public as $$
declare r contract_links; h json; n int;
begin
  select * into r from contract_links where token = p_token;
  if not found or r.status = 'void' or (r.status = 'sent' and r.expires_at < now()) then
    return json_build_object('ok', false, 'error', 'This signing link is no longer valid. Please ask the contractor for a new one.');
  end if;
  if r.status = 'signed' then return json_build_object('ok', false, 'error', 'This agreement has already been signed.'); end if;
  if length(trim(coalesce(p_name, ''))) < 2 then return json_build_object('ok', false, 'error', 'Please type your full name.'); end if;
  if p_signature is null or p_signature not like 'data:image/png;base64,%' or length(p_signature) > 400000 then
    return json_build_object('ok', false, 'error', 'Please draw your signature again.');
  end if;
  begin h := current_setting('request.headers', true)::json; exception when others then h := null; end;
  update contract_links set status = 'signed', signed_at = now(), signer_name = left(trim(p_name), 120), signature = p_signature,
    signer_meta = json_build_object('ip', split_part(coalesce(h->>'x-forwarded-for', h->>'cf-connecting-ip', ''), ',', 1), 'user_agent', left(coalesce(h->>'user-agent', ''), 300))
  where id = r.id and status = 'sent';
  get diagnostics n = row_count;
  if n = 0 then return json_build_object('ok', false, 'error', 'This agreement has already been signed.'); end if;
  return json_build_object('ok', true, 'signed_at', now());
end $$;

revoke execute on function public.get_contract(text), public.sign_contract(text, text, text) from public;
grant  execute on function public.get_contract(text), public.sign_contract(text, text, text) to anon, authenticated;
-- Note: if you ever re-run fix-permissions.sql (it removes all anon access to functions), run the two lines above again.

-- Check: should return two rows
select routine_name from information_schema.routines where routine_schema = 'public' and routine_name in ('get_contract', 'sign_contract');
