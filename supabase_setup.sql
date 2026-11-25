-- VoiceStamp CRM - databaseopzet voor Supabase (versie 3: beslisser, adresstatus, outbox)
-- Plak dit in Supabase > SQL Editor > New query > Run.
-- Werkt zowel voor een nieuwe database als voor een bestaande (voegt alleen toe wat mist).

-- 1) Leads-tabel
create table if not exists public.leads (
    id                uuid primary key default gen_random_uuid(),
    user_id           uuid not null references auth.users(id) default auth.uid(),
    company_name      text,
    segment           text,
    website           text,
    email             text,
    phone             text,
    score             int,
    hook              text,
    channel_advice    text,
    status            text default 'Nieuw',
    note              text,
    next_action       text,
    next_action_date  date,
    dnc_reason        text,
    last_mailed       date,
    email_status      text,
    contact_name      text,
    contact_role      text,
    contact_email     text,
    contact_email_status text,
    done              boolean default false,
    created_at        timestamptz default now(),
    updated_at        timestamptz default now()
);

-- Kolommen toevoegen als de tabel al bestond zonder deze velden:
alter table public.leads add column if not exists next_action      text;
alter table public.leads add column if not exists next_action_date date;
alter table public.leads add column if not exists dnc_reason       text;
alter table public.leads add column if not exists last_mailed      date;
alter table public.leads add column if not exists email_status         text;
alter table public.leads add column if not exists contact_name         text;
alter table public.leads add column if not exists contact_role         text;
alter table public.leads add column if not exists contact_email        text;
alter table public.leads add column if not exists contact_email_status text;

alter table public.leads enable row level security;
drop policy if exists "leads_select_own" on public.leads;
create policy "leads_select_own" on public.leads for select using (auth.uid() = user_id);
drop policy if exists "leads_insert_own" on public.leads;
create policy "leads_insert_own" on public.leads for insert with check (auth.uid() = user_id);
drop policy if exists "leads_update_own" on public.leads;
create policy "leads_update_own" on public.leads for update using (auth.uid() = user_id);
drop policy if exists "leads_delete_own" on public.leads;
create policy "leads_delete_own" on public.leads for delete using (auth.uid() = user_id);
create index if not exists leads_user_idx on public.leads (user_id);

-- 2) Activiteiten-tabel (contacthistorie per lead)
create table if not exists public.activities (
    id           uuid primary key default gen_random_uuid(),
    user_id      uuid not null references auth.users(id) default auth.uid(),
    lead_id      uuid references public.leads(id) on delete cascade,
    type         text,
    description  text,
    created_at   timestamptz default now()
);

alter table public.activities enable row level security;
drop policy if exists "act_select_own" on public.activities;
create policy "act_select_own" on public.activities for select using (auth.uid() = user_id);
drop policy if exists "act_insert_own" on public.activities;
create policy "act_insert_own" on public.activities for insert with check (auth.uid() = user_id);
drop policy if exists "act_delete_own" on public.activities;
create policy "act_delete_own" on public.activities for delete using (auth.uid() = user_id);
create index if not exists act_user_idx on public.activities (user_id);
create index if not exists act_lead_idx on public.activities (lead_id);
create index if not exists act_date_idx on public.activities (created_at);

-- 3) Outbox (mailreeks per lead: mail 1, 2 en 3 met eigen verzenddatum en goedkeuring)
create table if not exists public.outbox (
    id            uuid primary key default gen_random_uuid(),
    user_id       uuid not null references auth.users(id) default auth.uid(),
    lead_id       uuid references public.leads(id) on delete cascade,
    step          int  not null default 1,
    to_email      text,
    subject       text,
    body          text,
    status        text not null default 'pending_review',  -- pending_review | approved | sent | rejected | cancelled
    scheduled_for date,
    sent_at       timestamptz,
    message_id    text,
    error         text,
    created_at    timestamptz default now()
);

alter table public.outbox enable row level security;
drop policy if exists "ob_select_own" on public.outbox;
create policy "ob_select_own" on public.outbox for select using (auth.uid() = user_id);
drop policy if exists "ob_insert_own" on public.outbox;
create policy "ob_insert_own" on public.outbox for insert with check (auth.uid() = user_id);
drop policy if exists "ob_update_own" on public.outbox;
create policy "ob_update_own" on public.outbox for update using (auth.uid() = user_id);
drop policy if exists "ob_delete_own" on public.outbox;
create policy "ob_delete_own" on public.outbox for delete using (auth.uid() = user_id);
create index if not exists ob_user_idx   on public.outbox (user_id);
create index if not exists ob_lead_idx   on public.outbox (lead_id);
create index if not exists ob_status_idx on public.outbox (status, scheduled_for);


-- 4) Extra kolommen bij leads (kanalen, score-uitleg, techniek) en de contentplanning
alter table public.leads add column if not exists instagram    text;
alter table public.leads add column if not exists linkedin     text;
alter table public.leads add column if not exists facebook     text;
alter table public.leads add column if not exists plaats       text;
alter table public.leads add column if not exists score_reason text;
alter table public.leads add column if not exists tech         text;
alter table public.leads add column if not exists form_url     text;
alter table public.leads add column if not exists whatsapp     text;

create table if not exists public.content_posts (
    id          uuid primary key default gen_random_uuid(),
    user_id     uuid not null references auth.users(id) default auth.uid(),
    gepland_op  date,
    tijd        text,
    kanaal      text,
    sector      text,
    formaat     text,
    tekst       text,
    beeld       text,
    status      text not null default 'concept',  -- concept | geplaatst
    created_at  timestamptz default now()
);

alter table public.content_posts enable row level security;
drop policy if exists "cp_select_own" on public.content_posts;
create policy "cp_select_own" on public.content_posts for select using (auth.uid() = user_id);
drop policy if exists "cp_insert_own" on public.content_posts;
create policy "cp_insert_own" on public.content_posts for insert with check (auth.uid() = user_id);
drop policy if exists "cp_update_own" on public.content_posts;
create policy "cp_update_own" on public.content_posts for update using (auth.uid() = user_id);
drop policy if exists "cp_delete_own" on public.content_posts;
create policy "cp_delete_own" on public.content_posts for delete using (auth.uid() = user_id);
create index if not exists cp_user_idx on public.content_posts (user_id, gepland_op);
