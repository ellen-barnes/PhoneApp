create table if not exists public.ledger_snapshots (
  user_id uuid primary key references auth.users (id) on delete cascade,
  payload jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default timezone('utc', now())
);

alter table public.ledger_snapshots enable row level security;

drop policy if exists "ledger_snapshots_select_own" on public.ledger_snapshots;
create policy "ledger_snapshots_select_own"
on public.ledger_snapshots
for select
to authenticated
using ((select auth.uid()) = user_id);

drop policy if exists "ledger_snapshots_insert_own" on public.ledger_snapshots;
create policy "ledger_snapshots_insert_own"
on public.ledger_snapshots
for insert
to authenticated
with check ((select auth.uid()) = user_id);

drop policy if exists "ledger_snapshots_update_own" on public.ledger_snapshots;
create policy "ledger_snapshots_update_own"
on public.ledger_snapshots
for update
to authenticated
using ((select auth.uid()) = user_id)
with check ((select auth.uid()) = user_id);
