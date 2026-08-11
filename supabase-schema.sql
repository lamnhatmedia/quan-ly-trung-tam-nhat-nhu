-- Chạy toàn bộ file này trong Supabase > SQL Editor.
create table if not exists public.app_state (
  user_id uuid not null references auth.users(id) on delete cascade,
  key text not null,
  data jsonb not null,
  updated_at timestamptz not null default now(),
  primary key (user_id, key)
);

alter table public.app_state enable row level security;

drop policy if exists "Users can read own app state" on public.app_state;
drop policy if exists "Users can insert own app state" on public.app_state;
drop policy if exists "Users can update own app state" on public.app_state;
drop policy if exists "Users can delete own app state" on public.app_state;

create policy "Users can read own app state" on public.app_state
  for select using (auth.uid() = user_id);
create policy "Users can insert own app state" on public.app_state
  for insert with check (auth.uid() = user_id);
create policy "Users can update own app state" on public.app_state
  for update using (auth.uid() = user_id) with check (auth.uid() = user_id);
create policy "Users can delete own app state" on public.app_state
  for delete using (auth.uid() = user_id);
