-- รันไฟล์นี้ใน Supabase Dashboard > SQL Editor (ครั้งเดียว)
-- ข้อมูลของแต่ละห้องเรียนเก็บเป็น JSON 1 ก้อนต่อห้อง เจ้าของข้อมูลเห็นได้เฉพาะของตัวเอง (RLS)

create table if not exists public.school_settings (
  user_id    uuid primary key default auth.uid() references auth.users(id) on delete cascade,
  data       jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

create table if not exists public.classrooms (
  id         uuid primary key default gen_random_uuid(),
  user_id    uuid not null default auth.uid() references auth.users(id) on delete cascade,
  title      text not null,
  year       int,
  data       jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

create index if not exists classrooms_user_idx on public.classrooms(user_id, year desc);

alter table public.school_settings enable row level security;
alter table public.classrooms      enable row level security;

drop policy if exists "own school settings" on public.school_settings;
create policy "own school settings" on public.school_settings
  for all using (auth.uid() = user_id) with check (auth.uid() = user_id);

drop policy if exists "own classrooms" on public.classrooms;
create policy "own classrooms" on public.classrooms
  for all using (auth.uid() = user_id) with check (auth.uid() = user_id);
