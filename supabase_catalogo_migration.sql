-- ModuloCadastro catalog persistence
-- Safe for the shared Supabase project: creates only a cadastro_ table.

create table if not exists public.cadastro_catalogo (
    config_key text primary key,
    payload jsonb not null default '{}'::jsonb,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

drop trigger if exists cadastro_catalogo_touch_updated_at on public.cadastro_catalogo;
create trigger cadastro_catalogo_touch_updated_at
before update on public.cadastro_catalogo
for each row execute function public.cadastro_touch_updated_at();

alter table public.cadastro_catalogo enable row level security;
