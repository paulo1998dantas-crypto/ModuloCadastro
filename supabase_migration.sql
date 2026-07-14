-- ModuloCadastro Supabase migration
-- Safe for a shared production database: creates only new objects with the cadastro_ prefix.

create table if not exists public.cadastro_registros (
    id bigserial primary key,
    category_key text not null,
    category_label text not null default '',
    sheet text not null default '',
    sku text not null,
    descricao_primaria text not null default '',
    descricao_secundaria text not null default '',
    sufixo text not null default '',
    caracteres_primario integer not null default 0,
    caracteres_secundario integer not null default 0,
    form_values jsonb not null default '{}'::jsonb,
    field_values jsonb not null default '{}'::jsonb,
    field_codes jsonb not null default '{}'::jsonb,
    search_text text not null default '',
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now(),
    constraint cadastro_registros_category_sku_unique unique (category_key, sku)
);

create table if not exists public.cadastro_rascunhos (
    draft_id text primary key,
    category_key text not null,
    category_label text not null default '',
    sheet text not null default '',
    descricao_primaria text not null default '',
    payload jsonb not null default '{}'::jsonb,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

create index if not exists cadastro_registros_category_sku_idx
    on public.cadastro_registros (category_key, sku);

create index if not exists cadastro_registros_search_idx
    on public.cadastro_registros using gin (to_tsvector('simple', search_text));

create index if not exists cadastro_rascunhos_updated_idx
    on public.cadastro_rascunhos (updated_at desc);

create or replace function public.cadastro_touch_updated_at()
returns trigger
language plpgsql
as $$
begin
    new.updated_at = now();
    return new;
end;
$$;

drop trigger if exists cadastro_registros_touch_updated_at on public.cadastro_registros;
create trigger cadastro_registros_touch_updated_at
before update on public.cadastro_registros
for each row execute function public.cadastro_touch_updated_at();

drop trigger if exists cadastro_rascunhos_touch_updated_at on public.cadastro_rascunhos;
create trigger cadastro_rascunhos_touch_updated_at
before update on public.cadastro_rascunhos
for each row execute function public.cadastro_touch_updated_at();

alter table public.cadastro_registros enable row level security;
alter table public.cadastro_rascunhos enable row level security;
