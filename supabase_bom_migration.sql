-- ModuloCadastro B.O.M. migration
-- Safe for a shared production database: creates only cadastro_bom_* objects.

create table if not exists public.cadastro_bom_cabecalhos (
    id bigserial primary key,
    parent_sku text not null unique,
    parent_descricao text not null default '',
    parent_category_key text not null default '',
    parent_category_label text not null default '',
    registration_id bigint null,
    source text not null default '',
    search_text text not null default '',
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

create table if not exists public.cadastro_bom_componentes (
    id bigserial primary key,
    bom_id bigint not null references public.cadastro_bom_cabecalhos(id) on delete cascade,
    parent_sku text not null,
    component_sku text not null,
    component_descricao text not null default '',
    unidade text not null default '',
    quantidade numeric not null default 0,
    ordem integer not null default 0,
    search_text text not null default '',
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

create index if not exists cadastro_bom_cabecalhos_parent_idx
    on public.cadastro_bom_cabecalhos (parent_sku);

create index if not exists cadastro_bom_cabecalhos_category_idx
    on public.cadastro_bom_cabecalhos (parent_category_key, parent_sku);

create index if not exists cadastro_bom_cabecalhos_search_idx
    on public.cadastro_bom_cabecalhos using gin (to_tsvector('simple', search_text));

create index if not exists cadastro_bom_componentes_bom_idx
    on public.cadastro_bom_componentes (bom_id, ordem);

create index if not exists cadastro_bom_componentes_parent_child_idx
    on public.cadastro_bom_componentes (parent_sku, component_sku);

create index if not exists cadastro_bom_componentes_search_idx
    on public.cadastro_bom_componentes using gin (to_tsvector('simple', search_text));

create or replace function public.cadastro_touch_updated_at()
returns trigger
language plpgsql
as $$
begin
    new.updated_at = now();
    return new;
end;
$$;

drop trigger if exists cadastro_bom_cabecalhos_touch_updated_at on public.cadastro_bom_cabecalhos;
create trigger cadastro_bom_cabecalhos_touch_updated_at
before update on public.cadastro_bom_cabecalhos
for each row execute function public.cadastro_touch_updated_at();

drop trigger if exists cadastro_bom_componentes_touch_updated_at on public.cadastro_bom_componentes;
create trigger cadastro_bom_componentes_touch_updated_at
before update on public.cadastro_bom_componentes
for each row execute function public.cadastro_touch_updated_at();

alter table public.cadastro_bom_cabecalhos enable row level security;
alter table public.cadastro_bom_componentes enable row level security;
