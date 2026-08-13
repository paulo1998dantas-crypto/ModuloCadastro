-- Biblioteca compartilhada de layouts de O.S.
-- Mudança aditiva: não remove nem reinterpreta documentos existentes.

create table if not exists public.layout_arquivos (
    id uuid primary key default gen_random_uuid(),
    sha256 text not null unique,
    storage_bucket text not null default 'os-layouts',
    storage_path text not null unique,
    nome_original text not null,
    nome_exibicao text not null,
    mime_type text not null default 'application/pdf',
    tamanho_bytes bigint not null check (tamanho_bytes > 0),
    criado_por text,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

alter table public.layout_arquivos enable row level security;

insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values (
    'os-layouts',
    'os-layouts',
    false,
    10485760,
    array['application/pdf']::text[]
)
on conflict (id) do update
set public = excluded.public,
    file_size_limit = excluded.file_size_limit,
    allowed_mime_types = excluded.allowed_mime_types;

alter table public.suprimentos_documentos
    add column if not exists layout_arquivo_id uuid;

do $$
begin
    if not exists (
        select 1
        from pg_constraint
        where conname = 'suprimentos_documentos_layout_arquivo_id_fkey'
          and conrelid = 'public.suprimentos_documentos'::regclass
    ) then
        alter table public.suprimentos_documentos
            add constraint suprimentos_documentos_layout_arquivo_id_fkey
            foreign key (layout_arquivo_id)
            references public.layout_arquivos(id)
            on delete restrict;
    end if;
end $$;

create index if not exists suprimentos_documentos_layout_arquivo_id_idx
    on public.suprimentos_documentos (layout_arquivo_id)
    where layout_arquivo_id is not null;

create or replace function public.touch_layout_arquivos_updated_at()
returns trigger
language plpgsql
security invoker
set search_path = public
as $$
begin
    new.updated_at = now();
    return new;
end;
$$;

drop trigger if exists touch_layout_arquivos_updated_at on public.layout_arquivos;
create trigger touch_layout_arquivos_updated_at
before update on public.layout_arquivos
for each row execute function public.touch_layout_arquivos_updated_at();

revoke all on table public.layout_arquivos from anon, authenticated;
revoke all on function public.touch_layout_arquivos_updated_at() from public;
