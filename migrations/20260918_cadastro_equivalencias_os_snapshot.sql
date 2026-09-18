-- Códigos equivalentes por aplicação.
-- Um SKU continua sendo o registro físico, fiscal e rastreável. O grupo
-- somente consolida planejamento e libera escolhas controladas na O.S.

create table if not exists public.cadastro_grupos_equivalencia (
    id uuid primary key,
    codigo text not null unique,
    nome text not null,
    aplicacao text not null default '',
    unidade_funcional text not null default 'pc',
    observacoes text not null default '',
    ativo boolean not null default true,
    created_by text not null default 'SISTEMA',
    created_by_user_id bigint,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now(),
    constraint cadastro_grupos_equivalencia_codigo_ck check (btrim(codigo) <> ''),
    constraint cadastro_grupos_equivalencia_nome_ck check (btrim(nome) <> '')
);

create table if not exists public.cadastro_equivalencia_membros (
    id uuid primary key,
    grupo_id uuid not null references public.cadastro_grupos_equivalencia(id) on delete restrict,
    registration_id bigint references public.cadastro_registros(id) on delete restrict,
    sku text not null,
    fator_unidade_funcional numeric(18,6) not null default 1,
    prioridade integer not null default 100,
    ativo boolean not null default true,
    observacoes text not null default '',
    created_by text not null default 'SISTEMA',
    created_by_user_id bigint,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now(),
    constraint cadastro_equivalencia_membros_sku_ck check (btrim(sku) <> ''),
    constraint cadastro_equivalencia_membros_fator_ck check (fator_unidade_funcional > 0),
    constraint cadastro_equivalencia_membros_unique unique (grupo_id, sku)
);

create index if not exists cadastro_equivalencia_membros_sku_idx
    on public.cadastro_equivalencia_membros (sku)
    where ativo;
create index if not exists cadastro_equivalencia_membros_grupo_idx
    on public.cadastro_equivalencia_membros (grupo_id)
    where ativo;

-- Registro imutável da decisão tomada na composição de uma O.S. O documento
-- preserva a composição corrente; esta tabela preserva cada decisão anterior.
create table if not exists public.erp_component_equivalence_snapshots (
    id uuid primary key default gen_random_uuid(),
    documento_id bigint not null,
    work_order_id uuid,
    line_id text not null,
    snapshot_hash text not null,
    grupo_id uuid,
    grupo_codigo text not null default '',
    grupo_nome text not null default '',
    sku_planejado text not null,
    sku_selecionado text not null,
    quantidade_planejada numeric(18,6),
    quantidade_selecionada numeric(18,6),
    fator_planejado numeric(18,6),
    fator_selecionado numeric(18,6),
    motivo text not null default '',
    selecionado_por text not null default 'SISTEMA',
    selected_at timestamptz not null default now(),
    source_line jsonb not null default '{}'::jsonb,
    created_at timestamptz not null default now(),
    constraint erp_component_equivalence_snapshots_unique
        unique (documento_id, line_id, snapshot_hash)
);

create index if not exists erp_component_equivalence_snapshots_work_order_idx
    on public.erp_component_equivalence_snapshots (work_order_id, created_at desc);
create index if not exists erp_component_equivalence_snapshots_document_idx
    on public.erp_component_equivalence_snapshots (documento_id, created_at desc);

create or replace function public.erp_capture_component_equivalence_snapshot()
returns trigger
language plpgsql
set search_path = public
as $$
declare
    component jsonb;
    effective_line_id text;
    effective_hash text;
begin
    if lower(coalesce(new.tipo, '')) <> 'os'
       or jsonb_typeof(new.composicao) <> 'array' then
        return new;
    end if;

    for component in select value from jsonb_array_elements(new.composicao)
    loop
        if nullif(btrim(coalesce(component->>'equivalence_group_id', '')), '') is null then
            continue;
        end if;

        effective_line_id := nullif(btrim(component->>'line_id'), '');
        if effective_line_id is null then
            effective_line_id := md5(
                coalesce(component->>'item', '') || '|' ||
                coalesce(component->>'sku_planejado', component->>'codigo', '') || '|' ||
                coalesce(component->>'level', '0')
            );
        end if;
        effective_hash := md5(
            jsonb_build_object(
                'grupo', component->>'equivalence_group_id',
                'planejado', component->>'sku_planejado',
                'selecionado', component->>'sku_selecionado',
                'quantidade_planejada', component->>'quantidade_planejada',
                'quantidade_selecionada', component->>'qtd',
                'fator_planejado', component->>'equivalence_planned_factor',
                'fator_selecionado', component->>'equivalence_selected_factor',
                'motivo', component->>'equivalence_reason'
            )::text
        );

        insert into public.erp_component_equivalence_snapshots(
            documento_id, work_order_id, line_id, snapshot_hash,
            grupo_id, grupo_codigo, grupo_nome,
            sku_planejado, sku_selecionado,
            quantidade_planejada, quantidade_selecionada,
            fator_planejado, fator_selecionado,
            motivo, selecionado_por, source_line
        ) values (
            new.id, new.erp_work_order_id, effective_line_id, effective_hash,
            nullif(component->>'equivalence_group_id', '')::uuid,
            coalesce(component->>'equivalence_group_code', ''),
            coalesce(component->>'equivalence_group_name', ''),
            coalesce(component->>'sku_planejado', component->>'codigo', ''),
            coalesce(component->>'sku_selecionado', component->>'codigo', ''),
            nullif(component->>'quantidade_planejada', '')::numeric,
            nullif(component->>'qtd', '')::numeric,
            nullif(component->>'equivalence_planned_factor', '')::numeric,
            nullif(component->>'equivalence_selected_factor', '')::numeric,
            coalesce(component->>'equivalence_reason', ''),
            coalesce(component->>'equivalence_selected_by', 'SISTEMA'),
            component
        ) on conflict (documento_id, line_id, snapshot_hash) do update
          set work_order_id = coalesce(excluded.work_order_id, erp_component_equivalence_snapshots.work_order_id);
    end loop;
    return new;
end;
$$;

drop trigger if exists erp_capture_component_equivalence_snapshot on public.suprimentos_documentos;
create trigger erp_capture_component_equivalence_snapshot
after insert or update of composicao, erp_work_order_id on public.suprimentos_documentos
for each row execute function public.erp_capture_component_equivalence_snapshot();

-- O histórico não pode ser reescrito. A única atualização permitida é ligar
-- um snapshot já capturado ao UUID da O.S. quando esse vínculo é criado depois
-- do documento; o conteúdo da decisão continua imutável.
create or replace function public.erp_component_equivalence_snapshot_immutable()
returns trigger
language plpgsql
set search_path = public
as $$
begin
    if tg_op = 'DELETE' then
        raise exception 'erp_component_equivalence_snapshots e imutavel';
    end if;
    if new.documento_id is distinct from old.documento_id
       or new.line_id is distinct from old.line_id
       or new.snapshot_hash is distinct from old.snapshot_hash
       or new.grupo_id is distinct from old.grupo_id
       or new.grupo_codigo is distinct from old.grupo_codigo
       or new.grupo_nome is distinct from old.grupo_nome
       or new.sku_planejado is distinct from old.sku_planejado
       or new.sku_selecionado is distinct from old.sku_selecionado
       or new.quantidade_planejada is distinct from old.quantidade_planejada
       or new.quantidade_selecionada is distinct from old.quantidade_selecionada
       or new.fator_planejado is distinct from old.fator_planejado
       or new.fator_selecionado is distinct from old.fator_selecionado
       or new.motivo is distinct from old.motivo
       or new.selecionado_por is distinct from old.selecionado_por
       or new.selected_at is distinct from old.selected_at
       or new.source_line is distinct from old.source_line
       or new.created_at is distinct from old.created_at
       or (old.work_order_id is not null and new.work_order_id is distinct from old.work_order_id)
    then
        raise exception 'erp_component_equivalence_snapshots e imutavel';
    end if;
    return new;
end;
$$;

drop trigger if exists erp_component_equivalence_snapshot_immutable on public.erp_component_equivalence_snapshots;
create trigger erp_component_equivalence_snapshot_immutable
before update or delete on public.erp_component_equivalence_snapshots
for each row execute function public.erp_component_equivalence_snapshot_immutable();

alter table public.cadastro_grupos_equivalencia enable row level security;
alter table public.cadastro_equivalencia_membros enable row level security;
alter table public.erp_component_equivalence_snapshots enable row level security;
revoke all on table public.cadastro_grupos_equivalencia from anon, authenticated;
revoke all on table public.cadastro_equivalencia_membros from anon, authenticated;
revoke all on table public.erp_component_equivalence_snapshots from anon, authenticated;
grant select, insert, update on table public.cadastro_grupos_equivalencia to service_role;
grant select, insert, update on table public.cadastro_equivalencia_membros to service_role;
grant select, insert, update on table public.erp_component_equivalence_snapshots to service_role;
revoke execute on function public.erp_capture_component_equivalence_snapshot() from public, anon, authenticated;
revoke execute on function public.erp_component_equivalence_snapshot_immutable() from public, anon, authenticated;
