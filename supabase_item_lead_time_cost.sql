-- Parametrizacao aditiva de lead time e custo por item.
-- Nao altera cadastro, saldo, movimentos, pedidos ou recebimentos existentes.

create table if not exists public.cadastro_item_parametros (
    registration_id bigint primary key
        references public.cadastro_registros(id) on delete cascade,
    sku text not null,
    origem_fabricacao text not null
        check (origem_fabricacao in ('INTERNA', 'EXTERNA')),
    unidade_tempo text not null default 'DIA_UTIL'
        check (unidade_tempo = 'DIA_UTIL'),
    fornecimento_dias numeric(12,3) not null default 0 check (fornecimento_dias >= 0),
    transporte_dias numeric(12,3) not null default 0 check (transporte_dias >= 0),
    recebimento_dias numeric(12,3) not null default 0 check (recebimento_dias >= 0),
    inspecao_recebimento_dias numeric(12,3) not null default 0 check (inspecao_recebimento_dias >= 0),
    estocagem_dias numeric(12,3) not null default 0 check (estocagem_dias >= 0),
    expedicao_dias numeric(12,3) not null default 0 check (expedicao_dias >= 0),
    montagem_kit_dias numeric(12,3) not null default 0 check (montagem_kit_dias >= 0),
    setup_dias numeric(12,3) not null default 0 check (setup_dias >= 0),
    producao_dias numeric(12,3) not null default 0 check (producao_dias >= 0),
    liberacao_dias numeric(12,3) not null default 0 check (liberacao_dias >= 0),
    preco_compra numeric(18,4) null check (preco_compra is null or preco_compra >= 0),
    updated_by text not null default '',
    source_type text not null default 'MANUAL',
    source_file text not null default '',
    source_sheet text not null default '',
    source_row integer null,
    source_hash text not null default '',
    version integer not null default 1,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

create index if not exists cadastro_item_parametros_sku_idx
    on public.cadastro_item_parametros (sku);

create index if not exists cadastro_item_parametros_origem_idx
    on public.cadastro_item_parametros (origem_fabricacao, sku);

create table if not exists public.cadastro_item_parametros_historico (
    id bigserial primary key,
    registration_id bigint not null,
    sku text not null,
    action text not null check (action in ('INSERT', 'UPDATE', 'DELETE')),
    actor text not null default '',
    before_data jsonb null,
    after_data jsonb null,
    changed_at timestamptz not null default now()
);

create index if not exists cadastro_item_parametros_historico_registration_idx
    on public.cadastro_item_parametros_historico (registration_id, changed_at desc);

create or replace function public.cadastro_item_parametros_before_update()
returns trigger
language plpgsql
as $$
begin
    new.updated_at = now();
    new.version = old.version + 1;
    return new;
end;
$$;

drop trigger if exists cadastro_item_parametros_touch on public.cadastro_item_parametros;
create trigger cadastro_item_parametros_touch
before update on public.cadastro_item_parametros
for each row execute function public.cadastro_item_parametros_before_update();

create or replace function public.cadastro_item_parametros_audit()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
    if tg_op = 'INSERT' then
        insert into public.cadastro_item_parametros_historico (
            registration_id, sku, action, actor, before_data, after_data
        ) values (
            new.registration_id, new.sku, tg_op, coalesce(new.updated_by, ''),
            null, to_jsonb(new)
        );
    elsif tg_op = 'UPDATE' then
        insert into public.cadastro_item_parametros_historico (
            registration_id, sku, action, actor, before_data, after_data
        ) values (
            new.registration_id, new.sku, tg_op, coalesce(new.updated_by, ''),
            to_jsonb(old), to_jsonb(new)
        );
    else
        insert into public.cadastro_item_parametros_historico (
            registration_id, sku, action, actor, before_data, after_data
        ) values (
            old.registration_id, old.sku, tg_op, coalesce(old.updated_by, ''),
            to_jsonb(old), null
        );
    end if;
    return null;
end;
$$;

drop trigger if exists cadastro_item_parametros_audit on public.cadastro_item_parametros;
create trigger cadastro_item_parametros_audit
after insert or update or delete on public.cadastro_item_parametros
for each row execute function public.cadastro_item_parametros_audit();

create or replace function public.cadastro_calcular_preco_medio(
    p_sku text,
    p_limite integer default 10
)
returns table (
    preco_medio numeric,
    quantidade_total numeric,
    entradas_consideradas bigint,
    primeira_entrada timestamptz,
    ultima_entrada timestamptz
)
language sql
stable
security definer
set search_path = public
as $$
    with entradas_recentes as (
        select
            l.valor_unitario_real,
            l.quantidade_aprovada,
            r.data_recebimento
        from public.erp_goods_receipt_lines l
        join public.erp_goods_receipts r on r.id = l.goods_receipt_id
        where r.status = 'CONFIRMADO'
          and upper(trim(coalesce(l.sku_codigo, ''))) = upper(trim(coalesce(p_sku, '')))
          and l.quantidade_aprovada > 0
          and l.valor_unitario_real > 0
        order by r.data_recebimento desc, r.created_at desc, l.id desc
        limit greatest(1, least(coalesce(p_limite, 10), 100))
    )
    select
        round(
            sum(valor_unitario_real * quantidade_aprovada)
            / nullif(sum(quantidade_aprovada), 0),
            4
        ) as preco_medio,
        coalesce(sum(quantidade_aprovada), 0) as quantidade_total,
        count(*) as entradas_consideradas,
        min(data_recebimento) as primeira_entrada,
        max(data_recebimento) as ultima_entrada
    from entradas_recentes;
$$;

alter table public.cadastro_item_parametros enable row level security;
alter table public.cadastro_item_parametros_historico enable row level security;

revoke all on table public.cadastro_item_parametros from anon, authenticated;
revoke all on table public.cadastro_item_parametros_historico from anon, authenticated;
grant select, insert, update, delete on table public.cadastro_item_parametros to service_role;
grant select on table public.cadastro_item_parametros_historico to service_role;
grant usage, select on sequence public.cadastro_item_parametros_historico_id_seq to service_role;
revoke all on function public.cadastro_calcular_preco_medio(text, integer) from public, anon, authenticated;
grant execute on function public.cadastro_calcular_preco_medio(text, integer) to service_role;

comment on table public.cadastro_item_parametros is
    'Lead time padrao em dias uteis e preco de compra parametrizado por SKU. Nao movimenta estoque.';
comment on function public.cadastro_calcular_preco_medio(text, integer) is
    'Media ponderada das ultimas entradas confirmadas com quantidade aprovada e valor real positivo.';
