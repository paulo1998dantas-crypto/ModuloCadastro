-- Rastreabilidade imutável do Módulo Cadastro.
-- Registra eventos do cadastro principal e das B.O.M.; o histórico de
-- parâmetros continua sendo preservado em cadastro_item_parametros_historico
-- e é exibido na mesma consulta da aplicação.

create table if not exists public.cadastro_auditoria (
    id bigserial primary key,
    registration_id bigint null,
    sku text not null default '',
    previous_sku text not null default '',
    category_key text not null default '',
    category_label text not null default '',
    action text not null,
    actor text not null default 'sistema',
    summary text not null default '',
    details jsonb not null default '{}'::jsonb,
    created_at timestamptz not null default now()
);

create index if not exists cadastro_auditoria_registration_idx
    on public.cadastro_auditoria (registration_id, created_at desc);

create index if not exists cadastro_auditoria_sku_idx
    on public.cadastro_auditoria (sku, created_at desc);

create index if not exists cadastro_auditoria_action_idx
    on public.cadastro_auditoria (action, created_at desc);

create index if not exists cadastro_auditoria_created_idx
    on public.cadastro_auditoria (created_at desc, id desc);

alter table public.cadastro_auditoria enable row level security;
revoke all on table public.cadastro_auditoria from anon, authenticated;
grant select, insert on table public.cadastro_auditoria to service_role;
grant usage, select on sequence public.cadastro_auditoria_id_seq to service_role;

create or replace function public.cadastro_auditoria_immutable()
returns trigger
language plpgsql
set search_path = public
as $$
begin
    raise exception 'cadastro_auditoria e imutavel';
end;
$$;

drop trigger if exists cadastro_auditoria_immutable on public.cadastro_auditoria;
create trigger cadastro_auditoria_immutable
before update or delete on public.cadastro_auditoria
for each row execute function public.cadastro_auditoria_immutable();

comment on table public.cadastro_auditoria is
    'Trilha imutavel de criacao, alteracao, status, migracao, exclusao e B.O.M. do Modulo Cadastro.';

-- Marco inicial: aproveita as datas de criacao existentes sem inventar
-- alteracoes antigas que nao foram registradas antes desta funcionalidade.
insert into public.cadastro_auditoria (
    registration_id, sku, category_key, category_label, action, actor,
    summary, details, created_at
)
select
    r.id,
    coalesce(r.sku, ''),
    coalesce(r.category_key, ''),
    coalesce(r.category_label, ''),
    'criacao',
    'migracao',
    'Marco inicial do histórico; alterações anteriores à ativação não estavam disponíveis.',
    jsonb_build_object('antes', null, 'depois', to_jsonb(r), 'alteracoes', '{}'::jsonb),
    coalesce(r.created_at, now())
from public.cadastro_registros r
where not exists (
    select 1
    from public.cadastro_auditoria a
    where a.registration_id = r.id
      and a.action = 'criacao'
);

insert into public.cadastro_auditoria (
    registration_id, sku, category_key, category_label, action, actor,
    summary, details, created_at
)
select
    h.registration_id,
    coalesce(h.parent_sku, ''),
    coalesce(h.parent_category_key, ''),
    coalesce(h.parent_category_label, ''),
    'bom_criacao',
    'migracao',
    'Marco inicial da B.O.M.; alterações anteriores à ativação não estavam disponíveis.',
    jsonb_build_object(
        'antes', null,
        'depois', jsonb_build_object(
            'cabecalho', to_jsonb(h),
            'componentes', coalesce((
                select jsonb_agg(to_jsonb(c) order by c.ordem, c.id)
                from public.cadastro_bom_componentes c
                where c.bom_id = h.id
            ), '[]'::jsonb)
        ),
        'alteracoes', '{}'::jsonb
    ),
    coalesce(h.created_at, now())
from public.cadastro_bom_cabecalhos h
where not exists (
    select 1
    from public.cadastro_auditoria a
    where a.sku = h.parent_sku
      and a.action = 'bom_criacao'
);
