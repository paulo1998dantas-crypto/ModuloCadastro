begin;

with inconsistentes as materialized (
    select
        id,
        sku,
        category_key,
        case
            when jsonb_typeof(coalesce(form_values, '{}'::jsonb)->'grupo_codigo') = 'array'
                then coalesce(form_values->'grupo_codigo'->>0, '')
            else coalesce(form_values->>'grupo_codigo', '')
        end as grupo_anterior,
        left(trim(sku), 2) as grupo_correto
    from public.cadastro_registros
    where trim(sku) ~ '^(10|20|30)'
      and coalesce(
            substring(
                case
                    when jsonb_typeof(coalesce(form_values, '{}'::jsonb)->'grupo_codigo') = 'array'
                        then coalesce(form_values->'grupo_codigo'->>0, '')
                    else coalesce(form_values->>'grupo_codigo', '')
                end
                from '^\s*(\d+)'
            ),
            ''
          ) <> left(trim(sku), 2)
),
auditoria as (
insert into public.erp_audit_events (
    entity_type,
    entity_id,
    action,
    actor,
    origin,
    before_data,
    after_data,
    reason
)
select
    'CADASTRO_REGISTRO',
    null,
    'CORRECAO_GRUPO_POR_PREFIXO_SKU',
    'sistema:migracao-controlada',
    'CADASTRO',
    jsonb_build_object(
        'cadastro_registro_id', id,
        'sku', sku,
        'category_key', category_key,
        'grupo_codigo', grupo_anterior
    ),
    jsonb_build_object(
        'cadastro_registro_id', id,
        'sku', sku,
        'category_key', category_key,
        'grupo_codigo', grupo_correto
    ),
    'Regra estrutural restaurada: SKU 10 = INSUMO, 20 = PRODUTO PROCESSO e 30 = CONJUNTO / KIT.'
from inconsistentes
returning 1
),
atualizacao as (
update public.cadastro_registros r
set
    form_values = jsonb_set(
        coalesce(r.form_values, '{}'::jsonb),
        '{grupo_codigo}',
        jsonb_build_array(i.grupo_correto),
        true
    ),
    updated_at = now()
from inconsistentes i
where r.id = i.id
returning r.id
)
select
    (select count(*) from inconsistentes) as identificados,
    (select count(*) from auditoria) as auditados,
    (select count(*) from atualizacao) as atualizados;

commit;
