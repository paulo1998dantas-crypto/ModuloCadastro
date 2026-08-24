begin;

with affected as materialized (
    select
        id,
        sku,
        coalesce(form_values -> 'grupo_codigo', '[]'::jsonb) as grupo_anterior,
        case left(regexp_replace(sku, '[^0-9]', '', 'g'), 2)
            when '10' then '10'
            when '20' then '20'
            when '30' then '30'
        end as grupo_correto
    from public.cadastro_registros
    where category_key = 'cat_18_revestimento'
      and left(regexp_replace(sku, '[^0-9]', '', 'g'), 2) in ('10', '20', '30')
      and coalesce(form_values -> 'grupo_codigo' ->> 0, '') is distinct from
          left(regexp_replace(sku, '[^0-9]', '', 'g'), 2)
), audit as (
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
        'CADASTRO_REVESTIMENTO',
        null,
        'CORRECAO_GRUPO_POR_PREFIXO_SKU',
        'sistema:migracao-controlada',
        'CADASTRO',
        jsonb_build_object(
            'categoria', 'cat_18_revestimento',
            'registros', coalesce(
                jsonb_agg(
                    jsonb_build_object(
                        'id', id,
                        'sku', sku,
                        'grupo', grupo_anterior
                    ) order by sku
                ),
                '[]'::jsonb
            )
        ),
        jsonb_build_object(
            'categoria', 'cat_18_revestimento',
            'quantidade', count(*),
            'registros', coalesce(
                jsonb_agg(
                    jsonb_build_object(
                        'id', id,
                        'sku', sku,
                        'grupo', grupo_correto
                    ) order by sku
                ),
                '[]'::jsonb
            )
        ),
        'Reconciliação exclusiva de Revestimento: prefixos 10, 20 e 30 definem INSUMO, PRODUTO PROCESSO e CONJUNTO.'
    from affected
    having count(*) > 0
    returning id
)
update public.cadastro_registros as registration
set
    form_values = jsonb_set(
        registration.form_values,
        '{grupo_codigo}',
        jsonb_build_array(affected.grupo_correto),
        true
    ),
    updated_at = now()
from affected
where registration.id = affected.id
  and (select count(*) from audit) >= 0;

commit;
