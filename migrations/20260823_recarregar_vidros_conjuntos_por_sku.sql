-- Os registros legados de Vidros não trazem grupo_codigo no JSON.
-- Para eles, SKUs 30... representam CONJUNTO / KIT e recebem CJ na descrição.
begin;

with source as (
    select
        r.id,
        r.sku,
        regexp_replace(coalesce(r.field_values ->> 'prefixo', ''), E'^[0-9]+\\s*-\\s*', '') as prefixo,
        regexp_replace(coalesce(r.field_values ->> 'descritor_base', ''), E'^[0-9]+\\s*-\\s*', '') as nome_comercial,
        regexp_replace(coalesce(r.field_values ->> 'posicao_lado', ''), E'^[0-9]+\\s*-\\s*', '') as veiculo_modelo,
        regexp_replace(coalesce(r.field_values ->> 'fornecedor', ''), E'^[0-9]+\\s*-\\s*', '') as fornecedor
    from public.cadastro_registros r
    where r.category_key = 'cat_12_vidros'
      and r.sku like '30%'
), changes as (
    select
        source.*,
        concat_ws(
            ' ',
            case when upper(prefixo) <> 'CJ' then 'CJ' else null end,
            nullif(prefixo, ''),
            nullif(nome_comercial, ''),
            nullif(veiculo_modelo, ''),
            nullif(fornecedor, '')
        ) as descricao_primaria_nova
    from source
), updated as (
    update public.cadastro_registros r
       set descricao_primaria = c.descricao_primaria_nova,
           descricao_secundaria = c.descricao_primaria_nova,
           caracteres_primario = char_length(c.descricao_primaria_nova),
           caracteres_secundario = char_length(c.descricao_primaria_nova),
           search_text = lower(concat_ws(' ', r.sku, c.descricao_primaria_nova, r.field_values::text)),
           updated_at = timezone('utc', now())
      from changes c
     where r.id = c.id
       and (
           r.descricao_primaria is distinct from c.descricao_primaria_nova
           or r.descricao_secundaria is distinct from c.descricao_primaria_nova
       )
    returning r.id
)
insert into public.erp_audit_events (
    entity_type, entity_id, action, actor, origin, before_data, after_data, reason
)
select
    'CADASTRO_CATEGORIA', null, 'RECARREGAR_CONJUNTOS_VIDROS', 'sistema:cadastro', 'MIGRATION',
    jsonb_build_object('categoria', 'cat_12_vidros', 'registros_atualizados', count(*)),
    jsonb_build_object('regra', 'SKU 30... é grupo CONJUNTO e recebe CJ'),
    'Recarga complementar dos conjuntos legados de Vidros sem grupo_codigo.'
from updated;

commit;
