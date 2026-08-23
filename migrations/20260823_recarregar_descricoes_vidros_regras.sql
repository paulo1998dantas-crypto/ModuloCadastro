-- Recalcula apenas as descrições dos SKUs da categoria 12 - VIDROS,
-- sem modificar campos, opções, B.O.M. ou saldo de estoque.
begin;

with source as (
    select
        r.id,
        r.sku,
        r.descricao_primaria as descricao_primaria_anterior,
        r.descricao_secundaria as descricao_secundaria_anterior,
        coalesce(
            nullif(r.form_values -> 'grupo_codigo' ->> 0, ''),
            case when r.sku like '30%' then '30' else '10' end
        ) as grupo_codigo,
        regexp_replace(coalesce(r.field_values ->> 'prefixo', ''), E'^[0-9]+\\s*-\\s*', '') as prefixo,
        regexp_replace(coalesce(r.field_values ->> 'descritor_base', ''), E'^[0-9]+\\s*-\\s*', '') as nome_comercial,
        regexp_replace(coalesce(r.field_values ->> 'veiculo_modelo', ''), E'^[0-9]+\\s*-\\s*', '') as lado_aplicacao,
        regexp_replace(coalesce(r.field_values ->> 'posicao_lado', ''), E'^[0-9]+\\s*-\\s*', '') as veiculo_modelo,
        regexp_replace(coalesce(r.field_values ->> 'medida', ''), E'^[0-9]+\\s*-\\s*', '') as local_aplicacao,
        regexp_replace(coalesce(r.field_values ->> 'complemento_regra', ''), E'^[0-9]+\\s*-\\s*', '') as comprimento,
        regexp_replace(coalesce(r.field_values ->> 'fornecedor_referencia', ''), E'^[0-9]+\\s*-\\s*', '') as largura,
        regexp_replace(coalesce(r.field_values ->> 'espessura', ''), E'^[0-9]+\\s*-\\s*', '') as espessura,
        regexp_replace(coalesce(r.field_values ->> 'fornecedor', ''), E'^[0-9]+\\s*-\\s*', '') as fornecedor,
        regexp_replace(coalesce(r.field_values ->> 'especificidade', ''), E'^[0-9]+\\s*-\\s*', '') as especificidade,
        regexp_replace(coalesce(r.field_values ->> 'material_cor', ''), E'^[0-9]+\\s*-\\s*', '') as material_cor
    from public.cadastro_registros r
    where r.category_key = 'cat_12_vidros'
), descriptions as (
    select
        source.*,
        concat_ws(
            ' ',
            case
                when grupo_codigo = '30' and upper(prefixo) <> 'CJ' then 'CJ'
                else null
            end,
            nullif(prefixo, ''),
            nullif(nome_comercial, ''),
            case when grupo_codigo <> '30' then nullif(lado_aplicacao, '') end,
            case
                when grupo_codigo <> '30' then nullif(
                    concat_ws('X', nullif(comprimento, ''), nullif(largura, ''), nullif(espessura, '')),
                    ''
                )
                else null
            end,
            nullif(veiculo_modelo, ''),
            case when grupo_codigo <> '30' then nullif(local_aplicacao, '') end,
            nullif(fornecedor, ''),
            case
                when grupo_codigo <> '30' and upper(especificidade) <> 'N/A' then nullif(especificidade, '')
                else null
            end
        ) as descricao_primaria_nova
    from source
), changes as (
    select
        descriptions.*,
        case
            when grupo_codigo <> '30' and nullif(material_cor, '') is not null
                then concat_ws(' ', descricao_primaria_nova, material_cor)
            else descricao_primaria_nova
        end as descricao_secundaria_nova
    from descriptions
), updated as (
    update public.cadastro_registros r
       set descricao_primaria = c.descricao_primaria_nova,
           descricao_secundaria = c.descricao_secundaria_nova,
           caracteres_primario = char_length(c.descricao_primaria_nova),
           caracteres_secundario = char_length(c.descricao_secundaria_nova),
           search_text = lower(concat_ws(' ', r.sku, c.descricao_primaria_nova, c.descricao_secundaria_nova, r.field_values::text)),
           updated_at = timezone('utc', now())
      from changes c
     where r.id = c.id
       and (
           r.descricao_primaria is distinct from c.descricao_primaria_nova
           or r.descricao_secundaria is distinct from c.descricao_secundaria_nova
       )
    returning r.id
)
insert into public.erp_audit_events (
    entity_type, entity_id, action, actor, origin, before_data, after_data, reason
)
select
    'CADASTRO_CATEGORIA', null, 'RECARREGAR_DESCRICOES_VIDROS', 'sistema:cadastro', 'MIGRATION',
    jsonb_build_object('categoria', 'cat_12_vidros', 'registros_atualizados', count(*)),
    jsonb_build_object(
        'regra_na', 'ESPECIFICIDADE = N/A não compõe a descrição',
        'regra_conjunto', 'Grupo 30 recebe CJ antes do prefixo real'
    ),
    'Recarga controlada das descrições de Vidros conforme regras de catálogo.'
from updated;

commit;
