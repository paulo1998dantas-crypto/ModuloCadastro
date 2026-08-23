-- Regra auditável de Vidros:
-- 1) ESPECIFICIDADE = N/A permanece registrada, mas não integra a descrição;
-- 2) para o grupo 30 - CONJUNTO / KIT, a descrição inicia com CJ.
begin;

do $$
declare
    v_payload jsonb;
    v_before_category jsonb;
    v_after_category jsonb;
    v_conditional_rules jsonb;
    v_description_rules jsonb;
    v_categories jsonb;
begin
    select payload
      into v_payload
      from public.cadastro_catalogo
     where config_key = 'default'
     for update;

    if v_payload is null then
        raise exception 'Catálogo padrão não encontrado.';
    end if;

    select category
      into v_before_category
      from jsonb_array_elements(v_payload -> 'categories') as category
     where category ->> 'key' = 'cat_12_vidros'
     limit 1;

    if v_before_category is null then
        raise exception 'Categoria cat_12_vidros não encontrada no catálogo.';
    end if;

    select coalesce(jsonb_agg(rule), '[]'::jsonb)
      into v_conditional_rules
      from jsonb_array_elements(coalesce(v_before_category -> 'conditional_rules', '[]'::jsonb)) as rule
     where not (
         lower(coalesce(rule ->> 'action', '')) = 'hide'
         and coalesce(rule ->> 'source_type', 'field') = 'field'
         and coalesce(rule ->> 'source_field_key', '') = 'especificidade'
         and coalesce(rule ->> 'target_field_key', '') = 'especificidade'
     );

    v_conditional_rules := v_conditional_rules || jsonb_build_array(
        jsonb_build_object(
            'key', 'vidros_especificidade_na_oculta',
            'action', 'hide',
            'match_by', 'option',
            'source_type', 'field',
            'source_field_key', 'especificidade',
            'source_field_label', 'ESPECIFICIDADE',
            'source_field_scope', 'primaria',
            'source_values', jsonb_build_array('NA'),
            'target_field_key', 'especificidade',
            'target_field_label', 'ESPECIFICIDADE',
            'target_field_scope', 'primaria',
            'origin', 'system'
        )
    );

    select coalesce(jsonb_agg(rule), '[]'::jsonb)
      into v_description_rules
      from jsonb_array_elements(coalesce(v_before_category -> 'description_rules', '[]'::jsonb)) as rule
     where coalesce(rule ->> 'key', '') <> 'vidros_grupo_conjunto_prefixo_cj';

    v_description_rules := v_description_rules || jsonb_build_array(
        jsonb_build_object(
            'key', 'vidros_grupo_conjunto_prefixo_cj',
            'action', 'prepend_literal',
            'source_type', 'group',
            'source_field_key', 'grupo_codigo',
            'source_field_label', 'GRUPO DO SKU',
            'source_values', jsonb_build_array('30'),
            'literal', 'CJ',
            'origin', 'system'
        )
    );

    v_after_category := jsonb_set(
        jsonb_set(v_before_category, '{conditional_rules}', v_conditional_rules, true),
        '{description_rules}', v_description_rules,
        true
    );

    select jsonb_agg(
        case
            when category ->> 'key' = 'cat_12_vidros' then v_after_category
            else category
        end
    )
      into v_categories
      from jsonb_array_elements(v_payload -> 'categories') as category;

    update public.cadastro_catalogo
       set payload = jsonb_set(v_payload, '{categories}', v_categories, true),
           updated_at = timezone('utc', now())
     where config_key = 'default';

    insert into public.erp_audit_events (
        entity_type, entity_id, action, actor, origin, before_data, after_data, reason
    ) values (
        'CADASTRO_CATEGORIA', null, 'ATUALIZAR_REGRAS_VIDROS', 'sistema:cadastro', 'MIGRATION',
        v_before_category, v_after_category,
        'N/A deixa de compor a descrição; conjuntos de Vidros passam a receber CJ automaticamente.'
    );
end $$;

commit;
