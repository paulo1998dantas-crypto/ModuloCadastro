-- Reafirma as regras de Vidros após qualquer gravação concorrente de catálogo.
-- É idempotente e não altera cadastros nem movimentos.
begin;

do $$
declare
    v_payload jsonb;
    v_before jsonb;
    v_after jsonb;
begin
    select payload into v_payload
      from public.cadastro_catalogo
     where config_key = 'default'
     for update;

    select category into v_before
      from jsonb_array_elements(v_payload -> 'categories') as category
     where category ->> 'key' = 'cat_12_vidros';

    if v_before is null then
        raise exception 'Categoria cat_12_vidros não encontrada no catálogo.';
    end if;

    v_after := jsonb_set(
        v_before,
        '{conditional_rules}',
        (
            select coalesce(jsonb_agg(rule), '[]'::jsonb)
              from jsonb_array_elements(coalesce(v_before -> 'conditional_rules', '[]'::jsonb)) as rule
             where coalesce(rule ->> 'key', '') <> 'vidros_especificidade_na_oculta'
        ) || jsonb_build_array(jsonb_build_object(
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
        )),
        true
    );

    v_after := jsonb_set(
        v_after,
        '{description_rules}',
        (
            select coalesce(jsonb_agg(rule), '[]'::jsonb)
              from jsonb_array_elements(coalesce(v_after -> 'description_rules', '[]'::jsonb)) as rule
             where coalesce(rule ->> 'key', '') <> 'vidros_grupo_conjunto_prefixo_cj'
        ) || jsonb_build_array(jsonb_build_object(
            'key', 'vidros_grupo_conjunto_prefixo_cj',
            'action', 'prepend_literal',
            'source_type', 'group',
            'source_field_key', 'grupo_codigo',
            'source_field_label', 'GRUPO DO SKU',
            'source_values', jsonb_build_array('30'),
            'literal', 'CJ',
            'origin', 'system'
        )),
        true
    );

    update public.cadastro_catalogo
       set payload = jsonb_set(
               v_payload,
               '{categories}',
               (
                   select jsonb_agg(case when category ->> 'key' = 'cat_12_vidros' then v_after else category end)
                     from jsonb_array_elements(v_payload -> 'categories') as category
               ),
               true
           ),
           updated_at = timezone('utc', now())
     where config_key = 'default';

    insert into public.erp_audit_events (
        entity_type, entity_id, action, actor, origin, before_data, after_data, reason
    ) values (
        'CADASTRO_CATEGORIA', null, 'REAFIRMAR_REGRAS_VIDROS', 'sistema:cadastro', 'MIGRATION',
        v_before, v_after,
        'Reafirma a ocultação de ESPECIFICIDADE=N/A e o prefixo CJ para conjuntos de Vidros.'
    );
end $$;

commit;
