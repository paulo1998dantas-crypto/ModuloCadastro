-- Une COMPRIMENTO, LARGURA e ESPESSURA na medida principal dos vidros.
-- Atualiza somente a configuracao do catalogo; nenhum SKU e alterado.
do $$
declare
    v_category_index integer;
    v_before jsonb;
    v_after jsonb;
    v_rules jsonb;
begin
    select ordinality::integer - 1, category
      into v_category_index, v_before
      from public.cadastro_catalogo catalog
      cross join lateral jsonb_array_elements(catalog.payload->'categories')
           with ordinality as categories(category, ordinality)
     where catalog.config_key = 'default'
       and category->>'key' = 'cat_12_vidros';

    if v_before is null then
        raise exception 'Categoria cat_12_vidros nao encontrada no catalogo default.';
    end if;

    select coalesce(jsonb_agg(
        case
            when rule->>'key' = 'vidros_comprimento_x_largura' then rule || jsonb_build_object(
                'additional_target_field_keys', jsonb_build_array('espessura')
            )
            else rule
        end
        order by ordinality
    ), '[]'::jsonb)
      into v_rules
      from jsonb_array_elements(coalesce(v_before->'description_rules', '[]'::jsonb))
           with ordinality as rules(rule, ordinality);

    v_after := jsonb_set(v_before, '{description_rules}', v_rules);

    if v_after is distinct from v_before then
        update public.cadastro_catalogo
           set payload = jsonb_set(payload, array['categories', v_category_index::text], v_after),
               updated_at = now()
         where config_key = 'default';

        insert into public.erp_audit_events(
            entity_type, entity_id, action, actor, origin, before_data, after_data, reason
        ) values (
            'CADASTRO_CATALOGO', null, 'UNIR_MEDIDA_VIDROS_COMPRIMENTO_LARGURA_ESPESSURA',
            'sistema:migration', 'MODULO_CADASTRO', v_before, v_after,
            'Regra de composicao de vidros ajustada para COMPRIMENTOXLARGURAXESPESSURA.'
        );
    end if;
end $$;
