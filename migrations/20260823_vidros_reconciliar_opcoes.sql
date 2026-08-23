-- Reconciliacao controlada das opcoes da categoria 12 - VIDROS.
-- Preserva todos os demais campos, regras e cadastros existentes.
do $$
declare
    v_category_index integer;
    v_before jsonb;
    v_after jsonb;
    v_fields jsonb;
    v_rules jsonb;
    v_dimension_fields jsonb := jsonb_build_array(
        jsonb_build_object('key', 'complemento_regra', 'label', 'COMPRIMENTO', 'scope', 'primaria', 'selection_mode', 'unitaria', 'description_order', 4, 'options', to_jsonb(array[
            '1- 671','2- 676','3- 686','4- 743','5- 771','6- 792','7- 814','8- 815','9- 830','10- 838','11- 923','12- 976','13- 1073','14- 1078','15- 1082','16- 1089','17- 1096','18- 1171','19- 1177','20- 1195','21- 1210','22- 1291','23- 1352','24- 1357','25- 1365','26- 1373','27- 1383','28- 1389','29- 1400','30- 1404','31- 1418','32- 1422','33- 1425','34- 1431','35- 1437','36- 1637','37- 1664','38- 1666'
        ]::text[])),
        jsonb_build_object('key', 'fornecedor_referencia', 'label', 'LARGURA', 'scope', 'primaria', 'selection_mode', 'unitaria', 'description_order', 5, 'options', to_jsonb(array[
            '1- 495','2- 534','3- 556','4- 583','5- 630','6- 667','7- 686','8- 705','9- 716','10- 728','11- 730','12- 738','13- 739','14- 741','15- 761','16- 766','17- 769','18- 771','19- 793','20- 798','21- 800','22- 802','23- 806','24- 814','25- 830','26- 833'
        ]::text[])),
        jsonb_build_object('key', 'espessura', 'label', 'ESPESSURA', 'scope', 'primaria', 'selection_mode', 'unitaria', 'description_order', 6, 'options', to_jsonb(array['1- 3MM','2- 4MM','3- 5MM']::text[])),
        jsonb_build_object('key', 'fornecedor', 'label', 'FORNECEDOR', 'scope', 'primaria', 'selection_mode', 'unitaria', 'description_order', 7, 'options', to_jsonb(array['1- SALTON','2- STYLLUS','3- VIDROFORTE','4- VETROEX']::text[])),
        jsonb_build_object('key', 'especificidade', 'label', 'ESPECIFICIDADE', 'scope', 'primaria', 'selection_mode', 'unitaria', 'description_order', 8, 'options', to_jsonb(array['1- N/A','2- 2º VAO LD VIDRO FIXO']::text[]))
    );
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

    select jsonb_agg(
               case
                   when field->>'key' in ('complemento_regra','fornecedor_referencia','espessura','fornecedor','especificidade')
                   then field || (
                       select dimension_field - 'key'
                         from jsonb_array_elements(v_dimension_fields) dimension_field
                        where dimension_field->>'key' = field->>'key'
                   )
                   else field
               end
               order by ordinality
           )
      into v_fields
      from jsonb_array_elements(v_before->'fields')
           with ordinality as fields(field, ordinality);

    select v_fields || coalesce(jsonb_agg(dimension_field order by dimension_field->>'key'), '[]'::jsonb)
      into v_fields
      from jsonb_array_elements(v_dimension_fields) dimension_field
     where not exists (
        select 1
          from jsonb_array_elements(v_before->'fields') existing_field
         where existing_field->>'key' = dimension_field->>'key'
     );

    select coalesce(jsonb_agg(rule order by ordinality), '[]'::jsonb)
      into v_rules
      from jsonb_array_elements(coalesce(v_before->'description_rules', '[]'::jsonb))
           with ordinality as rules(rule, ordinality)
     where rule->>'key' <> 'vidros_comprimento_x_largura';

    v_rules := v_rules || jsonb_build_array(jsonb_build_object(
        'key', 'vidros_comprimento_x_largura',
        'action', 'join_fields',
        'source_field_key', 'complemento_regra',
        'source_field_label', 'COMPRIMENTO',
        'target_field_key', 'fornecedor_referencia',
        'target_field_label', 'LARGURA',
        'separator', 'X'
    ));
    v_after := jsonb_set(jsonb_set(v_before, '{fields}', v_fields), '{description_rules}', v_rules);

    if v_after is distinct from v_before then
        update public.cadastro_catalogo
           set payload = jsonb_set(payload, array['categories', v_category_index::text], v_after),
               updated_at = now()
         where config_key = 'default';

        insert into public.erp_audit_events(
            entity_type, entity_id, action, actor, origin, before_data, after_data, reason
        ) values (
            'CADASTRO_CATALOGO', null, 'RECONCILIAR_OPCOES_VIDROS', 'sistema:migration', 'MODULO_CADASTRO',
            v_before, v_after,
            'Reconciliacao controlada dos cinco campos de Vidros e regra COMPRIMENTOxLARGURA.'
        );
    end if;
end $$;
