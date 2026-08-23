-- Rollback da ultima aplicacao da correcao controlada de VIDROS.
-- Restaura somente a categoria armazenada no evento de auditoria.
do $$
declare
    v_category_index integer;
    v_before jsonb;
begin
    select before_data
      into v_before
      from public.erp_audit_events
     where entity_type = 'CADASTRO_CATALOGO'
       and action = 'ATUALIZAR_REGRAS_VIDROS'
       and origin = 'MODULO_CADASTRO'
     order by created_at desc
     limit 1;

    if v_before is null then
        raise exception 'Evento de auditoria da correcao de VIDROS nao encontrado.';
    end if;

    select ordinality::integer - 1
      into v_category_index
      from public.cadastro_catalogo catalog
      cross join lateral jsonb_array_elements(catalog.payload->'categories')
           with ordinality as categories(category, ordinality)
     where catalog.config_key = 'default'
       and category->>'key' = 'cat_12_vidros';

    if v_category_index is null then
        raise exception 'Categoria cat_12_vidros nao encontrada para rollback.';
    end if;

    update public.cadastro_catalogo
       set payload = jsonb_set(
               payload,
               array['categories', v_category_index::text],
               v_before
           ),
           updated_at = now()
     where config_key = 'default';

    insert into public.erp_audit_events(
        entity_type, entity_id, action, actor, origin,
        before_data, after_data, reason
    ) values (
        'CADASTRO_CATALOGO', null, 'ROLLBACK_REGRAS_VIDROS',
        'sistema:migration', 'MODULO_CADASTRO',
        '{}'::jsonb, v_before,
        'Rollback controlado da regra COMPRIMENTOxLARGURA e opcoes da categoria VIDROS.'
    );
end $$;
