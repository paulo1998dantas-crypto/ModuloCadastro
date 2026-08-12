-- Exclusao segura de cadastro/SKU compartilhado.
--
-- Esta rotina NAO exclui historico operacional. Ela so remove um cadastro que
-- ainda nao possui uso em Estoque, Compras, Producao, Forecast ou B.O.M. A
-- resposta traz os bloqueios encontrados para a interface orientar o ADMIN.

create or replace function public.erp_delete_catalog_sku(p_registration_id bigint)
returns jsonb
language plpgsql
security invoker
set search_path = public
as $$
declare
    v_registration public.cadastro_registros%rowtype;
    v_sku public.skus%rowtype;
    v_blockers text[] := array[]::text[];
begin
    select *
      into v_registration
      from public.cadastro_registros
     where id = p_registration_id
     for update;

    if not found then
        raise exception 'Cadastro nao encontrado.';
    end if;

    if exists (
        select 1
          from public.cadastro_registros other_registration
         where other_registration.sku = v_registration.sku
           and other_registration.id <> v_registration.id
    ) then
        v_blockers := array_append(v_blockers, 'existe outro cadastro com o mesmo SKU');
    end if;

    select *
      into v_sku
      from public.skus
     where sku = v_registration.sku
     for update;

    if not found then
        v_blockers := array_append(v_blockers, 'SKU operacional nao localizado para conciliacao');
    else
        if exists (select 1 from public.stock_balances where sku_id = v_sku.id and saldo_atual <> 0) then
            v_blockers := array_append(v_blockers, 'saldo de estoque diferente de zero');
        end if;
        if exists (select 1 from public.movements where sku_id = v_sku.id) then
            v_blockers := array_append(v_blockers, 'movimentacoes de estoque');
        end if;
        if exists (select 1 from public.inventory_counts where sku_id = v_sku.id) then
            v_blockers := array_append(v_blockers, 'contagens de inventario');
        end if;
        if exists (select 1 from public.label_print_jobs where sku_id = v_sku.id) then
            v_blockers := array_append(v_blockers, 'historico de etiquetas');
        end if;
        if exists (select 1 from public.bom_components where item_sku_id = v_sku.id or component_sku_id = v_sku.id) then
            v_blockers := array_append(v_blockers, 'B.O.M. operacional');
        end if;
        if exists (select 1 from public.erp_purchase_order_lines where sku_id = v_sku.id) then
            v_blockers := array_append(v_blockers, 'linhas de pedido de compra');
        end if;
        if exists (select 1 from public.erp_goods_receipt_lines where sku_id = v_sku.id) then
            v_blockers := array_append(v_blockers, 'linhas de recebimento');
        end if;
        if exists (select 1 from public.erp_production_orders where target_sku_id = v_sku.id) then
            v_blockers := array_append(v_blockers, 'ordens de producao');
        end if;
        if exists (select 1 from public.erp_production_order_inputs where source_sku_id = v_sku.id) then
            v_blockers := array_append(v_blockers, 'insumos de ordem de producao');
        end if;
        if exists (select 1 from public.suprimentos_forecast_itens where sku_id = v_sku.id) then
            v_blockers := array_append(v_blockers, 'itens de forecast');
        end if;
        if exists (select 1 from public.suprimentos_forecast_necessidades where sku_id = v_sku.id) then
            v_blockers := array_append(v_blockers, 'necessidades de forecast');
        end if;
    end if;

    if exists (
        select 1
          from public.cadastro_bom_cabecalhos
         where registration_id = v_registration.id
            or parent_sku = v_registration.sku
    ) then
        v_blockers := array_append(v_blockers, 'B.O.M. do cadastro');
    end if;
    if exists (
        select 1
          from public.cadastro_bom_componentes
         where parent_sku = v_registration.sku
            or component_sku = v_registration.sku
    ) then
        v_blockers := array_append(v_blockers, 'referencias em B.O.M. de cadastro');
    end if;
    if exists (
        select 1
          from public.erp_work_orders
         where codigo_banco = v_registration.sku
    ) then
        v_blockers := array_append(v_blockers, 'ordens de servico vinculadas');
    end if;
    if exists (
        select 1
          from public.suprimentos_documentos
         where coalesce(itens, '[]'::jsonb)::text like '%' || v_registration.sku || '%'
            or coalesce(componentes, '[]'::jsonb)::text like '%' || v_registration.sku || '%'
            or coalesce(composicao, '[]'::jsonb)::text like '%' || v_registration.sku || '%'
    ) then
        v_blockers := array_append(v_blockers, 'documentos de suprimentos');
    end if;

    if cardinality(v_blockers) > 0 then
        return jsonb_build_object(
            'deleted', false,
            'sku', v_registration.sku,
            'blockers', to_jsonb(v_blockers)
        );
    end if;

    -- Um saldo zerado sem qualquer movimento e apenas um registro tecnico de
    -- inicializacao; ele pode sair junto com o SKU, na mesma transacao.
    delete from public.stock_balances where sku_id = v_sku.id;
    delete from public.skus where id = v_sku.id;
    delete from public.cadastro_registros where id = v_registration.id;

    return jsonb_build_object('deleted', true, 'sku', v_registration.sku);
end;
$$;

revoke all on function public.erp_delete_catalog_sku(bigint) from public;
revoke all on function public.erp_delete_catalog_sku(bigint) from anon;
revoke all on function public.erp_delete_catalog_sku(bigint) from authenticated;
grant execute on function public.erp_delete_catalog_sku(bigint) to service_role;
