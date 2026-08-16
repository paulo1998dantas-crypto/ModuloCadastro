-- Reconcilia a marca form_values.possui_bom com a fonte de verdade:
-- cadastro_bom_cabecalhos.parent_sku.
--
-- Seguro para reexecução: só atualiza linhas cujo valor armazenado difere do
-- valor derivado. Não altera cabeçalhos/componentes de B.O.M., estoque,
-- pedidos, ordens de serviço ou cadastros além da própria marca B.O.M.

begin;

with bom_parent_skus as (
    select distinct btrim(parent_sku) as sku
    from public.cadastro_bom_cabecalhos
    where btrim(parent_sku) <> ''
),
desired as (
    select
        registro.id,
        exists (
            select 1
            from bom_parent_skus pai
            where pai.sku = btrim(registro.sku)
        ) as possui_bom
    from public.cadastro_registros registro
),
updated as (
    update public.cadastro_registros registro
       set form_values = jsonb_set(
               coalesce(registro.form_values, '{}'::jsonb),
               '{possui_bom}',
               to_jsonb(desired.possui_bom),
               true
           ),
           updated_at = now()
      from desired
     where registro.id = desired.id
       and (registro.form_values ->> 'possui_bom') is distinct from desired.possui_bom::text
    returning desired.possui_bom
)
select
    count(*) as cadastros_atualizados,
    count(*) filter (where possui_bom) as marcados_sim,
    count(*) filter (where not possui_bom) as marcados_nao
from updated;

-- A reconciliação só é válida se não restar divergência entre o indicador e
-- os itens-pai efetivos. Cabeçalhos sem cadastro são preservados como legado.
with bom_parent_skus as (
    select distinct btrim(parent_sku) as sku
    from public.cadastro_bom_cabecalhos
    where btrim(parent_sku) <> ''
)
select
    count(*) filter (
        where coalesce((form_values ->> 'possui_bom')::boolean, false)
          is distinct from (btrim(sku) in (select sku from bom_parent_skus))
    ) as divergencias_restantes,
    count(*) filter (where form_values ->> 'possui_bom' is null) as sem_definicao_restantes
from public.cadastro_registros;

commit;
