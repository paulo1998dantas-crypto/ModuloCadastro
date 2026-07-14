-- Sincroniza a unidade de medida dos componentes da B.O.M. com a unidade oficial do cadastro.
-- Seguro para banco compartilhado: altera somente public.cadastro_bom_componentes.

update public.cadastro_bom_componentes as c
set
    unidade = lower(trim(r.unidade)),
    search_text = trim(concat_ws(
        ' ',
        c.parent_sku,
        h.parent_sku,
        h.parent_descricao,
        c.component_sku,
        c.component_descricao,
        lower(trim(r.unidade)),
        h.source
    ))
from public.cadastro_registros as r,
     public.cadastro_bom_cabecalhos as h
where h.id = c.bom_id
  and r.sku = c.component_sku
  and trim(coalesce(r.unidade, '')) <> ''
  and coalesce(c.unidade, '') is distinct from lower(trim(r.unidade));

select
    c.component_sku,
    c.component_descricao,
    c.unidade as unidade_bom,
    r.unidade as unidade_cadastro
from public.cadastro_bom_componentes as c
join public.cadastro_registros as r on r.sku = c.component_sku
where trim(coalesce(r.unidade, '')) <> ''
  and coalesce(c.unidade, '') is distinct from lower(trim(r.unidade))
order by c.component_sku
limit 100;
