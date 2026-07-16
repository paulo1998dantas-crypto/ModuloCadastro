-- Marca como "Com B.O.M." todo cadastro que ja possui estrutura no Supabase.
-- Os demais campos de form_values e todas as composicoes sao preservados.

update public.cadastro_registros as registro
set form_values = jsonb_set(
    coalesce(registro.form_values, '{}'::jsonb),
    '{possui_bom}',
    'true'::jsonb,
    true
)
where exists (
    select 1
    from public.cadastro_bom_cabecalhos as bom
    where bom.parent_sku = registro.sku
)
and coalesce(registro.form_values ->> 'possui_bom', '') <> 'true';

select
    count(*) as estruturas_sem_preferencia
from public.cadastro_bom_cabecalhos as bom
join public.cadastro_registros as registro
    on registro.sku = bom.parent_sku
where coalesce(registro.form_values ->> 'possui_bom', '') <> 'true';
