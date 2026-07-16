-- Define como "Com B.O.M." somente os cadastros legados de VEICULO P.B.
-- Valores ja definidos explicitamente como true ou false sao preservados.

update public.cadastro_registros
set form_values = jsonb_set(
    coalesce(form_values, '{}'::jsonb),
    '{possui_bom}',
    'true'::jsonb,
    true
)
where category_key = 'cat_34_veiculo_p_b'
  and not (coalesce(form_values, '{}'::jsonb) ? 'possui_bom');

select
    count(*) as total_veiculo_pb,
    count(*) filter (
        where coalesce(form_values ->> 'possui_bom', '') = 'true'
    ) as com_bom,
    count(*) filter (
        where not (coalesce(form_values, '{}'::jsonb) ? 'possui_bom')
    ) as nao_definidos
from public.cadastro_registros
where category_key = 'cat_34_veiculo_p_b';
