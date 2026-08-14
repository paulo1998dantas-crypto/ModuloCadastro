-- Rollback da normalizacao controlada dos conjuntos de bancos LE (v2).
-- Restaura somente registros que ainda possuem o snapshot criado pela
-- normalizacao. Nao afeta cadastros unitarios, outras categorias, saldos ou B.O.M.

begin;

with restaurados as (
    update public.cadastro_registros r
    set
        descricao_primaria = r.form_values #>> '{_normalizacao_conjunto_bancos_le_v2_original,descricao_primaria}',
        descricao_secundaria = r.form_values #>> '{_normalizacao_conjunto_bancos_le_v2_original,descricao_secundaria}',
        caracteres_primario = length(
            r.form_values #>> '{_normalizacao_conjunto_bancos_le_v2_original,descricao_primaria}'
        ),
        caracteres_secundario = length(
            r.form_values #>> '{_normalizacao_conjunto_bancos_le_v2_original,descricao_secundaria}'
        ),
        field_values = r.form_values #> '{_normalizacao_conjunto_bancos_le_v2_original,field_values}',
        field_codes = r.form_values #> '{_normalizacao_conjunto_bancos_le_v2_original,field_codes}',
        form_values = (
            r.form_values #> '{_normalizacao_conjunto_bancos_le_v2_original,form_values}'
        ),
        search_text = r.form_values #>> '{_normalizacao_conjunto_bancos_le_v2_original,search_text}',
        updated_at = now()
    where r.category_key = 'bancos'
      and r.sku in (
          '30200025','30200032','30200033','30200034','30200035',
          '30200036','30200037','30200038','30200039','30200048'
      )
      and coalesce(r.form_values, '{}'::jsonb) ? '_normalizacao_conjunto_bancos_le_v2_original'
    returning r.sku
)
select count(*) as registros_restaurados, array_agg(sku order by sku) as skus
from restaurados;

commit;
