-- Reverte somente a normalizacao v3 usando o snapshot salvo em cada registro.

begin;

with restored as (
    update public.cadastro_registros r
    set
        descricao_primaria = r.form_values
            #>> '{_normalizacao_descricao_primaria_conjuntos_le_v3_original,descricao_primaria}',
        descricao_secundaria = r.form_values
            #>> '{_normalizacao_descricao_primaria_conjuntos_le_v3_original,descricao_secundaria}',
        search_text = r.form_values
            #>> '{_normalizacao_descricao_primaria_conjuntos_le_v3_original,search_text}',
        caracteres_primario = (
            r.form_values
                #>> '{_normalizacao_descricao_primaria_conjuntos_le_v3_original,caracteres_primario}'
        )::integer,
        caracteres_secundario = (
            r.form_values
                #>> '{_normalizacao_descricao_primaria_conjuntos_le_v3_original,caracteres_secundario}'
        )::integer,
        form_values = r.form_values - '_normalizacao_descricao_primaria_conjuntos_le_v3_original',
        updated_at = clock_timestamp()
    where r.sku in (
        '30200025', '30200032', '30200033', '30200034', '30200035',
        '30200036', '30200037', '30200038', '30200039', '30200048'
    )
      and r.form_values ? '_normalizacao_descricao_primaria_conjuntos_le_v3_original'
    returning r.sku
)
select count(*) as registros_restaurados,
       array_agg(sku order by sku) as skus_restaurados
from restored;

commit;

