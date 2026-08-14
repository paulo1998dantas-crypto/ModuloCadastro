-- Normaliza exclusivamente os conjuntos de bancos LE executivos informados.
-- A descricao abaixo e a saida da mesma regra usada pelo cadastro normal.
-- O estado anterior de cada linha fica preservado no proprio registro para rollback.

begin;

with targets(sku, descricao) as (
    values
        ('30200025', 'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN MARROM/DIAMANTE/LINHA DOURADA - E/S/ J - EXECUTIVO'),
        ('30200032', 'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN MARROM/BOOMERANG/LINHA DOURADA - E/S/ J - EXECUTIVO'),
        ('30200033', 'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN PRETO/BOOMERANG/LINHA PRETA - E/S/ J - EXECUTIVO'),
        ('30200034', 'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN PRETO/RETILINEA/LINHA PRETA - E/S/ J - EXECUTIVO'),
        ('30200035', 'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN PRETO/DIAMANTE/LINHA BRANCA - E/S/ J - EXECUTIVO'),
        ('30200036', 'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN PRETO/CINZA/DIAMANTE/LINHA CINZA - E/S/ J - EXECUTIVO'),
        ('30200037', 'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN PRETO/DIAMANTE/LINHA PRETA - E/S/ J - EXECUTIVO'),
        ('30200038', 'CJ BANCOS REC - CS - LE - 3,2 - 3P - COURVIN PRETO/DIAMANTE/LINHA PRETA - E/S/ J - EXECUTIVO'),
        ('30200039', 'CJ BANCOS REC - CS - LE - 3,2-1 - 3P - COURVIN PRETO/BOOMERANG/LINHA PRETA - E/S/ J - EXECUTIVO'),
        ('30200048', 'CJ BANCOS REC - STF - LE - 4,2-1,2-1,3 - 3P - COURVIN PRETO/ST02/LINHA BRANCA - MASTER - PME - EXECUTIVO')
),
eligible as (
    select r.id, r.sku, t.descricao
    from public.cadastro_registros r
    join targets t on t.sku = r.sku
    where r.category_key = 'bancos'
      and r.ativo is true
      and r.form_values->'grupo_codigo' @> '["30"]'::jsonb
      and r.form_values->'linha' @> '["2- LE"]'::jsonb
      and r.form_values->'especificidade' @> '["EXECUTIVO"]'::jsonb
),
updated as (
    update public.cadastro_registros r
    set
        form_values = jsonb_set(
            r.form_values,
            '{_normalizacao_descricao_primaria_conjuntos_le_v3_original}',
            coalesce(
                r.form_values->'_normalizacao_descricao_primaria_conjuntos_le_v3_original',
                jsonb_build_object(
                    'descricao_primaria', r.descricao_primaria,
                    'descricao_secundaria', r.descricao_secundaria,
                    'search_text', r.search_text,
                    'caracteres_primario', r.caracteres_primario,
                    'caracteres_secundario', r.caracteres_secundario,
                    'capturado_em', clock_timestamp()
                )
            ),
            true
        ),
        descricao_primaria = e.descricao,
        descricao_secundaria = e.descricao,
        caracteres_primario = length(e.descricao),
        caracteres_secundario = length(e.descricao),
        search_text = concat_ws(
            ' ', r.sku, r.category_label, e.descricao,
            r.field_values->>'fornecedor', r.field_values->>'linha'
        ),
        updated_at = clock_timestamp()
    from eligible e
    where r.id = e.id
    returning r.sku
)
select count(*) as registros_atualizados,
       array_agg(sku order by sku) as skus_atualizados
from updated;

commit;

