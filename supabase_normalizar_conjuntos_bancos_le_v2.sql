-- Normalizacao controlada dos conjuntos de bancos LE informados em 14/08/2026.
-- Escopo deliberadamente restrito a 10 SKUs da categoria bancos/grupo 30.
-- Nao altera cadastros unitarios 1020, outras categorias, saldos ou B.O.M.

begin;

with alvo (
    sku,
    descricao_primaria,
    descricao_secundaria,
    especificidade,
    especificidade_texto,
    especificidade_codigos,
    cor_revestimento,
    cor_revestimento_texto,
    cor_revestimento_codigo,
    tipo_costura,
    tipo_costura_texto,
    tipo_costura_codigo,
    cor_linha,
    cor_linha_texto,
    cor_linha_codigo
) as (
    values
        (
            '30200025',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO | REVESTIMENTO: MARROM/DIAMANTE/LINHA DOURADA',
            '["11- E/S/J", "EXECUTIVO"]'::jsonb, '11- E/S/J | EXECUTIVO', '11',
            '["3- CAPA LE MARROM"]'::jsonb, '3- CAPA LE MARROM', '3',
            '["3- CS COSTURA DIAMANTE"]'::jsonb, '3- CS COSTURA DIAMANTE', '3',
            '["6- COR LINHA DOURADO"]'::jsonb, '6- COR LINHA DOURADO', '6'
        ),
        (
            '30200032',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO | REVESTIMENTO: MARROM/BOOMERANG/LINHA DOURADA',
            '["11- E/S/J", "EXECUTIVO"]'::jsonb, '11- E/S/J | EXECUTIVO', '11',
            '["3- CAPA LE MARROM"]'::jsonb, '3- CAPA LE MARROM', '3',
            '["4- CS COSTURA BOOMERANG"]'::jsonb, '4- CS COSTURA BOOMERANG', '4',
            '["6- COR LINHA DOURADO"]'::jsonb, '6- COR LINHA DOURADO', '6'
        ),
        (
            '30200033',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO | REVESTIMENTO: PRETO/BOOMERANG/LINHA PRETA',
            '["11- E/S/J", "EXECUTIVO"]'::jsonb, '11- E/S/J | EXECUTIVO', '11',
            '["4- CAPA LE PRETA"]'::jsonb, '4- CAPA LE PRETA', '4',
            '["4- CS COSTURA BOOMERANG"]'::jsonb, '4- CS COSTURA BOOMERANG', '4',
            '["4- COR LINHA PRETA"]'::jsonb, '4- COR LINHA PRETA', '4'
        ),
        (
            '30200034',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO | REVESTIMENTO: PRETO/RETILINEA/LINHA PRETA',
            '["11- E/S/J", "EXECUTIVO"]'::jsonb, '11- E/S/J | EXECUTIVO', '11',
            '["4- CAPA LE PRETA"]'::jsonb, '4- CAPA LE PRETA', '4',
            '["5- CS COSTURA RETILINEA"]'::jsonb, '5- CS COSTURA RETILINEA', '5',
            '["4- COR LINHA PRETA"]'::jsonb, '4- COR LINHA PRETA', '4'
        ),
        (
            '30200035',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO | REVESTIMENTO: PRETO/DIAMANTE/LINHA BRANCA',
            '["11- E/S/J", "EXECUTIVO"]'::jsonb, '11- E/S/J | EXECUTIVO', '11',
            '["4- CAPA LE PRETA"]'::jsonb, '4- CAPA LE PRETA', '4',
            '["3- CS COSTURA DIAMANTE"]'::jsonb, '3- CS COSTURA DIAMANTE', '3',
            '["3- COR LINHA BRANCA"]'::jsonb, '3- COR LINHA BRANCA', '3'
        ),
        (
            '30200036',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO | REVESTIMENTO: PRETO/CINZA/DIAMANTE/LINHA CINZA',
            '["11- E/S/J", "EXECUTIVO"]'::jsonb, '11- E/S/J | EXECUTIVO', '11',
            '["5- CAPA LE PRETA/CINZA"]'::jsonb, '5- CAPA LE PRETA/CINZA', '5',
            '["3- CS COSTURA DIAMANTE"]'::jsonb, '3- CS COSTURA DIAMANTE', '3',
            '["7- COR LINHA CINZA"]'::jsonb, '7- COR LINHA CINZA', '7'
        ),
        (
            '30200037',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO',
            'CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN - E/S/ J - EXECUTIVO | REVESTIMENTO: PRETO/DIAMANTE/LINHA PRETA',
            '["11- E/S/J", "EXECUTIVO"]'::jsonb, '11- E/S/J | EXECUTIVO', '11',
            '["4- CAPA LE PRETA"]'::jsonb, '4- CAPA LE PRETA', '4',
            '["3- CS COSTURA DIAMANTE"]'::jsonb, '3- CS COSTURA DIAMANTE', '3',
            '["4- COR LINHA PRETA"]'::jsonb, '4- COR LINHA PRETA', '4'
        ),
        (
            '30200038',
            'CJ BANCOS REC - CS - LE - 3,2 - 3P - COURVIN - E/S/ J - EXECUTIVO',
            'CJ BANCOS REC - CS - LE - 3,2 - 3P - COURVIN - E/S/ J - EXECUTIVO | REVESTIMENTO: PRETO/DIAMANTE/LINHA PRETA',
            '["11- E/S/J", "EXECUTIVO"]'::jsonb, '11- E/S/J | EXECUTIVO', '11',
            '["4- CAPA LE PRETA"]'::jsonb, '4- CAPA LE PRETA', '4',
            '["3- CS COSTURA DIAMANTE"]'::jsonb, '3- CS COSTURA DIAMANTE', '3',
            '["4- COR LINHA PRETA"]'::jsonb, '4- COR LINHA PRETA', '4'
        ),
        (
            '30200039',
            'CJ BANCOS REC - CS - LE - 3,2-1 - 3P - COURVIN - E/S/ J - EXECUTIVO',
            'CJ BANCOS REC - CS - LE - 3,2-1 - 3P - COURVIN - E/S/ J - EXECUTIVO | REVESTIMENTO: PRETO/BOOMERANG/LINHA PRETA',
            '["11- E/S/J", "EXECUTIVO"]'::jsonb, '11- E/S/J | EXECUTIVO', '11',
            '["4- CAPA LE PRETA"]'::jsonb, '4- CAPA LE PRETA', '4',
            '["4- CS COSTURA BOOMERANG"]'::jsonb, '4- CS COSTURA BOOMERANG', '4',
            '["4- COR LINHA PRETA"]'::jsonb, '4- COR LINHA PRETA', '4'
        ),
        (
            '30200048',
            'CJ BANCOS REC - STF - LE - 4,2-1,2-1,3 - 3P - COURVIN - EXECUTIVO',
            'CJ BANCOS REC - STF - LE - 4,2-1,2-1,3 - 3P - COURVIN - EXECUTIVO | REVESTIMENTO: PRETO/ST02/LINHA BRANCA | ESPECIFICIDADE LEGADA: MASTER - PME',
            '["EXECUTIVO", "MASTER - PME"]'::jsonb, 'EXECUTIVO | MASTER - PME', '',
            '["4- CAPA LE PRETA"]'::jsonb, '4- CAPA LE PRETA', '4',
            '["7- COSTURA ST02 STF"]'::jsonb, '7- COSTURA ST02 STF', '7',
            '["3- COR LINHA BRANCA"]'::jsonb, '3- COR LINHA BRANCA', '3'
        )
),
registros as (
    select r.id as registro_id, a.*
    from public.cadastro_registros r
    join alvo a on a.sku = r.sku
    where r.category_key = 'bancos'
      and coalesce(r.form_values->'grupo_codigo', '[]'::jsonb) ? '30'
),
atualizados as (
    update public.cadastro_registros r
    set
        descricao_primaria = x.descricao_primaria,
        descricao_secundaria = x.descricao_secundaria,
        caracteres_primario = length(x.descricao_primaria),
        caracteres_secundario = length(x.descricao_secundaria),
        form_values =
            (coalesce(r.form_values, '{}'::jsonb) - 'cj_especificidade' - 'cj_observacao')
            || jsonb_build_object(
                'especificidade', x.especificidade,
                'tipo_revestimento', '["2- COURVIN"]'::jsonb,
                'cor_do_revestimento', x.cor_revestimento,
                'tipo_costura', x.tipo_costura,
                'cor_da_linha', x.cor_linha
            )
            || case
                when coalesce(r.form_values, '{}'::jsonb) ? '_normalizacao_conjunto_bancos_le_v2_original'
                    then '{}'::jsonb
                else jsonb_build_object(
                    '_normalizacao_conjunto_bancos_le_v2_original',
                    jsonb_build_object(
                        'descricao_primaria', r.descricao_primaria,
                        'descricao_secundaria', r.descricao_secundaria,
                        'form_values', r.form_values,
                        'field_values', r.field_values,
                        'field_codes', r.field_codes,
                        'search_text', r.search_text,
                        'capturado_em', now()
                    )
                )
                end,
        field_values =
            (coalesce(r.field_values, '{}'::jsonb) - 'cj_especificidade' - 'cj_observacao')
            || jsonb_build_object(
                'especificidade', x.especificidade_texto,
                'tipo_revestimento', '2- COURVIN',
                'cor_do_revestimento', x.cor_revestimento_texto,
                'tipo_costura', x.tipo_costura_texto,
                'cor_da_linha', x.cor_linha_texto
            ),
        field_codes =
            (coalesce(r.field_codes, '{}'::jsonb) - 'cj_especificidade' - 'cj_observacao')
            || jsonb_build_object(
                'especificidade', x.especificidade_codigos,
                'tipo_revestimento', '2',
                'cor_do_revestimento', x.cor_revestimento_codigo,
                'tipo_costura', x.tipo_costura_codigo,
                'cor_da_linha', x.cor_linha_codigo
            ),
        search_text = upper(concat_ws(
            ' ', r.sku, r.category_label, x.descricao_primaria, x.descricao_secundaria,
            r.unidade, x.especificidade_texto, x.cor_revestimento_texto,
            x.tipo_costura_texto, x.cor_linha_texto
        )),
        updated_at = now()
    from registros x
    where r.id = x.registro_id
    returning r.sku
)
select count(*) as registros_atualizados, array_agg(sku order by sku) as skus
from atualizados;

commit;

-- Verificacao esperada: 10 registros, sem cj_especificidade/cj_observacao e
-- com especificidade, revestimento, costura e cor em suas chaves canonicas.
select
    sku,
    descricao_primaria,
    descricao_secundaria,
    form_values->'especificidade' as especificidade,
    form_values->'cor_do_revestimento' as cor_do_revestimento,
    form_values->'tipo_costura' as tipo_costura,
    form_values->'cor_da_linha' as cor_da_linha,
    form_values ? 'cj_especificidade' as possui_cj_especificidade,
    form_values ? 'cj_observacao' as possui_cj_observacao
from public.cadastro_registros
where category_key = 'bancos'
  and sku in (
      '30200025','30200032','30200033','30200034','30200035',
      '30200036','30200037','30200038','30200039','30200048'
  )
order by sku;
