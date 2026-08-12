-- Normalizacao aditiva dos conjuntos (3020) da categoria 20 - BANCOS.
--
-- Escopo deliberadamente limitado: nao toca 1020 (bancos unitarios), B.O.M.,
-- estoque, movimentos ou historicos. O detalhe legado de revestimento deixa
-- de compor a descricao primaria; quando necessario e preservado em
-- OBSERVACAO TECNICA.

WITH correcao_manual (
    sku,
    fornecedor,
    linha,
    layout,
    tipo_revestimento,
    especificidade,
    acessibilidade,
    observacao
) AS (
    VALUES
        ('30200014', NULL, NULL, '1E,1E,1E,1D', NULL, NULL, NULL, NULL),
        ('30200016', NULL, NULL, '4,2,2,3,3,3', NULL, NULL, NULL, NULL),
        ('30200017', NULL, NULL, '4,2,3,3,3', NULL, NULL, NULL, NULL),
        ('30200018', NULL, NULL, '1E,1E,1D,1D,3,3', NULL, NULL, NULL, NULL),
        ('30200020', NULL, NULL, '1E,1E,1E,1D', NULL, NULL, NULL, NULL),
        ('30200025', NULL, 'LE', NULL, NULL, NULL, NULL, 'REVESTIMENTO: MARROM/DIAMANTE/LINHA DOURADA'),
        ('30200027', NULL, NULL, NULL, 'TECIDO', NULL, NULL, NULL),
        ('30200028', NULL, NULL, NULL, 'TECIDO', NULL, NULL, NULL),
        ('30200032', NULL, 'LE', NULL, NULL, NULL, NULL, 'REVESTIMENTO: MARROM/BOOMERANG/LINHA DOURADA'),
        ('30200033', NULL, 'LE', NULL, NULL, NULL, NULL, 'REVESTIMENTO: PRETO/BOOMERANG/LINHA PRETA'),
        ('30200034', NULL, 'LE', NULL, NULL, NULL, NULL, 'REVESTIMENTO: PRETO/RETILINEA/LINHA PRETA'),
        ('30200035', NULL, 'LE', NULL, NULL, NULL, NULL, 'REVESTIMENTO: PRETO/DIAMANTE/LINHA BRANCA'),
        ('30200036', NULL, 'LE', NULL, NULL, NULL, NULL, 'REVESTIMENTO: PRETO/CINZA/DIAMANTE/LINHA CINZA'),
        ('30200037', NULL, 'LE', NULL, NULL, NULL, NULL, 'REVESTIMENTO: PRETO/DIAMANTE/LINHA PRETA'),
        ('30200038', NULL, 'LE', NULL, NULL, NULL, NULL, 'REVESTIMENTO: PRETO/DIAMANTE/LINHA PRETA'),
        ('30200039', NULL, 'LE', NULL, NULL, NULL, NULL, 'REVESTIMENTO: PRETO/BOOMERANG/LINHA PRETA'),
        ('30200041', NULL, NULL, NULL, NULL, ARRAY['TRILHO'], 'ELEVITTA', NULL),
        ('30200042', NULL, NULL, '4,2,3,3,3', NULL, NULL, NULL, NULL),
        ('30200047', 'MC', NULL, '4,3,3,3', NULL, ARRAY['BJD', '4L REC'], NULL, NULL),
        ('30200048', NULL, NULL, NULL, NULL, ARRAY['EXECUTIVO'], NULL, 'REVESTIMENTO: PRETO/ST02/LINHA BRANCA | ESPECIFICIDADE LEGADA: MASTER - PME'),
        ('30200050', NULL, NULL, NULL, NULL, ARRAY['E/S/ J'], 'FOCA', NULL)
), origem AS (
    SELECT
        r.id,
        r.sku,
        r.field_values,
        r.form_values,
        c.fornecedor,
        c.linha,
        c.layout,
        c.tipo_revestimento,
        c.especificidade,
        c.acessibilidade,
        c.observacao,
        (COALESCE(r.form_values, '{}'::jsonb) - 'cj_detalhe_revestimento') AS valores_sem_detalhe
    FROM public.cadastro_registros AS r
    LEFT JOIN correcao_manual AS c ON c.sku = r.sku
    WHERE r.category_key = 'bancos'
      AND r.sku ~ '^3020[0-9]+$'
), campos_normalizados AS (
    SELECT
        origem.*,
        CASE
            WHEN fornecedor IS NOT NULL THEN fornecedor
            WHEN upper(COALESCE(valores_sem_detalhe #>> '{cj_fornecedor,0}', '')) IN ('MC REC', 'MC RECLINAVEL') THEN 'MC'
            ELSE COALESCE(valores_sem_detalhe #>> '{cj_fornecedor,0}', '')
        END AS fornecedor_final,
        COALESCE(linha, valores_sem_detalhe #>> '{cj_linha,0}', '') AS linha_final,
        replace(COALESCE(layout, valores_sem_detalhe #>> '{cj_layout,0}', ''), ';', ',') AS layout_final,
        COALESCE(tipo_revestimento, valores_sem_detalhe #>> '{cj_tipo_revestimento,0}', '') AS revestimento_final,
        CASE WHEN especificidade IS NULL THEN COALESCE(valores_sem_detalhe -> 'cj_especificidade', '[]'::jsonb)
             ELSE to_jsonb(especificidade) END AS especificidade_final,
        COALESCE(acessibilidade, valores_sem_detalhe #>> '{cj_acessibilidade,0}', '') AS acessibilidade_final,
        CASE
            WHEN COALESCE(observacao, '') = '' THEN COALESCE(valores_sem_detalhe #>> '{cj_observacao,0}', '')
            WHEN COALESCE(valores_sem_detalhe #>> '{cj_observacao,0}', '') = '' THEN observacao
            WHEN position(observacao IN (valores_sem_detalhe #>> '{cj_observacao,0}')) > 0 THEN valores_sem_detalhe #>> '{cj_observacao,0}'
            ELSE (valores_sem_detalhe #>> '{cj_observacao,0}') || ' | ' || observacao
        END AS observacao_final
    FROM origem
), valores_reconstruidos AS (
    SELECT
        campos_normalizados.*,
        jsonb_set(
            jsonb_set(
                jsonb_set(
                    jsonb_set(
                        jsonb_set(
                            jsonb_set(
                                jsonb_set(
                                    jsonb_set(
                                        valores_sem_detalhe || jsonb_build_object(
                                            'cj_sufixo', jsonb_build_array('CJ'),
                                            'grupo_codigo', jsonb_build_array('30')
                                        ),
                                        '{cj_fornecedor}', jsonb_build_array(fornecedor_final), true
                                    ),
                                    '{cj_linha}', jsonb_build_array(linha_final), true
                                ),
                                '{cj_layout}', jsonb_build_array(layout_final), true
                            ),
                            '{cj_tipo_revestimento}', jsonb_build_array(revestimento_final), true
                        ),
                        '{cj_especificidade}', especificidade_final, true
                    ),
                    '{cj_acessibilidade}', jsonb_build_array(acessibilidade_final), true
                ),
                '{cj_observacao}', jsonb_build_array(observacao_final), true
            ),
            '{cj_sufixo}', jsonb_build_array('CJ'), true
        ) AS valores_finais
    FROM campos_normalizados
), descricoes AS (
    SELECT
        valores_reconstruidos.*,
        concat_ws(
            ' - ',
            'CJ BANCOS ' || CASE upper(COALESCE(valores_finais #>> '{cj_encosto,0}', ''))
                WHEN 'FIXO' THEN 'FIXOS'
                WHEN 'FIXOS' THEN 'FIXOS'
                WHEN 'RECLINAVEL' THEN 'REC'
                WHEN 'REC' THEN 'REC'
                ELSE COALESCE(valores_finais #>> '{cj_encosto,0}', '')
            END,
            NULLIF(valores_finais #>> '{cj_fornecedor,0}', ''),
            NULLIF(valores_finais #>> '{cj_linha,0}', ''),
            NULLIF(valores_finais #>> '{cj_layout,0}', ''),
            NULLIF(valores_finais #>> '{cj_tipo_cinto,0}', ''),
            NULLIF(valores_finais #>> '{cj_tipo_revestimento,0}', ''),
            CASE WHEN (valores_finais -> 'cj_especificidade') ? 'NORMAL' THEN 'NORMAL' END,
            CASE WHEN (valores_finais -> 'cj_especificidade') ? 'TRILHO' THEN 'TRILHO' END,
            CASE WHEN (valores_finais -> 'cj_especificidade') ? 'BJD' THEN 'BJD' END,
            CASE WHEN (valores_finais -> 'cj_especificidade') ? 'E/S/ J'
                    OR (valores_finais -> 'cj_especificidade') ? 'ESJ'
                    OR (valores_finais -> 'cj_especificidade') ? 'E/S/J' THEN 'E/S/ J' END,
            CASE WHEN (valores_finais -> 'cj_especificidade') ? 'PME 1A' THEN 'PME 1A' END,
            CASE WHEN (valores_finais -> 'cj_especificidade') ? 'PME 2A' THEN 'PME 2A' END,
            CASE WHEN (valores_finais -> 'cj_especificidade') ? 'PME 3A' THEN 'PME 3A' END,
            CASE WHEN (valores_finais -> 'cj_especificidade') ? 'EXECUTIVO' THEN 'EXECUTIVO' END,
            CASE WHEN (valores_finais -> 'cj_especificidade') ? '4L REC' THEN '4L REC' END,
            NULLIF(NULLIF(upper(valores_finais #>> '{cj_acessibilidade,0}'), 'N/A'), '')
        ) AS primaria
    FROM valores_reconstruidos
), atualizados AS (
    UPDATE public.cadastro_registros AS r
    SET
        form_values = d.valores_finais,
        field_values = (COALESCE(d.field_values, '{}'::jsonb) - 'cj_detalhe_revestimento') || jsonb_build_object(
            'cj_sufixo', 'CJ',
            'cj_fornecedor', d.fornecedor_final,
            'cj_linha', d.linha_final,
            'cj_layout', d.layout_final,
            'cj_tipo_revestimento', d.revestimento_final,
            'cj_especificidade', array_to_string(ARRAY(SELECT jsonb_array_elements_text(d.valores_finais -> 'cj_especificidade')), ' | '),
            'cj_acessibilidade', d.acessibilidade_final,
            'cj_observacao', d.observacao_final
        ),
        descricao_primaria = d.primaria,
        descricao_secundaria = CASE WHEN NULLIF(d.observacao_final, '') IS NULL THEN d.primaria ELSE d.primaria || ' ' || d.observacao_final END,
        sufixo = 'CJ'
    FROM descricoes AS d
    WHERE r.id = d.id
    RETURNING r.sku
)
SELECT count(*) AS conjuntos_normalizados FROM atualizados;
