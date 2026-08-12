-- Unifica a categoria legada "20 - CJ-BCO" na categoria canônica
-- "20 - BANCOS", preservando os SKUs 3020, saldos, movimentos e componentes.
--
-- Escopo deliberadamente restrito:
--   * atualiza somente metadados cadastrais e descrições dos conjuntos;
--   * não altera quantidades, saldo, estoque mínimo, movimentos ou itens de B.O.M.;
--   * mantém uma cópia dos metadados originais em form_values para rollback.
--
-- Execute somente depois de publicar a versão do ModuloCadastro que conhece
-- o grupo 30 (CONJUNTO / KIT) dentro da categoria "bancos".

begin;

with origem as (
    select
        r.id,
        r.sku,
        r.category_key as category_key_original,
        r.category_label as category_label_original,
        r.sheet as sheet_original,
        r.descricao_primaria as primaria_original,
        r.descricao_secundaria as secundaria_original,
        r.sufixo as sufixo_original,
        r.form_values as form_values_original,
        r.field_values as field_values_original,
        nullif(trim(r.field_values ->> 'descritor_base'), '') as descritor_base,
        nullif(trim(r.field_values ->> 'complemento_regra'), '') as complemento_regra,
        nullif(trim(r.field_values ->> 'posicao_lado'), '') as posicao_lado
    from public.cadastro_registros r
    where r.category_key = 'cat_20_bco'
), montado as (
    select
        origem.*,
        case
            -- Os primeiros registros tiveram a descrição duplicada no campo
            -- DESCRITOR BASE. A regra de complemento é a fonte correta.
            when sku in ('30200001', '30200002', '30200029')
                 and complemento_regra is not null then
                'CJ ' ||
                case
                    when descritor_base ilike '%BANCOS FIXOS%' then 'BANCOS FIXOS'
                    when descritor_base ilike '%BANCOS REC%' then 'BANCOS REC'
                    else regexp_replace(coalesce(descritor_base, primaria_original), '^CJ[[:space:]]+', '', 'i')
                end || ' - ' || complemento_regra
            else primaria_original
        end as primaria_base
    from origem
), normalizado as (
    select
        montado.*,
        trim(
            regexp_replace(
                regexp_replace(
                    regexp_replace(
                        regexp_replace(primaria_base, ';', ',', 'g'),
                        '([[:alpha:]])[[:space:]]*-[[:space:]]*', '\1 - ', 'g'
                    ),
                    'E/S/[[:space:]]*J', 'E/S/ J', 'g'
                ),
                'DIAMENTE', 'DIAMANTE', 'gi'
            ),
            '[[:space:]]+', ' ', 'g'
        ) as primaria_nova
    from montado
), preparado as (
    select
        normalizado.*,
        coalesce((regexp_match(primaria_nova, '(?:^| - )(MC|CS|MD|STF|INC|JL|ORI)(?: - |$)'))[1], '') as fornecedor,
        coalesce((regexp_match(primaria_nova, '(?:^| - )(LB|LE|LS|LL|ORI)(?: - |$)'))[1], '') as linha,
        coalesce((regexp_match(primaria_nova, ' - ([0-9ED][0-9ED,;-]*) - (?:2P|3P)'))[1], '') as layout,
        coalesce((regexp_match(primaria_nova, ' - (2P|3P) - '))[1], '') as tipo_cinto,
        coalesce((regexp_match(primaria_nova, ' - (TECIDO|COURVIN|MISTO)(?:[[:space:]]| - )'))[1], '') as tipo_revestimento,
        coalesce((regexp_match(primaria_nova, ' - COURVIN ([A-Z0-9/[:space:]]+) - (?:E/S|ESJ|TRILHO|NORMAL|BJD|PME|EXECUTIVO|FOCA|ELEVITTA|MASTER)'))[1], '') as detalhe_revestimento,
        nullif(trim(btrim(replace(secundaria_original, primaria_original, ''), ' |')), '') as observacao_legada
    from normalizado
), atualizado as (
    update public.cadastro_registros r
       set category_key = 'bancos',
           category_label = '20 - BANCOS',
           sheet = '20 - BANCOS',
           descricao_primaria = p.primaria_nova,
           -- A descrição secundária sempre começa pela primária. Informações
           -- históricas que não repetem a composição técnica são preservadas.
           descricao_secundaria = p.primaria_nova || case
               when p.observacao_legada is not null
                    and p.observacao_legada !~ '(MC|CS|MD|STF|INC|JL|ORI)[[:space:]]*-[[:space:]]*(LB|LE|LS|LL|ORI)'
                   then ' | ' || p.observacao_legada
               else ''
           end,
           sufixo = 'CJ',
           caracteres_primario = char_length(p.primaria_nova),
           caracteres_secundario = char_length(
               p.primaria_nova || case
                   when p.observacao_legada is not null
                        and p.observacao_legada !~ '(MC|CS|MD|STF|INC|JL|ORI)[[:space:]]*-[[:space:]]*(LB|LE|LS|LL|ORI)'
                       then ' | ' || p.observacao_legada
                   else ''
               end
           ),
           form_values = coalesce(r.form_values, '{}'::jsonb)
               || jsonb_build_object(
                    'migracao_20_bancos_original', jsonb_build_object(
                        'category_key', p.category_key_original,
                        'category_label', p.category_label_original,
                        'sheet', p.sheet_original,
                        'descricao_primaria', p.primaria_original,
                        'descricao_secundaria', p.secundaria_original,
                        'sufixo', p.sufixo_original
                    ),
                    'grupo_codigo', jsonb_build_array('30'),
                    'cj_sufixo', jsonb_build_array('CJ'),
                    'cj_encosto', jsonb_build_array(case when p.primaria_nova ilike 'CJ BANCOS FIXOS%' then 'FIXO' else 'RECLINAVEL' end),
                    'cj_fornecedor', jsonb_build_array(p.fornecedor),
                    'cj_linha', jsonb_build_array(p.linha),
                    'cj_layout', jsonb_build_array(p.layout),
                    'cj_tipo_cinto', jsonb_build_array(p.tipo_cinto),
                    'cj_tipo_revestimento', jsonb_build_array(p.tipo_revestimento),
                    'cj_detalhe_revestimento', jsonb_build_array(p.detalhe_revestimento),
                    'cj_especificidade', to_jsonb(array_remove(array[
                        case when p.primaria_nova ~* '(E/S/[[:space:]]*J|ESJ)' then 'E/S/ J' end,
                        case when p.primaria_nova ~* 'TRILHO' then 'TRILHO' end,
                        case when p.primaria_nova ~* 'BJD' then 'BJD' end,
                        case when p.primaria_nova ~* 'NORMAL' then 'NORMAL' end,
                        case when p.primaria_nova ~* 'PME[[:space:]]*1A' then 'PME 1A' end,
                        case when p.primaria_nova ~* 'PME[[:space:]]*2A' then 'PME 2A' end,
                        case when p.primaria_nova ~* 'PME[[:space:]]*3A' then 'PME 3A' end,
                        case when p.primaria_nova ~* 'EXECUTIVO' then 'EXECUTIVO' end,
                        case when p.primaria_nova ~* '4L[[:space:]]*REC' then '4L REC' end
                    ], null)),
                    'cj_acessibilidade', jsonb_build_array(case
                        when p.primaria_nova ~* 'PLATAFORMA BI[ -]?PARTIDA' then 'PLATAFORMA BI-PARTIDA'
                        when p.primaria_nova ~* 'PLATAFORMA FECHADA' then 'PLATAFORMA FECHADA'
                        when p.primaria_nova ~* 'ELEVITTA' then 'ELEVITTA'
                        when p.primaria_nova ~* 'FOCA' then 'FOCA'
                        else '' end),
                    'cj_acessibilidade_secundaria', '[]'::jsonb,
                    'cj_observacao', jsonb_build_array(coalesce(p.observacao_legada, ''))
                ),
           field_values = coalesce(r.field_values, '{}'::jsonb)
               || jsonb_build_object(
                    'cj_sufixo', 'CJ',
                    'cj_encosto', case when p.primaria_nova ilike 'CJ BANCOS FIXOS%' then 'FIXO' else 'RECLINAVEL' end,
                    'cj_fornecedor', p.fornecedor,
                    'cj_linha', p.linha,
                    'cj_layout', p.layout,
                    'cj_tipo_cinto', p.tipo_cinto,
                    'cj_tipo_revestimento', p.tipo_revestimento,
                    'cj_detalhe_revestimento', p.detalhe_revestimento
                ),
           search_text = upper(concat_ws(' ', r.sku, p.primaria_nova, r.descricao_secundaria, '20 BANCOS CONJUNTO KIT CJ')),
           updated_at = now()
      from preparado p
     where r.id = p.id
 returning r.sku, r.descricao_primaria
)
update public.skus s
   set categoria = '20 - BANCOS',
       grupo = '30 - CONJUNTO / KIT',
       descricao = a.descricao_primaria,
       updated_at = now()
  from atualizado a
 where s.sku = a.sku;

-- Atualiza somente os metadados dos cabeçalhos de B.O.M.; os componentes e
-- suas quantidades permanecem inalterados.
update public.cadastro_bom_cabecalhos h
   set parent_category_key = 'bancos',
       parent_category_label = '20 - BANCOS',
       parent_descricao = r.descricao_primaria,
       search_text = upper(concat_ws(' ', h.parent_sku, r.descricao_primaria, '20 BANCOS CONJUNTO KIT CJ')),
       updated_at = now()
  from public.cadastro_registros r
 where h.parent_sku = r.sku
   and r.category_key = 'bancos'
   and r.form_values ? 'migracao_20_bancos_original';

-- A categoria antiga deixa de ser apresentada no catálogo publicado. Os
-- campos específicos do conjunto são fornecidos pela versão do aplicativo.
update public.cadastro_catalogo c
   set payload = jsonb_set(
           c.payload,
           '{categories}',
           coalesce((
               select jsonb_agg(item order by ordinalidade)
               from jsonb_array_elements(c.payload -> 'categories') with ordinality as x(item, ordinalidade)
               where item ->> 'key' <> 'cat_20_bco'
           ), '[]'::jsonb)
       ),
       updated_at = now()
 where c.payload @> '{"categories": [{"key": "cat_20_bco"}]}'::jsonb;

commit;

-- Verificação pós-migração (somente leitura):
-- select category_key, category_label, count(*)
--   from public.cadastro_registros
--  where sku like '3020%'
--  group by category_key, category_label;
--
-- Rollback lógico (usar somente após avaliação):
-- update public.cadastro_registros
--    set category_key = form_values #>> '{migracao_20_bancos_original,category_key}',
--        category_label = form_values #>> '{migracao_20_bancos_original,category_label}',
--        sheet = form_values #>> '{migracao_20_bancos_original,sheet}',
--        descricao_primaria = form_values #>> '{migracao_20_bancos_original,descricao_primaria}',
--        descricao_secundaria = form_values #>> '{migracao_20_bancos_original,descricao_secundaria}',
--        sufixo = form_values #>> '{migracao_20_bancos_original,sufixo}'
--  where form_values ? 'migracao_20_bancos_original';
