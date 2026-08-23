-- Atualiza exclusivamente o catálogo da categoria 18 - REVESTIMENTO conforme
-- planilha de campos/opções. O campo legado fornecedor_referencia deixa de ser
-- exibido; dados históricos existentes não são apagados.

with revestimento_fields as (
  select $$[
    {"key":"prefixo","label":"ESTAGIO","scope":"primaria","selection_mode":"unitaria","description_order":1,"required":true,"free_text":false,"options":["1- CJ","2- PP","3- N/A"]},
    {"key":"descritor_base","label":"IDENTIFICAÇÃO","scope":"primaria","selection_mode":"unitaria","description_order":2,"required":true,"free_text":false,"options":["1- ACAB.","2- ALCA PQP","3- ALETA AR","4- ARO JAN","5- BATENTE","6- BOCAL","7- CAPUCINO","8- COIFA","9- COLUNA A","10- COLUNA B","11- COLUNA B SALAO","12- COLUNA C","13- COLUNA D","14- CONSOLE TV","16- CX. DE RODAS","17- DESEMBACADOR","18- DUTO AR COND.","19- FORRO","20- LATERAL","21- LENTE","22- LUMINARIA","23- MACANETA PORTA LATERAL","24- MOLDURA INTERNA","25- QUEBRA SOL","26- REFORCO","27- SOLEIRA","28- TAMPA INSPECAO","29- TETO","30- TETO C/ DUTO CENTRAL","31- TETO C/ DUTO LATERAL","32- TETO S/ DUTO"]},
    {"key":"veiculo_modelo","label":"NIVEL","scope":"primaria","selection_mode":"unitaria","description_order":3,"required":true,"free_text":false,"options":["1- INFERIOR","2- PARTE 1","3- PARTE 2","4- PARTE 2 / PARTE 4","5- PARTE 3","6- SUPERIOR"]},
    {"key":"posicao_lado","label":"LOCAL","scope":"secundaria","selection_mode":"unitaria","description_order":4,"required":true,"free_text":false,"options":["1- 1 VAO","2- 1/2 VAO","3- 2 VAO","4- 2/3 VAO","5- 3 VAO","6- AR COND.","7- ASSOALHO","8- ASSOALHO CABINE/SALAO","9- CABINE","10- CINTO DE SEGURANCA","11- COLUNA CORREDICA","12- COLUNA D","13- CX. AR COND.","14- DUTO AR COND.","16- MICROFONE SOM.","17- PORTA LATERAL","18- PORTA TRASIERA","19- QUEBRA SOL","20- TETO"]},
    {"key":"medida","label":"LADO","scope":"secundaria","selection_mode":"unitaria","description_order":5,"required":true,"free_text":false,"options":["1- LD","2- LE","3- LE/LD"]},
    {"key":"material_cor","label":"VEICULO","scope":"secundaria","selection_mode":"unitaria","description_order":6,"required":true,"free_text":false,"options":["1- B/J/D","2- E/S/J","3- IVECO","4- MASTER","5- MASTER L2H2","6- MASTER L2H2/L3H2","7- MASTER L3H2","8- SPRINTER","9- SPRINTER 10/14","10- SPRINTER 10/15","11- SPRINTER 10","12- SPRINTER 14/15","13- SPRINTER 14","14- SPRINTER 15","15- SPRINTER VITRE REV.","17- TRANSIT","18- TRANSIT L2H3/L3H3","19- TRANSIT L3H2 VITRE","20- TRANSIT L3H3","21- TRANSIT L4H3"]},
    {"key":"complemento_regra","label":"FORNECEDOR","scope":"secundaria","selection_mode":"unitaria","description_order":7,"required":true,"free_text":false,"options":["1- JI","2- JI/PF ES","3- OMZ","4- ORIGINAL","5- PILAR PF","6- PILAR PF ES","7- PILAR PF PLUS","8- TERMO"]},
    {"key":"tipo1","label":"TIPO","scope":"secundaria","selection_mode":"unitaria","description_order":8,"required":true,"free_text":false,"options":["1- ESSENCIAL","2- PLUS"]},
    {"key":"material1","label":"MATERIAL/COR","scope":"secundaria","selection_mode":"unitaria","description_order":9,"required":true,"free_text":false,"options":["1- ABS","2- ACO","3- MADEIRA REVESTIDA","4- PLASTICO","5- PRFV","6- PVC"]},
    {"key":"acabamento1","label":"ACABAMENTO","scope":"secundaria","selection_mode":"unitaria","description_order":10,"required":false,"free_text":false,"options":["1- VELUDO"]},
    {"key":"cor1","label":"COR","scope":"secundaria","selection_mode":"unitaria","description_order":11,"required":false,"free_text":false,"options":["1- BEGE","2- BEGE RAL 1013 PU - TI.01.J85","3- CINZA RAL 7021 PU - TI.01.J04","4- PRETO","5- PRETO FOSCO VINILICA"]},
    {"key":"especificidade1","label":"ESPECIFICIDADE","scope":"secundaria","selection_mode":"unitaria","description_order":12,"required":false,"free_text":false,"options":["1- C/ CORTINA VELUDO PRETA","2- C/ D CLASS","3- NOVA VERSAO"]}
  ]$$::jsonb as fields
)
update public.cadastro_catalogo as catalogo
set
  payload = jsonb_set(
    catalogo.payload,
    '{categories}',
    (
      select jsonb_agg(
        case
          when categoria ->> 'key' = 'cat_18_revestimento'
            then jsonb_set(categoria, '{fields}', revestimento_fields.fields, true)
          else categoria
        end
      )
      from jsonb_array_elements(catalogo.payload -> 'categories') as categoria
      cross join revestimento_fields
    ),
    true
  ),
  updated_at = timezone('utc', now())
where catalogo.config_key = 'default'
  and exists (
    select 1
    from jsonb_array_elements(catalogo.payload -> 'categories') as categoria
    where categoria ->> 'key' = 'cat_18_revestimento'
  );
