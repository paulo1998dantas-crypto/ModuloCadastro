# Auditoria — correção controlada da categoria Bancos

Data da execução: 28/08/2026

Base de destino: Supabase produtivo compartilhado

Categoria: `bancos`

Origem: `ARQUIVO UPLOAD BANCOS.xlsx`

SHA-256 da origem: `0fcb188c749266efb0c2145d1302c7e77bf2a23b8fbeba2bfd2a4d5c553b48dc`

## Resultado

- 277 SKUs únicos encontrados na planilha e no Cadastro.
- 215 cadastros tiveram campos técnicos reconciliados.
- 62 cadastros marcados como `EXCLUIR` foram inativados, sem exclusão física.
- 0 novos SKUs foram criados.
- 118 descrições primárias foram alteradas pelas regras vigentes.
- 97 descrições primárias permaneceram semanticamente iguais.
- Nenhuma movimentação ou saldo de estoque foi alterado.

## Reconciliação de opções

As opções foram associadas pelo texto completo, desconsiderando prefixos numéricos possivelmente incorretos na planilha. Nenhuma opção nova foi criada fora do campo de medidas de pé/vãos.

Normalizações explícitas:

- `X - DPM` → `19- DPM` — 14 ocorrências.
- `1- PE BCO MCA` → `1- PE BCO MCA NORMAL` — 24 ocorrências.
- `7- PE MCA ANTIGO` → `7- PE MCA ANTIGO NORMAL` — 6 ocorrências.
- `RCEIRO VAO 1191 MM` → `31- TERCEIRO VAO 1191 MM` — 6 ocorrências.

Novas medidas autorizadas, criadas sequencialmente:

- `82- SEGUNDO VAO 165 MM`
- `83- PRIMEIRO VAO 280 MM`
- `84- PRIMEIRO VAO 250 MM`
- `85- TERCEIRO VAO 1130 MM`
- `86- QUARTO VAO 1360 MM`
- `87- PRIMEIRO VAO 295 MM`
- `88- SEGUNDO VAO 855 MM`
- `89- QUARTO VAO 1460 MM`

## Método de aplicação

A execução utilizou uma tabela transitória de estágio e uma única transação PostgreSQL. Antes da gravação foram validados:

- existência e unicidade dos 277 SKUs;
- correspondência do `updated_at` com o snapshot de validação;
- integridade do catálogo de Bancos;
- contagens esperadas de atualização e inativação.

Em caso de concorrência ou divergência, toda a transação seria abortada. Após a aplicação, a tabela transitória foi removida.

## Auditoria gravada na base

Foram registrados eventos em `erp_audit_events` com origem `CADASTRO` e ator `sistema:codex`:

- 215 eventos `ATUALIZACAO_CONTROLADA_BANCOS`;
- 62 eventos `INATIVACAO_CONTROLADA_BANCOS`;
- 1 evento `INCLUSAO_CONTROLADA_MEDIDAS_PE`.

Cada evento referencia o nome e o SHA-256 do arquivo de origem.

## Validação posterior

- 277 registros reconciliados na categoria correta.
- 215 ativos e 62 inativos.
- 0 descrições primárias vazias.
- 0 descrições secundárias vazias.
- 277 SKUs encontrados na tabela operacional `skus`.
- 277 status sincronizados entre Cadastro e Estoque.
- 277 descrições sincronizadas entre Cadastro e Estoque.
- As duas diferenças de unidade são apenas representação equivalente: string vazia no Cadastro e `NULL` no Estoque.
- Estrutura transitória de estágio ausente após o commit da transação.

## Observações

Os registros `10200172` (`BCO SUCATA`) e `10200173` (`BCO PROTOTIPO`) não possuíam estrutura técnica na planilha. As descrições de origem foram preservadas e os campos técnicos permaneceram vazios.

Este documento registra uma correção de dados já aplicada. Ele não executa novamente a operação e não contém credenciais.
