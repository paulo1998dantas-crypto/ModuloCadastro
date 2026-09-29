# ModuloCadastro

Modulo de dados mestres da JI Montadora.

## Login compartilhado

Em producao, o modulo autentica na mesma tabela `public.users` usada pelo
ModuloEstoque:

- `CADASTRO_AUTH_MODE=shared`
- `CADASTRO_SHARED_RBAC_ENABLED=1` ativa o RBAC multi-perfil depois da migration
  compartilhada.

Com o RBAC ativo, usuarios com perfil `ADMIN` ou permissao `cadastro.access`
entram no modulo. O perfil `COMERCIAL` recebe somente essa permissao e o
backend limita sua navegação à consulta dos cadastros e à gestão de pessoas em
Suprimentos; as demais rotas são bloqueadas também no servidor. Alteração de
senha, perfis, permissão ou estado ativo incrementa `users.auth_version` e
revoga a sessão anterior.

Durante o corte, mantenha `CADASTRO_SHARED_RBAC_ENABLED=0`; o rollback consiste
em voltar a flag para `0`, sem apagar tabelas ou usuarios. Com a flag desligada
o modo de compatibilidade reconhece `ADMIN`, `ENGENHARIA`, `PCP` e `COMERCIAL`.
O papel COMERCIAL também é reconhecido como perfil adicional quando o schema
RBAC compartilhado está instalado. Com a flag ligada, a RPC
`erp_get_user_access`, `auth_version` e
o schema RBAC são obrigatórios; falhas bloqueiam o acesso e deixam `/healthz`
em `503`.

## Lead time e custo por item

A tela **Cadastros** possui uma ação por SKU para parametrizar a origem de
fabricação e os tempos padrão em dias úteis:

- externa: fornecimento, transporte, recebimento, inspeção, estocagem,
  expedição e montagem de kit;
- interna: setup, produção e liberação;
- externa também permite informar o preço de compra e consultar a média
  ponderada das 10 entradas confirmadas mais recentes.

O preço médio é somente consultivo. Ele ignora recebimentos rejeitados,
estornados, sem quantidade aprovada ou sem valor unitário real e não altera
saldo, custo ou movimentação de Estoque.

Antes de liberar a função, aplique a migration aditiva
`supabase_item_lead_time_cost.sql` em staging e valide o rollback pelo backup.
Ela cria apenas as tabelas de parâmetros/histórico e a RPC de consulta de
preço médio.

O importador da planilha padrão inicia sempre em dry-run:

```powershell
python leadtime_import.py "C:\caminho\leadtime.xlsx" --report outputs\leadtime_dry_run.json
```

Depois da conferência do relatório e da aplicação da migration, a gravação é
habilitada explicitamente com `--apply`. Parâmetros existentes são preservados;
`--overwrite-existing` deve ser usado somente após revisão e autorização.
