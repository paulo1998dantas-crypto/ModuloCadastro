# ModuloCadastro

Modulo de dados mestres da JI Montadora.

## Login compartilhado

Em producao, o modulo autentica na mesma tabela `public.users` usada pelo
ModuloEstoque:

- `CADASTRO_AUTH_MODE=shared`
- `CADASTRO_SHARED_RBAC_ENABLED=1` ativa o RBAC multi-perfil depois da migration
  compartilhada.

Com o RBAC ativo, somente usuarios com perfil `ADMIN` ou permissao
`cadastro.access` (concedida ao perfil `ENGENHARIA`) entram no modulo. Alteracao
de senha, perfis, permissao ou estado ativo incrementa `users.auth_version` e
revoga a sessao anterior.

Durante o corte, mantenha `CADASTRO_SHARED_RBAC_ENABLED=0`; o rollback consiste
em voltar a flag para `0`, sem apagar tabelas ou usuarios. Com a flag desligada
o modo de compatibilidade permite apenas os perfis legados `ADMIN` e
`ENGENHARIA`. Com a flag ligada, a RPC `erp_get_user_access`, `auth_version` e
o schema RBAC são obrigatórios; falhas bloqueiam o acesso e deixam `/healthz`
em `503`.
