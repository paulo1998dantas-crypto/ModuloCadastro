# ModuloCadastro no Supabase

## Ordem segura de ativacao

1. Abra o SQL Editor do Supabase do projeto `rodtxswtqbsbtukmvobn`.
2. Rode o arquivo `supabase_migration.sql`.
3. No Render, configure:
   - `CADASTRO_SAVE_MODE=supabase`
   - `SUPABASE_URL=https://rodtxswtqbsbtukmvobn.supabase.co`
   - `SUPABASE_SERVICE_ROLE_KEY=<service role key>`
   - `CADASTRO_REQUIRE_LOGIN=1`
   - `CADASTRO_REQUIRE_PERSISTENCE=0`
   - `CADASTRO_ADMIN_USER=admin`
   - `CADASTRO_ADMIN_PASSWORD=<senha>`
   - `CADASTRO_SESSION_SECRET=<segredo longo>`
4. Rebuild/deploy do Render.
5. Importe a planilha atual uma vez:

```powershell
$env:CADASTRO_SAVE_MODE='supabase'
$env:SUPABASE_URL='https://rodtxswtqbsbtukmvobn.supabase.co'
$env:SUPABASE_SERVICE_ROLE_KEY='<service role key>'
python importar_planilha_supabase.py
```

Para conferir sem inserir:

```powershell
python importar_planilha_supabase.py --dry-run
```

## Tabelas criadas

- `public.cadastro_registros`
- `public.cadastro_rascunhos`

Essas tabelas usam prefixo `cadastro_` para nao colidir com o `moduloestoque`.
