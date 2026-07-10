# Deploy do Módulo Cadastro no Render com Ponte Local

Este projeto pode rodar online no Render sem perder a gravação na planilha local usando a **Ponte Local**.

## Como funciona

1. O usuário abre o Módulo Cadastro online no Render.
2. Ao salvar um cadastro, o Render cria um job em uma fila.
3. A Ponte Local, rodando no PC da produção, consulta essa fila.
4. A Ponte Local salva o cadastro na planilha configurada no PC.
5. Para conjunto/PP, a Ponte Local também gera a planilha B.O.M. localmente.

## Variáveis no Render

Crie o serviço pelo `render.yaml` e configure:

- `CADASTRO_SAVE_MODE=bridge`
- `CADASTRO_DATA_DIR=/var/data`
- `CADASTRO_BRIDGE_TOKEN=<um token grande e secreto>`
- `CADASTRO_REQUIRE_LOGIN=1`
- `CADASTRO_REQUIRE_PERSISTENCE=1`
- `CADASTRO_ADMIN_USER=admin`
- `CADASTRO_ADMIN_PASSWORD=<senha de acesso ao app online>`
- `CADASTRO_SESSION_SECRET=<outro segredo longo>`
- `CADASTRO_MASTER_WORKBOOK_PATH=<caminho local da planilha-mãe no PC da ponte>`
- `CADASTRO_MASTER_WORKBOOK_URL=<link SharePoint da planilha-mãe>`

O token do Render precisa ser o mesmo token usado pela Ponte Local.

## Persistência obrigatória

Para alterações de categorias, campos, opções, regras e fila não sumirem após restart/deploy, o Render precisa ter um **Persistent Disk** montado no caminho configurado em:

```bash
CADASTRO_DATA_DIR=/var/data
```

Com `CADASTRO_REQUIRE_PERSISTENCE=1`, o app bloqueia gravações quando o diretório persistente não está ativo. A tela **Planilha** e o endpoint `/healthz` mostram se a persistência está OK.

## Comando de start no Render

```bash
uvicorn main:app --host 0.0.0.0 --port $PORT
```

## Rodar a Ponte Local

No PC que acessa a planilha:

```bash
python local_bridge.py --server https://SEU-APP.onrender.com --token SEU_TOKEN
```

Ou crie um arquivo `ponte_config.json` ao lado do `local_bridge.py`:

```json
{
  "server_url": "https://SEU-APP.onrender.com",
  "token": "SEU_TOKEN"
}
```

E rode:

```bash
python local_bridge.py
```

## Observações importantes

- O Render não acessa diretamente arquivos `C:\...`; quem faz isso é a Ponte Local.
- A Ponte Local usa a mesma configuração de planilha do app local (`config.json`).
- A busca de componentes online usa o catálogo sincronizado pela Ponte Local.
- Se a Ponte Local estiver fechada, os cadastros ficam pendentes na fila online.
- Para persistir a fila no Render, use o disco persistente definido no `render.yaml`.
