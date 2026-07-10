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

O token do Render precisa ser o mesmo token usado pela Ponte Local.

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
