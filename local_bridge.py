import argparse
import json
import os
import platform
import sys
import time
import traceback
from pathlib import Path
from urllib import error, request

import bridge_store
import excel_bancos


APP_NAME = "Ponte Local - Módulo Cadastro"
DEFAULT_INTERVAL_SECONDS = 3
PRODUCT_SYNC_INTERVAL_SECONDS = 180


def clean_url(value: str) -> str:
    return (value or "").strip().rstrip("/")


def load_config(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def http_json(method: str, url: str, token: str, payload: dict | None = None, timeout: int = 30) -> dict:
    data = None
    headers = {"Authorization": f"Bearer {token}"}
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = request.Request(url, data=data, headers=headers, method=method.upper())
    with request.urlopen(req, timeout=timeout) as response:
        body = response.read().decode("utf-8")
        return json.loads(body) if body else {}


def heartbeat(server_url: str, token: str, message: str = "") -> None:
    http_json(
        "POST",
        f"{server_url}/api/ponte/heartbeat",
        token,
        {
            "machine": platform.node() or os.environ.get("COMPUTERNAME", ""),
            "workbook_path": excel_bancos.template_path(),
            "message": message,
        },
        timeout=20,
    )


def sync_products(server_url: str, token: str) -> None:
    products = excel_bancos._product_catalog_from_workbook()
    http_json(
        "POST",
        f"{server_url}/api/ponte/produtos",
        token,
        {"products": products},
        timeout=120,
    )
    print(f"Catálogo sincronizado: {len(products)} produto(s).", flush=True)


def process_save_registration(job: dict) -> dict:
    payload = job.get("payload") or {}
    pairs = payload.get("form_pairs") or []
    form = bridge_store.FormPayload(pairs)

    fields = excel_bancos.get_banco_fields(excel_bancos.clean_text(form.get("categoria")))
    needs_bom = excel_bancos.requires_component_bom(fields, form)
    components = excel_bancos.parse_component_lines(form) if needs_bom else []
    if needs_bom and not components:
        raise ValueError("Inclua pelo menos um componente para conjunto ou produto em processo.")

    result = excel_bancos.save_banco_registration(form)
    if needs_bom:
        bom_path = excel_bancos.generate_bom_workbook(
            result.get("sku") or excel_bancos.clean_text(form.get("bom_item_codigo")),
            result.get("descricao_primaria") or result.get("sku") or "",
            components,
        )
        result["bom_path"] = str(bom_path)
    return result


def process_job(server_url: str, token: str, job: dict) -> None:
    job_id = job.get("id")
    print(f"Processando job {job_id} ({job.get('type')})...", flush=True)
    try:
        if job.get("type") != "save_registration":
            raise ValueError(f"Tipo de job não suportado: {job.get('type')}")
        result = process_save_registration(job)
        http_json("POST", f"{server_url}/api/ponte/jobs/{job_id}/complete", token, {"result": result}, timeout=60)
        print(f"Job {job_id} concluído. SKU: {result.get('sku') or '-'}", flush=True)
    except Exception as exc:
        traceback.print_exc()
        http_json(
            "POST",
            f"{server_url}/api/ponte/jobs/{job_id}/fail",
            token,
            {"error": str(exc), "retry": True},
            timeout=60,
        )
        print(f"Job {job_id} falhou: {exc}", flush=True)


def run(server_url: str, token: str, interval: int) -> None:
    if not server_url:
        raise SystemExit("Informe a URL do app online. Ex: --server https://modulo-cadastro.onrender.com")
    if not token:
        raise SystemExit("Informe o token da ponte. Ex: --token SEU_TOKEN")

    print(APP_NAME, flush=True)
    print(f"Servidor online: {server_url}", flush=True)
    print(f"Planilha local: {excel_bancos.template_path()}", flush=True)

    last_product_sync = 0.0
    while True:
        try:
            heartbeat(server_url, token, "online")
            now = time.time()
            if now - last_product_sync >= PRODUCT_SYNC_INTERVAL_SECONDS:
                sync_products(server_url, token)
                last_product_sync = now

            payload = http_json("GET", f"{server_url}/api/ponte/jobs/next", token, timeout=30)
            job = payload.get("job")
            if job:
                process_job(server_url, token, job)
                last_product_sync = 0.0
            else:
                time.sleep(interval)
        except KeyboardInterrupt:
            print("Ponte encerrada pelo usuário.", flush=True)
            return
        except error.HTTPError as exc:
            print(f"Erro HTTP na ponte: {exc.code} {exc.reason}", flush=True)
            time.sleep(max(interval, 5))
        except Exception as exc:
            print(f"Erro na ponte: {exc}", flush=True)
            time.sleep(max(interval, 5))


def main() -> None:
    default_config = Path(__file__).resolve().with_name("ponte_config.json")
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--config", default=str(default_config), help="Caminho do arquivo ponte_config.json.")
    parser.add_argument("--server", default="", help="URL do app online no Render.")
    parser.add_argument("--token", default="", help="Token configurado em CADASTRO_BRIDGE_TOKEN no Render.")
    parser.add_argument("--interval", type=int, default=DEFAULT_INTERVAL_SECONDS, help="Intervalo de consulta em segundos.")
    args = parser.parse_args()

    config = load_config(Path(args.config))
    server_url = clean_url(args.server or os.environ.get("CADASTRO_BRIDGE_SERVER", "") or config.get("server_url", ""))
    token = args.token or os.environ.get("CADASTRO_BRIDGE_TOKEN", "") or config.get("token", "")
    run(server_url, token, max(1, args.interval))


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Falha ao iniciar ponte: {exc}", file=sys.stderr)
        raise
