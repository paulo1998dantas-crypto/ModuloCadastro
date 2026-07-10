import os
import socket
import sys
import base64
import hashlib
import hmac
import json
import time
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, Form, Header, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
from fastapi.templating import Jinja2Templates

import bridge_store
import excel_bancos


HOST = os.environ.get("HOST", "127.0.0.1")
PORT = int(os.environ.get("PORT", "8001"))
app = FastAPI(title="Módulo de Cadastro")
SESSION_COOKIE = "cadastro_session"
SESSION_MAX_AGE_SECONDS = 12 * 60 * 60
def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = _app_dir()
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))


def _template_dir() -> Path:
    candidates = [
        BASE_DIR / "templates",
        BASE_DIR / "_internal" / "templates",
        RESOURCE_DIR / "templates",
        RESOURCE_DIR / "_internal" / "templates",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return BASE_DIR / "_internal" / "templates"


templates = Jinja2Templates(directory=str(_template_dir()))


def _resource_file(name: str) -> Path:
    for base in (BASE_DIR, RESOURCE_DIR):
        candidate = base / name
        if candidate.exists():
            return candidate
    return BASE_DIR / name


def _login_required() -> bool:
    value = os.environ.get("CADASTRO_REQUIRE_LOGIN")
    if value is not None:
        return value.strip().lower() not in {"0", "false", "nao", "não", "no"}
    return bridge_store.save_via_bridge()


def _persistence_required() -> bool:
    value = os.environ.get("CADASTRO_REQUIRE_PERSISTENCE")
    if value is not None:
        return value.strip().lower() not in {"0", "false", "nao", "não", "no"}
    return bridge_store.save_via_bridge()


def _persistence_ok() -> bool:
    return bool(bridge_store.persistence_info().get("persistent"))


def _auth_user() -> str:
    return os.environ.get("CADASTRO_ADMIN_USER", "admin").strip() or "admin"


def _auth_password() -> str:
    return (
        os.environ.get("CADASTRO_ADMIN_PASSWORD", "").strip()
        or os.environ.get("CADASTRO_BRIDGE_TOKEN", "").strip()
    )


def _session_secret() -> bytes:
    secret = (
        os.environ.get("CADASTRO_SESSION_SECRET", "").strip()
        or os.environ.get("CADASTRO_BRIDGE_TOKEN", "").strip()
        or "dev-local-session-secret"
    )
    return secret.encode("utf-8")


def _b64encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode("ascii").rstrip("=")


def _b64decode(data: str) -> bytes:
    padding = "=" * (-len(data) % 4)
    return base64.urlsafe_b64decode((data + padding).encode("ascii"))


def _sign_session(payload: str) -> str:
    return hmac.new(_session_secret(), payload.encode("utf-8"), hashlib.sha256).hexdigest()


def _make_session(username: str) -> str:
    payload = {
        "u": username,
        "exp": int(time.time()) + SESSION_MAX_AGE_SECONDS,
    }
    encoded = _b64encode(json.dumps(payload, separators=(",", ":")).encode("utf-8"))
    return f"{encoded}.{_sign_session(encoded)}"


def _read_session(request: Request) -> str:
    raw = request.cookies.get(SESSION_COOKIE, "")
    if "." not in raw:
        return ""
    encoded, signature = raw.rsplit(".", 1)
    if not hmac.compare_digest(_sign_session(encoded), signature):
        return ""
    try:
        payload = json.loads(_b64decode(encoded).decode("utf-8"))
    except Exception:
        return ""
    if int(payload.get("exp") or 0) < int(time.time()):
        return ""
    return str(payload.get("u") or "")


def _is_public_path(path: str) -> bool:
    return (
        path in {"/login", "/healthz", "/favicon.ico"}
        or path.startswith("/api/ponte/")
    )


@app.middleware("http")
async def require_login_middleware(request: Request, call_next):
    if not _login_required() or _is_public_path(request.url.path):
        return await call_next(request)
    if _read_session(request):
        if (
            bridge_store.save_via_bridge()
            and _persistence_required()
            and request.method.upper() in {"POST", "PUT", "PATCH", "DELETE"}
            and request.url.path not in {"/login", "/logout"}
            and not request.url.path.startswith("/api/ponte/")
            and not _persistence_ok()
        ):
            message = "Persistência do Render não está ativa. Configure o Persistent Disk antes de salvar alterações."
            if request.url.path.startswith("/api/"):
                return JSONResponse({"ok": False, "error": message}, status_code=503)
            return HTMLResponse(message, status_code=503)
        return await call_next(request)
    if request.url.path.startswith("/api/"):
        return JSONResponse({"ok": False, "error": "Login obrigatório."}, status_code=401)
    next_url = quote(str(request.url.path))
    return RedirectResponse(url=f"/login?next={next_url}", status_code=303)


def port_is_in_use() -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.2)
        return sock.connect_ex((HOST, PORT)) == 0


def _enrich_fields(fields, form_data=None):
    normalized = form_data or {}
    enriched = []
    for field in fields:
        selected_values = normalized.get(field["key"], [])
        selected_value = selected_values[0] if selected_values else ""
        enriched.append(
            {
                **field,
                "selected_values": selected_values,
                "selected_value": selected_value,
            }
        )
    return enriched


def _normalize_form_data(fields, form_data):
    values = {}
    for field in fields:
        selected = excel_bancos._serialize_field_values(field, form_data)
        values[field["key"]] = selected
    return values


def _sync_active_workbook(category_key: str) -> None:
    return None


def _workbook_display_path() -> str:
    if bridge_store.save_via_bridge():
        return "Modo online: cadastros serão gravados pela Ponte Local"
    return excel_bancos.template_path()


def _online_draft_payload(category_key: str, draft_payload: str, draft_id: str = "") -> dict:
    category = excel_bancos.selected_category(category_key)
    fields = excel_bancos.get_banco_fields(category["key"])
    groups = excel_bancos._draft_payload_groups(draft_payload)
    if not groups:
        raise ValueError("Rascunho vazio.")
    descriptions = excel_bancos.build_descriptions(fields, groups, category["key"])
    return bridge_store.save_draft(
        category["key"],
        category["label"],
        category.get("sheet_name") or category["label"],
        descriptions.get("primaria") or "",
        {"category": category["key"], "groups": groups},
        draft_id,
    )


def _render_cadastro_page(
    request: Request,
    categoria: str = "",
    sucesso: str = "",
    erro: str = "",
    form_data=None,
    draft_id: str = "",
):
    active_draft = None
    online_mode = bridge_store.save_via_bridge()
    if draft_id and form_data is None and online_mode:
        active_draft = bridge_store.get_draft(draft_id)
        if active_draft:
            categoria = active_draft["category_key"]
            form_data = active_draft["groups"]
    elif draft_id and form_data is None:
        active_draft = excel_bancos.get_registration_draft(draft_id)
        if active_draft:
            categoria = active_draft["category_key"]
            form_data = active_draft["groups"]
    elif draft_id and form_data is not None:
        groups = {}
        if hasattr(form_data, "multi_items"):
            for key, value in form_data.multi_items():
                groups.setdefault(excel_bancos.clean_text(key), []).append(excel_bancos.clean_text(value))
        active_draft = {"draft_id": draft_id, "groups": groups}
    selected_category = excel_bancos.selected_category(categoria)
    fields = excel_bancos.get_banco_fields(selected_category["key"])
    normalized_form = _normalize_form_data(fields, form_data) if form_data is not None else {}
    return templates.TemplateResponse(
        request=request,
        name="cadastro_bancos.html",
        context={
            "request": request,
            "categories": excel_bancos.list_categories(),
            "selected_category": selected_category,
            "fields": _enrich_fields(fields, normalized_form),
            "ordered_fields": _enrich_fields(
                excel_bancos.get_banco_fields_for_display(selected_category["key"]),
                normalized_form,
            ),
            "conditional_rules": excel_bancos.get_conditional_rules_for_form(selected_category["key"]),
            "workbook_path": _workbook_display_path(),
            "save_via_bridge": online_mode,
            "sucesso": sucesso,
            "erro": erro,
            "form_data": normalized_form,
            "drafts": bridge_store.list_drafts() if online_mode else excel_bancos.list_registration_drafts(),
            "active_draft": active_draft or ({"draft_id": draft_id} if draft_id else None),
            "active_page": "cadastro",
        },
    )


def _render_opcoes_page(request: Request, categoria: str = "", sucesso: str = "", erro: str = ""):
    selected_category = excel_bancos.selected_category(categoria)
    return templates.TemplateResponse(
        request=request,
        name="opcoes.html",
        context={
            "request": request,
            "categories": excel_bancos.list_categories(),
            "selected_category": selected_category,
            "fields": excel_bancos.get_banco_fields(selected_category["key"]),
            "ordered_fields": excel_bancos.get_banco_fields_for_display(selected_category["key"]),
            "conditional_rules": excel_bancos.get_conditional_rules(selected_category["key"]),
            "workbook_path": _workbook_display_path(),
            "save_via_bridge": bridge_store.save_via_bridge(),
            "sucesso": sucesso,
            "erro": erro,
            "active_page": "opcoes",
        },
    )


@app.get("/", response_class=HTMLResponse)
async def home():
    return RedirectResponse(url="/cadastro/bancos", status_code=303)


@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request, next: str = "/cadastro/bancos", erro: str = ""):
    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={
            "request": request,
            "erro": erro,
            "next_url": next or "/cadastro/bancos",
            "username": _auth_user(),
            "auth_configured": bool(_auth_password()),
        },
    )


@app.post("/login")
async def login_post(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    next_url: str = Form("/cadastro/bancos"),
):
    expected_password = _auth_password()
    if not expected_password:
        return RedirectResponse(url="/login?erro=Login não configurado.", status_code=303)
    if username != _auth_user() or not hmac.compare_digest(password, expected_password):
        return RedirectResponse(url=f"/login?erro={quote('Usuário ou senha inválidos.')}", status_code=303)
    response = RedirectResponse(url=next_url or "/cadastro/bancos", status_code=303)
    response.set_cookie(
        SESSION_COOKIE,
        _make_session(username),
        max_age=SESSION_MAX_AGE_SECONDS,
        httponly=True,
        secure=request.url.scheme == "https",
        samesite="lax",
    )
    return response


@app.get("/logout")
async def logout():
    response = RedirectResponse(url="/login", status_code=303)
    response.delete_cookie(SESSION_COOKIE)
    return response


@app.head("/")
async def home_head():
    return Response(status_code=200)


@app.get("/healthz")
async def healthz():
    return {"ok": True, "mode": bridge_store.save_mode(), "persistence": bridge_store.persistence_info()}


@app.head("/healthz")
async def healthz_head():
    return Response(status_code=200)


@app.get("/cadastro/bancos", response_class=HTMLResponse)
async def cadastro_bancos_page(
    request: Request,
    categoria: str = "",
    sucesso: str = "",
    erro: str = "",
    draft_id: str = "",
):
    return _render_cadastro_page(request, categoria=categoria, sucesso=sucesso, erro=erro, draft_id=draft_id)


@app.post("/cadastro/bancos", response_class=HTMLResponse)
async def cadastro_bancos_post(request: Request):
    form_data = await request.form()
    category_key = excel_bancos.clean_text(form_data.get("categoria"))
    draft_id = excel_bancos.clean_text(form_data.get("draft_id"))
    try:
        fields = excel_bancos.get_banco_fields(category_key)
        needs_bom = excel_bancos.requires_component_bom(fields, form_data)
        components = excel_bancos.parse_component_lines(form_data) if needs_bom else []
        bom_item_code = excel_bancos.clean_text(form_data.get("bom_item_codigo"))
        if needs_bom and not components:
            raise ValueError("Inclua pelo menos um componente para conjunto ou produto em processo.")

        if bridge_store.save_via_bridge():
            if not bridge_store.token_configured():
                raise ValueError("Modo online ativo, mas CADASTRO_BRIDGE_TOKEN não foi configurado no Render.")
            job = bridge_store.enqueue_registration(form_data, category_key)
            if draft_id:
                try:
                    bridge_store.delete_draft(draft_id)
                except Exception:
                    pass
            message = (
                f"Cadastro enviado para a ponte local. Protocolo: {job['id']}. "
                "Quando a ponte estiver aberta no PC da produção, ela gravará na planilha local."
            )
            return RedirectResponse(
                url=f"/cadastro/bancos?categoria={quote(category_key)}&sucesso={quote(message)}",
                status_code=303,
            )

        result = excel_bancos.save_banco_registration(form_data)
        if draft_id:
            try:
                excel_bancos.delete_registration_draft(draft_id)
            except Exception:
                pass
        if needs_bom:
            bom_item_code = result.get("sku") or bom_item_code
            bom_path = excel_bancos.generate_bom_workbook(
                bom_item_code,
                result.get("descricao_primaria") or bom_item_code,
                components,
            )
            return FileResponse(
                bom_path,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=bom_path.name,
            )
        message = f"Cadastro salvo na linha {result['row']}. SKU: {result.get('sku') or '-'}."
        return RedirectResponse(
            url=f"/cadastro/bancos?categoria={quote(result['category_key'])}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return _render_cadastro_page(request, categoria=category_key, erro=str(exc), form_data=form_data, draft_id=draft_id)


@app.post("/rascunhos/salvar")
async def rascunhos_salvar_post(
    category_key: str = Form(...),
    draft_payload: str = Form(...),
    draft_id: str = Form(""),
):
    try:
        if bridge_store.save_via_bridge():
            result = _online_draft_payload(category_key, draft_payload, draft_id)
        else:
            result = excel_bancos.save_registration_draft(category_key, draft_payload, draft_id)
        return JSONResponse({"ok": True, "draft": result})
    except Exception as exc:
        return JSONResponse({"ok": False, "error": str(exc)}, status_code=400)


@app.post("/rascunhos/excluir")
async def rascunhos_excluir_post(draft_id: str = Form(...), category_key: str = Form("")):
    try:
        if bridge_store.save_via_bridge():
            result = bridge_store.delete_draft(draft_id)
        else:
            result = excel_bancos.delete_registration_draft(draft_id)
        message = f"Rascunho excluído: {result.get('category_label') or draft_id}."
        return RedirectResponse(
            url=f"/cadastro/bancos?categoria={quote(category_key)}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/cadastro/bancos?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.post("/api/rascunhos/excluir")
async def api_rascunhos_excluir_post(draft_id: str = Form(...)):
    try:
        if bridge_store.save_via_bridge():
            result = bridge_store.delete_draft(draft_id)
        else:
            result = excel_bancos.delete_registration_draft(draft_id)
        return JSONResponse({"ok": True, "draft": result})
    except Exception as exc:
        return JSONResponse({"ok": False, "error": str(exc)}, status_code=400)


@app.get("/api/produtos")
async def api_produtos(q: str = ""):
    if bridge_store.save_via_bridge():
        return {"items": _search_bridge_products(q)}
    return {"items": excel_bancos.search_products(q)}


def _search_bridge_products(query: str, limit: int = 25):
    term = excel_bancos.clean_text(query)
    if len(term) < 1:
        return []
    normalized_term = excel_bancos.normalize_label(term)
    compact_term = normalized_term.replace(" ", "")
    matches = []
    for product in bridge_store.products():
        code = excel_bancos.clean_text(product.get("codigo"))
        description = excel_bancos.clean_text(product.get("descricao"))
        haystack = excel_bancos.normalize_label(f"{code} {description} {product.get('categoria')}")
        compact_haystack = haystack.replace(" ", "")
        if compact_term not in compact_haystack:
            continue
        score = 0
        if code.startswith(term):
            score -= 30
        if excel_bancos.normalize_label(description).startswith(normalized_term):
            score -= 15
        score += len(description)
        matches.append((score, product))
    matches.sort(key=lambda item: (item[0], item[1]["codigo"]))
    return [product for _, product in matches[:limit]]


def _require_bridge_token(authorization: str = "") -> None:
    if not bridge_store.verify_token(authorization):
        raise HTTPException(status_code=401, detail="Token da ponte inválido ou ausente.")


@app.get("/ponte", response_class=HTMLResponse)
async def ponte_page(request: Request, sucesso: str = "", erro: str = ""):
    return templates.TemplateResponse(
        request=request,
        name="ponte.html",
        context={
            "request": request,
            "status": bridge_store.status(),
            "save_via_bridge": bridge_store.save_via_bridge(),
            "sucesso": sucesso,
            "erro": erro,
            "active_page": "ponte",
        },
    )


@app.get("/ponte/download")
async def ponte_download():
    return FileResponse(
        _resource_file("local_bridge.py"),
        media_type="text/x-python",
        filename="ponte_local_modulo_cadastro.py",
    )


@app.get("/ponte/download-bat")
async def ponte_download_bat():
    return FileResponse(
        _resource_file("iniciar_ponte_local.bat"),
        media_type="application/x-bat",
        filename="iniciar_ponte_local.bat",
    )


@app.get("/api/ponte/status")
async def api_ponte_status():
    public_status = bridge_store.status(limit=10)
    public_status["store_path"] = "" if bridge_store.save_via_bridge() else public_status.get("store_path", "")
    return public_status


@app.post("/api/ponte/heartbeat")
async def api_ponte_heartbeat(request: Request, authorization: str = Header("")):
    _require_bridge_token(authorization)
    payload = await request.json()
    return {"ok": True, "bridge": bridge_store.heartbeat(payload if isinstance(payload, dict) else {})}


@app.get("/api/ponte/jobs/next")
async def api_ponte_jobs_next(authorization: str = Header("")):
    _require_bridge_token(authorization)
    job = bridge_store.next_job()
    return {"ok": True, "job": job}


@app.post("/api/ponte/jobs/{job_id}/complete")
async def api_ponte_jobs_complete(job_id: str, request: Request, authorization: str = Header("")):
    _require_bridge_token(authorization)
    payload = await request.json()
    result = payload.get("result") if isinstance(payload, dict) else {}
    return {"ok": True, "job": bridge_store.complete_job(job_id, result or {})}


@app.post("/api/ponte/jobs/{job_id}/fail")
async def api_ponte_jobs_fail(job_id: str, request: Request, authorization: str = Header("")):
    _require_bridge_token(authorization)
    payload = await request.json()
    error = payload.get("error") if isinstance(payload, dict) else "Erro não informado."
    retry = bool(payload.get("retry", True)) if isinstance(payload, dict) else True
    return {"ok": True, "job": bridge_store.fail_job(job_id, error, retry)}


@app.post("/api/ponte/produtos")
async def api_ponte_produtos(request: Request, authorization: str = Header("")):
    _require_bridge_token(authorization)
    payload = await request.json()
    products = payload.get("products") if isinstance(payload, dict) else []
    if not isinstance(products, list):
        products = []
    return {"ok": True, "catalog": bridge_store.replace_products(products)}


@app.get("/opcoes", response_class=HTMLResponse)
async def opcoes_page(request: Request, categoria: str = "", sucesso: str = "", erro: str = ""):
    return _render_opcoes_page(request, categoria=categoria, sucesso=sucesso, erro=erro)


@app.post("/opcoes")
async def opcoes_post(
    category_key: str = Form(...),
    field_key: str = Form(...),
    option_value: str = Form(...),
):
    try:
        result = excel_bancos.add_field_option(category_key, field_key, option_value)
        message = f"Opção {result['option']} adicionada em {result['field']}."
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.post("/opcoes/editar")
async def opcoes_editar_post(
    category_key: str = Form(...),
    field_key: str = Form(...),
    option_row: int = Form(...),
    option_value: str = Form(...),
):
    try:
        result = excel_bancos.update_field_option(category_key, field_key, option_row, option_value)
        message = f"Opção atualizada: {result['option']}."
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.post("/opcoes/salvar-lote")
async def opcoes_salvar_lote_post(
    category_key: str = Form(...),
    field_key: str = Form(...),
    option_row: list[int] = Form(...),
    option_value: list[str] = Form(...),
):
    try:
        result = excel_bancos.update_field_options(category_key, field_key, option_row, option_value)
        message = f"{result['count']} opção(ões) atualizada(s) em {result['field']}."
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.post("/opcoes/excluir")
async def opcoes_excluir_post(
    category_key: str = Form(...),
    field_key: str = Form(...),
    option_row: int = Form(...),
):
    try:
        result = excel_bancos.delete_field_option(category_key, field_key, option_row)
        message = f"Opção excluída: {result['option']}."
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.post("/campos/reordenar")
async def campos_reordenar_post(
    category_key: str = Form(...),
    scope: str = Form(...),
    ordered_field_keys: list[str] = Form(...),
):
    try:
        result = excel_bancos.reorder_fields_by_description(category_key, scope, ordered_field_keys)
        message = f"Ordem atualizada em {result['scope']}."
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.post("/regras/adicionar")
async def regras_adicionar_post(
    category_key: str = Form(...),
    source_field_key: str = Form(...),
    source_value: str = Form(...),
    target_field_key: str = Form(""),
    target_field_label: str = Form(""),
    target_field_scope: str = Form("secundaria"),
    action: str = Form("hide"),
):
    try:
        result = excel_bancos.add_conditional_rule(
            category_key,
            source_field_key,
            source_value,
            target_field_key,
            target_field_label,
            action,
        )
        if not excel_bancos.clean_text(target_field_key) and excel_bancos.clean_text(target_field_label):
            catalog = excel_bancos.load_catalog()
            category = next((item for item in catalog["categories"] if item["key"] == category_key), None)
            if category is not None:
                for rule in category.get("conditional_rules") or []:
                    if rule.get("key") == result["rule"]["key"]:
                        if excel_bancos.clean_text(target_field_label):
                            rule["target_field_label"] = excel_bancos.clean_text(target_field_label)
                        rule["target_field_scope"] = excel_bancos.clean_text(target_field_scope) or "secundaria"
                        excel_bancos.save_catalog(catalog)
                        break
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&sucesso={quote('Regra condicional criada.')}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.post("/regras/excluir")
async def regras_excluir_post(category_key: str = Form(...), rule_key: str = Form(...)):
    try:
        excel_bancos.delete_conditional_rule(category_key, rule_key)
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&sucesso={quote('Regra condicional excluída.')}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.post("/categorias/adicionar")
async def categorias_adicionar_post(category_label: str = Form(...)):
    try:
        result = excel_bancos.add_category(category_label)
        message = f"Categoria criada: {result['category']}."
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(result['category_key'])}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(url=f"/opcoes?erro={quote(str(exc))}", status_code=303)


@app.post("/categorias/editar")
async def categorias_editar_post(category_key: str = Form(...), category_label: str = Form(...)):
    try:
        result = excel_bancos.update_category(category_key, category_label)
        message = f"Categoria atualizada: {result['category']}."
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(result['category_key'])}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.post("/categorias/excluir")
async def categorias_excluir_post(category_key: str = Form(...)):
    try:
        result = excel_bancos.delete_category(category_key)
        message = f"Categoria excluída: {result['category']}."
        active_category = excel_bancos.selected_category("")
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(active_category['key'])}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.post("/campos/adicionar")
async def campos_adicionar_post(
    category_key: str = Form(...),
    field_label: str = Form(...),
    field_scope: str = Form(...),
    field_selection_mode: str = Form(...),
):
    try:
        result = excel_bancos.add_field(category_key, field_label, field_scope, field_selection_mode)
        message = f"Campo criado: {result['field']}."
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.post("/campos/editar")
async def campos_editar_post(
    category_key: str = Form(...),
    field_key: str = Form(...),
    field_label: str = Form(...),
    field_scope: str = Form(...),
    field_selection_mode: str = Form(...),
):
    try:
        result = excel_bancos.update_field(
            category_key,
            field_key,
            field_label,
            field_scope,
            field_selection_mode,
        )
        message = f"Campo atualizado: {result['field']}."
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.post("/campos/excluir")
async def campos_excluir_post(category_key: str = Form(...), field_key: str = Form(...)):
    try:
        result = excel_bancos.delete_field(category_key, field_key)
        message = f"Campo excluído: {result['field']}."
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.get("/planilha", response_class=HTMLResponse)
async def planilha_page(request: Request, sucesso: str = "", erro: str = ""):
    return templates.TemplateResponse(
        request=request,
        name="planilha.html",
        context={
            "request": request,
            "workbooks": [] if bridge_store.save_via_bridge() else excel_bancos.list_workbooks(),
            "folders": [] if bridge_store.save_via_bridge() else excel_bancos.list_folders(),
            "workbook_path": _workbook_display_path(),
            "save_via_bridge": bridge_store.save_via_bridge(),
            "bridge_status": bridge_store.status(limit=5),
            "app_config": bridge_store.app_config(),
            "sucesso": sucesso,
            "erro": erro,
            "active_page": "planilha",
        },
    )


@app.post("/planilha")
async def planilha_post(
    workbook_select: str = Form(""),
    manual_workbook_path: str = Form(""),
    folder_select: str = Form(""),
    manual_folder_path: str = Form(""),
    workbook_name: str = Form(""),
    master_workbook_path: str = Form(""),
    master_workbook_url: str = Form(""),
):
    try:
        if bridge_store.save_via_bridge():
            config = bridge_store.save_app_config(master_workbook_path, master_workbook_url)
            message = "Planilha-mãe definida para o modo online."
            if config.get("master_workbook_path"):
                message += f" Caminho: {config['master_workbook_path']}"
            return RedirectResponse(url=f"/planilha?sucesso={quote(message)}", status_code=303)

        selected_workbook = excel_bancos.clean_text(manual_workbook_path) or excel_bancos.clean_text(workbook_select)
        selected_folder = excel_bancos.clean_text(manual_folder_path) or excel_bancos.clean_text(folder_select)

        if selected_workbook:
            path = excel_bancos.set_active_workbook(selected_workbook)
            message = f"Planilha existente definida: {path.name}"
        elif selected_folder:
            path = excel_bancos.set_active_workbook_from_folder(selected_folder, workbook_name)
            message = f"Planilha criada/definida: {path.name}"
        else:
            raise ValueError("Selecione uma planilha existente ou uma pasta para criar a planilha.")

        return RedirectResponse(url=f"/planilha?sucesso={quote(message)}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/planilha?erro={quote(str(exc))}", status_code=303)


if __name__ == "__main__":
    if port_is_in_use():
        print(
            f"O Módulo de Cadastro já está aberto em http://{HOST}:{PORT}. "
            "Feche o servidor anterior se quiser iniciar novamente."
        )
    else:
        import uvicorn

        uvicorn.run("main:app", host=HOST, port=PORT, reload=False)
