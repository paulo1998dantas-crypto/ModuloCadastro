import os
import secrets
import socket
import sys
import base64
import hashlib
import hmac
import json
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
from fastapi.templating import Jinja2Templates
from werkzeug.security import check_password_hash

import bridge_store
import excel_bancos
import supabase_store
import supabase_suprimentos


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
    if supabase_store.enabled():
        return True
    return bool(bridge_store.persistence_info().get("persistent"))


PASSWORD_HASH_ALGORITHM = "pbkdf2_sha256"
PASSWORD_HASH_ITERATIONS = 260000


def _env_auth_password() -> str:
    return (
        os.environ.get("CADASTRO_ADMIN_PASSWORD", "").strip()
        or os.environ.get("CADASTRO_BRIDGE_TOKEN", "").strip()
    )


def _shared_auth_enabled() -> bool:
    return os.environ.get("CADASTRO_AUTH_MODE", "").strip().lower() in {
        "shared",
        "supabase_users",
        "estoque",
    }


def _shared_auth_configured() -> bool:
    return bool(supabase_store.configured())


def _shared_user_lookup(username: str) -> dict[str, str | int | bool] | None:
    username = excel_bancos.clean_text(username)
    if not username or not _shared_auth_configured():
        return None
    query = urllib.parse.urlencode(
        [
            ("select", "id,username,password_hash,role,active"),
            ("username", f"eq.{username}"),
            ("limit", "1"),
        ]
    )
    url = f"{supabase_store._supabase_url()}/rest/v1/users?{query}"
    request = urllib.request.Request(url, headers=supabase_store._headers(), method="GET")
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            body = response.read().decode("utf-8")
            rows = json.loads(body) if body else []
            return rows[0] if rows else None
    except Exception:
        return None


def _verify_shared_login(username: str, password: str) -> bool:
    user = _shared_user_lookup(username)
    if not user or not bool(user.get("active", True)):
        return False
    password_hash = str(user.get("password_hash") or "")
    if not password_hash:
        return False
    try:
        return check_password_hash(password_hash, password)
    except Exception:
        return False


def _hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        PASSWORD_HASH_ITERATIONS,
    )
    return "$".join(
        [
            PASSWORD_HASH_ALGORITHM,
            str(PASSWORD_HASH_ITERATIONS),
            _b64encode(salt),
            _b64encode(digest),
        ]
    )


def _verify_password_hash(password_hash: str, password: str) -> bool:
    try:
        algorithm, iterations_text, salt_text, digest_text = password_hash.split("$", 3)
        if algorithm != PASSWORD_HASH_ALGORITHM:
            return False
        expected = _b64decode(digest_text)
        digest = hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            _b64decode(salt_text),
            int(iterations_text),
        )
    except Exception:
        return False
    return hmac.compare_digest(digest, expected)


def _auth_record() -> dict[str, str | bool]:
    if _shared_auth_enabled():
        return {
            "username": "",
            "source": "shared",
            "configured": _shared_auth_configured(),
            "updated_at": "",
        }
    stored = bridge_store.auth_config()
    if stored.get("username") and stored.get("password_hash"):
        return {
            "username": stored["username"],
            "password_hash": stored["password_hash"],
            "source": "store",
            "configured": True,
            "updated_at": stored.get("updated_at", ""),
        }
    env_password = _env_auth_password()
    return {
        "username": os.environ.get("CADASTRO_ADMIN_USER", "admin").strip() or "admin",
        "password": env_password,
        "source": "environment" if env_password else "missing",
        "configured": bool(env_password),
        "updated_at": "",
    }


def _auth_user() -> str:
    return str(_auth_record().get("username") or "")


def _auth_configured() -> bool:
    return bool(_auth_record().get("configured"))


def _verify_login(username: str, password: str) -> bool:
    if _shared_auth_enabled():
        return _verify_shared_login(username, password)
    record = _auth_record()
    if username != str(record.get("username") or ""):
        return False
    password_hash = str(record.get("password_hash") or "")
    if password_hash:
        return _verify_password_hash(password_hash, password)
    expected_password = str(record.get("password") or "")
    return bool(expected_password) and hmac.compare_digest(password, expected_password)


def _auth_status() -> dict[str, str | bool]:
    record = _auth_record()
    source = str(record.get("source") or "missing")
    labels = {
        "store": "Arquivo persistente do app",
        "environment": "Variável de ambiente do servidor",
        "shared": "Tabela users compartilhada do ModuloEstoque",
        "missing": "Não configurado",
    }
    return {
        "username": str(record.get("username") or ""),
        "configured": bool(record.get("configured")),
        "source": source,
        "source_label": labels.get(source, source),
        "updated_at": str(record.get("updated_at") or ""),
        "editable": not _shared_auth_enabled(),
    }


def _save_admin_credentials(username: str, password: str, password_confirm: str) -> dict[str, str]:
    username = excel_bancos.clean_text(username)
    password = password.strip()
    password_confirm = password_confirm.strip()
    if not username:
        raise ValueError("Informe o usuário administrador.")
    if len(password) < 6:
        raise ValueError("A senha precisa ter pelo menos 6 caracteres.")
    if password != password_confirm:
        raise ValueError("A confirmação da senha não confere.")
    return bridge_store.save_auth_config(username, _hash_password(password))


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
        or (path.startswith("/admin/setup") and not _shared_auth_enabled() and not _auth_configured())
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


def _form_first_value(form_data, key: str) -> str:
    if not form_data:
        return ""
    if isinstance(form_data, dict):
        value = form_data.get(key)
        if isinstance(value, list):
            return excel_bancos.clean_text(value[0] if value else "")
        return excel_bancos.clean_text(value)
    if hasattr(form_data, "getlist"):
        values = form_data.getlist(key)
        if values:
            return excel_bancos.clean_text(values[0])
    if hasattr(form_data, "get"):
        return excel_bancos.clean_text(form_data.get(key))
    return ""


def _sync_active_workbook(category_key: str) -> None:
    return None


def _supabase_mode() -> bool:
    return supabase_store.enabled()


def _workbook_display_path() -> str:
    if _supabase_mode():
        return supabase_store.display_target()
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
    supabase_mode = _supabase_mode()
    if draft_id and form_data is None and supabase_mode:
        active_draft = supabase_store.get_draft(draft_id)
        if active_draft:
            categoria = active_draft["category_key"]
            form_data = active_draft["groups"]
    elif draft_id and form_data is None and online_mode:
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
            "supabase_mode": supabase_mode,
            "pn_groups": excel_bancos.list_pn_groups(),
            "selected_group_code": excel_bancos._pn_group_code(
                _form_first_value(form_data, excel_bancos.PN_GROUP_FORM_KEY)
            ),
            "unit_options": supabase_store.unidade_options(),
            "selected_unit": supabase_store.normalize_unit(_form_first_value(form_data, "unidade")),
            "selected_bom_option": _form_first_value(form_data, excel_bancos.BOM_FORM_KEY),
            "sucesso": sucesso,
            "erro": erro,
            "form_data": normalized_form,
            "drafts": (
                supabase_store.list_drafts()
                if supabase_mode
                else bridge_store.list_drafts()
                if online_mode
                else excel_bancos.list_registration_drafts()
            ),
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
            "pn_groups": excel_bancos.list_pn_groups(),
            "workbook_path": _workbook_display_path(),
            "save_via_bridge": bridge_store.save_via_bridge(),
            "supabase_mode": _supabase_mode(),
            "sucesso": sucesso,
            "erro": erro,
            "active_page": "opcoes",
        },
    )


@app.get("/", response_class=HTMLResponse)
async def home():
    if _supabase_mode():
        return RedirectResponse(url="/cadastros", status_code=303)
    return RedirectResponse(url="/cadastro/bancos", status_code=303)


@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request, next: str = "/cadastro/bancos", erro: str = "", sucesso: str = ""):
    auth_status = _auth_status()
    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={
            "request": request,
            "erro": erro,
            "sucesso": sucesso,
            "next_url": next or "/cadastro/bancos",
            "username": auth_status["username"],
            "auth_configured": auth_status["configured"],
            "setup_available": (not auth_status["configured"]) and (not _shared_auth_enabled()),
        },
    )


@app.post("/login")
async def login_post(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    next_url: str = Form("/cadastro/bancos"),
):
    auth_is_configured = _auth_configured()
    if not auth_is_configured:
        if _shared_auth_enabled():
            return RedirectResponse(url=f"/login?erro={quote('Login compartilhado não configurado.')}", status_code=303)
        return RedirectResponse(url="/admin/setup", status_code=303)
    if not _verify_login(username, password):
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


@app.get("/admin/setup", response_class=HTMLResponse)
async def admin_setup_page(request: Request, erro: str = ""):
    if _shared_auth_enabled():
        return RedirectResponse(url="/cadastros", status_code=303)
    if _auth_configured():
        return RedirectResponse(url="/admin", status_code=303)
    return templates.TemplateResponse(
        request=request,
        name="admin_setup.html",
        context={
            "request": request,
            "erro": erro,
            "username": _auth_user(),
        },
    )


@app.post("/admin/setup")
async def admin_setup_post(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    password_confirm: str = Form(...),
):
    if _shared_auth_enabled():
        return RedirectResponse(url="/cadastros", status_code=303)
    if _auth_configured():
        return RedirectResponse(url="/admin", status_code=303)
    try:
        _save_admin_credentials(username, password, password_confirm)
        response = RedirectResponse(
            url=f"/login?sucesso={quote('Administrador configurado. Entre com a nova senha.')}",
            status_code=303,
        )
        response.delete_cookie(SESSION_COOKIE)
        return response
    except Exception as exc:
        return RedirectResponse(url=f"/admin/setup?erro={quote(str(exc))}", status_code=303)


@app.get("/admin", response_class=HTMLResponse)
async def admin_page(request: Request, sucesso: str = "", erro: str = ""):
    if _shared_auth_enabled():
        return RedirectResponse(url="/cadastros", status_code=303)
    return templates.TemplateResponse(
        request=request,
        name="admin.html",
        context={
            "request": request,
            "auth_status": _auth_status(),
            "save_via_bridge": bridge_store.save_via_bridge(),
            "bridge_status": bridge_store.status(limit=5),
            "sucesso": sucesso,
            "erro": erro,
            "active_page": "admin",
        },
    )


@app.post("/admin/credenciais")
async def admin_credentials_post(
    username: str = Form(...),
    password: str = Form(...),
    password_confirm: str = Form(...),
):
    if _shared_auth_enabled():
        return RedirectResponse(url="/cadastros", status_code=303)
    try:
        _save_admin_credentials(username, password, password_confirm)
        response = RedirectResponse(
            url=f"/admin?sucesso={quote('Credenciais do administrador atualizadas. Entre novamente com a nova senha.')}",
            status_code=303,
        )
        response.delete_cookie(SESSION_COOKIE)
        return response
    except Exception as exc:
        return RedirectResponse(url=f"/admin?erro={quote(str(exc))}", status_code=303)


@app.get("/suprimentos", response_class=HTMLResponse)
async def suprimentos_page(
    request: Request,
    sucesso: str = "",
    erro: str = "",
    editar_pessoa: int = 0,
):
    pessoa_edicao = {}
    try:
        pessoas = supabase_suprimentos.listar_pessoas()
        if editar_pessoa:
            pessoa_edicao = supabase_suprimentos.obter_pessoa(editar_pessoa) or {}
            if not pessoa_edicao:
                erro = erro or "Pessoa nao encontrada para edicao."
        processos = supabase_suprimentos.listar_processos()
        regras = supabase_suprimentos.listar_regras()
        relacoes = supabase_suprimentos.listar_relacoes()
    except Exception as exc:
        pessoas, processos, regras, relacoes = [], {}, [], {}
        erro = erro or str(exc)
    return templates.TemplateResponse(
        request=request,
        name="suprimentos.html",
        context={
            "request": request,
            "pessoas": pessoas,
            "pessoa_edicao": pessoa_edicao,
            "processos": processos,
            "regras": regras,
            "relacoes": relacoes,
            "sucesso": sucesso,
            "erro": erro,
            "active_page": "suprimentos",
        },
    )


@app.post("/suprimentos/pessoas")
async def suprimentos_pessoas_post(request: Request):
    try:
        form = await request.form()
        campos = (
            "data_registro", "nome_fantasia", "razao_social", "cnpj_cpf",
            "codigo_identificador_unico", "rg", "ie", "identificador",
            "logradouro", "logradouro_numero", "complemento", "bairro", "cidade",
            "codigo_municipio", "pais", "codigo_pais", "cep", "uf", "codigo_uf",
            "telefone", "whatsapp", "celular", "email", "site", "pessoa_grupo",
            "vendedor_padrao", "categoria", "tabela_preco", "observacoes",
            "limite_credito", "periodicidade_venda_compra_dias", "validation",
            "valor_minimo_compra", "data_nascimento_fundacao",
        )
        pessoa = {campo: str(form.get(campo, "") or "").strip() for campo in campos}
        for campo in ("pessoa_fisica", "cliente", "fornecedor", "colaborador", "transportadora"):
            pessoa[campo] = bool(form.get(campo))
        pessoa_id = int(str(form.get("pessoa_id", "0") or "0"))
        if pessoa_id:
            count = supabase_suprimentos.atualizar_pessoa(pessoa_id, pessoa)
            mensagem = f"{count} pessoa(s) atualizada(s)."
        else:
            count = supabase_suprimentos.salvar_pessoas([pessoa])
            mensagem = f"{count} pessoa(s) salva(s)."
        return RedirectResponse(url=f"/suprimentos?sucesso={quote(mensagem)}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/suprimentos?erro={quote(str(exc))}", status_code=303)


@app.post("/suprimentos/pessoas/upload")
async def suprimentos_pessoas_upload(arquivo_pessoas: UploadFile = File(...)):
    try:
        count = supabase_suprimentos.importar_pessoas_xlsx(await arquivo_pessoas.read())
        return RedirectResponse(url=f"/suprimentos?sucesso={quote(f'{count} pessoa(s) importada(s).')}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/suprimentos?erro={quote(str(exc))}", status_code=303)


@app.post("/suprimentos/processos/upload")
async def suprimentos_processos_upload(arquivo_processos: list[UploadFile] = File(...)):
    try:
        total = 0
        for arquivo in arquivo_processos:
            total += supabase_suprimentos.importar_processos_xlsx(
                await arquivo.read(),
                arquivo.filename or "",
            )
        return RedirectResponse(url=f"/suprimentos?sucesso={quote(f'{total} processo(s)/atividade(s) importado(s).')}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/suprimentos?erro={quote(str(exc))}", status_code=303)


@app.post("/suprimentos/regras")
async def suprimentos_regras_post(
    gatilho: str = Form(...),
    opcoes: str = Form(...),
    quantidade: str = Form("1"),
    quantidade_editavel: str = Form(""),
):
    try:
        regras = supabase_suprimentos.listar_regras()
        next_id = max([int("".join(ch for ch in regra["id"] if ch.isdigit()) or 0) for regra in regras] or [0]) + 1
        regras.append(
            {
                "id": f"regra-{next_id}",
                "gatilho": gatilho,
                "opcoes": [parte.strip() for parte in opcoes.replace("\n", ";").split(";") if parte.strip()],
                "quantidade": float((quantidade or "1").replace(",", ".")),
                "quantidade_editavel": bool(quantidade_editavel),
            }
        )
        count = supabase_suprimentos.salvar_regras(regras)
        return RedirectResponse(url=f"/suprimentos?sucesso={quote(f'{count} regra(s) salva(s).')}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/suprimentos?erro={quote(str(exc))}", status_code=303)


@app.post("/suprimentos/regras/upload")
async def suprimentos_regras_upload(arquivo_regras: UploadFile = File(...)):
    try:
        count = supabase_suprimentos.importar_regras_xlsx(await arquivo_regras.read())
        return RedirectResponse(url=f"/suprimentos?sucesso={quote(f'{count} regra(s) persistida(s).')}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/suprimentos?erro={quote(str(exc))}", status_code=303)


@app.post("/suprimentos/regras/excluir")
async def suprimentos_regras_excluir(rule_id: str = Form(...)):
    try:
        regras = [regra for regra in supabase_suprimentos.listar_regras() if regra["id"] != rule_id]
        supabase_suprimentos.salvar_regras(regras)
        return RedirectResponse(url=f"/suprimentos?sucesso={quote('Regra excluída.')}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/suprimentos?erro={quote(str(exc))}", status_code=303)


@app.post("/suprimentos/relacoes")
async def suprimentos_relacoes_post(item_codigo: str = Form(...), processos: str = Form(...)):
    try:
        relacoes = supabase_suprimentos.listar_relacoes()
        relacoes[item_codigo] = [parte.strip() for parte in processos.replace("\n", ";").split(";") if parte.strip()]
        count = supabase_suprimentos.salvar_relacoes(relacoes)
        return RedirectResponse(url=f"/suprimentos?sucesso={quote(f'{count} relação(ões) salva(s).')}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/suprimentos?erro={quote(str(exc))}", status_code=303)


@app.post("/suprimentos/relacoes/upload")
async def suprimentos_relacoes_upload(arquivo_relacoes: UploadFile = File(...)):
    try:
        count = supabase_suprimentos.importar_relacoes_xlsx(await arquivo_relacoes.read())
        return RedirectResponse(url=f"/suprimentos?sucesso={quote(f'{count} relação(ões) persistida(s).')}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/suprimentos?erro={quote(str(exc))}", status_code=303)


@app.post("/suprimentos/relacoes/excluir")
async def suprimentos_relacoes_excluir(item_codigo: str = Form(...)):
    try:
        relacoes = supabase_suprimentos.listar_relacoes()
        relacoes.pop(item_codigo, None)
        supabase_suprimentos.salvar_relacoes(relacoes)
        return RedirectResponse(url=f"/suprimentos?sucesso={quote('Relação excluída.')}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/suprimentos?erro={quote(str(exc))}", status_code=303)


@app.head("/")
async def home_head():
    return Response(status_code=200)


@app.get("/healthz")
async def healthz():
    return {
        "ok": True,
        "mode": supabase_store.save_mode() or bridge_store.save_mode(),
        "persistence": bridge_store.persistence_info(),
        "supabase": supabase_store.status(),
    }


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
            raise ValueError("O item foi definido com B.O.M. Inclua pelo menos um componente.")

        if _supabase_mode():
            result = supabase_store.save_registration(form_data)
            if draft_id:
                try:
                    supabase_store.delete_draft(draft_id)
                except Exception:
                    pass
            if needs_bom:
                bom_item_code = result.get("sku") or bom_item_code
                supabase_store.save_bom(
                    bom_item_code,
                    result.get("descricao_primaria") or bom_item_code,
                    components,
                    category_key=result.get("category_key") or category_key,
                    category_label=result.get("category") or "",
                    registration_id=result.get("id"),
                    source="cadastro",
                )
                message = f"Cadastro e B.O.M. salvos no Supabase. SKU: {result.get('sku') or '-'}."
            else:
                message = f"Cadastro salvo no Supabase. SKU: {result.get('sku') or '-'}."
            return RedirectResponse(
                url=f"/cadastro/bancos?categoria={quote(result['category_key'])}&sucesso={quote(message)}",
                status_code=303,
            )

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
        if _supabase_mode():
            result = supabase_store.save_draft(category_key, draft_payload, draft_id)
        elif bridge_store.save_via_bridge():
            result = _online_draft_payload(category_key, draft_payload, draft_id)
        else:
            result = excel_bancos.save_registration_draft(category_key, draft_payload, draft_id)
        return JSONResponse({"ok": True, "draft": result})
    except Exception as exc:
        return JSONResponse({"ok": False, "error": str(exc)}, status_code=400)


@app.post("/rascunhos/excluir")
async def rascunhos_excluir_post(draft_id: str = Form(...), category_key: str = Form("")):
    try:
        if _supabase_mode():
            result = supabase_store.delete_draft(draft_id)
        elif bridge_store.save_via_bridge():
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
        if _supabase_mode():
            result = supabase_store.delete_draft(draft_id)
        elif bridge_store.save_via_bridge():
            result = bridge_store.delete_draft(draft_id)
        else:
            result = excel_bancos.delete_registration_draft(draft_id)
        return JSONResponse({"ok": True, "draft": result})
    except Exception as exc:
        return JSONResponse({"ok": False, "error": str(exc)}, status_code=400)


@app.get("/api/produtos")
async def api_produtos(q: str = ""):
    if _supabase_mode():
        return {"items": supabase_store.search_products(q)}
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
    if _supabase_mode():
        raise HTTPException(status_code=410, detail="Ponte local desativada no modo Supabase.")
    if not bridge_store.verify_token(authorization):
        raise HTTPException(status_code=401, detail="Token da ponte inválido ou ausente.")


@app.get("/cadastros", response_class=HTMLResponse)
async def cadastros_page(
    request: Request,
    categoria: str = "",
    grupo: str = "",
    q: str = "",
    sem_unidade: str = "",
    mostrar_inativos: str = "",
    sucesso: str = "",
    erro: str = "",
):
    all_categories = supabase_store.all_categories_key(categoria)
    selected_category = (
        {"key": supabase_store.ALL_CATEGORIES_KEY, "label": "Todas as categorias"}
        if all_categories
        else excel_bancos.selected_category(categoria)
    )
    fields = [] if all_categories else excel_bancos.get_banco_fields_for_display(selected_category["key"])
    filters = {
        key[2:]: excel_bancos.clean_text(value)
        for key, value in request.query_params.items()
        if key.startswith("f_") and excel_bancos.clean_text(value)
    }
    items = []
    unit_pending_count = 0
    if _supabase_mode():
        only_missing_unit = excel_bancos.clean_text(sem_unidade) == "1"
        include_inactive = excel_bancos.clean_text(mostrar_inativos) == "1"
        items = supabase_store.list_registrations(
            selected_category["key"],
            query=q,
            filters=filters,
            group_code=grupo,
            missing_unit=only_missing_unit,
            include_inactive=include_inactive,
            limit=1000,
        )
        unit_pending_count = supabase_store.count_registrations_without_unit(
            selected_category["key"],
            include_inactive=include_inactive,
        )
        inactive_count = supabase_store.count_inactive_registrations(selected_category["key"])
    else:
        only_missing_unit = False
        include_inactive = False
        inactive_count = 0
    nav_category = excel_bancos.selected_category("")
    return templates.TemplateResponse(
        request=request,
        name="cadastros.html",
        context={
            "request": request,
            "categories": excel_bancos.list_categories(),
            "pn_groups": excel_bancos.list_pn_groups(),
            "selected_category": selected_category,
            "nav_category_key": nav_category["key"],
            "all_categories": all_categories,
            "grupo": excel_bancos._pn_group_code(grupo),
            "fields": fields,
            "items": items,
            "q": q,
            "filters": filters,
            "sem_unidade": "1" if only_missing_unit else "",
            "mostrar_inativos": "1" if include_inactive else "",
            "unit_pending_count": unit_pending_count,
            "inactive_count": inactive_count,
            "workbook_path": _workbook_display_path(),
            "save_via_bridge": bridge_store.save_via_bridge(),
            "supabase_mode": _supabase_mode(),
            "sucesso": sucesso,
            "erro": erro,
            "active_page": "cadastros",
        },
    )


@app.get("/cadastros/exportar")
async def cadastros_exportar(
    request: Request,
    categoria: str = "",
    grupo: str = "",
    q: str = "",
    sem_unidade: str = "",
    mostrar_inativos: str = "",
):
    if not _supabase_mode():
        raise HTTPException(status_code=400, detail="Exportação pela base está disponível no modo Supabase.")
    all_categories = supabase_store.all_categories_key(categoria)
    selected_category = (
        {"key": supabase_store.ALL_CATEGORIES_KEY, "label": "Todas as categorias"}
        if all_categories
        else excel_bancos.selected_category(categoria)
    )
    filters = {
        key[2:]: excel_bancos.clean_text(value)
        for key, value in request.query_params.items()
        if key.startswith("f_") and excel_bancos.clean_text(value)
    }
    output = supabase_store.export_registrations(
        selected_category["key"],
        query=q,
        filters=filters,
        group_code=grupo,
        missing_unit=excel_bancos.clean_text(sem_unidade) == "1",
        include_inactive=excel_bancos.clean_text(mostrar_inativos) == "1",
    )
    return FileResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output.name,
    )


@app.get("/bom", response_class=HTMLResponse)
async def bom_page(
    request: Request,
    categoria: str = "",
    item_pai: str = "",
    item_filho: str = "",
    revisao: str = "",
    sucesso: str = "",
    erro: str = "",
):
    if not _supabase_mode():
        return RedirectResponse(url="/cadastro/bancos", status_code=303)
    selected_category = excel_bancos.selected_category(categoria)
    items = supabase_store.list_boms(
        category_key=selected_category["key"] if categoria else "",
        parent_query=item_pai,
        component_query=item_filho,
        limit=1000,
    )
    review_count = sum(1 for item in items if item.get("needs_review"))
    if revisao:
        items = [item for item in items if item.get("needs_review")]
    return templates.TemplateResponse(
        request=request,
        name="bom.html",
        context={
            "request": request,
            "categories": excel_bancos.list_categories(),
            "selected_category": selected_category,
            "items": items,
            "item_pai": item_pai,
            "item_filho": item_filho,
            "revisao": revisao,
            "review_count": review_count,
            "workbook_path": _workbook_display_path(),
            "supabase_mode": True,
            "sucesso": sucesso,
            "erro": erro,
            "active_page": "bom",
        },
    )


@app.post("/bom/upload")
async def bom_upload(arquivo_bom: UploadFile = File(...)):
    if not _supabase_mode():
        raise HTTPException(status_code=400, detail="Upload de B.O.M. disponivel apenas no modo Supabase.")
    try:
        content = await arquivo_bom.read()
        result = supabase_store.import_bom_workbook(content, arquivo_bom.filename or "")
        message = f"B.O.M. importada: {result['parents']} item(ns) pai, {result['components']} componente(s), {result.get('review_parents', 0)} para revisao."
        return RedirectResponse(url=f"/bom?sucesso={quote(message)}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/bom?erro={quote(str(exc))}", status_code=303)


@app.post("/bom/copiar")
async def bom_copiar(
    source_parent_sku: str = Form(...),
    target_parent_sku: str = Form(...),
):
    if not _supabase_mode():
        raise HTTPException(status_code=400, detail="Copia de B.O.M. disponivel apenas no modo Supabase.")
    try:
        result = supabase_store.copy_bom(source_parent_sku, target_parent_sku)
        message = (
            f"B.O.M. copiada de {result['source_parent_sku']} para {result['target_parent_sku']}: "
            f"{result['components_count']} componente(s)."
        )
        return RedirectResponse(url=f"/bom?item_pai={quote(result['target_parent_sku'])}&sucesso={quote(message)}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/bom?erro={quote(str(exc))}", status_code=303)


@app.get("/bom/exportar")
async def bom_exportar(categoria: str = "", item_pai: str = "", item_filho: str = ""):
    if not _supabase_mode():
        raise HTTPException(status_code=400, detail="Exportacao de B.O.M. disponivel apenas no modo Supabase.")
    selected_category = excel_bancos.selected_category(categoria) if categoria else {"key": ""}
    output = supabase_store.export_boms(
        category_key=selected_category["key"] if categoria else "",
        parent_query=item_pai,
        component_query=item_filho,
    )
    return FileResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output.name,
    )


@app.get("/bom/{bom_id}", response_class=HTMLResponse)
async def bom_detalhe_page(request: Request, bom_id: int, sucesso: str = "", erro: str = ""):
    if not _supabase_mode():
        return RedirectResponse(url="/cadastro/bancos", status_code=303)
    try:
        bom = supabase_store.get_bom(bom_id)
        selected_category = excel_bancos.selected_category(excel_bancos.clean_text(bom.get("parent_category_key")))
        return templates.TemplateResponse(
            request=request,
            name="bom_detalhe.html",
            context={
                "request": request,
                "categories": excel_bancos.list_categories(),
                "selected_category": selected_category,
                "bom": bom,
                "workbook_path": _workbook_display_path(),
                "supabase_mode": True,
                "sucesso": sucesso,
                "erro": erro,
                "active_page": "bom",
            },
        )
    except Exception as exc:
        return RedirectResponse(url=f"/bom?erro={quote(str(exc))}", status_code=303)


@app.post("/bom/{bom_id}", response_class=HTMLResponse)
async def bom_detalhe_post(request: Request, bom_id: int):
    form_data = await request.form()
    try:
        components = excel_bancos.parse_component_lines(form_data, allow_incomplete=True)
        result = supabase_store.update_bom(
            bom_id,
            excel_bancos.clean_text(form_data.get("parent_descricao")),
            components,
            parent_sku=excel_bancos.clean_text(form_data.get("parent_sku")),
        )
        message = f"B.O.M. atualizada: {result.get('parent_sku') or bom_id}."
        return RedirectResponse(url=f"/bom/{bom_id}?sucesso={quote(message)}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/bom/{bom_id}?erro={quote(str(exc))}", status_code=303)


@app.post("/bom/{bom_id}/excluir")
async def bom_excluir(bom_id: int):
    if not _supabase_mode():
        return RedirectResponse(url="/cadastro/bancos", status_code=303)
    try:
        result = supabase_store.delete_bom(bom_id)
        message = f"B.O.M. excluida: {result.get('parent_sku') or bom_id}."
        return RedirectResponse(url=f"/bom?sucesso={quote(message)}", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/bom?erro={quote(str(exc))}", status_code=303)


@app.get("/cadastros/{registration_id}/editar", response_class=HTMLResponse)
async def cadastro_editar_page(request: Request, registration_id: int, sucesso: str = "", erro: str = ""):
    if not _supabase_mode():
        return RedirectResponse(url="/cadastros", status_code=303)
    try:
        editable = supabase_store.editable_registration(registration_id)
        fields = editable["fields"]
        groups = editable["groups"]
        category = {"key": editable["category"]["key"], "label": editable["category"]["label"]}
        return templates.TemplateResponse(
            request=request,
            name="editar_cadastro.html",
            context={
                "request": request,
                "record": editable["record"],
                "categories": excel_bancos.list_categories(),
                "selected_category": category,
                "fields": _enrich_fields(fields, groups),
                "ordered_fields": _enrich_fields(excel_bancos.get_banco_fields_for_display(category["key"]), groups),
                "conditional_rules": excel_bancos.get_conditional_rules_for_form(category["key"]),
                "workbook_path": _workbook_display_path(),
                "supabase_mode": True,
                "unit_options": supabase_store.unidade_options(),
                "sucesso": sucesso,
                "erro": erro,
                "active_page": "cadastros",
            },
        )
    except Exception as exc:
        return RedirectResponse(url=f"/cadastros?erro={quote(str(exc))}", status_code=303)


@app.post("/cadastros/{registration_id}/editar", response_class=HTMLResponse)
async def cadastro_editar_post(request: Request, registration_id: int):
    form_data = await request.form()
    try:
        result = supabase_store.update_registration(registration_id, form_data)
        message = f"Cadastro atualizado. SKU: {result.get('sku') or '-'}."
        return RedirectResponse(
            url=f"/cadastros?categoria={quote(result.get('category_key') or '')}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/cadastros/{registration_id}/editar?erro={quote(str(exc))}",
            status_code=303,
        )


@app.get("/ponte", response_class=HTMLResponse)
async def ponte_page(request: Request, sucesso: str = "", erro: str = ""):
    if _supabase_mode():
        return RedirectResponse(url="/cadastros", status_code=303)
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
    if _supabase_mode():
        return RedirectResponse(url="/cadastros", status_code=303)
    return FileResponse(
        _resource_file("local_bridge.py"),
        media_type="text/x-python",
        filename="ponte_local_modulo_cadastro.py",
    )


@app.get("/ponte/download-bat")
async def ponte_download_bat():
    if _supabase_mode():
        return RedirectResponse(url="/cadastros", status_code=303)
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


@app.post("/grupos/adicionar")
async def grupos_adicionar_post(
    category_key: str = Form(""),
    group_code: str = Form(...),
    group_label: str = Form(...),
    group_prefixes: str = Form(""),
):
    try:
        result = excel_bancos.add_pn_group(group_code, group_label, group_prefixes)
        message = f"Grupo criado: {result['code']} - {result['label']}."
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&erro={quote(str(exc))}",
            status_code=303,
        )


@app.post("/grupos/editar")
async def grupos_editar_post(
    category_key: str = Form(""),
    group_code: str = Form(...),
    group_label: str = Form(...),
    group_prefixes: str = Form(""),
):
    try:
        result = excel_bancos.update_pn_group(group_code, group_label, group_prefixes)
        message = f"Grupo atualizado: {result['code']} - {result['label']}."
        return RedirectResponse(
            url=f"/opcoes?categoria={quote(category_key)}&sucesso={quote(message)}",
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
    if _supabase_mode():
        return RedirectResponse(url="/cadastros", status_code=303)
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
    if _supabase_mode():
        return RedirectResponse(url="/cadastros", status_code=303)
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
