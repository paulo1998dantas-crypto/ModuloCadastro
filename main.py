import socket
import sys
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

import excel_bancos


HOST = "127.0.0.1"
PORT = 8001
app = FastAPI(title="Módulo de Cadastro")
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


def _render_cadastro_page(request: Request, categoria: str = "", sucesso: str = "", erro: str = "", form_data=None):
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
            "workbook_path": excel_bancos.template_path(),
            "sucesso": sucesso,
            "erro": erro,
            "form_data": normalized_form,
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
            "workbook_path": excel_bancos.template_path(),
            "sucesso": sucesso,
            "erro": erro,
            "active_page": "opcoes",
        },
    )


@app.get("/", response_class=HTMLResponse)
async def home():
    return RedirectResponse(url="/cadastro/bancos", status_code=303)


@app.get("/cadastro/bancos", response_class=HTMLResponse)
async def cadastro_bancos_page(request: Request, categoria: str = "", sucesso: str = "", erro: str = ""):
    return _render_cadastro_page(request, categoria=categoria, sucesso=sucesso, erro=erro)


@app.post("/cadastro/bancos", response_class=HTMLResponse)
async def cadastro_bancos_post(request: Request):
    form_data = await request.form()
    category_key = excel_bancos.clean_text(form_data.get("categoria"))
    try:
        result = excel_bancos.save_banco_registration(form_data)
        message = f"Cadastro salvo na linha {result['row']}."
        return RedirectResponse(
            url=f"/cadastro/bancos?categoria={quote(result['category_key'])}&sucesso={quote(message)}",
            status_code=303,
        )
    except Exception as exc:
        return _render_cadastro_page(request, categoria=category_key, erro=str(exc), form_data=form_data)


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
            "workbooks": excel_bancos.list_workbooks(),
            "folders": excel_bancos.list_folders(),
            "workbook_path": excel_bancos.template_path(),
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
):
    try:
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
