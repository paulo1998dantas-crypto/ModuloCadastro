"""Shared Commercial queue via private PostgREST RPC. No master-data mutation."""
import json
import uuid
import supabase_store as store


class WorkflowError(ValueError):
    def __init__(self,message,status=400):
        super().__init__(message);self.status=status


def call(method,table,query=None,payload=None):
    try:return store._request(method,table,query,payload)
    except store.SupabaseStoreError as exc:
        message=str(exc);start=message.find("{")
        try:data=json.loads(message[start:]) if start>=0 else {}
        except ValueError:data={}
        if data.get("code")=="P0001":
            text=data.get("message","Solicitação inválida.")
            status=409 if any(word in text for word in ["alterada","encerrada","assumida"]) else 400
            raise WorkflowError(text,status) from exc
        raise WorkflowError("Fila indisponível. Confirme a publicação da migração 202610100008 no Comercial e a conexão com o Supabase.",503) from exc


def request_id(value):
    try:return str(uuid.UUID(str(value)))
    except (ValueError,TypeError):raise WorkflowError("Solicitação inválida.")


def listing(q="",status="",page=1):
    if status and status not in {"SOLICITADA","EM ANÁLISE","CONCLUÍDA","CANCELADA"}:raise WorkflowError("Status inválido.")
    try:page=int(page)
    except (ValueError,TypeError):raise WorkflowError("Página inválida.")
    if page<1:raise WorkflowError("Página inválida.")
    params=[("select","*"),("order","created_at.desc,id"),("limit","51"),("offset",str((page-1)*50))]
    if status:params.append(("status","eq."+status))
    # Quoted PostgREST patterns: punctuation cannot inject another filter.
    q=str(q or "").strip()[:100]
    if q:
        pattern=json.dumps("*"+q.replace("*","").replace("%","")+"*",ensure_ascii=False)
        params.append(("or","("+",".join(k+".ilike."+pattern for k in ["sku","specification","created_by","result_sku","response"])+")"))
    rows=call("GET","cadastro_solicitacoes",params) or []
    return {"rows":rows[:50],"has_next":len(rows)>50,"page":page}


def pending_count():
    rows=call("GET","cadastro_solicitacoes",[("select","id"),("status","in.(SOLICITADA,EM ANÁLISE)"),("limit","1000")]) or []
    return {"pending":len(rows),"capped":len(rows)>=1000}


def detail(value):
    rid=request_id(value)
    rows=call("GET","cadastro_solicitacoes",[("select","*"),("id","eq."+rid),("limit","1")]) or []
    if not rows:raise WorkflowError("Solicitação não encontrada.",404)
    events=call("GET","cadastro_solicitacao_eventos",[("select","*"),("request_id","eq."+rid),("order","created_at.asc,id")]) or []
    return {**rows[0],"events":events}


def respond(value,data,actor,user_id):
    rid=request_id(value);version=data.get("version")
    if not isinstance(version,int) or isinstance(version,bool) or version<1:raise WorkflowError("Atualize a solicitação antes de responder.",409)
    action=data.get("action")
    if action not in {"ASSUMIR","RESPONDER"}:raise WorkflowError("Ação inválida.")
    response=str(data.get("response") or "").strip();sku=str(data.get("sku") or "").strip()
    if len(response)>6000 or len(sku)>80:raise WorkflowError("Devolutiva ou SKU excede o limite.")
    if action=="RESPONDER" and not response and not sku:raise WorkflowError("Informe uma devolutiva ou um SKU para concluir.")
    if not actor or not user_id:raise WorkflowError("Sessão sem usuário de auditoria. Entre novamente.",401)
    return call("POST","rpc/cadastro_responder_solicitacao",payload={"p_id":rid,"p_version":version,
        "p_action":action,"p_response":response,"p_sku":sku,"p_actor":actor,"p_user_id":str(user_id)})
