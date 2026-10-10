import asyncio
import json
import unittest
from types import SimpleNamespace
from unittest.mock import patch
from uuid import uuid4
from starlette.requests import Request
import cadastro_requests as workflow
import main
import supabase_store


class RegistrationRequestTests(unittest.TestCase):
    def test_commercial_can_consult_but_never_respond(self):
        for path in ["/cadastros/solicitacoes","/api/cadastro/solicitacoes","/api/cadastro/solicitacoes/notificacoes","/api/cadastro/solicitacoes/"+str(uuid4())]:
            self.assertTrue(main._commercial_route_allowed("GET",path))
            self.assertFalse(main._commercial_route_allowed("POST",path+"/responder"))

    def test_rpc_uses_trusted_actor_and_requires_response(self):
        rid=str(uuid4())
        with patch.object(supabase_store,"_request",return_value={"status":"CONCLUÍDA"}) as call:
            result=workflow.respond(rid,{"version":1,"action":"RESPONDER","sku":"10240089","actor":"FALSO","user_id":99},"PAULO",1)
            self.assertEqual(result["status"],"CONCLUÍDA")
            self.assertEqual(call.call_args.args[:2],("POST","rpc/cadastro_responder_solicitacao"))
            payload=call.call_args.args[3]
            self.assertEqual(payload["p_actor"],"PAULO")
            self.assertEqual(payload["p_user_id"],"1")
        for data in [
            {"version":1,"action":"RESPONDER"},{"version":True,"action":"ASSUMIR"},
            {"version":0,"action":"ASSUMIR"},{"version":1,"action":"DELETE"}]:
            with self.assertRaises(workflow.WorkflowError):
                workflow.respond(rid,data,"PAULO",1)

    def test_errors_do_not_leak_server_details(self):
        with patch.object(supabase_store,"_request",side_effect=supabase_store.SupabaseStoreError('Erro Supabase 500: SECRET')):
            with self.assertRaises(workflow.WorkflowError) as error:workflow.listing()
            self.assertEqual(error.exception.status,503)
            self.assertNotIn("SECRET",str(error.exception))
        with patch.object(supabase_store,"_request",side_effect=supabase_store.SupabaseStoreError('Erro Supabase 400: {"code":"P0001","message":"Solicitação alterada por outro usuário. Atualize a tela."}')):
            with self.assertRaises(workflow.WorkflowError) as error:workflow.listing()
            self.assertEqual(error.exception.status,409)

    def test_pagination_and_safe_search(self):
        with patch.object(supabase_store,"_request",return_value=[{"id":str(n)} for n in range(51)]) as call:
            data=workflow.listing('foo"),status.eq.CONCLUÍDA',"SOLICITADA",2)
            self.assertEqual(len(data["rows"]),50)
            self.assertTrue(data["has_next"])
            query=dict(call.call_args.args[2])
            self.assertEqual(query["offset"],"50")
            self.assertIn('foo\\\")',query["or"])

    def test_response_endpoint_enforces_permission_csrf_and_identity(self):
        rid=str(uuid4())
        async def run():
            payload={"version":1,"action":"RESPONDER","response":"Conferido"}
            body=json.dumps(payload).encode()
            async def receive():return {"type":"http.request","body":body,"more_body":False}
            request=Request({"type":"http","method":"POST","path":"/api/cadastro/solicitacoes/"+rid+"/responder",
                "headers":[(b"x-cadastro-csrf",b"token")],"scheme":"http","server":("testserver",80),
                "query_string":b""},receive)
            with patch.object(main,"_cadastro_write_allowed",return_value=False):
                result=await main.registration_request_response(request,rid)
                self.assertEqual(result.status_code,403)
            with patch.object(main,"_cadastro_write_allowed",return_value=True),patch.object(main,"_read_session_payload",return_value={"uid":1}),patch.object(main,"_request_workflow_csrf",return_value="other"):
                result=await main.registration_request_response(request,rid)
                self.assertEqual(result.status_code,403)
            with patch.object(main,"_cadastro_write_allowed",return_value=True),patch.object(main,"_read_session_payload",return_value={"uid":1}),patch.object(main,"_request_workflow_csrf",return_value="token"),patch.object(main,"_audit_actor",return_value="PAULO"),patch.object(main,"_audit_actor_user_id",return_value=1),patch.object(workflow,"respond",return_value={"status":"CONCLUÍDA"}) as respond:
                result=await main.registration_request_response(request,rid)
                self.assertEqual(result["status"],"CONCLUÍDA")
                self.assertEqual(respond.call_args.args[-2:],("PAULO",1))
        asyncio.run(run())

    def test_template_renders_for_reader_without_write_actions(self):
        request=Request({"type":"http","headers":[],"path":"/cadastros/solicitacoes"})
        html=main.templates.get_template("cadastro_solicitacoes.html").render(request=request,can_write=False,workflow_csrf="test")
        self.assertIn("const canWrite=false",html)
        self.assertIn("Solicitações de cadastro",html)

    def test_http_page_route_uses_current_template_response_signature(self):
        from fastapi.testclient import TestClient
        with patch.object(main,"_login_required",return_value=False),patch.object(main,"_read_session_payload",return_value={"uid":1}),patch.object(main,"_cadastro_write_allowed",return_value=False),patch.object(main,"_request_workflow_csrf",return_value="test"):
            response=TestClient(main.app).get("/cadastros/solicitacoes")
            self.assertEqual(response.status_code,200)
            self.assertIn("const canWrite=false",response.text)

    def test_queue_requires_signed_session_even_when_legacy_login_is_disabled(self):
        from fastapi.testclient import TestClient
        with patch.object(main,"_login_required",return_value=False),patch.object(main,"_read_session_payload",return_value={}):
            client=TestClient(main.app)
            for path in ["/api/cadastro/solicitacoes","/api/cadastro/solicitacoes/notificacoes","/api/cadastro/solicitacoes/"+str(uuid4())]:
                self.assertEqual(client.get(path).status_code,401)
            response=client.get("/cadastros/solicitacoes",follow_redirects=False)
            self.assertEqual(response.status_code,303)
