-- A função de proteção é usada apenas por trigger; não deve ser uma RPC pública.
alter function public.cadastro_auditoria_immutable()
    set search_path = public;

revoke execute on function public.cadastro_auditoria_immutable() from public, anon, authenticated;
