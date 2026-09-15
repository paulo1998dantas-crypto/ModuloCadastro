-- Autoria dos cadastros e dos eventos de criação.
-- Os campos são nulos de propósito: não existe evidência confiável para
-- atribuir um usuário aos registros criados antes desta correção.

alter table public.cadastro_registros
    add column if not exists created_by text null,
    add column if not exists created_by_user_id integer null;

alter table public.cadastro_auditoria
    add column if not exists actor_user_id integer null;

do $$
begin
    if not exists (
        select 1
        from pg_constraint
        where conname = 'cadastro_registros_created_by_user_fk'
          and conrelid = 'public.cadastro_registros'::regclass
    ) then
        alter table public.cadastro_registros
            add constraint cadastro_registros_created_by_user_fk
            foreign key (created_by_user_id)
            references public.users(id)
            on delete set null;
    end if;

    if not exists (
        select 1
        from pg_constraint
        where conname = 'cadastro_auditoria_actor_user_fk'
          and conrelid = 'public.cadastro_auditoria'::regclass
    ) then
        alter table public.cadastro_auditoria
            add constraint cadastro_auditoria_actor_user_fk
            foreign key (actor_user_id)
            references public.users(id)
            on delete set null;
    end if;
end;
$$;

create index if not exists cadastro_registros_created_by_idx
    on public.cadastro_registros (created_by, created_at desc);

create index if not exists cadastro_registros_created_by_user_idx
    on public.cadastro_registros (created_by_user_id, created_at desc);

create index if not exists cadastro_auditoria_actor_user_idx
    on public.cadastro_auditoria (actor_user_id, created_at desc);

-- Reconciliação conservadora: só preenche autoria quando já houver um evento
-- de criação com ator operacional. Eventos marcados como migração/sistema não
-- são convertidos em usuário, pois não identificam quem criou o cadastro.
update public.cadastro_registros r
set
    created_by = nullif(trim(a.actor), ''),
    created_by_user_id = u.id
from public.cadastro_auditoria a
left join public.users u
    on lower(trim(u.username)) = lower(trim(a.actor))
where a.registration_id = r.id
  and a.action = 'criacao'
  and nullif(trim(a.actor), '') is not null
  and lower(trim(a.actor)) not in (
      'migracao',
      'sistema',
      'sistema:cadastro',
      'cadastro:opcoes'
  )
  and r.created_by is null;
