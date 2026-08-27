begin;

select 1
from public.cadastro_catalogo
where config_key = 'default'
for update;

with field_location as (
    select
        category_position - 1 as category_index,
        field_position - 1 as field_index
    from public.cadastro_catalogo catalog
    cross join lateral jsonb_array_elements(catalog.payload -> 'categories')
        with ordinality as category(category_data, category_position)
    cross join lateral jsonb_array_elements(category.category_data -> 'fields')
        with ordinality as field(field_data, field_position)
    where catalog.config_key = 'default'
      and category.category_data ->> 'key' = 'bancos'
      and field.field_data ->> 'key' = 'especificidade'
)
update public.cadastro_catalogo catalog
set payload = jsonb_set(
        jsonb_set(
            catalog.payload,
            array[
                'categories', field_location.category_index::text,
                'fields', field_location.field_index::text,
                'options'
            ],
            '["1- NORMAL","2- (2REC / 1 REB)","3- BJD","4- TRILHO","5- FOCA","6- PME","7- ELEVITTA","8- SPRINTER","9- MASTER","10- ESCOLAR","11- E/S/J","12- (2FIX / 1REB)","13- TRAS CX RODAS","14- PME 1A","15- PME 2A","16- PME 3A","17- EXECUTIVO","18- 4L REC","19- MASTER - PME"]'::jsonb,
            false
        ),
        array[
            'categories', field_location.category_index::text,
            'fields', field_location.field_index::text,
            'conjunto_only_options'
        ],
        '["14- PME 1A","15- PME 2A","16- PME 3A","17- EXECUTIVO","18- 4L REC","19- MASTER - PME"]'::jsonb,
        true
    ),
    updated_at = now()
from field_location
where catalog.config_key = 'default';

with canonical_values as (
    select
        registration.id,
        jsonb_agg(to_jsonb(
            case selected.value
                when 'PME 1A' then '14- PME 1A'
                when 'PME 2A' then '15- PME 2A'
                when 'PME 3A' then '16- PME 3A'
                when 'EXECUTIVO' then '17- EXECUTIVO'
                when '4L REC' then '18- 4L REC'
                when 'MASTER - PME' then '19- MASTER - PME'
                else selected.value
            end
        ) order by selected.position) as form_value,
        string_agg(
            case selected.value
                when 'PME 1A' then '14- PME 1A'
                when 'PME 2A' then '15- PME 2A'
                when 'PME 3A' then '16- PME 3A'
                when 'EXECUTIVO' then '17- EXECUTIVO'
                when '4L REC' then '18- 4L REC'
                when 'MASTER - PME' then '19- MASTER - PME'
                else selected.value
            end,
            ' | ' order by selected.position
        ) as field_value,
        string_agg(
            btrim(split_part(
                case selected.value
                    when 'PME 1A' then '14- PME 1A'
                    when 'PME 2A' then '15- PME 2A'
                    when 'PME 3A' then '16- PME 3A'
                    when 'EXECUTIVO' then '17- EXECUTIVO'
                    when '4L REC' then '18- 4L REC'
                    when 'MASTER - PME' then '19- MASTER - PME'
                    else selected.value
                end,
                '-', 1
            )),
            ' | ' order by selected.position
        ) as field_code
    from public.cadastro_registros registration
    cross join lateral jsonb_array_elements_text(
        coalesce(registration.form_values -> 'especificidade', '[]'::jsonb)
    ) with ordinality as selected(value, position)
    where registration.category_key = 'bancos'
      and exists (
          select 1
          from jsonb_array_elements_text(
              coalesce(registration.form_values -> 'especificidade', '[]'::jsonb)
          ) current_value(value)
          where current_value.value in (
              'PME 1A', 'PME 2A', 'PME 3A',
              'EXECUTIVO', '4L REC', 'MASTER - PME'
          )
      )
    group by registration.id
)
update public.cadastro_registros registration
set form_values = jsonb_set(
        coalesce(nullif(registration.form_values, 'null'::jsonb), '{}'::jsonb),
        '{especificidade}', canonical_values.form_value, true
    ),
    field_values = jsonb_set(
        coalesce(nullif(registration.field_values, 'null'::jsonb), '{}'::jsonb),
        '{especificidade}', to_jsonb(canonical_values.field_value), true
    ),
    field_codes = jsonb_set(
        coalesce(nullif(registration.field_codes, 'null'::jsonb), '{}'::jsonb),
        '{especificidade}', to_jsonb(canonical_values.field_code), true
    ),
    updated_at = now()
from canonical_values
where registration.id = canonical_values.id;

insert into public.erp_audit_events (
    id, entity_type, entity_id, action, actor, origin,
    before_data, after_data, reason, created_at
)
select
    gen_random_uuid(),
    'CADASTRO_CATALOGO',
    null,
    'RECONCILIAR_ESPECIFICIDADE_BANCOS_20260827',
    'sistema:correcao-controlada',
    'MODULO_CADASTRO',
    jsonb_build_object(
        'catalogo_terminava_em', '13- TRAS CX RODAS',
        'valores_sem_codigo', jsonb_build_array('PME 2A', 'EXECUTIVO', 'MASTER - PME')
    ),
    jsonb_build_object(
        'catalogo_termina_em', '19- MASTER - PME',
        'novas_opcoes', jsonb_build_array(
            '14- PME 1A', '15- PME 2A', '16- PME 3A',
            '17- EXECUTIVO', '18- 4L REC', '19- MASTER - PME'
        )
    ),
    'Reconciliação idempotente do catálogo fechado de especificidades de Bancos.',
    now()
where not exists (
    select 1
    from public.erp_audit_events
    where action = 'RECONCILIAR_ESPECIFICIDADE_BANCOS_20260827'
);

commit;
