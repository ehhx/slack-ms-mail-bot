-- Server-side state for the Deno mail bridge. The key encoding mirrors the
-- existing Deno KV keys so migration can be cursor-based and idempotent.
create table if not exists public.mail_persistent_state (
  state_key text primary key,
  state_value jsonb not null,
  state_version bigint not null default 1,
  expires_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint mail_persistent_state_version_positive check (state_version > 0)
);

create index if not exists mail_persistent_state_expires_at_idx
  on public.mail_persistent_state (expires_at)
  where expires_at is not null;

create index if not exists mail_persistent_state_key_prefix_idx
  on public.mail_persistent_state (state_key text_pattern_ops);

alter table public.mail_persistent_state enable row level security;

revoke all on table public.mail_persistent_state from public, anon, authenticated;
grant select, insert, update, delete on table public.mail_persistent_state
  to service_role;
grant usage on schema public to service_role;

create or replace function public.mail_state_get(p_state_key text)
returns jsonb
language sql
security invoker
set search_path = ''
as $$
  select coalesce(
    (
      select jsonb_build_object(
        'value', state_value,
        'version', state_version
      )
      from public.mail_persistent_state
      where state_key = p_state_key
        and (expires_at is null or expires_at > now())
      limit 1
    ),
    jsonb_build_object('value', null, 'version', null)
  );
$$;

create or replace function public.mail_state_get_many(p_state_keys text[])
returns table (
  state_key text,
  state_value jsonb,
  state_version bigint
)
language sql
security invoker
set search_path = ''
as $$
  select requested.state_key, stored.state_value, stored.state_version
  from unnest(p_state_keys) with ordinality as requested(state_key, position)
  left join public.mail_persistent_state as stored
    on stored.state_key = requested.state_key
   and (stored.expires_at is null or stored.expires_at > now())
  order by requested.position;
$$;

create or replace function public.mail_state_list(
  p_prefix text,
  p_limit integer default 1000,
  p_after text default null
)
returns table (
  state_key text,
  state_value jsonb,
  state_version bigint
)
language sql
security invoker
set search_path = ''
as $$
  select state_key, state_value, state_version
  from public.mail_persistent_state
  where (
      state_key = p_prefix
      or state_key like p_prefix || '/%'
    )
    and (expires_at is null or expires_at > now())
    and (p_after is null or state_key > p_after)
  order by state_key
  limit greatest(1, least(coalesce(p_limit, 1000), 1000));
$$;

create or replace function public.mail_state_set(
  p_state_key text,
  p_state_value jsonb,
  p_expires_at timestamptz default null
)
returns jsonb
language plpgsql
security invoker
set search_path = ''
as $$
declare
  result jsonb;
begin
  delete from public.mail_persistent_state
  where state_key = p_state_key
    and expires_at is not null
    and expires_at <= now();

  if p_expires_at is not null and p_expires_at <= now() then
    delete from public.mail_persistent_state where state_key = p_state_key;
    return jsonb_build_object('value', null, 'version', null);
  end if;

  insert into public.mail_persistent_state (
    state_key,
    state_value,
    state_version,
    expires_at,
    created_at,
    updated_at
  )
  values (p_state_key, p_state_value, 1, p_expires_at, now(), now())
  on conflict (state_key) do update
    set state_value = excluded.state_value,
        state_version = public.mail_persistent_state.state_version + 1,
        expires_at = excluded.expires_at,
        updated_at = now()
  returning jsonb_build_object(
    'value', state_value,
    'version', state_version
  ) into result;

  return result;
end;
$$;

create or replace function public.mail_state_set_if_absent(
  p_state_key text,
  p_state_value jsonb,
  p_expires_at timestamptz default null
)
returns boolean
language plpgsql
security invoker
set search_path = ''
as $$
declare
  inserted_count integer;
begin
  delete from public.mail_persistent_state
  where state_key = p_state_key
    and expires_at is not null
    and expires_at <= now();

  if p_expires_at is not null and p_expires_at <= now() then
    return false;
  end if;

  insert into public.mail_persistent_state (
    state_key,
    state_value,
    state_version,
    expires_at,
    created_at,
    updated_at
  )
  values (p_state_key, p_state_value, 1, p_expires_at, now(), now())
  on conflict (state_key) do nothing;

  get diagnostics inserted_count = row_count;
  return inserted_count = 1;
end;
$$;

create or replace function public.mail_state_delete(p_state_key text)
returns boolean
language plpgsql
security invoker
set search_path = ''
as $$
declare
  deleted_count integer;
begin
  delete from public.mail_persistent_state where state_key = p_state_key;
  get diagnostics deleted_count = row_count;
  return deleted_count = 1;
end;
$$;

drop function if exists public.mail_state_atomic_batch(jsonb);

create or replace function public.mail_state_prune_expired(
  p_limit integer default 100
)
returns integer
language plpgsql
security invoker
set search_path = ''
as $$
declare
  deleted_count integer;
begin
  with expired as (
    select ctid
    from public.mail_persistent_state
    where expires_at is not null
      and expires_at <= now()
    order by expires_at
    limit greatest(1, least(coalesce(p_limit, 100), 1000))
  )
  delete from public.mail_persistent_state
  where ctid in (select ctid from expired);

  get diagnostics deleted_count = row_count;
  return deleted_count;
end;
$$;

revoke execute on function public.mail_state_get(text)
  from public, anon, authenticated;
revoke execute on function public.mail_state_get_many(text[])
  from public, anon, authenticated;
revoke execute on function public.mail_state_list(text, integer, text)
  from public, anon, authenticated;
revoke execute on function public.mail_state_set(text, jsonb, timestamptz)
  from public, anon, authenticated;
revoke execute on function public.mail_state_set_if_absent(text, jsonb, timestamptz)
  from public, anon, authenticated;
revoke execute on function public.mail_state_delete(text)
  from public, anon, authenticated;
revoke execute on function public.mail_state_prune_expired(integer)
  from public, anon, authenticated;

grant execute on function public.mail_state_get(text) to service_role;
grant execute on function public.mail_state_get_many(text[]) to service_role;
grant execute on function public.mail_state_list(text, integer, text)
  to service_role;
grant execute on function public.mail_state_set(text, jsonb, timestamptz)
  to service_role;
grant execute on function public.mail_state_set_if_absent(text, jsonb, timestamptz)
  to service_role;
grant execute on function public.mail_state_delete(text) to service_role;
grant execute on function public.mail_state_prune_expired(integer)
  to service_role;
