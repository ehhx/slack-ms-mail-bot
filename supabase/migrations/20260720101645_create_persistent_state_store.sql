-- Supabase stores the durable state that previously lived in Deno KV.  The
-- application only reaches this table through narrowly scoped RPC functions;
-- no browser or anonymous client can read task payloads or lock metadata.
create table public.bot_persistent_state (
  state_key text primary key,
  state_value jsonb not null,
  state_version bigint not null default 1,
  expires_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint bot_persistent_state_version_positive check (state_version > 0)
);
create index bot_persistent_state_expires_at_idx
  on public.bot_persistent_state (expires_at)
  where expires_at is not null;
create index bot_persistent_state_deferred_schedule_idx
  on public.bot_persistent_state (
    state_key,
    (state_value ->> 'nextRetryAt')
  );
alter table public.bot_persistent_state enable row level security;
revoke all on table public.bot_persistent_state from public, anon, authenticated;
grant select, insert, update, delete on table public.bot_persistent_state to service_role;
-- A normal read treats expired state as absent. Expired rows are also deleted
-- by every write path, so TTL-based locks and de-duplication keys converge
-- without a background full-table scan.
create or replace function public.bot_state_get(p_state_key text)
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
      from public.bot_persistent_state
      where state_key = p_state_key
        and (expires_at is null or expires_at > now())
      limit 1
    ),
    jsonb_build_object('value', null, 'version', null)
  );
$$;
create or replace function public.bot_state_list(
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
  from public.bot_persistent_state
  where left(state_key, char_length(p_prefix)) = p_prefix
    and (
      char_length(state_key) = char_length(p_prefix)
      or substr(state_key, char_length(p_prefix) + 1, 1) = '/'
    )
    and (expires_at is null or expires_at > now())
    and (p_after is null or state_key > p_after)
  order by state_key
  limit greatest(1, least(coalesce(p_limit, 1000), 1000));
$$;
create or replace function public.bot_state_set(
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
  delete from public.bot_persistent_state
  where state_key = p_state_key
    and expires_at is not null
    and expires_at <= now();

  if p_expires_at is not null and p_expires_at <= now() then
    delete from public.bot_persistent_state where state_key = p_state_key;
    return jsonb_build_object('value', null, 'version', null);
  end if;

  insert into public.bot_persistent_state (
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
        state_version = public.bot_persistent_state.state_version + 1,
        expires_at = excluded.expires_at,
        updated_at = now()
  returning jsonb_build_object(
    'value', state_value,
    'version', state_version
  ) into result;

  return result;
end;
$$;
create or replace function public.bot_state_delete(p_state_key text)
returns boolean
language plpgsql
security invoker
set search_path = ''
as $$
declare
  deleted_count integer;
begin
  delete from public.bot_persistent_state where state_key = p_state_key;
  get diagnostics deleted_count = row_count;
  return deleted_count = 1;
end;
$$;
-- Deferred Bitable records have a timestamp in their JSON payload. Querying it
-- in Postgres instead of listing the first N keys prevents a large backlog
-- from starving records whose client token sorts later in the keyspace.
create or replace function public.bot_state_list_due_deferred(
  p_prefix text,
  p_now text,
  p_limit integer default 100
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
  from public.bot_persistent_state
  where left(state_key, char_length(p_prefix)) = p_prefix
    and substr(state_key, char_length(p_prefix) + 1, 1) = '/'
    and (expires_at is null or expires_at > now())
    and coalesce(state_value ->> 'nextRetryAt', '') <= p_now
  order by state_value ->> 'nextRetryAt', state_key
  limit greatest(1, least(coalesce(p_limit, 100), 1000));
$$;
-- TTL rows are invisible to reads immediately. Delete a small indexed batch on
-- each worker pass so expired values do not accumulate in table storage.
create or replace function public.bot_state_prune_expired(
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
    from public.bot_persistent_state
    where expires_at is not null
      and expires_at <= now()
    order by expires_at
    limit greatest(1, least(coalesce(p_limit, 100), 1000))
  )
  delete from public.bot_persistent_state
  where ctid in (select ctid from expired);

  get diagnostics deleted_count = row_count;
  return deleted_count;
end;
$$;
-- This is the compare-and-set primitive formerly supplied by Deno KV atomic
-- checks. It protects queue claims and leases when Deno Deploy runs more than
-- one instance at the same time.
create or replace function public.bot_state_compare_and_write(
  p_state_key text,
  p_expected_version bigint,
  p_state_value jsonb default null,
  p_expires_at timestamptz default null,
  p_delete boolean default false
)
returns boolean
language plpgsql
security invoker
set search_path = ''
as $$
declare
  changed_count integer;
begin
  delete from public.bot_persistent_state
  where state_key = p_state_key
    and expires_at is not null
    and expires_at <= now();

  if p_expected_version is null then
    if p_delete then
      return true;
    end if;

    if p_expires_at is not null and p_expires_at <= now() then
      return true;
    end if;

    insert into public.bot_persistent_state (
      state_key,
      state_value,
      state_version,
      expires_at,
      created_at,
      updated_at
    )
    values (p_state_key, p_state_value, 1, p_expires_at, now(), now())
    on conflict (state_key) do nothing;

    get diagnostics changed_count = row_count;
    return changed_count = 1;
  end if;

  if p_delete then
    delete from public.bot_persistent_state
    where state_key = p_state_key
      and state_version = p_expected_version;
  elsif p_expires_at is not null and p_expires_at <= now() then
    delete from public.bot_persistent_state
    where state_key = p_state_key
      and state_version = p_expected_version;
  else
    update public.bot_persistent_state
    set state_value = p_state_value,
        state_version = state_version + 1,
        expires_at = p_expires_at,
        updated_at = now()
    where state_key = p_state_key
      and state_version = p_expected_version;
  end if;

  get diagnostics changed_count = row_count;
  return changed_count = 1;
end;
$$;
revoke execute on function public.bot_state_get(text) from public, anon, authenticated;
revoke execute on function public.bot_state_list(text, integer, text) from public, anon, authenticated;
revoke execute on function public.bot_state_set(text, jsonb, timestamptz) from public, anon, authenticated;
revoke execute on function public.bot_state_delete(text) from public, anon, authenticated;
revoke execute on function public.bot_state_compare_and_write(text, bigint, jsonb, timestamptz, boolean) from public, anon, authenticated;
revoke execute on function public.bot_state_list_due_deferred(text, text, integer) from public, anon, authenticated;
revoke execute on function public.bot_state_prune_expired(integer) from public, anon, authenticated;
grant execute on function public.bot_state_get(text) to service_role;
grant execute on function public.bot_state_list(text, integer, text) to service_role;
grant execute on function public.bot_state_set(text, jsonb, timestamptz) to service_role;
grant execute on function public.bot_state_delete(text) to service_role;
grant execute on function public.bot_state_compare_and_write(text, bigint, jsonb, timestamptz, boolean) to service_role;
grant execute on function public.bot_state_list_due_deferred(text, text, integer) to service_role;
grant execute on function public.bot_state_prune_expired(integer) to service_role;
