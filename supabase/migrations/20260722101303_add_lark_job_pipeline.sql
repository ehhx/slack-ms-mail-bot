begin;
-- One row represents one Lark message. Lark documents message_id as the duplicate-delivery key;
-- event_id is retained only for diagnostics because retries may use a different event ID. The
-- related media_jobs rows split parsing, transfer, and reply work into independently leased units.
create table public.lark_requests (
  id uuid primary key default gen_random_uuid(),
  event_id text not null check (char_length(event_id) between 1 and 160),
  message_id text not null unique check (char_length(message_id) between 1 and 200),
  chat_id text not null check (char_length(chat_id) between 1 and 200),
  sender_open_id text,
  source_text text not null check (char_length(source_text) between 1 and 20000),
  link_count integer not null check (link_count between 1 and 20),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
alter table public.lark_requests enable row level security;
revoke all on table public.lark_requests from public, anon, authenticated, service_role;
grant select, insert, update on table public.lark_requests to service_role;
alter table public.media_jobs
  add column job_type text not null default 'media_transfer'
    check (job_type in ('media_transfer', 'lark_parse', 'lark_reply')),
  add column lark_request_id uuid references public.lark_requests(id),
  add column parent_job_id uuid references public.media_jobs(id),
  add column item_index integer check (item_index is null or item_index >= 0);
create index media_jobs_lark_request_idx
  on public.media_jobs (lark_request_id, job_type, status, item_index)
  where lark_request_id is not null;
create index media_jobs_parent_idx
  on public.media_jobs (parent_job_id, item_index)
  where parent_job_id is not null;
-- This helper is called only from the service-role RPCs below. The unique source/job key makes
-- scheduling idempotent even when two terminal transitions race or a caller retries an RPC.
create function public.enqueue_lark_reply_if_ready(p_request_id uuid)
returns void
language plpgsql
security invoker
set search_path = pg_catalog, public, pg_temp
as $$
declare
  v_message_id text;
begin
  if p_request_id is null then
    return;
  end if;

  if exists (
    select 1
    from public.media_jobs
    where lark_request_id = p_request_id
      and job_type <> 'lark_reply'
      and status in ('queued', 'processing')
  ) then
    return;
  end if;

  select message_id
  into v_message_id
  from public.lark_requests
  where id = p_request_id;

  if not found then
    return;
  end if;

  insert into public.media_jobs (
    source,
    source_event_id,
    payload,
    job_type,
    lark_request_id,
    max_attempts
  )
  values (
    'lark',
    p_request_id::text || ':reply',
    jsonb_build_object('message_id', v_message_id),
    'lark_reply',
    p_request_id,
    3
  )
  on conflict (source, source_event_id) do nothing;
end;
$$;
-- The ingress writes the request and every per-link parse job in one transaction. A duplicate Lark
-- delivery returns the existing request without resetting jobs that are already running or done.
create function public.enqueue_lark_request(
  p_event_id text,
  p_message_id text,
  p_chat_id text,
  p_sender_open_id text,
  p_source_text text,
  p_links jsonb
)
returns jsonb
language plpgsql
security invoker
set search_path = pg_catalog, public, pg_temp
as $$
declare
  v_request_id uuid;
  v_inserted boolean := false;
  v_link_count integer;
  v_item record;
  v_link text;
begin
  if nullif(btrim(p_event_id), '') is null or char_length(btrim(p_event_id)) > 160 then
    raise exception 'event_id must be between 1 and 160 characters';
  end if;
  if nullif(btrim(p_message_id), '') is null or char_length(btrim(p_message_id)) > 200 then
    raise exception 'message_id must be between 1 and 200 characters';
  end if;
  if nullif(btrim(p_chat_id), '') is null or char_length(btrim(p_chat_id)) > 200 then
    raise exception 'chat_id must be between 1 and 200 characters';
  end if;
  if nullif(btrim(p_source_text), '') is null or char_length(btrim(p_source_text)) > 20000 then
    raise exception 'source_text must be between 1 and 20000 characters';
  end if;
  if p_sender_open_id is not null and char_length(btrim(p_sender_open_id)) > 200 then
    raise exception 'sender_open_id cannot exceed 200 characters';
  end if;
  if jsonb_typeof(p_links) <> 'array' then
    raise exception 'links must be a JSON array';
  end if;

  v_link_count := jsonb_array_length(p_links);
  if v_link_count < 1 or v_link_count > 20 then
    raise exception 'links must contain between 1 and 20 items';
  end if;

  insert into public.lark_requests (
    event_id,
    message_id,
    chat_id,
    sender_open_id,
    source_text,
    link_count
  )
  values (
    btrim(p_event_id),
    btrim(p_message_id),
    btrim(p_chat_id),
    nullif(btrim(p_sender_open_id), ''),
    btrim(p_source_text),
    v_link_count
  )
  on conflict (message_id) do nothing
  returning id into v_request_id;

  if found then
    v_inserted := true;
  else
    select id into v_request_id
    from public.lark_requests
    where message_id = btrim(p_message_id);
  end if;

  for v_item in
    select value, ordinality - 1 as item_index
    from jsonb_array_elements(p_links) with ordinality
  loop
    if jsonb_typeof(v_item.value) <> 'object' then
      raise exception 'each link must be a JSON object';
    end if;

    v_link := nullif(btrim(v_item.value ->> 'url'), '');
    if v_link is null or char_length(v_link) > 4000 or left(v_link, 8) <> 'https://' then
      raise exception 'each link URL must be an HTTPS URL up to 4000 characters';
    end if;

    insert into public.media_jobs (
      source,
      source_event_id,
      payload,
      job_type,
      lark_request_id,
      item_index,
      max_attempts
    )
    values (
      'lark',
      v_request_id::text || ':' || v_item.item_index::text,
      jsonb_build_object(
        'share_url', v_link,
        'source_text', btrim(p_source_text),
        'is_homepage', coalesce((v_item.value ->> 'is_homepage')::boolean, false),
        'link_index', v_item.item_index
      ),
      'lark_parse',
      v_request_id,
      v_item.item_index,
      3
    )
    on conflict (source, source_event_id) do nothing;
  end loop;

  return jsonb_build_object(
    'request_id', v_request_id,
    'duplicate', not v_inserted,
    'job_count', v_link_count
  );
end;
$$;
-- A parse task expands into either more parse tasks (homepage entries) or one transfer task per
-- media file. Expansion and parent completion are atomic, so a crash cannot expose a half-built
-- task tree. Child source IDs are derived from the parent UUID and remain stable across retries.
create function public.expand_lark_job(
  p_job_id uuid,
  p_worker_id text,
  p_items jsonb,
  p_result jsonb default '{}'::jsonb
)
returns integer
language plpgsql
security invoker
set search_path = pg_catalog, public, pg_temp
as $$
declare
  v_job public.media_jobs%rowtype;
  v_item record;
  v_job_type text;
  v_payload jsonb;
  v_inserted integer := 0;
  v_row_count integer;
begin
  if jsonb_typeof(p_items) <> 'array' then
    raise exception 'items must be a JSON array';
  end if;
  if jsonb_array_length(p_items) < 1 or jsonb_array_length(p_items) > 500 then
    raise exception 'items must contain between 1 and 500 entries';
  end if;
  if jsonb_typeof(coalesce(p_result, '{}'::jsonb)) <> 'object' then
    raise exception 'result must be a JSON object';
  end if;

  select * into v_job
  from public.media_jobs
  where id = p_job_id
    and job_type = 'lark_parse'
    and status = 'processing'
    and locked_by = btrim(p_worker_id)
    and locked_until > now()
  for update;

  if not found then
    raise exception 'parse job is not owned by this active worker lease';
  end if;

  perform pg_advisory_xact_lock(hashtextextended(v_job.lark_request_id::text, 0));

  for v_item in
    select value, ordinality - 1 as item_index
    from jsonb_array_elements(p_items) with ordinality
  loop
    if jsonb_typeof(v_item.value) <> 'object' then
      raise exception 'each expanded item must be a JSON object';
    end if;

    v_job_type := nullif(btrim(v_item.value ->> 'job_type'), '');
    v_payload := v_item.value -> 'payload';
    if v_job_type not in ('lark_parse', 'media_transfer') then
      raise exception 'expanded job_type must be lark_parse or media_transfer';
    end if;
    if jsonb_typeof(v_payload) <> 'object' then
      raise exception 'expanded payload must be a JSON object';
    end if;

    insert into public.media_jobs (
      source,
      source_event_id,
      payload,
      job_type,
      lark_request_id,
      parent_job_id,
      item_index,
      max_attempts
    )
    values (
      'lark',
      v_job.id::text || ':' || v_job_type || ':' || v_item.item_index::text,
      v_payload,
      v_job_type,
      v_job.lark_request_id,
      v_job.id,
      v_item.item_index,
      3
    )
    on conflict (source, source_event_id) do nothing;

    get diagnostics v_row_count = row_count;
    v_inserted := v_inserted + v_row_count;
  end loop;

  update public.media_jobs
  set status = 'succeeded',
      locked_by = null,
      locked_until = null,
      result = coalesce(p_result, '{}'::jsonb) || jsonb_build_object(
        'expanded_job_count', jsonb_array_length(p_items)
      ),
      last_error = null,
      completed_at = now(),
      updated_at = now()
  where id = v_job.id;

  perform public.enqueue_lark_reply_if_ready(v_job.lark_request_id);
  return v_inserted;
end;
$$;
-- Replace the original function without changing its signature. Lark jobs acquire a per-request
-- advisory lock before a terminal transition so concurrent children cannot both miss the moment
-- when every sibling has finished.
create or replace function public.finish_media_job(
  p_job_id uuid,
  p_worker_id text,
  p_success boolean,
  p_result jsonb default '{}'::jsonb,
  p_error text default null,
  p_retry_delay_seconds integer default 30
)
returns text
language plpgsql
security invoker
set search_path = pg_catalog, public, pg_temp
as $$
declare
  v_job public.media_jobs%rowtype;
  v_status text;
  v_retry_delay_seconds integer := greatest(5, least(coalesce(p_retry_delay_seconds, 30), 3600));
begin
  if p_success is null then
    raise exception 'success is required';
  end if;

  select * into v_job
  from public.media_jobs
  where id = p_job_id
    and status = 'processing'
    and locked_by = btrim(p_worker_id)
    and locked_until > now()
  for update;

  if not found then
    raise exception 'job is not owned by this active worker lease';
  end if;

  if v_job.lark_request_id is not null then
    perform pg_advisory_xact_lock(hashtextextended(v_job.lark_request_id::text, 0));
  end if;

  update public.media_jobs
  set status = case
        when p_success then 'succeeded'
        when attempt_count >= max_attempts then 'failed'
        else 'queued'
      end,
      available_at = case
        when p_success or attempt_count >= max_attempts then available_at
        else now() + make_interval(secs => v_retry_delay_seconds)
      end,
      locked_by = null,
      locked_until = null,
      result = case when p_success then coalesce(p_result, '{}'::jsonb) else result end,
      last_error = case when p_success then null else left(nullif(p_error, ''), 4000) end,
      completed_at = case when p_success or attempt_count >= max_attempts then now() else null end,
      updated_at = now()
  where id = v_job.id
  returning status into v_status;

  if v_job.lark_request_id is not null
     and v_job.job_type <> 'lark_reply'
     and v_status in ('succeeded', 'failed') then
    perform public.enqueue_lark_reply_if_ready(v_job.lark_request_id);
  end if;

  return v_status;
end;
$$;
-- Reclaim behavior remains compatible with the original queue. Final expired leases also pass
-- through reply scheduling so a worker crash on the last attempt cannot strand a Lark request.
create function public.claim_pipeline_jobs(
  p_worker_id text,
  p_limit integer default 1,
  p_lease_seconds integer default 300
)
returns setof public.media_jobs
language plpgsql
security invoker
set search_path = pg_catalog, public, pg_temp
as $$
declare
  v_limit integer := greatest(1, least(coalesce(p_limit, 1), 5));
  v_lease_seconds integer := greatest(30, least(coalesce(p_lease_seconds, 300), 900));
  v_request_ids uuid[];
  v_request_id uuid;
begin
  if nullif(btrim(p_worker_id), '') is null then
    raise exception 'worker ID is required';
  end if;

  with expired as (
    update public.media_jobs
    set status = 'failed',
        locked_by = null,
        locked_until = null,
        last_error = coalesce(last_error, 'Worker lease expired after final attempt'),
        completed_at = now(),
        updated_at = now()
    where status = 'processing'
      and locked_until < now()
      and attempt_count >= max_attempts
    returning lark_request_id
  )
  select array_agg(distinct lark_request_id)
  into v_request_ids
  from expired
  where lark_request_id is not null;

  if v_request_ids is not null then
    foreach v_request_id in array v_request_ids
    loop
      perform pg_advisory_xact_lock(hashtextextended(v_request_id::text, 0));
      perform public.enqueue_lark_reply_if_ready(v_request_id);
    end loop;
  end if;

  return query
  with candidates as (
    select id
    from public.media_jobs
    where (status = 'queued' and available_at <= now())
       or (status = 'processing' and locked_until < now() and attempt_count < max_attempts)
    order by available_at, created_at
    for update skip locked
    limit v_limit
  )
  update public.media_jobs as job
  set status = 'processing',
      attempt_count = job.attempt_count + 1,
      locked_by = btrim(p_worker_id),
      locked_until = now() + make_interval(secs => v_lease_seconds),
      updated_at = now(),
      last_error = null
  from candidates
  where job.id = candidates.id
  returning job.*;
end;
$$;
-- The retained Supabase media-worker remains a transfer-only fallback. It must never claim parse
-- or reply jobs that only the Cloudflare pipeline understands.
create or replace function public.claim_media_jobs(
  p_worker_id text,
  p_limit integer default 1,
  p_lease_seconds integer default 300
)
returns setof public.media_jobs
language plpgsql
security invoker
set search_path = pg_catalog, public, pg_temp
as $$
declare
  v_limit integer := greatest(1, least(coalesce(p_limit, 1), 5));
  v_lease_seconds integer := greatest(30, least(coalesce(p_lease_seconds, 300), 900));
  v_request_ids uuid[];
  v_request_id uuid;
begin
  if nullif(btrim(p_worker_id), '') is null then
    raise exception 'worker ID is required';
  end if;

  with expired as (
    update public.media_jobs
    set status = 'failed',
        locked_by = null,
        locked_until = null,
        last_error = coalesce(last_error, 'Worker lease expired after final attempt'),
        completed_at = now(),
        updated_at = now()
    where job_type = 'media_transfer'
      and status = 'processing'
      and locked_until < now()
      and attempt_count >= max_attempts
    returning lark_request_id
  )
  select array_agg(distinct lark_request_id)
  into v_request_ids
  from expired
  where lark_request_id is not null;

  if v_request_ids is not null then
    foreach v_request_id in array v_request_ids
    loop
      perform pg_advisory_xact_lock(hashtextextended(v_request_id::text, 0));
      perform public.enqueue_lark_reply_if_ready(v_request_id);
    end loop;
  end if;

  return query
  with candidates as (
    select id
    from public.media_jobs
    where job_type = 'media_transfer'
      and (
        (status = 'queued' and available_at <= now())
        or (status = 'processing' and locked_until < now() and attempt_count < max_attempts)
      )
    order by available_at, created_at
    for update skip locked
    limit v_limit
  )
  update public.media_jobs as job
  set status = 'processing',
      attempt_count = job.attempt_count + 1,
      locked_by = btrim(p_worker_id),
      locked_until = now() + make_interval(secs => v_lease_seconds),
      updated_at = now(),
      last_error = null
  from candidates
  where job.id = candidates.id
  returning job.*;
end;
$$;
-- The reply worker receives one bounded aggregate instead of reading exposed tables directly.
create function public.get_lark_reply_context(
  p_job_id uuid,
  p_worker_id text
)
returns jsonb
language plpgsql
security invoker
set search_path = pg_catalog, public, pg_temp
as $$
declare
  v_request_id uuid;
  v_context jsonb;
begin
  select lark_request_id into v_request_id
  from public.media_jobs
  where id = p_job_id
    and job_type = 'lark_reply'
    and status = 'processing'
    and locked_by = btrim(p_worker_id)
    and locked_until > now();

  if not found or v_request_id is null then
    raise exception 'reply job is not owned by this active worker lease';
  end if;

  select jsonb_build_object(
    'request_id', request.id,
    'event_id', request.event_id,
    'message_id', request.message_id,
    'chat_id', request.chat_id,
    'source_text', request.source_text,
    'jobs', coalesce((
      select jsonb_agg(jsonb_build_object(
        'id', job.id,
        'job_type', job.job_type,
        'status', job.status,
        'item_index', job.item_index,
        'payload', job.payload,
        'result', job.result,
        'last_error', job.last_error
      ) order by job.created_at, job.item_index, job.id)
      from public.media_jobs as job
      where job.lark_request_id = request.id
        and job.job_type <> 'lark_reply'
    ), '[]'::jsonb)
  )
  into v_context
  from public.lark_requests as request
  where request.id = v_request_id;

  if v_context is null then
    raise exception 'Lark request context was not found';
  end if;

  return v_context;
end;
$$;
revoke all on function public.enqueue_lark_reply_if_ready(uuid) from public, anon, authenticated;
revoke all on function public.enqueue_lark_request(text, text, text, text, text, jsonb) from public, anon, authenticated;
revoke all on function public.expand_lark_job(uuid, text, jsonb, jsonb) from public, anon, authenticated;
revoke all on function public.get_lark_reply_context(uuid, text) from public, anon, authenticated;
revoke all on function public.claim_pipeline_jobs(text, integer, integer) from public, anon, authenticated;
grant execute on function public.enqueue_lark_reply_if_ready(uuid) to service_role;
grant execute on function public.enqueue_lark_request(text, text, text, text, text, jsonb) to service_role;
grant execute on function public.expand_lark_job(uuid, text, jsonb, jsonb) to service_role;
grant execute on function public.get_lark_reply_context(uuid, text) to service_role;
grant execute on function public.claim_pipeline_jobs(text, integer, integer) to service_role;
-- CREATE OR REPLACE preserves existing ACLs, but repeat the explicit boundary for auditability.
revoke all on function public.claim_media_jobs(text, integer, integer) from public, anon, authenticated;
revoke all on function public.finish_media_job(uuid, text, boolean, jsonb, text, integer) from public, anon, authenticated;
grant execute on function public.claim_media_jobs(text, integer, integer) to service_role;
grant execute on function public.finish_media_job(uuid, text, boolean, jsonb, text, integer) to service_role;
commit;
