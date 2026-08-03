begin;
-- The table is deliberately private: Edge Functions call the narrowly-scoped RPCs below using
-- service_role, while no browser role can read payloads, leases, errors, or WebDAV destinations.
create table public.media_jobs (
  id uuid primary key default gen_random_uuid(),
  source text not null check (char_length(source) between 1 and 80),
  source_event_id text not null check (char_length(source_event_id) between 1 and 200),
  payload jsonb not null check (jsonb_typeof(payload) = 'object'),
  status text not null default 'queued' check (status in ('queued', 'processing', 'succeeded', 'failed')),
  attempt_count integer not null default 0 check (attempt_count >= 0),
  max_attempts integer not null default 3 check (max_attempts between 1 and 10),
  available_at timestamptz not null default now(),
  locked_by text,
  locked_until timestamptz,
  result jsonb,
  last_error text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  completed_at timestamptz,
  unique (source, source_event_id),
  check (
    (status = 'processing' and locked_by is not null and locked_until is not null)
    or (status <> 'processing' and locked_by is null and locked_until is null)
  )
);
create index media_jobs_available_idx
  on public.media_jobs (available_at, created_at)
  where status = 'queued';
create index media_jobs_expired_lease_idx
  on public.media_jobs (locked_until)
  where status = 'processing';
alter table public.media_jobs enable row level security;
revoke all on table public.media_jobs from public, anon, authenticated, service_role;
grant select, insert, update on table public.media_jobs to service_role;
-- Returns the pre-existing ID on duplicate delivery without resetting a job that is already running
-- or completed. The no-op update makes the result available to the single caller under concurrency.
create function public.enqueue_media_job(
  p_source text,
  p_source_event_id text,
  p_payload jsonb
)
returns uuid
language plpgsql
security invoker
set search_path = pg_catalog, public, pg_temp
as $$
declare
  v_job_id uuid;
begin
  if nullif(btrim(p_source), '') is null or nullif(btrim(p_source_event_id), '') is null then
    raise exception 'source and source_event_id are required';
  end if;

  if jsonb_typeof(p_payload) <> 'object' then
    raise exception 'payload must be a JSON object';
  end if;

  insert into public.media_jobs (source, source_event_id, payload)
  values (btrim(p_source), btrim(p_source_event_id), p_payload)
  on conflict (source, source_event_id) do update
    set updated_at = public.media_jobs.updated_at
  returning id into v_job_id;

  return v_job_id;
end;
$$;
-- SKIP LOCKED permits multiple worker invocations without double-processing a task. An expired
-- processing lease is reclaimable, so an interrupted Edge Function cannot leave a job stuck.
create function public.claim_media_jobs(
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
begin
  if nullif(btrim(p_worker_id), '') is null then
    raise exception 'worker ID is required';
  end if;

  -- A worker that dies during its final permitted attempt cannot call finish_media_job. Mark that
  -- expired lease terminal first, otherwise repeated reclaiming can exceed max_attempts forever.
  update public.media_jobs
  set status = 'failed',
      locked_by = null,
      locked_until = null,
      last_error = coalesce(last_error, 'Worker lease expired after final attempt'),
      completed_at = now(),
      updated_at = now()
  where status = 'processing'
    and locked_until < now()
    and attempt_count >= max_attempts;

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
-- Only the worker that owns a still-valid lease can finish a job. Failed attempts are requeued
-- with a bounded delay until max_attempts is reached, at which point the job becomes terminal.
create function public.finish_media_job(
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
  v_status text;
  v_retry_delay_seconds integer := greatest(5, least(coalesce(p_retry_delay_seconds, 30), 3600));
begin
  if p_success is null then
    raise exception 'success is required';
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
  where id = p_job_id
    and status = 'processing'
    and locked_by = btrim(p_worker_id)
    and locked_until > now()
  returning status into v_status;

  if not found then
    raise exception 'job is not owned by this active worker lease';
  end if;

  return v_status;
end;
$$;
revoke all on function public.enqueue_media_job(text, text, jsonb) from public, anon, authenticated;
revoke all on function public.claim_media_jobs(text, integer, integer) from public, anon, authenticated;
revoke all on function public.finish_media_job(uuid, text, boolean, jsonb, text, integer) from public, anon, authenticated;
grant execute on function public.enqueue_media_job(text, text, jsonb) to service_role;
grant execute on function public.claim_media_jobs(text, integer, integer) to service_role;
grant execute on function public.finish_media_job(uuid, text, boolean, jsonb, text, integer) to service_role;
commit;
