begin;
-- Cloudflare Queues carries only this stable (source, source_event_id) pair. The database remains
-- the source of truth: a duplicate, stale, or delayed queue message merely finds no eligible row
-- and is acknowledged by the consumer. Keeping this as a separate single-job claim avoids an
-- unnecessary ordered queue scan on the low-latency path while preserving the existing Cron sweep.
create function public.claim_pipeline_job_by_source(
  p_source text,
  p_source_event_id text,
  p_worker_id text,
  p_lease_seconds integer default 300
)
returns setof public.media_jobs
language plpgsql
security invoker
set search_path = pg_catalog, public, pg_temp
as $$
declare
  v_lease_seconds integer := greatest(30, least(coalesce(p_lease_seconds, 300), 900));
  v_expired_request_id uuid;
begin
  if nullif(btrim(p_source), '') is null or char_length(btrim(p_source)) > 80 then
    raise exception 'source must be between 1 and 80 characters';
  end if;
  if nullif(btrim(p_source_event_id), '') is null
     or char_length(btrim(p_source_event_id)) > 200 then
    raise exception 'source_event_id must be between 1 and 200 characters';
  end if;
  if nullif(btrim(p_worker_id), '') is null then
    raise exception 'worker ID is required';
  end if;

  -- Match the normal sweep's terminal-lease handling for this exact message. This makes a late
  -- Queue delivery harmless and ensures the final Lark reply can still be scheduled immediately.
  update public.media_jobs
  set status = 'failed',
      locked_by = null,
      locked_until = null,
      last_error = coalesce(last_error, 'Worker lease expired after final attempt'),
      completed_at = now(),
      updated_at = now()
  where source = btrim(p_source)
    and source_event_id = btrim(p_source_event_id)
    and status = 'processing'
    and locked_until < now()
    and attempt_count >= max_attempts
  returning lark_request_id into v_expired_request_id;

  if v_expired_request_id is not null then
    perform pg_advisory_xact_lock(hashtextextended(v_expired_request_id::text, 0));
    perform public.enqueue_lark_reply_if_ready(v_expired_request_id);
  end if;

  return query
  with candidate as (
    select id
    from public.media_jobs
    where source = btrim(p_source)
      and source_event_id = btrim(p_source_event_id)
      and (
        (status = 'queued' and available_at <= now())
        or (status = 'processing' and locked_until < now() and attempt_count < max_attempts)
      )
    for update skip locked
  )
  update public.media_jobs as job
  set status = 'processing',
      attempt_count = job.attempt_count + 1,
      locked_by = btrim(p_worker_id),
      locked_until = now() + make_interval(secs => v_lease_seconds),
      updated_at = now(),
      last_error = null
  from candidate
  where job.id = candidate.id
  returning job.*;
end;
$$;
revoke all on function public.claim_pipeline_job_by_source(text, text, text, integer)
  from public, anon, authenticated;
grant execute on function public.claim_pipeline_job_by_source(text, text, text, integer)
  to service_role;
commit;
