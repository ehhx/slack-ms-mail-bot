begin;
-- A zero retry delay is an explicit permanent-failure signal from the Worker. Positive values
-- retain the existing delayed retry behavior, so older callers remain fully compatible.
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
  v_retryable boolean := coalesce(p_retry_delay_seconds, 30) > 0;
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
        when not v_retryable or attempt_count >= max_attempts then 'failed'
        else 'queued'
      end,
      available_at = case
        when p_success or not v_retryable or attempt_count >= max_attempts then available_at
        else now() + make_interval(secs => v_retry_delay_seconds)
      end,
      locked_by = null,
      locked_until = null,
      result = case when p_success then coalesce(p_result, '{}'::jsonb) else result end,
      last_error = case when p_success then null else left(nullif(p_error, ''), 4000) end,
      completed_at = case
        when p_success or not v_retryable or attempt_count >= max_attempts then now()
        else null
      end,
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
commit;
