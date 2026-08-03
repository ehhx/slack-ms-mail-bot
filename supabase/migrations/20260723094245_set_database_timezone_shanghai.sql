begin;
-- Keep all columns as timestamptz so leases and cross-region comparisons remain correct. This
-- changes the default session display timezone used by Supabase SQL clients and the dashboard.
alter database postgres set timezone to 'Asia/Shanghai';
commit;
