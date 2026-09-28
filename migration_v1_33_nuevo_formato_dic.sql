-- DIC ITESO · V1.33
-- Nuevo formato de captura del Informe Mensual DIC (versión 24 sep 2026)
-- Ejecutar una sola vez antes de publicar app.py V1.33.

alter table public.activities
    add column if not exists action_purpose text,
    add column if not exists action_type text,
    add column if not exists inclusion_criteria jsonb not null default '[]'::jsonb,
    add column if not exists location text,
    add column if not exists target_population jsonb not null default '[]'::jsonb,
    add column if not exists target_external_name text,
    add column if not exists dic_collaboration boolean,
    add column if not exists dic_collaboration_units text,
    add column if not exists relevance_note text,
    add column if not exists detail_data jsonb not null default '{}'::jsonb;

alter table public.reports
    add column if not exists monthly_highlights jsonb not null default '[]'::jsonb,
    add column if not exists learning_planning_advances text,
    add column if not exists learning_risks text,
    add column if not exists learning_opportunity text,
    add column if not exists media_monthly_summary jsonb not null default '{}'::jsonb,
    add column if not exists form_version text;

grant select, insert, update, delete on table public.activities to service_role;
grant select, insert, update, delete on table public.reports to service_role;
