
-- DIC / ITESO - esquema inicial para Supabase
-- Ejecutar en SQL Editor de Supabase.

create extension if not exists pgcrypto;

create table if not exists public.units (
  code text primary key,
  name text not null,
  active boolean not null default true
);

insert into public.units (code, name) values
('CUE','Centro Universidad Empresa'),
('COINCIDE','Centro Universitario de Incidencia Social'),
('CUDJ','Centro Universitario por la Dignidad y la Justicia Francisco Suárez, SJ'),
('CUI','Centro Universitario Ignaciano'),
('CEJUVEN','Centro de Acompañamiento y Estudios Juveniles'),
('CPC','Centro de Promoción Cultural'),
('CEFSI','Centro de Educación Física y Salud Integral')
on conflict (code) do update set name = excluded.name;

create table if not exists public.reports (
  id uuid primary key default gen_random_uuid(),
  unit_code text not null references public.units(code),
  month text not null,
  year integer not null,
  status text not null default 'BORRADOR'
    check (status in ('BORRADOR','ENVIADO','REVISADO','CERRADO')),
  sender_email text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  submitted_at timestamptz,
  unique(unit_code, month, year)
);

create table if not exists public.activities (
  id uuid primary key default gen_random_uuid(),
  report_id uuid not null references public.reports(id) on delete cascade,
  title text not null,
  description_original text not null,
  description_edited text,
  category text,
  ranking integer check (ranking is null or ranking between 1 and 3),
  activity_date date,
  participants integer,
  order_index integer not null default 1,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- Evita que un mismo reporte tenga dos Top 1, dos Top 2, etc.
create unique index if not exists uq_report_ranking
on public.activities(report_id, ranking)
where ranking is not null;

create table if not exists public.activity_photos (
  id uuid primary key default gen_random_uuid(),
  activity_id uuid not null references public.activities(id) on delete cascade,
  storage_path text not null,
  original_filename text,
  caption text,
  created_at timestamptz not null default now()
);

create table if not exists public.audit_log (
  id bigint generated always as identity primary key,
  user_email text,
  action text not null,
  entity_type text not null,
  entity_id text,
  details jsonb,
  created_at timestamptz not null default now()
);

-- Para una prueba inicial puede habilitarse RLS posteriormente.
-- En producción conviene activar Row Level Security y políticas por usuario/unidad.
