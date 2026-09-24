-- Run once in the Neon SQL editor.
create table if not exists coils (
  coil_no            text primary key,
  entry_date         date,
  so_no              text,
  do_no              text,
  form               text,
  grade              text,
  thk                numeric,
  width              integer,
  length             integer,
  weight             numeric,
  destination        text,
  transporter_name   text,
  transporter_mobile text,
  status             text not null default 'Under Loading',
  vehicle_no         text,
  lr_no              text,
  dispatch_date      date,
  expected_arrival   date,
  actual_arrival     date,
  remarks            text,
  dispatch_id        integer,
  excel_row          integer,          -- 0-based row index inside tbl_Dispatch at last sync
  excel_hash         text,             -- hash of BSC-owned fields as last seen in Excel
  updated_at         timestamptz not null default now(),
  mirrored_at        timestamptz,
  created_at         timestamptz not null default now()
);
create table if not exists dispatches (
  id                 serial primary key,
  token              text unique not null,
  do_no              text,
  transporter_name   text,
  transporter_mobile text,
  created_at         timestamptz not null default now(),
  notified_at        timestamptz,
  notify_error       text
);
create table if not exists events (
  id        serial primary key,
  coil_no   text,
  action    text not null,
  actor     text,
  detail    jsonb,
  at        timestamptz not null default now()
);
create index if not exists coils_dispatch_idx on coils(dispatch_id);
create index if not exists coils_status_idx on coils(status);

-- v2: transporter self-service
create table if not exists transporters (
  id        serial primary key,
  name      text not null,
  mobile    text unique not null,
  token     text unique not null,
  active    boolean not null default true,
  excel_row integer,
  link_sent_at timestamptz,
  notify_error text,
  created_at timestamptz not null default now()
);
alter table coils add column if not exists source text not null default 'bsc';          -- 'bsc' | 'transporter'
alter table coils add column if not exists transporter_id integer;
alter table coils add column if not exists entered_by text;
alter table dispatches add column if not exists transporter_id integer;
create index if not exists coils_transporter_idx on coils(transporter_id);

-- v3: consignment lines (one row = invoice + vehicle + size), keyed by system Ref stored in coil_no
alter table coils add column if not exists n_coils integer;
alter table coils add column if not exists mirror_hash text;

-- v4: sales-order master for ordered vs dispatched vs pending
create table if not exists sos (
  so_no          text primary key,
  so_date        date,
  description    text,
  ordered_mt     numeric,
  ordered_coils  integer,
  remarks        text,
  excel_row      integer,
  updated_at     timestamptz not null default now()
);
