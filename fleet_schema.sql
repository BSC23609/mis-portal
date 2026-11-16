-- fleet_schema.sql  —  Fleet MIS module tables. Run once in the Neon SQL editor.

-- Vehicle master: driver salary is per MONTH; insurance/road-tax/fc/dep/interest are per YEAR.
create table if not exists fleet_vehicles (
  vehicle_no   text primary key,
  active       boolean not null default true,
  driver       text,
  salary_mo    numeric not null default 0,
  insurance_yr numeric not null default 0,
  roadtax_yr   numeric not null default 0,
  fc_yr        numeric not null default 0,
  dep_yr       numeric not null default 0,
  interest_yr  numeric not null default 0,
  sort_order   integer not null default 0,
  updated_at   timestamptz not null default now()
);

create table if not exists fleet_drivers (
  name    text primary key,
  vehicle text,
  active  boolean not null default true
);

-- One row per trip (entered in the app).
create table if not exists fleet_trips (
  id         serial primary key,
  trip_date  date not null,
  vehicle_no text not null,
  driver     text,
  trip_type  text not null default 'Loaded',   -- Loaded | Empty/Return | Inbound | Other
  from_place text,
  to_place   text,
  load_mt    numeric,
  km         numeric,
  income     numeric,                           -- null/blank for non-earning trips
  remarks    text,
  created_by text,
  created_at timestamptz not null default now()
);
create index if not exists fleet_trips_vd_idx on fleet_trips(vehicle_no, trip_date);
create index if not exists fleet_trips_date_idx on fleet_trips(trip_date);

-- Monthly manual entries per vehicle (toll).
create table if not exists fleet_monthly (
  month      date not null,                     -- first day of month
  vehicle_no text not null,
  toll       numeric not null default 0,
  updated_at timestamptz not null default now(),
  primary key (month, vehicle_no)
);

-- SAP-derived running costs per vehicle per month (pushed by fleet_refresh.ps1).
create table if not exists fleet_costs (
  month        date not null,
  vehicle_no   text not null,
  diesel       numeric not null default 0,
  bata         numeric not null default 0,
  maintenance  numeric not null default 0,
  unclassified numeric not null default 0,
  synced_at    timestamptz not null default now(),
  primary key (month, vehicle_no)
);

-- Confirmed vendor -> category map (seed below; editable in the app later).
create table if not exists fleet_vendor_map (
  vendor_code text primary key,
  vendor_name text,
  category    text not null                     -- Diesel | Bata | Maintenance | Insurance
);

-- Seed the nine vehicles (edit costs in the app; order controls column order).
insert into fleet_vehicles (vehicle_no, sort_order) values
  ('TN04AD3336',1),('TN04AE3484',2),('TN04AF6394',3),('TN04AH1119',4),
  ('TN04AT3201',5),('TN18BM8443',6),('TN18BV5968',7),('TN18CA0415',8),('TN30AY3818',9)
on conflict (vehicle_no) do nothing;

-- Seed the confirmed vendor map.
insert into fleet_vendor_map (vendor_code, vendor_name, category) values
  ('BSV00843','NATIONAL INSURANCE COMPANY LIMITED','Insurance'),
  ('BSV00773','ROYAL SUNDARAM GENERAL INSURANCE CO.LIMITED','Insurance'),
  ('BSV00190','KUMAAR AUTOMOBILES','Maintenance'),
  ('BSV00133','MILLENNIUM TYRES','Maintenance'),
  ('BSV01312','SRI ANCHANEYA AUTOMOBILES','Maintenance'),
  ('BSV01409','TATA MOTORS LIMITED','Maintenance'),
  ('BSV00761','S.M. MOGANRAAJ M.KARTHICK TRANSPORT CONSULTANT','Maintenance'),
  ('BSV00763','SRI HI-TECH ENTERPRISES','Maintenance'),
  ('BSV00220','VICTOR PAUL G','Bata'),
  ('BSV01187','ABDHUL RAHAMAN U','Bata'),
  ('BSV01253','MUNUSAMY KASI','Bata'),
  ('BSV00784','R.KANAGARAJ','Bata'),
  ('BSV01378','ELUMALAI.V','Bata'),
  ('BSV00221','MOHAN VEL','Bata')
on conflict (vendor_code) do update set vendor_name=excluded.vendor_name, category=excluded.category;
