// api/fleet.js — Fleet / Vehicle P&L MIS module backend (Neon-backed).
// Routes (all under /api/fleet):
//   GET  ?fmt=bootstrap                 -> vehicles + drivers (for dropdowns)
//   GET  ?fmt=dashboard&month=YYYY-MM-01 -> the P&L matrix for one month
//   GET  ?fmt=trips&from=&to=&vehicle=  -> trip rows
//   POST {action:'save_trip'|'delete_trip'|'save_vehicle'|'save_driver'|'save_toll'}
//   POST ?fmt=sync   (Bearer CRON_SECRET) -> upsert fleet_costs from fleet_refresh.ps1
// build 2026-11-16a — initial Fleet module (marker: FLEET-1116A)
import { neon } from "@neondatabase/serverless";
const sql = neon(process.env.DATABASE_URL);

const json = (res, code, obj) => { res.setHeader("Cache-Control", "no-store"); res.status(code).json(obj); };
const firstOfMonth = (s) => {
  const m = /^(\d{4})-(\d{2})/.exec(String(s || ""));
  const d = m ? new Date(Date.UTC(+m[1], +m[2] - 1, 1)) : new Date();
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-01`;
};
const monthEnd = (first) => { const [y, mo] = first.split("-").map(Number); const d = new Date(Date.UTC(y, mo, 0)); return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`; };
const num = (v) => { const n = parseFloat(v); return isFinite(n) ? n : 0; };
async function readBody(req) {
  if (req.body && typeof req.body === "object") return req.body;
  let raw = ""; await new Promise(r => { req.on("data", c => raw += c); req.on("end", r); req.on("error", r); });
  try { return JSON.parse(raw || "{}"); } catch { return {}; }
}

async function handler(req, res) {
  const url = req.url || "";
  const qp = (k) => { try { return new URL(url, "http://x").searchParams.get(k); } catch { return (new RegExp("[?&]" + k + "=([^&]*)").exec(url) || [])[1] || null; } };
  const fmt = qp("fmt") || "";

  try {
    // ---- SAP cost sync (called by fleet_refresh.ps1) ----
    if (fmt === "sync") {
      if (!process.env.CRON_SECRET) return json(res, 503, { error: "CRON_SECRET not set" });
      const auth = req.headers.authorization || "";
      if (auth !== "Bearer " + process.env.CRON_SECRET) return json(res, 401, { error: "unauthorized" });
      const body = await readBody(req);
      const rows = Array.isArray(body.rows) ? body.rows : [];
      let n = 0;
      for (const r of rows) {
        const month = firstOfMonth(r.month || r.Mon);
        const veh = String(r.vehicle || r.Vehicle || "").replace(/\s+/g, "").toUpperCase();
        if (!veh) continue;
        await sql`insert into fleet_costs (month, vehicle_no, diesel, bata, maintenance, unclassified, synced_at)
          values (${month}, ${veh}, ${num(r.diesel ?? r.Diesel)}, ${num(r.bata ?? r.Bata)}, ${num(r.maintenance ?? r.Maintenance)}, ${num(r.unclassified ?? r.Unclassified)}, now())
          on conflict (month, vehicle_no) do update set
            diesel=excluded.diesel, bata=excluded.bata, maintenance=excluded.maintenance,
            unclassified=excluded.unclassified, synced_at=now()`;
        n++;
      }
      return json(res, 200, { ok: true, upserted: n });
    }

    // ---- dropdowns ----
    if (fmt === "bootstrap") {
      const vehicles = await sql`select vehicle_no, driver, active from fleet_vehicles order by sort_order, vehicle_no`;
      const drivers = await sql`select name, vehicle from fleet_drivers where active order by name`;
      return json(res, 200, { vehicles, drivers });
    }

    // ---- full vehicle master (for the Vehicle Master tab) ----
    if (fmt === "master") {
      const rows = await sql`select * from fleet_vehicles order by sort_order, vehicle_no`;
      return json(res, 200, { rows });
    }

    // ---- the P&L matrix for one month ----
    if (fmt === "dashboard") {
      const month = firstOfMonth(qp("month"));
      const end = monthEnd(month);
      const vehicles = await sql`select * from fleet_vehicles where active order by sort_order, vehicle_no`;
      const trips = await sql`
        select vehicle_no,
               count(*)::int as trips,
               coalesce(sum(km),0) as km,
               coalesce(sum(income),0) as income
        from fleet_trips where trip_date >= ${month} and trip_date <= ${end}
        group by vehicle_no`;
      const costs = await sql`select * from fleet_costs where month = ${month}`;
      const tolls = await sql`select vehicle_no, toll from fleet_monthly where month = ${month}`;
      const tripBy = Object.fromEntries(trips.map(r => [r.vehicle_no, r]));
      const costBy = Object.fromEntries(costs.map(r => [r.vehicle_no, r]));
      const tollBy = Object.fromEntries(tolls.map(r => [r.vehicle_no, r]));
      const out = vehicles.map(v => {
        const t = tripBy[v.vehicle_no] || {}, c = costBy[v.vehicle_no] || {}, tl = tollBy[v.vehicle_no] || {};
        const fixed = {
          salary: num(v.salary_mo),
          insurance: num(v.insurance_yr) / 12,
          roadtax: num(v.roadtax_yr) / 12,
          fc: num(v.fc_yr) / 12,
          dep: num(v.dep_yr) / 12,
          interest: num(v.interest_yr) / 12,
        };
        const totFixed = fixed.salary + fixed.insurance + fixed.roadtax + fixed.fc + fixed.dep + fixed.interest;
        const variable = { diesel: num(c.diesel), bata: num(c.bata), toll: num(tl.toll), maintenance: num(c.maintenance) };
        const totVar = variable.diesel + variable.bata + variable.toll + variable.maintenance;
        const income = num(t.income), totExp = totFixed + totVar;
        return {
          vehicle_no: v.vehicle_no, driver: v.driver || "",
          trips: t.trips || 0, km: num(t.km),
          fixed, totFixed, variable, totVar,
          unclassified: num(c.unclassified),
          income, totExp, pl: income - totExp,
          synced: c.synced_at || null,
        };
      });
      return json(res, 200, { month, rows: out });
    }

    // ---- trip list ----
    if (fmt === "trips") {
      const from = firstOfMonth(qp("from") || qp("month"));
      const to = qp("to") ? monthEnd(firstOfMonth(qp("to"))) : monthEnd(from);
      const veh = qp("vehicle");
      const rows = veh
        ? await sql`select * from fleet_trips where trip_date >= ${from} and trip_date <= ${to} and vehicle_no = ${veh} order by trip_date desc, id desc`
        : await sql`select * from fleet_trips where trip_date >= ${from} and trip_date <= ${to} order by trip_date desc, id desc`;
      return json(res, 200, { rows });
    }

    // ---- writes ----
    if (req.method === "POST") {
      const b = await readBody(req);
      const action = b.action || "";

      if (action === "save_trip") {
        const t = b.trip || b;
        const veh = String(t.vehicle_no || "").replace(/\s+/g, "").toUpperCase();
        if (!t.trip_date || !veh) return json(res, 400, { error: "trip_date and vehicle_no required" });
        const income = (t.income === "" || t.income == null) ? null : num(t.income);
        if (t.id) {
          await sql`update fleet_trips set trip_date=${t.trip_date}, vehicle_no=${veh}, driver=${t.driver || null},
            trip_type=${t.trip_type || "Loaded"}, from_place=${t.from_place || null}, to_place=${t.to_place || null},
            load_mt=${t.load_mt == null || t.load_mt === "" ? null : num(t.load_mt)}, km=${t.km == null || t.km === "" ? null : num(t.km)},
            income=${income}, remarks=${t.remarks || null} where id=${t.id}`;
          return json(res, 200, { ok: true, id: t.id });
        }
        const r = await sql`insert into fleet_trips (trip_date, vehicle_no, driver, trip_type, from_place, to_place, load_mt, km, income, remarks, created_by)
          values (${t.trip_date}, ${veh}, ${t.driver || null}, ${t.trip_type || "Loaded"}, ${t.from_place || null}, ${t.to_place || null},
            ${t.load_mt == null || t.load_mt === "" ? null : num(t.load_mt)}, ${t.km == null || t.km === "" ? null : num(t.km)},
            ${income}, ${t.remarks || null}, ${b.user || null}) returning id`;
        return json(res, 200, { ok: true, id: r[0].id });
      }

      if (action === "delete_trip") {
        if (!b.id) return json(res, 400, { error: "id required" });
        await sql`delete from fleet_trips where id=${b.id}`;
        return json(res, 200, { ok: true });
      }

      if (action === "save_vehicle") {
        const v = b.vehicle || b;
        const veh = String(v.vehicle_no || "").replace(/\s+/g, "").toUpperCase();
        if (!veh) return json(res, 400, { error: "vehicle_no required" });
        await sql`insert into fleet_vehicles (vehicle_no, driver, salary_mo, insurance_yr, roadtax_yr, fc_yr, dep_yr, interest_yr, active, sort_order, updated_at)
          values (${veh}, ${v.driver || null}, ${num(v.salary_mo)}, ${num(v.insurance_yr)}, ${num(v.roadtax_yr)}, ${num(v.fc_yr)}, ${num(v.dep_yr)}, ${num(v.interest_yr)}, ${v.active !== false}, ${num(v.sort_order)}, now())
          on conflict (vehicle_no) do update set driver=excluded.driver, salary_mo=excluded.salary_mo,
            insurance_yr=excluded.insurance_yr, roadtax_yr=excluded.roadtax_yr, fc_yr=excluded.fc_yr,
            dep_yr=excluded.dep_yr, interest_yr=excluded.interest_yr, active=excluded.active, updated_at=now()`;
        if (v.driver) await sql`insert into fleet_drivers (name, vehicle) values (${v.driver}, ${veh}) on conflict (name) do update set vehicle=excluded.vehicle`;
        return json(res, 200, { ok: true });
      }

      if (action === "save_driver") {
        if (!b.name) return json(res, 400, { error: "name required" });
        await sql`insert into fleet_drivers (name, vehicle, active) values (${b.name}, ${b.vehicle || null}, ${b.active !== false})
          on conflict (name) do update set vehicle=excluded.vehicle, active=excluded.active`;
        return json(res, 200, { ok: true });
      }

      if (action === "save_toll") {
        const month = firstOfMonth(b.month);
        const veh = String(b.vehicle_no || "").replace(/\s+/g, "").toUpperCase();
        if (!veh) return json(res, 400, { error: "vehicle_no required" });
        await sql`insert into fleet_monthly (month, vehicle_no, toll, updated_at) values (${month}, ${veh}, ${num(b.toll)}, now())
          on conflict (month, vehicle_no) do update set toll=excluded.toll, updated_at=now()`;
        return json(res, 200, { ok: true });
      }

      return json(res, 400, { error: "unknown action: " + action });
    }

    return json(res, 404, { error: "unknown fmt: " + fmt });
  } catch (e) {
    return json(res, 500, { error: String(e && e.message || e) });
  }
}

export { handler as default };
