// NMDC backend for MIS portal — sync (Excel->Neon) + dispatch (read). Transporter/WATI/arrive disabled (routes removed; WATI env vars unset).
// original single-file build (dispatch, sync, tform, arrive) — build 2026-09-21g (consignment lines, Ref key, SO master)

// lib/db.js
import { neon } from "@neondatabase/serverless";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
var sql = neon(process.env.DATABASE_URL);
var STATUSES = ["Under Loading", "Loaded", "In-Transit", "Arrived"];
function json(res, status, body) {
  res.setHeader("Cache-Control", "no-store");
  res.status(status).json(body);
}
var d = (v) => v ? String(v).slice(0, 10) : null;
function cors(req, res) {
  const norm2 = (u) => String(u || "").trim().toLowerCase().replace(/\/+$/, "");
  const allowed = (process.env.ALLOWED_ORIGINS || "").split(",").map(norm2).filter(Boolean);
  const origin = req.headers.origin;
  const ok = origin && (allowed.includes("*") || allowed.some((a) => a === norm2(origin) || a === norm2(origin).replace(/^https?:\/\//, "")));
  if (ok) {
    res.setHeader("Access-Control-Allow-Origin", origin);
    res.setHeader("Vary", "Origin");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, x-dash-pin");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader("Access-Control-Max-Age", "86400");
  }
  if (req.method === "OPTIONS") {
    res.status(204).end();
    return true;
  }
  return false;
}

// api/dispatch.js
var lastSyncAt = 0;
async function handler(req, res) {
  if (cors(req, res)) return;
  // Optional live pull: ?sync=1 runs the Excel->Neon import first (rate-limited to
  // once per 20s so a refresh-spam can't hammer Graph). Uses the server-side secret;
  // the browser never sees it.
  const wantSync = /[?&]sync=1/.test(req.url || "");
  if (wantSync && process.env.CRON_SECRET && (Date.now() - lastSyncAt > 20000)) {
    lastSyncAt = Date.now();
    try {
      const base = (process.env.PUBLIC_URL || `https://${req.headers.host}`).replace(/\/$/, "");
      await fetch(base + "/api/sync", { headers: { authorization: "Bearer " + process.env.CRON_SECRET } });
    } catch (e) { /* if sync fails, still return whatever Neon has */ }
  }
  try {
    const c = await sql`select * from coils where coalesce(remarks,'') not ilike '%cancelled%' order by excel_row nulls last, created_at`;
    const rows = c.map((r) => ({
      "Entry Date": r.entry_date,
      "NMDC SO No": r.so_no,
      "NMDC Invoice / DO No": r.do_no,
      "Ref": r.coil_no,
      "No of Coils": r.n_coils ?? "",
      "Form": r.form,
      "Grade": r.grade,
      "Thk (mm)": r.thk == null ? "" : Number(r.thk),
      "Width (mm)": r.width ?? "",
      "Length (mm)": r.length ?? "",
      "Weight (MT)": r.weight == null ? "" : Number(r.weight),
      "Status": r.status,
      "Vehicle No": r.vehicle_no || "",
      "Transporter Name": r.transporter_name || "",
      "Transporter Mobile": r.transporter_mobile || "",
      "LR No": r.lr_no || "",
      "Dispatch Date (ex NMDC)": r.dispatch_date,
      "Expected Arrival": r.expected_arrival,
      "Actual Arrival": r.actual_arrival,
      "Receiving Unit": r.destination || "",
      "Remarks": r.remarks || "",
      "Entered By": r.entered_by || "BSC"
    }));
    const sos = (await sql`select so_no, so_date, description, ordered_mt, ordered_coils, remarks from sos order by so_no`).map((x) => ({ so_no: x.so_no, so_date: x.so_date, description: x.description || "", ordered_mt: x.ordered_mt == null ? null : Number(x.ordered_mt), ordered_coils: x.ordered_coils, remarks: x.remarks || "" }));
    json(res, 200, { rows, sos, fetchedAt: (/* @__PURE__ */ new Date()).toISOString() });
  } catch (e) {
    json(res, 502, { error: e.message });
  }
}

// api/sync.js
import crypto from "node:crypto";

// lib/graph.js
var G = "https://graph.microsoft.com/v1.0";
var tok = { v: null, exp: 0 };
var itemRef = null;
var layout = null;
async function token() {
  if (tok.v && Date.now() < tok.exp - 6e4) return tok.v;
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials"
  });
  const r = await fetch(
    `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body }
  );
  const j = await r.json();
  if (!r.ok) throw new Error("Token error: " + (j.error_description || r.status));
  tok = { v: j.access_token, exp: Date.now() + j.expires_in * 1e3 };
  return j.access_token;
}
async function graph(path, t, opts = {}) {
  const r = await fetch(G + path, {
    method: opts.method || "GET",
    body: opts.body,
    headers: { Authorization: "Bearer " + t, "Cache-Control": "no-cache", "Content-Type": "application/json", ...opts.headers || {} }
  });
  if (!r.ok) throw new Error(`Graph ${r.status} ${path.slice(0, 80)}: ${(await r.text()).slice(0, 300)}`);
  return r.status === 204 ? null : r.json();
}
var shareId = (u) => "u!" + Buffer.from(u).toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
var TABLE = () => encodeURIComponent(process.env.TABLE_NAME || "tbl_Dispatch");
async function workbook(t) {
  if (!itemRef) {
    let it;
    if (process.env.DRIVE_USER && process.env.FILE_PATH) {
      const p = process.env.FILE_PATH.replace(/^\/+/, "").split("/").map(encodeURIComponent).join("/");
      it = await graph(`/users/${encodeURIComponent(process.env.DRIVE_USER)}/drive/root:/${p}:?$select=id,parentReference`, t);
    } else {
      it = await graph(`/shares/${shareId(process.env.SHARE_LINK)}/driveItem?$select=id,parentReference`, t);
    }
    itemRef = { driveId: it.parentReference.driveId, itemId: it.id };
  }
  return `/drives/${itemRef.driveId}/items/${itemRef.itemId}/workbook`;
}
async function withSession(persist, fn) {
  const t = await token();
  const wb = await workbook(t);
  const s2 = await graph(wb + "/createSession", t, { method: "POST", body: JSON.stringify({ persistChanges: persist }) });
  const h = { "workbook-session-id": s2.id };
  try {
    return await fn({ t, wb, h, get: (p) => graph(wb + p, t, { headers: h }), patch: (p, body) => graph(wb + p, t, { method: "PATCH", headers: h, body: JSON.stringify(body) }) });
  } finally {
    graph(wb + "/closeSession", t, { method: "POST", headers: h }).catch(() => {
    });
  }
}
async function tableLayout(s2) {
  if (layout) return layout;
  const r = await s2.get(`/tables/${TABLE()}/range?$select=address`);
  const m = /^(.*)!\$?([A-Z]+)\$?(\d+):/.exec(r.address);
  layout = { sheet: m[1].replace(/^'|'$/g, ""), col: m[2], headerRow: parseInt(m[3], 10) };
  return layout;
}
async function readTable(s2) {
  const [hdr, body] = await Promise.all([s2.get(`/tables/${TABLE()}/headerRowRange?$select=values`), s2.get(`/tables/${TABLE()}/rows?$select=values`)]);
  const headers = hdr.values[0].map((h) => String(h).trim());
  return { headers, rows: body.value.map((x, i) => ({ i, raw: x.values[0], obj: Object.fromEntries(headers.map((h, k) => [h, x.values[0][k]])) })) };
}
var colLetter = (n2) => {
  let s2 = "";
  n2 += 1;
  while (n2 > 0) {
    const m = (n2 - 1) % 26;
    s2 = String.fromCharCode(65 + m) + s2;
    n2 = Math.floor((n2 - 1) / 26);
  }
  return s2;
};
var serial = (iso) => iso ? Math.round(Date.parse(String(iso).slice(0, 10) + "T00:00:00Z") / 864e5 + 25569) : "";
var fromSerial = (v) => typeof v === "number" && v > 2e4 ? new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10) : typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v) ? v.slice(0, 10) : null;

// lib/wati.js
function watiBase(){ return process.env.WATI_API_URL || process.env.WATI_ENDPOINT || process.env.WATI_BASE_URL || ""; }
function watiKey(){ return (process.env.WATI_API_KEY || process.env.WATI_ACCESS_TOKEN || process.env.WATI_TOKEN || "").replace(/^Bearer\s+/i, ""); }
async function send(mobile, template, parameters) {
  const base = watiBase(), key = watiKey();
  if (!base || !key) return { skipped: "WATI not configured" };
  const num = String(mobile).replace(/\D/g, "").replace(/^0+/, "");
  const wa = num.length === 10 ? "91" + num : num;
  const r = await fetch(`${base.replace(/\/$/, "")}/api/v1/sendTemplateMessage?whatsappNumber=${wa}`, {
    method: "POST",
    headers: { Authorization: "Bearer " + key, "Content-Type": "application/json" },
    body: JSON.stringify({ template_name: template, broadcast_name: template, parameters })
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || j.result === false) throw new Error("WATI: " + (j.info || j.message || r.status));
  return { ok: true };
}
var sendDispatchLink = ({ mobile, name, coils, doNo, link }) => send(mobile, process.env.WATI_TEMPLATE || "nmdc_dispatch_link", [{ name: "name", value: name || "Partner" }, { name: "coils", value: String(coils) }, { name: "do_no", value: doNo || "-" }, { name: "link", value: link }]);
var sendTransporterLink = ({ mobile, name, link }) => send(mobile, process.env.WATI_TEMPLATE_TRANSPORTER || "nmdc_transporter_link", [{ name: "name", value: name || "Partner" }, { name: "link", value: link }]);

// api/sync.js
var OWN = ["Entry Date", "NMDC SO No", "NMDC Invoice / DO No", "No of Coils", "Form", "Grade", "Thk (mm)", "Width (mm)", "Length (mm)", "Weight (MT)", "Receiving Unit", "Transporter Name", "Transporter Mobile", "Remarks"];
var MIRROR = ["Status", "Vehicle No", "LR No", "Dispatch Date (ex NMDC)", "Expected Arrival", "Actual Arrival"];
var PRESENT = ["Entry Date", "NMDC SO No", "NMDC Invoice / DO No", "No of Coils", "Thk (mm)", "Weight (MT)", "Status", "Vehicle No"];
var s = (v) => v == null ? "" : String(v).trim();
var n = (v) => typeof v === "number" ? v : v === "" || v == null ? null : isNaN(Number(v)) ? null : Number(v);
var norm = (o) => {
  if (!("Receiving Unit" in o) && "Destination" in o) o["Receiving Unit"] = o["Destination"];
  return o;
};
var sha = (o, keys) => crypto.createHash("sha1").update(keys.map((k) => s(o[k])).join("|")).digest("hex");
var formOf = (v) => /^ctl/i.test(s(v)) ? "CTL" : s(v) ? "Coil" : "";
var statusOf = (v, o) => STATUSES.includes(s(v)) ? s(v) : s(o["Actual Arrival"]) ? "Arrived" : s(o["Vehicle No"]) && s(o["Dispatch Date (ex NMDC)"]) ? "In-Transit" : "Under Loading";
async function handler2(req, res) {
  const auth = req.headers.authorization || "";
  if (!process.env.CRON_SECRET) return json(res, 503, { error: "CRON_SECRET not set" });
  const key = decodeURIComponent((/[?&]key=([^&]*)/.exec(req.url || "") || [])[1] || req.query && req.query.key || "");
  if (auth !== "Bearer " + process.env.CRON_SECRET && key !== process.env.CRON_SECRET) return json(res, 401, { error: "unauthorized" });
  const out = { sos: 0, transporters: 0, imported: 0, updated: 0, adopted: 0, dispatches: 0, notified: 0, mirrored: 0, appended: 0, errors: [] };
  const base = (process.env.PUBLIC_URL || `https://${req.headers.host}`).replace(/\/$/, "");
  try {
    await withSession(true, async (sess) => {
      try {
        const tr = await sess.get(`/tables/tbl_Transporters/rows?$select=values`);
        for (const [i, row] of tr.value.entries()) {
          const [name, mobile, link, sent, active] = row.values[0].map((v) => s(v));
          const mob = mobile.replace(/\D/g, "");
          if (!name) continue;
          let t = (await sql`select * from transporters where lower(name)=lower(${name}) or (mobile<>'' and mobile=${mob}) order by id limit 1`)[0];
          const isActive = active.toLowerCase() !== "no";
          if (!t) {
            const token2 = crypto.randomBytes(9).toString("base64url");
            t = (await sql`insert into transporters (name, mobile, token, active, excel_row) values (${name}, ${mob || "x-" + token2}, ${token2}, ${isActive}, ${i}) returning *`)[0];
            out.transporters++;
          } else if (t.name !== name || t.active !== isActive || t.excel_row !== i || mob && t.mobile !== mob) {
            await sql`update transporters set name=${name}, active=${isActive}, excel_row=${i}, mobile=${mob || t.mobile} where id=${t.id}`;
            t = { ...t, name, active: isActive, mobile: mob || t.mobile };
          }
          const wantLink = `${base}/t/${t.token}`;
          let sentVal = sent;
          if (!t.link_sent_at && isActive && mob.length >= 10) {
            try {
              const r = await sendTransporterLink({ mobile: mob, name, link: wantLink });
              if (!r.skipped) {
                await sql`update transporters set link_sent_at=now(), notify_error=null where id=${t.id}`;
                sentVal = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
                out.notified++;
              } else await sql`update transporters set notify_error=${r.skipped} where id=${t.id}`;
            } catch (e) {
              out.errors.push(e.message);
              await sql`update transporters set notify_error=${e.message} where id=${t.id}`;
            }
          }
          if (link !== wantLink || sentVal !== sent) await sess.patch(`/tables/tbl_Transporters/rows/itemAt(index=${i})`, { values: [[name, mobile, wantLink, sentVal, active || "Yes"]] });
        }
      } catch (e) {
        if (!/404|ItemNotFound/i.test(e.message)) throw e;
      }
      try {
        const so = await sess.get(`/tables/tbl_SO/rows?$select=values`);
        for (const [i, row] of so.value.entries()) {
          const [soNo, soDate, desc, mt, coils, rem] = row.values[0];
          if (!s(soNo)) continue;
          const r = await sql`insert into sos (so_no, so_date, description, ordered_mt, ordered_coils, remarks, excel_row) values (${s(soNo)}, ${fromSerial(soDate)}, ${s(desc)}, ${n(mt)}, ${n(coils)}, ${s(rem)}, ${i})
            on conflict (so_no) do update set so_date=excluded.so_date, description=excluded.description, ordered_mt=excluded.ordered_mt, ordered_coils=excluded.ordered_coils, remarks=excluded.remarks, excel_row=excluded.excel_row, updated_at=now()
            where (sos.so_date, sos.description, sos.ordered_mt, sos.ordered_coils, sos.remarks) is distinct from (excluded.so_date, excluded.description, excluded.ordered_mt, excluded.ordered_coils, excluded.remarks) returning so_no`;
          if (r.length) out.sos++;
        }
      } catch (e) {
        if (!/404|ItemNotFound/i.test(e.message)) throw e;
      }
      const tlist = await sql`select id, name, mobile from transporters`;
      const tByMobile = Object.fromEntries(tlist.filter((t) => /^\d{10,}$/.test(t.mobile)).map((t) => [t.mobile, t]));
      const tByName = Object.fromEntries(tlist.map((t) => [t.name.toLowerCase(), t]));
      const { headers, rows } = await readTable(sess);
      rows.forEach((r) => norm(r.obj));
      const hidx = Object.fromEntries(headers.map((h, i) => [h, i]));
      if (!("Receiving Unit" in hidx) && "Destination" in hidx) hidx["Receiving Unit"] = hidx["Destination"];
      for (const k of [...OWN, ...MIRROR, "Ref", "Entered By"]) if (!(k in hidx)) throw new Error(`Column "${k}" not found in ${TABLE()}`);
      const lay = await tableLayout(sess);
      const refCol = colLetter(hidx["Ref"]);
      const live = rows.filter((r) => PRESENT.some((k) => s(r.obj[k]) !== ""));
      const existing = Object.fromEntries((await sql`select coil_no, excel_hash, mirror_hash, excel_row, dispatch_id, status, source, updated_at, mirrored_at from coils`).map((c) => [c.coil_no, c]));
      let nextNo = Math.max(0, ...Object.keys(existing).map((k) => parseInt((/^D-(\d+)$/.exec(k) || [])[1] || "0", 10)), ...live.map((r) => parseInt((/^D-(\d+)$/.exec(s(r.obj["Ref"])) || [])[1] || "0", 10))) + 1;
      const fresh = [];
      for (const r of live) {
        const o = r.obj, sheetRow = lay.headerRow + 1 + r.i;
        let ref = s(o["Ref"]);
        if (!ref || !/^D-\d+$/.test(ref)) {
          ref = "D-" + String(nextNo++).padStart(4, "0");
          await sess.patch(`/worksheets('${encodeURIComponent(lay.sheet)}')/range(address='${refCol}${sheetRow}')`, { values: [[ref]], numberFormat: [["@"]] });
          o["Ref"] = ref;
        }
        const ex = existing[ref];
        if (ex && ex.source === "transporter") {
          if (ex.excel_row !== r.i) await sql`update coils set excel_row=${r.i} where coil_no=${ref}`;
          continue;
        }
        const h = sha(o, OWN), mh = sha(o, MIRROR);
        const own = {
          entry_date: fromSerial(o["Entry Date"]),
          so_no: s(o["NMDC SO No"]),
          do_no: s(o["NMDC Invoice / DO No"]),
          n_coils: n(o["No of Coils"]),
          form: formOf(o["Form"]),
          grade: s(o["Grade"]),
          thk: n(o["Thk (mm)"]),
          width: n(o["Width (mm)"]),
          length: n(o["Length (mm)"]),
          weight: n(o["Weight (MT)"]),
          destination: s(o["Receiving Unit"]),
          transporter_name: s(o["Transporter Name"]),
          transporter_mobile: s(o["Transporter Mobile"]).replace(/\D/g, ""),
          remarks: s(o["Remarks"])
        };
        if (!own.transporter_mobile && own.transporter_name && tByName[own.transporter_name.toLowerCase()]) {
          const m = tByName[own.transporter_name.toLowerCase()].mobile;
          if (/^\d{10,}$/.test(m)) own.transporter_mobile = m;
        }
        const tid = tByMobile[own.transporter_mobile]?.id ?? tByName[own.transporter_name.toLowerCase()]?.id ?? null;
        const mir = { status: statusOf(o["Status"], o), vehicle_no: s(o["Vehicle No"]).toUpperCase(), lr_no: s(o["LR No"]), dispatch_date: fromSerial(o["Dispatch Date (ex NMDC)"]), expected_arrival: fromSerial(o["Expected Arrival"]), actual_arrival: fromSerial(o["Actual Arrival"]) };
        if (!ex) {
          await sql`insert into coils (coil_no, entry_date, so_no, do_no, n_coils, form, grade, thk, width, length, weight, destination, transporter_name, transporter_mobile, remarks,
              status, vehicle_no, lr_no, dispatch_date, expected_arrival, actual_arrival, excel_row, excel_hash, mirror_hash, mirrored_at, source, transporter_id, entered_by)
            values (${ref}, ${own.entry_date}, ${own.so_no}, ${own.do_no}, ${own.n_coils}, ${own.form}, ${own.grade}, ${own.thk}, ${own.width}, ${own.length}, ${own.weight}, ${own.destination},
              ${own.transporter_name}, ${own.transporter_mobile}, ${own.remarks}, ${mir.status}, ${mir.vehicle_no}, ${mir.lr_no}, ${mir.dispatch_date}, ${mir.expected_arrival}, ${mir.actual_arrival},
              ${r.i}, ${h}, ${mh}, now(), 'bsc', ${tid}, ${s(o["Entered By"]) || "BSC"})`;
          await sql`insert into events (coil_no, action, actor, detail) values (${ref}, 'imported', 'sync', ${JSON.stringify({ row: r.i })})`;
          out.imported++;
          if (own.transporter_mobile) fresh.push({ ref, ...own });
          continue;
        }
        if (ex.excel_hash !== h || ex.excel_row !== r.i) {
          await sql`update coils set entry_date=${own.entry_date}, so_no=${own.so_no}, do_no=${own.do_no}, n_coils=${own.n_coils}, form=${own.form}, grade=${own.grade}, thk=${own.thk}, width=${own.width}, length=${own.length},
              weight=${own.weight}, destination=${own.destination}, transporter_name=${own.transporter_name}, transporter_mobile=${own.transporter_mobile}, remarks=${own.remarks},
              excel_row=${r.i}, excel_hash=${h}, transporter_id=${tid} where coil_no=${ref}`;
          if (ex.excel_hash !== h) out.updated++;
          if (!ex.dispatch_id && own.transporter_mobile) fresh.push({ ref, ...own });
        }
        const neonNewer = ex.mirrored_at && ex.updated_at && new Date(ex.updated_at) > new Date(ex.mirrored_at);
        if (ex.mirror_hash !== mh && !neonNewer) {
          await sql`update coils set status=${mir.status}, vehicle_no=${mir.vehicle_no}, lr_no=${mir.lr_no}, dispatch_date=${mir.dispatch_date}, expected_arrival=${mir.expected_arrival}, actual_arrival=${mir.actual_arrival}, mirror_hash=${mh}, mirrored_at=now(), updated_at=now() where coil_no=${ref}`;
          await sql`insert into events (coil_no, action, actor, detail) values (${ref}, 'excel-edit', 'bsc', ${JSON.stringify(mir)})`;
          out.adopted++;
        }
      }
      const groups = {};
      for (const f of fresh) {
        const k = (f.do_no || f.so_no) + "|" + f.transporter_mobile;
        (groups[k] ||= []).push(f);
      }
      for (const k of Object.keys(groups)) {
        const g = groups[k], f = g[0], dn = f.do_no || f.so_no;
        let disp = (await sql`select * from dispatches where do_no=${dn} and transporter_mobile=${f.transporter_mobile} order by id desc limit 1`)[0];
        if (!disp) {
          disp = (await sql`insert into dispatches (token, do_no, transporter_name, transporter_mobile, transporter_id) values (${crypto.randomBytes(9).toString("base64url")}, ${dn}, ${f.transporter_name}, ${f.transporter_mobile}, ${tByMobile[f.transporter_mobile]?.id ?? null}) returning *`)[0];
          out.dispatches++;
        }
        await sql`update coils set dispatch_id=${disp.id} where coil_no = any(${g.map((x) => x.ref)})`;
        try {
          const count = (await sql`select coalesce(sum(coalesce(n_coils,1)),0)::int as c from coils where dispatch_id=${disp.id}`)[0].c;
          const standing = tByMobile[disp.transporter_mobile];
          const r = await sendDispatchLink({ mobile: disp.transporter_mobile, name: disp.transporter_name, coils: count, doNo: disp.do_no, link: `${base}/t/${standing ? standing.token : disp.token}` });
          await sql`update dispatches set notified_at=now(), notify_error=${r.skipped || null} where id=${disp.id}`;
          if (!r.skipped) out.notified++;
        } catch (e) {
          out.errors.push(e.message);
          await sql`update dispatches set notify_error=${e.message} where id=${disp.id}`;
        }
      }
      const fullRow = (c) => headers.map((h) => ({
        "Sl No": "",
        "Entry Date": serial(c.entry_date),
        "NMDC SO No": c.so_no || "",
        "NMDC Invoice / DO No": c.do_no || "",
        "No of Coils": c.n_coils ?? "",
        "Form": c.form || "",
        "Grade": c.grade || "",
        "Thk (mm)": c.thk ?? "",
        "Width (mm)": c.width ?? "",
        "Length (mm)": c.length ?? "",
        "Weight (MT)": c.weight ?? "",
        "Status": c.status,
        "Vehicle No": c.vehicle_no || "",
        "Transporter Name": c.transporter_name || "",
        "Transporter Mobile": c.transporter_mobile || "",
        "LR No": c.lr_no || "",
        "Dispatch Date (ex NMDC)": serial(c.dispatch_date),
        "Expected Arrival": serial(c.expected_arrival),
        "Actual Arrival": serial(c.actual_arrival),
        "Receiving Unit": c.destination || "",
        "Destination": c.destination || "",
        "Remarks": c.remarks || "",
        "Entered By": c.entered_by || c.transporter_name || "",
        "Ref": c.coil_no,
        "Last Updated": ""
      })[h] ?? "");
      const slFormula = (sheetRow) => `=IF(COUNTA(${colLetter(hidx["Entry Date"])}${sheetRow}:${colLetter(hidx["Entered By"])}${sheetRow})=0,"",ROW()-${lay.headerRow})`;
      const luFormula = (sheetRow) => `=IF(COUNTA(${colLetter(hidx["Entry Date"])}${sheetRow}:${colLetter(hidx["Entered By"])}${sheetRow})=0,"",NOW())`;
      const rowFmt = headers.map((h) => /Date|Arrival|Updated/.test(h) ? "dd-mmm-yy" : /Weight/.test(h) ? "#,##0.000" : /Mobile|SO No|Ref|Vehicle|LR|Invoice/.test(h) ? "@" : "General");
      const news = await sql`select * from coils where source='transporter' and excel_row is null order by created_at`;
      let nextRow = live.length;
      for (const c of news) {
        const rowIdx = nextRow++, sheetRow = lay.headerRow + 1 + rowIdx;
        if (rowIdx >= rows.length) {
          out.errors.push("tbl_Dispatch is full \u2014 extend the table");
          break;
        }
        const vals = fullRow(c);
        vals[hidx["Sl No"]] = slFormula(sheetRow);
        if ("Last Updated" in hidx) vals[hidx["Last Updated"]] = luFormula(sheetRow);
        await sess.patch(`/worksheets('${encodeURIComponent(lay.sheet)}')/range(address='${colLetter(0)}${sheetRow}:${colLetter(headers.length - 1)}${sheetRow}')`, { values: [vals], numberFormat: [rowFmt] });
        await sql`update coils set excel_row=${rowIdx}, mirrored_at=now() where coil_no=${c.coil_no}`;
        live.push({ i: rowIdx, obj: { Ref: c.coil_no } });
        out.appended++;
      }
      const dirty = await sql`select * from coils where excel_row is not null and (mirrored_at is null or updated_at > mirrored_at)`;
      const byRow = Object.fromEntries(live.map((r) => [r.i, r]));
      const firstCol = hidx["Status"], lastCol = hidx["Actual Arrival"];
      for (const c of dirty) {
        const r = byRow[c.excel_row];
        if (!r || s(r.obj["Ref"]) !== c.coil_no) {
          out.errors.push(`row moved for ${c.coil_no}`);
          continue;
        }
        const sheetRow = lay.headerRow + 1 + c.excel_row;
        if (c.source === "transporter") {
          const vals = fullRow(c);
          vals[hidx["Sl No"]] = slFormula(sheetRow);
          if ("Last Updated" in hidx) vals[hidx["Last Updated"]] = luFormula(sheetRow);
          await sess.patch(`/worksheets('${encodeURIComponent(lay.sheet)}')/range(address='${colLetter(0)}${sheetRow}:${colLetter(headers.length - 1)}${sheetRow}')`, { values: [vals] });
        } else {
          const map = {
            "Status": c.status,
            "Vehicle No": c.vehicle_no || "",
            "Transporter Name": c.transporter_name || "",
            "Transporter Mobile": c.transporter_mobile || "",
            "LR No": c.lr_no || "",
            "Dispatch Date (ex NMDC)": serial(c.dispatch_date),
            "Expected Arrival": serial(c.expected_arrival),
            "Actual Arrival": serial(c.actual_arrival)
          };
          const vals = headers.slice(firstCol, lastCol + 1).map((h) => h in map ? map[h] : r.obj[h] ?? "");
          const fmt = headers.slice(firstCol, lastCol + 1).map((h) => /Date|Arrival/.test(h) ? "dd-mmm-yy" : /Mobile|Vehicle|LR/.test(h) ? "@" : "General");
          await sess.patch(`/worksheets('${encodeURIComponent(lay.sheet)}')/range(address='${colLetter(firstCol)}${sheetRow}:${colLetter(lastCol)}${sheetRow}')`, { values: [vals], numberFormat: [fmt] });
        }
        const o2 = { ...r.obj, "Status": c.status, "Vehicle No": c.vehicle_no || "", "LR No": c.lr_no || "", "Dispatch Date (ex NMDC)": serial(c.dispatch_date), "Expected Arrival": serial(c.expected_arrival), "Actual Arrival": serial(c.actual_arrival) };
        await sql`update coils set mirrored_at=now(), mirror_hash=${sha(o2, MIRROR)} where coil_no=${c.coil_no}`;
        out.mirrored++;
      }
    });
    json(res, 200, { ok: true, ...out });
  } catch (e) {
    json(res, 500, { ok: false, ...out, error: e.message });
  }
}

// api/tform.js
import crypto2 from "node:crypto";
var VEH = /^[A-Z]{2}[ -]?\d{1,2}[ -]?[A-Z]{0,3}[ -]?\d{4}$/i;
var destCache = { at: 0, list: [] };
async function destinations() {
  if (Date.now() - destCache.at < 10 * 6e4 && destCache.list.length) return destCache.list;
  try {
    const list = await withSession(false, async (s2) => (await s2.get(`/worksheets('Lists')/range(address='C2:C31')?$select=values`)).values.flat().map((v) => String(v).trim()).filter(Boolean));
    destCache = { at: Date.now(), list };
  } catch {
  }
  return destCache.list;
}
var group = (lines) => {
  const g = {};
  for (const c of lines) (g[c.do_no || (c.so_no ? "SO " + c.so_no : "(no invoice yet)")] ||= []).push(c);
  return Object.entries(g).map(([do_no, coils]) => ({ do_no, coils }));
};
async function handler3(req, res) {
  const qtoken = decodeURIComponent((/[?&]token=([^&]*)/.exec(req.url || "") || [])[1] || req.query && req.query.token || "");
  const token2 = String((req.method === "GET" ? qtoken : req.body?.token) || "");
  if (!token2) return json(res, 400, { error: "token missing" });
  let tp = (await sql`select * from transporters where token=${token2} and active`)[0];
  let disp = null;
  if (!tp) {
    disp = (await sql`select * from dispatches where token=${token2}`)[0];
    if (!disp) return json(res, 404, { error: "This link is not valid. Please contact Bharat Steel." });
    if (disp.transporter_id) tp = (await sql`select * from transporters where id=${disp.transporter_id} and active`)[0] || null;
  }
  const name = tp?.name || disp?.transporter_name || "";
  const mobile = tp?.mobile || disp?.transporter_mobile || "";
  const actor = "transporter:" + mobile;
  const mine = () => tp ? sql`select coil_no as ref, do_no, so_no, n_coils, form, grade, thk, width, length, weight, destination, status, vehicle_no, lr_no, dispatch_date, expected_arrival, actual_arrival, source, entered_by
          from coils where (transporter_id=${tp.id} or (transporter_mobile<>'' and transporter_mobile=${tp.mobile})) and coalesce(remarks,'') not ilike '%cancelled%' order by created_at desc, coil_no` : sql`select coil_no as ref, do_no, so_no, n_coils, form, grade, thk, width, length, weight, destination, status, vehicle_no, lr_no, dispatch_date, expected_arrival, actual_arrival, source, entered_by
          from coils where dispatch_id=${disp.id} order by coil_no`;
  if (req.method === "GET") return json(res, 200, { mode: tp ? "transporter" : "dispatch", name, do_no: disp?.do_no || null, groups: group(await mine()), destinations: await destinations() });
  if (req.method !== "POST") return json(res, 405, { error: "method" });
  const b = req.body || {}, action = String(b.action || ""), today = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
  if (action === "create") {
    if (!tp) return json(res, 403, { error: "This link cannot create new dispatches" });
    const doNo = String(b.do_no || "").trim(), soNo = String(b.so_no || "").trim(), dest = String(b.destination || "").trim();
    const veh = String(b.vehicle_no || "").toUpperCase().replace(/\s+/g, " ").trim(), lr = String(b.lr_no || "").trim();
    const dd = d(b.dispatch_date) || today, ea = d(b.expected_arrival);
    const lines = (Array.isArray(b.lines) ? b.lines : []).map((c) => ({
      n_coils: c.n_coils === "" || c.n_coils == null ? null : Number(c.n_coils),
      form: /^ctl/i.test(String(c.form || "")) ? "CTL" : "Coil",
      grade: String(c.grade || "").trim(),
      thk: Number(c.thk) || null,
      width: Number(c.width) || null,
      length: Number(c.length) || null,
      weight: Number(c.weight) || null
    }));
    if (!doNo && !soNo) return json(res, 400, { error: "Enter the NMDC invoice / DO number (or the SO number if the invoice is not issued yet)" });
    if (!dest) return json(res, 400, { error: "Select where the material is going" });
    if (!lines.length) return json(res, 400, { error: "Add at least one line" });
    if (lines.some((c) => !(c.thk > 0) || !(c.width > 0))) return json(res, 400, { error: "Enter thickness and width for every line" });
    if (lines.some((c) => c.form === "Coil" && !(c.n_coils > 0))) return json(res, 400, { error: "Enter the number of coils for each coil line" });
    if (veh && !VEH.test(veh)) return json(res, 400, { error: "Enter a valid vehicle number, e.g. TN 20 BX 7719" });
    if (veh && !ea) return json(res, 400, { error: "Enter the expected arrival date" });
    const status = veh ? "Loaded" : "Under Loading";
    const dn = doNo || soNo;
    let dsp = (await sql`select * from dispatches where do_no=${dn} and transporter_mobile=${tp.mobile} order by id desc limit 1`)[0];
    if (!dsp) dsp = (await sql`insert into dispatches (token, do_no, transporter_name, transporter_mobile, transporter_id, notified_at) values (${crypto2.randomBytes(9).toString("base64url")}, ${dn}, ${tp.name}, ${tp.mobile}, ${tp.id}, now()) returning *`)[0];
    const last = (await sql`select max(cast(substring(coil_no from 3) as integer)) as m from coils where coil_no ~ '^D-[0-9]+$'`)[0].m || 0;
    let no = last + 1;
    for (const c of lines) {
      const ref = "D-" + String(no++).padStart(4, "0");
      await sql`insert into coils (coil_no, entry_date, so_no, do_no, n_coils, form, grade, thk, width, length, weight, destination, transporter_name, transporter_mobile, transporter_id, status, vehicle_no, lr_no, dispatch_date, expected_arrival, dispatch_id, source, entered_by)
                values (${ref}, ${today}, ${soNo}, ${doNo}, ${c.n_coils}, ${c.form}, ${c.grade}, ${c.thk}, ${c.width}, ${c.length}, ${c.weight}, ${dest}, ${tp.name}, ${/^\d{10,}$/.test(tp.mobile) ? tp.mobile : ""}, ${tp.id}, ${status}, ${veh}, ${lr}, ${veh ? dd : null}, ${ea}, ${dsp.id}, 'transporter', ${tp.name})`;
      await sql`insert into events (coil_no, action, actor, detail) values (${ref}, 'created', ${actor}, ${JSON.stringify({ doNo, soNo, dest, veh, lr })})`;
    }
    return json(res, 200, { ok: true, groups: group(await mine()) });
  }
  const sel = Array.isArray(b.coils) ? b.coils.map(String) : [];
  if (!sel.length) return json(res, 400, { error: "Tick at least one line" });
  const own = tp ? await sql`select coil_no, status, source from coils where coil_no = any(${sel}) and (transporter_id=${tp.id} or (transporter_mobile<>'' and transporter_mobile=${tp.mobile}))` : await sql`select coil_no, status, source from coils where coil_no = any(${sel}) and dispatch_id=${disp.id}`;
  if (own.length !== sel.length) return json(res, 403, { error: "One of the lines is not on this link" });
  if (action === "load") {
    const veh = String(b.vehicle_no || "").toUpperCase().replace(/\s+/g, " ").trim(), lr = String(b.lr_no || "").trim();
    const dd = d(b.dispatch_date) || today, ea = d(b.expected_arrival);
    if (!VEH.test(veh)) return json(res, 400, { error: "Enter a valid vehicle number, e.g. TN 20 BX 7719" });
    if (!ea) return json(res, 400, { error: "Enter the expected arrival date" });
    await sql`update coils set vehicle_no=${veh}, lr_no=${lr}, dispatch_date=${dd}, expected_arrival=${ea}, status = case when status='Under Loading' then 'Loaded' else status end, updated_at=now() where coil_no = any(${sel}) and status <> 'Arrived'`;
    await sql`insert into events (coil_no, action, actor, detail) select unnest(${sel}::text[]), 'load', ${actor}, ${JSON.stringify({ veh, lr, dd, ea })}`;
  } else if (action === "gateout") {
    const blocked = own.filter((c) => c.status === "Under Loading");
    if (blocked.length) return json(res, 400, { error: "Enter truck details first for " + blocked.map((c) => c.coil_no).join(", ") });
    await sql`update coils set status='In-Transit', dispatch_date=coalesce(dispatch_date, ${today}), updated_at=now() where coil_no = any(${sel}) and status='Loaded'`;
    await sql`insert into events (coil_no, action, actor) select unnest(${sel}::text[]), 'gateout', ${actor}`;
  } else if (action === "delivered") {
    await sql`update coils set status='Arrived', actual_arrival=coalesce(actual_arrival, ${d(b.actual_arrival) || today}), updated_at=now() where coil_no = any(${sel}) and status in ('Loaded','In-Transit')`;
    await sql`insert into events (coil_no, action, actor) select unnest(${sel}::text[]), 'delivered', ${actor}`;
  } else if (action === "edit") {
    const editable = own.filter((c) => c.source === "transporter" && c.status !== "Arrived");
    if (editable.length !== own.length) return json(res, 403, { error: "Only your own undelivered entries can be edited" });
    const w = b.weight == null ? null : Number(b.weight), dest = b.destination == null ? null : String(b.destination).trim(), inv = b.do_no == null ? null : String(b.do_no).trim();
    await sql`update coils set weight=coalesce(${w}, weight), destination=coalesce(${dest}, destination), do_no=coalesce(${inv}, do_no), updated_at=now() where coil_no = any(${sel})`;
    await sql`insert into events (coil_no, action, actor, detail) select unnest(${sel}::text[]), 'edit', ${actor}, ${JSON.stringify({ w, dest })}`;
  } else return json(res, 400, { error: "unknown action" });
  json(res, 200, { ok: true, groups: group(await mine()) });
}

// api/arrive.js
async function handler4(req, res) {
  if (cors(req, res)) return;
  if (req.method !== "POST") return json(res, 405, { error: "method" });
  const b = req.body || {};
  if (!process.env.YARD_PIN) return json(res, 503, { error: "YARD_PIN not set" });
  if (String(b.pin || "") !== process.env.YARD_PIN) return json(res, 401, { error: "Wrong PIN" });
  const today = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
  let rows;
  if (b.vehicle_no) rows = await sql`update coils set status='Arrived', actual_arrival=coalesce(actual_arrival, ${today}), updated_at=now()
                                       where upper(regexp_replace(vehicle_no,'\s','','g')) = upper(regexp_replace(${String(b.vehicle_no)},'\s','','g')) and status in ('Loaded','In-Transit') returning coil_no`;
  else if (Array.isArray(b.coils)) rows = await sql`update coils set status='Arrived', actual_arrival=coalesce(actual_arrival, ${today}), updated_at=now() where coil_no = any(${b.coils.map(String)}) and status in ('Loaded','In-Transit') returning coil_no`;
  else return json(res, 400, { error: "vehicle_no or coils required" });
  for (const r of rows) await sql`insert into events (coil_no, action, actor) values (${r.coil_no}, 'delivered', 'yard')`;
  json(res, 200, { ok: true, updated: rows.map((r) => r.coil_no) });
}


// wheels.js — settings + WIP exclusions for the Wheels India module (Neon-backed).
async function ensureWheelsTable(){
  await sql`create table if not exists wheels_settings (
    id int primary key default 1,
    mail_to text[] default '{}', mail_cc text[] default '{}',
    send_time text default '10:00', enabled boolean default true,
    wip text[] default '{}', mail_subject text default '', mail_body text default '', updated_at timestamptz default now()
  )`;
  await sql`insert into wheels_settings (id) values (1) on conflict (id) do nothing`;
  await sql`alter table wheels_settings add column if not exists mail_subject text default ''`;
  await sql`alter table wheels_settings add column if not exists mail_body text default ''`;
  await sql`alter table wheels_settings add column if not exists wa_to text[] default '{}'`;
  await sql`alter table wheels_settings add column if not exists wa_enabled boolean default true`;
  await sql`alter table wheels_settings add column if not exists prod_mail_to text[] default '{}'`;
  await sql`alter table wheels_settings add column if not exists prod_wa_to text[] default '{}'`;
  await sql`alter table wheels_settings add column if not exists prod_enabled boolean default true`;
  await sql`alter table wheels_settings add column if not exists prod_format text default 'image'`;
}
async function handlerWheels(req, res){
  if (cors(req, res)) return;
  try{
    await ensureWheelsTable();
    // GET /api/wheels?fmt=pdf  ->  the statement as a PDF (same data/structure as the email).
    // Public, no-auth: this is the URL WATI fetches for the WhatsApp document template.
    const fmt = (req.query && req.query.fmt) || ((/[?&]fmt=(pdf|xlsx|prodpdf|prodpng)\b/.exec(req.url || "")) || [])[1] || "";
    if (req.method === "GET" && fmt === "pdf"){
      const pdf = await buildWheelsPdf();
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", 'inline; filename="WheelsIndia_Stock_' + new Date().toISOString().slice(0,10).replace(/-/g,"") + '.pdf"');
      res.setHeader("Cache-Control", "no-store");
      return res.status(200).end(pdf);
    }
    if (req.method === "GET" && fmt === "xlsx"){
      const b64 = await buildWheelsXlsx(await wheelsRows());
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", 'attachment; filename="WheelsIndia_Stock_' + new Date().toISOString().slice(0,10).replace(/-/g,"") + '.xlsx"');
      res.setHeader("Cache-Control", "no-store");
      return res.status(200).end(Buffer.from(b64, "base64"));
    }
    if (req.method === "GET" && fmt === "prodpdf"){
      const pdf = await buildWheelsProdPdf();
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", 'inline; filename="WheelsIndia_Production_' + new Date().toISOString().slice(0,10).replace(/-/g,"") + '.pdf"');
      res.setHeader("Cache-Control", "no-store");
      return res.status(200).end(pdf);
    }
    if (req.method === "GET" && fmt === "prodpng"){
      const png = await buildWheelsProdPng();
      res.setHeader("Content-Type", "image/png");
      res.setHeader("Content-Disposition", 'inline; filename="WheelsIndia_Production_' + new Date().toISOString().slice(0,10).replace(/-/g,"") + '.png"');
      res.setHeader("Cache-Control", "no-store");
      return res.status(200).end(png);
    }
    // Diagnostic: which WhatsApp env vars does this deployment actually see? (names/booleans only, no secrets)
    if (req.method === "GET" && fmt === "wadiag"){
      return json(res, 200, {
        WHEELS_WA_TEMPLATE: process.env.WHEELS_WA_TEMPLATE || null,
        WATI_API_URL: !!watiBase(),
        WATI_key_present: !!watiKey(),
        CRON_SECRET_present: !!process.env.CRON_SECRET,
        wati_url_names: { WATI_API_URL: !!process.env.WATI_API_URL, WATI_ENDPOINT: !!process.env.WATI_ENDPOINT, WATI_BASE_URL: !!process.env.WATI_BASE_URL },
        wati_key_names: { WATI_API_KEY: !!process.env.WATI_API_KEY, WATI_ACCESS_TOKEN: !!process.env.WATI_ACCESS_TOKEN, WATI_TOKEN: !!process.env.WATI_TOKEN }
      });
    }
    if (req.method === "POST"){
      let body = req.body;
      if (typeof body === "string") { try { body = JSON.parse(body); } catch { body = {}; } }
      body = body || {};
      // WhatsApp send trigger (fired by the daily script after the email). Guarded by CRON_SECRET.
      // Creds + template come from Vercel env (WATI_API_URL, WATI_API_KEY, WHEELS_WA_TEMPLATE);
      // recipient numbers come from the in-app settings (wa_to). PDF is served at /api/wheels?fmt=pdf.
      if (body.whatsapp) {
        const auth = req.headers.authorization || "";
        if (!process.env.CRON_SECRET) return json(res, 503, { error: "CRON_SECRET not set" });
        if (auth !== "Bearer " + process.env.CRON_SECRET) return json(res, 401, { error: "unauthorized" });
        const tmpl = process.env.WHEELS_WA_TEMPLATE;
        if (!tmpl) return json(res, 400, { error: "WHEELS_WA_TEMPLATE not set" });
        const r3 = await sql`select wa_to, wa_enabled from wheels_settings where id=1`;
        const s3 = r3[0] || {};
        if (s3.wa_enabled === false) return json(res, 200, { ok: true, skipped: "whatsapp disabled in settings" });
        const nums = (s3.wa_to || []).map(String).map(x => x.trim()).filter(Boolean);
        if (!nums.length) return json(res, 200, { ok: true, skipped: "no whatsapp recipients" });
        const base = (process.env.PUBLIC_URL || `https://${req.headers.host}`).replace(/\/$/, "");
        const sent = [];
        for (const num of nums) {
          try { const r = await send(num, tmpl, []); sent.push({ num, ...(r || { ok: true }) }); }   // static URL buttons -> no params
          catch (e) { sent.push({ num, error: String(e.message || e) }); }
        }
        return json(res, 200, { ok: true, pdf: base + "/wheels.pdf", xlsx: base + "/wheels.xlsx", template: tmpl, sent });
      }
      // WhatsApp test send (triggered from the settings page, like the email test — no CRON_SECRET).
      // Sends to wa_test_to if given, else the first saved number. Uses the same env creds/template.
      if (body.whatsapp_test) {
        const tmpl = process.env.WHEELS_WA_TEMPLATE;
        if (!tmpl) return json(res, 200, { ok: true, skipped: "WHEELS_WA_TEMPLATE not set on server" });
        if (!watiBase() || !watiKey()) return json(res, 200, { ok: true, skipped: "WATI not configured on server (need WATI_API_URL + token)" });
        let num = String(body.wa_test_to || "").trim();
        if (!num) {
          const r4 = await sql`select wa_to from wheels_settings where id=1`;
          num = ((r4[0]?.wa_to) || []).map(String)[0] || "";
        }
        if (!num) return json(res, 200, { ok: true, skipped: "no test number (enter one or save a number first)" });
        const base = (process.env.PUBLIC_URL || `https://${req.headers.host}`).replace(/\/$/, "");
        const sent = [];
        try { const r = await send(num, tmpl, []); sent.push({ num, ...(r || { ok: true }) }); }   // static URL buttons -> no params
        catch (e) { sent.push({ num, error: String(e.message || e) }); }
        return json(res, 200, { ok: true, pdf: base + "/wheels.pdf", xlsx: base + "/wheels.xlsx", template: tmpl, sent });
      }
      // Production WhatsApp — daily trigger (CRON_SECRET). Template WHEELS_WA_PROD_TEMPLATE, numbers from prod_wa_to,
      // media = /wheels-production.(png|pdf) per prod_format. Header variable must be named "report".
      if (body.prod_whatsapp) {
        const auth = req.headers.authorization || "";
        if (!process.env.CRON_SECRET) return json(res, 503, { error: "CRON_SECRET not set" });
        if (auth !== "Bearer " + process.env.CRON_SECRET) return json(res, 401, { error: "unauthorized" });
        const tmpl = process.env.WHEELS_WA_PROD_TEMPLATE;
        if (!tmpl) return json(res, 400, { error: "WHEELS_WA_PROD_TEMPLATE not set" });
        const r5 = await sql`select prod_wa_to, prod_enabled, prod_format from wheels_settings where id=1`;
        const s5 = r5[0] || {};
        if (s5.prod_enabled === false) return json(res, 200, { ok: true, skipped: "production whatsapp disabled in settings" });
        const nums = (s5.prod_wa_to || []).map(String).map(x => x.trim()).filter(Boolean);
        if (!nums.length) return json(res, 200, { ok: true, skipped: "no production whatsapp recipients" });
        const base = (process.env.PUBLIC_URL || `https://${req.headers.host}`).replace(/\/$/, "");
        const url = base + (s5.prod_format === "pdf" ? "/wheels-production.pdf" : "/wheels-production.png");
        const sent = [];
        for (const num of nums) { try { const r = await send(num, tmpl, [{ name: "report", value: url }]); sent.push({ num, ...(r || { ok: true }) }); } catch (e) { sent.push({ num, error: String(e.message || e) }); } }
        return json(res, 200, { ok: true, url, template: tmpl, sent });
      }
      // Production email — test send (and reused by nothing else). Renders the chosen format and emails it.
      if (body.prod_mail_test) {
        const r7 = await sql`select prod_mail_to, prod_format from wheels_settings where id=1`;
        const s7 = r7[0] || {};
        const to = body.test_to ? [String(body.test_to)] : (s7.prod_mail_to || []);
        if (!to.length) return json(res, 400, { error: "no recipient" });
        const fmtv = (s7.prod_format) || "image";
        const stamp = new Date().toISOString().slice(0,10).replace(/-/g,"");
        let attach, html;
        if (fmtv === "pdf") {
          attach = { name:"WheelsIndia_Production_"+stamp+".pdf", b64:(await buildWheelsProdPdf()).toString("base64"), contentType:"application/pdf" };
          html = "<p>Dear Team,</p><p>Please find attached the Wheels India Production Report (PDF).</p><p>Bharat Steel (Chennai) Pvt. Ltd.</p>";
        } else {
          attach = { name:"WheelsIndia_Production_"+stamp+".png", b64:(await buildWheelsProdPng()).toString("base64"), contentType:"image/png" };
          html = `<p>Dear Team,</p><p>Wheels India Production Report:</p><img src="cid:prodimg" style="max-width:100%;border:1px solid #d0d7dd"/><p>Bharat Steel (Chennai) Pvt. Ltd.</p>`;
          attach.cid = "prodimg";
        }
        await sendWheelsMail(to, [], html, (body.test_to?"[TEST] ":"")+"Wheels India — Production Report — "+new Date().toLocaleDateString("en-IN"), attach);
        return json(res, 200, { ok:true, sent_to:to, format:fmtv });
      }
      // Production WhatsApp — test send from the settings page (no CRON_SECRET)
      if (body.prod_whatsapp_test) {
        const tmpl = process.env.WHEELS_WA_PROD_TEMPLATE;
        if (!tmpl) return json(res, 200, { ok: true, skipped: "WHEELS_WA_PROD_TEMPLATE not set on server" });
        if (!watiBase() || !watiKey()) return json(res, 200, { ok: true, skipped: "WATI not configured on server (need WATI_API_URL + token)" });
        const r6 = await sql`select prod_wa_to, prod_format from wheels_settings where id=1`;
        const s6 = r6[0] || {};
        let num = String(body.prod_wa_test_to || "").trim();
        if (!num) num = (s6.prod_wa_to || []).map(String)[0] || "";
        if (!num) return json(res, 200, { ok: true, skipped: "no test number (enter one or save a number first)" });
        const base = (process.env.PUBLIC_URL || `https://${req.headers.host}`).replace(/\/$/, "");
        const url = base + (s6.prod_format === "pdf" ? "/wheels-production.pdf" : "/wheels-production.png");
        const sent = [];
        try { const r = await send(num, tmpl, [{ name: "report", value: url }]); sent.push({ num, ...(r || { ok: true }) }); } catch (e) { sent.push({ num, error: String(e.message || e) }); }
        return json(res, 200, { ok: true, url, template: tmpl, sent });
      }
      // update only the keys provided (settings save and WIP save are separate posts)
      if (Array.isArray(body.mail_to)) await sql`update wheels_settings set mail_to=${body.mail_to} where id=1`;
      if (Array.isArray(body.mail_cc)) await sql`update wheels_settings set mail_cc=${body.mail_cc} where id=1`;
      if (typeof body.send_time === "string") await sql`update wheels_settings set send_time=${body.send_time} where id=1`;
      if (typeof body.enabled === "boolean") await sql`update wheels_settings set enabled=${body.enabled} where id=1`;
      if (typeof body.mail_subject === "string") await sql`update wheels_settings set mail_subject=${body.mail_subject} where id=1`;
      if (typeof body.mail_body === "string") await sql`update wheels_settings set mail_body=${body.mail_body} where id=1`;
      if (Array.isArray(body.wip)) await sql`update wheels_settings set wip=${body.wip.map(String)} where id=1`;
      if (Array.isArray(body.wa_to)) await sql`update wheels_settings set wa_to=${body.wa_to.map(String)} where id=1`;
      if (typeof body.wa_enabled === "boolean") await sql`update wheels_settings set wa_enabled=${body.wa_enabled} where id=1`;
      if (Array.isArray(body.prod_mail_to)) await sql`update wheels_settings set prod_mail_to=${body.prod_mail_to.map(String)} where id=1`;
      if (Array.isArray(body.prod_wa_to)) await sql`update wheels_settings set prod_wa_to=${body.prod_wa_to.map(String)} where id=1`;
      if (typeof body.prod_enabled === "boolean") await sql`update wheels_settings set prod_enabled=${body.prod_enabled} where id=1`;
      if (body.prod_format==="pdf"||body.prod_format==="image") await sql`update wheels_settings set prod_format=${body.prod_format} where id=1`;
      // test send: render from the last CSV and email to the given address (or configured To)
      if (body.send) {
        const rows = await sql`select mail_to, mail_cc from wheels_settings where id=1`;
        const s2 = rows[0] || {};
        const wipSet = new Set();   // WIP exclusion removed — all stock is shown
        const base = (process.env.PUBLIC_URL || `https://${req.headers.host}`).replace(/\/$/, "");
        const note = typeof body.mail_body === "string" ? body.mail_body : (s2.mail_body||"");
        const { html, gb, gq, rows: wrows } = await renderWheelsHtml(base, wipSet, note);
        const to = body.test_to ? [String(body.test_to)] : (s2.mail_to||[]);
        const cc = body.test_to ? [] : (s2.mail_cc||[]);
        if (!to.length) return json(res, 400, { error: "no recipient" });
        const baseSubj = (body.mail_subject || s2.mail_subject || "BSC — Wheels India Job Work Stock Statement — " + new Date().toLocaleDateString("en-IN"));
        const subj = (body.test_to ? "[TEST] " : "") + baseSubj;
        let attach = null, attachErr = null;
        try {
          const xlsxB64 = await buildWheelsXlsx(wrows);
          attach = { name: "WheelsIndia_Stock_" + new Date().toISOString().slice(0,10).replace(/-/g,"") + ".xlsx", b64: xlsxB64 };
        } catch(e){ attachErr = String(e.message||e); }
        await sendWheelsMail(to, cc, html, subj, attach);
        await sql`update wheels_settings set updated_at=now() where id=1`;
        return json(res, 200, { ok: true, sent_to: to, items: gb, mt: Number(gq.toFixed(3)), attached: !!attach, attach_error: attachErr });
      }
      await sql`update wheels_settings set updated_at=now() where id=1`;
      return json(res, 200, { ok: true });
    }
    const rows = await sql`select mail_to, mail_cc, send_time, enabled, wip, mail_subject, mail_body, wa_to, wa_enabled, prod_mail_to, prod_wa_to, prod_enabled, prod_format from wheels_settings where id=1`;
    const s = rows[0] || {};
    return json(res, 200, {
      mail_to: s.mail_to || [], mail_cc: s.mail_cc || [],
      send_time: s.send_time || "10:00", enabled: s.enabled !== false,
      wip: s.wip || [], mail_subject: s.mail_subject || "", mail_body: s.mail_body || "",
      wa_to: s.wa_to || [], wa_enabled: s.wa_enabled !== false,
      prod_mail_to: s.prod_mail_to || [], prod_wa_to: s.prod_wa_to || [],
      prod_enabled: s.prod_enabled !== false, prod_format: s.prod_format || "image"
    });
  }catch(e){
    return json(res, 500, { error: String(e.message || e) });
  }
}


// wheels-mail.js — render the branded statement from the last-pushed CSV and send via Graph.
function wheelsFormOf(n){ n=(n||"").toUpperCase();
  if(n.includes("SCRAP"))return "SCRAP";
  if(n.includes("BLANK")||n.includes("FLAT")||n.includes("STRIP"))return "STRIP";
  if(n.includes("PLATE"))return "PLATE";
  if(n.includes("COIL"))return "COIL"; return "OTHER"; }
function parseCsvSrv(t){ if(!t)return[]; if(t.charCodeAt(0)===0xFEFF)t=t.slice(1);
  const rows=[];let row=[],cur="",q=false;
  for(let i=0;i<t.length;i++){const c=t[i];
    if(q){if(c==='"'){if(t[i+1]==='"'){cur+='"';i++;}else q=false;}else cur+=c;}
    else{if(c==='"')q=true;else if(c===",")row.push(cur),cur="";else if(c==="\n")row.push(cur),rows.push(row),row=[],cur="";else if(c==="\r"){}else cur+=c;}}
  if(cur!==""||row.length){row.push(cur);rows.push(row);}
  const h=rows.shift().map(x=>x.trim());
  return rows.filter(r=>r.length>1).map(r=>{const o={};h.forEach((k,i)=>o[k]=(r[i]??"").trim());return o;}); }
function dimS(v){ const n=parseFloat(v); return isNaN(n)?"":(""+ +n.toFixed(2)); }
function lenS(v){ const n=parseFloat(v); return (n&&n>0)?(""+ +n.toFixed(0)):"—"; }
function esc2(s){ return String(s??"").replace(/[&<>]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c])); }

async function renderWheelsHtml(baseUrl, wipSet, note){
  const DATA_BASE = process.env.DATA_BASE || "https://raw.githubusercontent.com/BSC23609/bsc-stock-data/main";
  const csvRes = await fetch(DATA_BASE + "/wheels_stock.csv?t=" + Date.now());
  const rows = parseCsvSrv(csvRes.ok ? await csvRes.text() : "").map(r=>({...r, Quantity:parseFloat(r.Quantity)||0, form:wheelsFormOf(r.ItemName)}))
    .filter(r=>!wipSet.has(String(r.BatchNum)))
    .filter(r=>!((r.form==="PLATE"&&r.Quantity<0.1)||(r.form==="STRIP"&&r.Quantity<0.011))); // plate <0.1 MT, strip <0.011 MT not listed
  const SEC=[["PLATE","Plate Stock (WIP / FG)","#0e7a4b"],["STRIP","Strip / Blank Stock","#b5651d"],["COIL","Coil Stock","#1367a6"]]; // WIP/FG (plate+strip) first, coil last
  let body="", gb=0, gq=0;
  for(const [k,t,c] of SEC){
    const rs=rows.filter(r=>r.form===k).sort((a,b)=>b.Quantity-a.Quantity); if(!rs.length)continue;
    const tq=rs.reduce((a,r)=>a+r.Quantity,0); gb+=rs.length; gq+=tq;
    body+=`<tr><td colspan="9" style="background:${c};color:#fff;font-weight:700;padding:7px 10px;font-size:13px">${t}<span style="float:right">${rs.length} items &middot; ${tq.toFixed(3)} MT</span></td></tr>`+
      `<tr style="background:#eef3f7"><th>#</th><th>Mill</th><th>Grade</th><th>Thk</th><th>Width</th><th>Length</th><th>Coil / Batch No</th><th>Qty (MT)</th><th>Age (d)</th></tr>`;
    rs.forEach((x,i)=>{ const mill=(x.ItemName||"").split(" ")[0];
      const bc=(+x.BatchCount>1)?` <span style="color:#999">(${x.BatchCount} batches)</span>`:"";
      body+=`<tr><td style="text-align:center;color:#888">${i+1}</td><td style="text-align:center">${esc2(mill)}</td><td style="text-align:center">${esc2(x.Grade)}</td><td style="text-align:center">${dimS(x.Thick)}</td><td style="text-align:center">${dimS(x.Width)}</td><td style="text-align:center">${lenS(x.Length)}</td><td style="text-align:center;font-family:monospace;font-size:11px">${esc2(x.CoilNo||x.BatchNum)}${bc}</td><td style="text-align:center;font-weight:600">${x.Quantity.toFixed(3)}</td><td style="text-align:center;color:#888">${esc2(x.AgeDays)}</td></tr>`; });
    body+=`<tr><td colspan="8" style="text-align:right;font-weight:700;padding:6px 10px;border-top:2px solid #333">${t} total</td><td style="text-align:right;font-weight:700;border-top:2px solid #333">${tq.toFixed(3)}</td><td style="border-top:2px solid #333"></td></tr>`;
  }
  const asof=new Date().toLocaleString("en-IN",{timeZone:"Asia/Kolkata",day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"});
  const noteHtml = note && note.trim() ? `<div style="padding:14px 22px 0;font-size:13px;color:#222;line-height:1.6">${esc2(note).replace(/\n/g,"<br>")}</div>` : "";
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>@media only screen and (max-width:600px){.wk-tbl{font-size:10px !important}.wk-tbl th,.wk-tbl td{padding:2px 3px !important}.wk-pad{padding:12px 8px 4px !important}.wk-hd{padding:12px 12px !important}}</style></head><body style="margin:0;padding:0;background:#fff;-webkit-text-size-adjust:100%">`+
    `<div style="font-family:Arial,Helvetica,sans-serif;color:#1a1a1a;max-width:920px;width:100%;margin:auto;border:1px solid #d0d7dd;box-sizing:border-box">
    <div style="background:#0e2a47;color:#fff;padding:14px 22px"><div style="font-size:18px;font-weight:700">Job Work Stock Statement &mdash; Wheels India Limited</div><div style="font-size:12px;opacity:.85;margin-top:3px">Bharat Steel (Chennai) Pvt. Ltd. &middot; Material held at BSC on customer's account</div></div>
    ${noteHtml}
    <div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap;padding:11px 22px;background:#f7f9fb;border-bottom:1px solid #e3e8ec;font-size:13px"><div>Customer: <b>Wheels India Limited</b></div><div>As on: <b>${asof}</b></div><div>Warehouse: <b>Job Work (WH 45)</b></div></div>
    <div class="wk-pad" style="padding:14px 22px 4px;overflow-x:auto;-webkit-overflow-scrolling:touch"><table class="wk-tbl" style="width:100%;border-collapse:collapse;font-size:12px">${body}</table></div>
    <div style="margin:14px 22px;padding:12px 16px;background:#0e2a47;color:#fff;border-radius:5px;display:flex;justify-content:space-between;font-size:15px;font-weight:700"><span>TOTAL MATERIAL HELD</span><span>${gb} items &middot; ${gq.toFixed(3)} MT</span></div>
    <div style="margin:0 22px 14px;padding:9px 13px;background:#fff8e1;border:1px solid #f0dc9a;color:#6b5200;font-size:11px;border-radius:5px">Plates currently under process (awaiting shearing) are excluded from this statement.</div>
    <div style="padding:6px 22px;font-size:9px;color:#aaa;font-style:italic">Computer generated Statement - may contain mistakes</div><div style="padding:16px 22px;font-size:11px;color:#667;border-top:1px solid #e3e8ec;line-height:1.6"><b>Bharat Steel (Chennai) Pvt. Ltd.</b> &middot; No.147, Survey No.133, Thirunilai Village, Ponneri Taluk, Chennai 600103 &middot; 044 6791 7800 &middot; info@bharatsteels.in<br>System-generated statement of your material held at BSC as on the date/time shown.</div></div></body></html>`;
  return { html, gb, gq, rows };
}

async function buildWheelsXlsx(rows){
  let XLSX;
  try { XLSX = await import("xlsx"); } catch(e){ throw new Error("xlsx module not available on server: "+e.message); }
  const SEC=[["PLATE","Plate Stock (WIP / FG)"],["STRIP","Strip / Blank Stock"],["COIL","Coil Stock"]]; // WIP/FG (plate+strip) first, coil last
  const aoa=[["BHARAT STEEL (CHENNAI) PRIVATE LIMITED"],
             ["Job Work Stock Statement — Wheels India Limited"],
             ["As on: "+new Date().toLocaleString("en-IN",{timeZone:"Asia/Kolkata"})+"   |   Customer: Wheels India Limited   |   Warehouse: Job Work (WH 45)"],
             []];
  const hdr=["#","Mill","Grade","Thk","Width","Length","Coil / Batch No","Mother Coil","Qty (MT)","Age (d)"];
  let gGrand=0, secRanges=[];
  for(const [k,t] of SEC){
    const rs=rows.filter(r=>r.form===k).sort((a,b)=>b.Quantity-a.Quantity); if(!rs.length)continue;
    const tq=rs.reduce((a,r)=>a+r.Quantity,0); gGrand+=tq;
    aoa.push([`${t}   (${rs.length} items, ${tq.toFixed(3)} MT)`]);
    aoa.push(hdr);
    rs.forEach((x,i)=>{ const mill=(x.ItemName||"").split(" ")[0];
      const T=parseFloat(x.Thick), W=parseFloat(x.Width), L=parseFloat(x.Length);
      aoa.push([i+1, mill, x.Grade, isNaN(T)?"":T, isNaN(W)?"":W, (L>0?L:""), (x.CoilNo||x.BatchNum), (x.Mothercoil||"-"), Number(x.Quantity.toFixed(3)), x.AgeDays?parseInt(x.AgeDays):""]); });
    aoa.push(["","","","","","","","",Number(tq.toFixed(3)),""]);
    aoa.push([]);
  }
  aoa.push(["TOTAL MATERIAL HELD","","","","","","","",Number(gGrand.toFixed(3)),""]);
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"]=[{wch:4},{wch:10},{wch:16},{wch:7},{wch:8},{wch:9},{wch:22},{wch:22},{wch:11},{wch:7}];
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Stock Statement");
  return XLSX.write(wb,{type:"base64",bookType:"xlsx"});
}
// Fetch + filter the current Wheels India stock rows from the data repo (shared by PDF/XLSX).
async function wheelsRows(){
  const DATA_BASE = process.env.DATA_BASE || "https://raw.githubusercontent.com/BSC23609/bsc-stock-data/main";
  const csvRes = await fetch(DATA_BASE + "/wheels_stock.csv?t=" + Date.now());
  return parseCsvSrv(csvRes.ok ? await csvRes.text() : "").map(r=>({...r, Quantity:parseFloat(r.Quantity)||0, form:wheelsFormOf(r.ItemName)}))
    .filter(r=>!((r.form==="PLATE"&&r.Quantity<0.1)||(r.form==="STRIP"&&r.Quantity<0.011))); // plate <0.1 MT, strip <0.011 MT not listed
}
// Build the statement as a PDF (same sections/columns as the email). Pure pdfkit,
// no headless browser — safe on Vercel Hobby. Served at /api/wheels?fmt=pdf (and /wheels.pdf).
async function buildWheelsPdf(){
  let PDFDocument;
  try { PDFDocument = (await import("pdfkit")).default; } catch(e){ throw new Error("pdfkit module not available on server: "+e.message); }
  const rows = await wheelsRows();
  const SECDEF = {
    COIL:  ["Coil Stock","#1367a6"],
    PLATE: ["Plate Stock (WIP / FG)","#0e7a4b"],
    STRIP: ["Strip / Blank Stock","#b5651d"]
  };
  // Page 1 = WIP / FG (plate + strip) ; Page 2 = Coil Stock
  const PAGES = [
    { title:"WIP / FG Stock", forms:["PLATE","STRIP"] },
    { title:"Coil Stock",     forms:["COIL"] }
  ];

  const doc = new PDFDocument({ size:"A4", margin:28 });
  const chunks=[]; doc.on("data",c=>chunks.push(c));
  const done = new Promise(r=>doc.on("end",r));

  const L=28, R=567, W=R-L;                        // A4 = 595pt wide, 28pt margins
  const cols=[
    {k:"n",   t:"#",              w:22, a:"center"},
    {k:"mill",t:"Mill",           w:50, a:"center"},
    {k:"grade",t:"Grade",         w:92, a:"left"},
    {k:"thk", t:"Thk",            w:38, a:"center"},
    {k:"wid", t:"Width",          w:44, a:"center"},
    {k:"len", t:"Length",         w:48, a:"center"},
    {k:"coil",t:"Coil / Batch No",w:118,a:"left"},
    {k:"qty", t:"Qty (MT)",       w:56, a:"right"},
    {k:"age", t:"Age",            w:34, a:"center"}
  ];
  const sumW=cols.reduce((a,c)=>a+c.w,0), scale=W/sumW; let cx=L;
  cols.forEach(c=>{ c.w=c.w*scale; c.x=cx; cx+=c.w; });
  const qtyCol=cols.find(c=>c.k==="qty");

  const asof=new Date().toLocaleString("en-IN",{timeZone:"Asia/Kolkata",day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"});
  const rowH=13.5, headH=15;
  let y=28;

  function drawHeadRow(){
    doc.rect(L,y,W,headH).fill("#eef3f7");
    doc.fillColor("#1a2a3a").font("Helvetica-Bold").fontSize(7.5);
    cols.forEach(c=>doc.text(c.t, c.x+2, y+4, {width:c.w-4, align:c.a, lineBreak:false}));
    y+=headH;
  }
  function ensure(h){ if(y+h>812){ doc.addPage(); y=28; return true; } return false; }

  function drawBrandHeader(pageTitle){
    y=28;
    doc.rect(L,y,W,42).fill("#0e2a47");
    doc.fillColor("#fff").font("Helvetica-Bold").fontSize(14).text("Job Work Stock Statement — Wheels India Limited", L+10, y+9, {width:W-20});
    doc.font("Helvetica").fontSize(8).fillColor("#cfe0ef").text("Bharat Steel (Chennai) Pvt. Ltd.  ·  Material held at BSC on customer's account", L+10, y+28, {width:W-20});
    y+=50;
    doc.fillColor("#333").font("Helvetica").fontSize(8.5)
       .text("Customer: Wheels India Limited        As on: "+asof+"        Warehouse: Job Work (WH 45)", L, y, {width:W});
    y+=16;
    doc.rect(L,y,W,20).fill("#0e2a47");
    doc.fillColor("#fff").font("Helvetica-Bold").fontSize(11).text(pageTitle, L+8, y+5, {lineBreak:false});
    y+=26;
  }

  function drawSection(form){
    const [t,hex]=SECDEF[form];
    const rs=rows.filter(r=>r.form===form).sort((a,b)=>b.Quantity-a.Quantity);
    if(!rs.length) return { n:0, q:0 };
    const tq=rs.reduce((a,r)=>a+r.Quantity,0);
    ensure(16+headH+rowH*2);
    doc.rect(L,y,W,16).fill(hex);
    doc.fillColor("#fff").font("Helvetica-Bold").fontSize(9).text(t, L+8, y+4, {lineBreak:false});
    doc.font("Helvetica").fontSize(8).text(rs.length+" items · "+tq.toFixed(3)+" MT", L, y+4, {width:W-8, align:"right"});
    y+=16;
    drawHeadRow();
    rs.forEach((x,i)=>{
      if(ensure(rowH)) drawHeadRow();
      if(i%2){ doc.rect(L,y,W,rowH).fill("#f7fafc"); }
      doc.fillColor("#222").font("Helvetica").fontSize(7.5);
      const mill=(x.ItemName||"").split(" ")[0];
      const vals={ n:String(i+1), mill, grade:x.Grade||"", thk:dimS(x.Thick), wid:dimS(x.Width), len:lenS(x.Length), coil:(x.CoilNo||x.BatchNum||""), qty:x.Quantity.toFixed(3), age:(x.AgeDays||"") };
      cols.forEach(c=>doc.text(String(vals[c.k]), c.x+2, y+3.3, {width:c.w-4, align:c.a, lineBreak:false, ellipsis:true}));
      y+=rowH;
    });
    ensure(rowH+4);
    doc.moveTo(L,y).lineWidth(1).lineTo(R,y).strokeColor("#333").stroke();
    doc.fillColor("#000").font("Helvetica-Bold").fontSize(8);
    doc.text(t+" total", L, y+3, {width:qtyCol.x-L-4, align:"right", lineBreak:false});
    doc.text(tq.toFixed(3), qtyCol.x+2, y+3, {width:qtyCol.w-4, align:"right", lineBreak:false});
    y+=rowH+8;
    return { n:rs.length, q:tq };
  }

  function pageTotal(label, n, q){
    ensure(30);
    doc.rect(L,y,W,22).fill("#0e2a47");
    doc.fillColor("#fff").font("Helvetica-Bold").fontSize(10).text(label, L+10, y+6, {lineBreak:false});
    doc.text(n+" items · "+q.toFixed(3)+" MT", L, y+6, {width:W-10, align:"right"});
    y+=30;
  }

  function footerNote(){
    ensure(26);
    doc.fillColor("#6b5200").font("Helvetica").fontSize(7).text("Plates currently under process (awaiting shearing) are excluded from this statement.", L, y, {width:W});
    y+=11;
    doc.fillColor("#999").font("Helvetica-Oblique").fontSize(6.5).text("Computer generated statement — may contain mistakes.  Bharat Steel (Chennai) Pvt. Ltd. · No.147, Survey No.133, Thirunilai Village, Ponneri Taluk, Chennai 600103 · 044 6791 7800 · info@bharatsteels.in", L, y, {width:W});
  }

  PAGES.forEach((pg, idx)=>{
    if(idx>0) doc.addPage();
    drawBrandHeader(pg.title);
    let pn=0, pq=0;
    pg.forms.forEach(f=>{ const r=drawSection(f); pn+=r.n; pq+=r.q; });
    pageTotal(pg.title+" — Total", pn, pq);
    footerNote();
  });

  doc.end();
  await done;
  return Buffer.concat(chunks);
}
// ================= Wheels India PRODUCTION report (periods + monthly) =================
async function wheelsProdData(){
  const DATA_BASE = process.env.DATA_BASE || "https://raw.githubusercontent.com/BSC23609/bsc-stock-data/main";
  const [pTxt,mTxt] = await Promise.all([
    fetch(DATA_BASE+"/wheels_production.csv?t="+Date.now()).then(r=>r.ok?r.text():"").catch(()=>""),
    fetch(DATA_BASE+"/wheels_production_monthly.csv?t="+Date.now()).then(r=>r.ok?r.text():"").catch(()=>"")
  ]);
  return { periods: parseCsvSrv(pTxt), monthly: parseCsvSrv(mTxt) };
}
function istYMD(){ const o={}; new Intl.DateTimeFormat("en-CA",{timeZone:"Asia/Kolkata",year:"numeric",month:"2-digit",day:"2-digit"}).formatToParts(new Date()).forEach(x=>o[x.type]=x.value); return {y:+o.year,m:+o.month,d:+o.day}; }
async function prodModel(){
  const { periods, monthly } = await wheelsProdData();
  const {y,m,d}=istYMD();
  const base=new Date(Date.UTC(y,m-1,d,12));
  const yd=new Date(base); yd.setUTCDate(base.getUTCDate()-1);
  const lmT=new Date(base); lmT.setUTCMonth(base.getUTCMonth()-1);
  const lmY=new Date(yd);   lmY.setUTCMonth(yd.getUTCMonth()-1);
  const dd=x=>String(x.getUTCDate()).padStart(2,"0")+"-"+String(x.getUTCMonth()+1).padStart(2,"0")+"-"+x.getUTCFullYear();
  const MON=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const cols=[
    {k:"Daily Production Qty",           t:"Today",                  s:dd(base)},
    {k:"Yesterday Production Qty",       t:"Yesterday",              s:dd(yd)},
    {k:"Last Month Upto Yesterday Qty",  t:"Last Mo. upto yesterday",s:dd(lmY)},
    {k:"Monthly Production Qty",         t:"Monthly",                s:MON[base.getUTCMonth()]+"-"+base.getUTCFullYear()},
    {k:"Yearly Production Qty",          t:"Yearly",                 s:""+base.getUTCFullYear()}
  ];
  const pick=t=>periods.find(r=>String(r["Production Type"]||"").toUpperCase()===t);
  const q=(r,k)=>{ const n=parseFloat(r&&r[k]); return isNaN(n)?"":n.toFixed(3); };
  const rows=[["Plate",pick("PLATE TOTAL")],["Strip",pick("STRIP TOTAL")],["Total",periods.find(r=>/GRAND TOTAL/i.test(r["Production Type"]||""))]]
    .filter(x=>x[1]).map(([label,r])=>({label, cells: cols.map(c=>q(r,c.k))}));
  const mrows=monthly.map(r=>({ month:String(r["Month With Year"]||""), qty:(parseFloat(r["Grand Production Qty"])||0).toFixed(3) }));
  const mtot=monthly.reduce((a,r)=>a+(parseFloat(r["Grand Production Qty"])||0),0).toFixed(3);
  const asof=new Date().toLocaleString("en-IN",{timeZone:"Asia/Kolkata",day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"});
  return { cols, rows, mrows, mtot, asof };
}
function prodFontFiles(){
  const dirs=[];
  try{ dirs.push(path.join(process.cwd(),"api","fonts")); }catch(e){}
  try{ dirs.push(path.join(path.dirname(fileURLToPath(import.meta.url)),"fonts")); }catch(e){}
  for(const dir of dirs){ try{ if(fs.existsSync(path.join(dir,"DejaVuSans.ttf"))) return [path.join(dir,"DejaVuSans.ttf"), path.join(dir,"DejaVuSans-Bold.ttf")]; }catch(e){} }
  return [];
}
function prodSvg(m){
  const W=940, L=16, R=W-16, typeW=90, nC=m.cols.length, colW=(R-L-typeW)/nC;
  const els=[]; let y=0;
  const rect=(x,yy,w,h,fill)=>els.push(`<rect x="${x}" y="${yy}" width="${w}" height="${h}" fill="${fill}"/>`);
  const txt=(x,yy,s,o={})=>els.push(`<text x="${x}" y="${yy}" font-family="DejaVu Sans" font-size="${o.size||12}"${o.bold?' font-weight="bold"':''} fill="${o.fill||'#222'}" text-anchor="${o.anchor||'start'}">${esc2(s)}</text>`);
  rect(0,0,W,58,"#0e2a47");
  txt(L,26,"Wheels India — Production Report",{size:18,bold:true,fill:"#fff"});
  txt(L,46,"Bharat Steel (Chennai) Pvt. Ltd.  ·  Plate = MACHINE01 (MDS) · Strip = MACHINE15 (MDS) · Qty in MT",{size:10,fill:"#cfe0ef"});
  y=58; txt(L,y+18,"As on: "+m.asof,{size:10,fill:"#445"}); y+=28;
  rect(L,y,R-L,22,"#0e2a47"); txt(L+8,y+15,"Production — Plate / Strip",{size:12,bold:true,fill:"#fff"}); y+=22;
  const hH=34; rect(L,y,R-L,hH,"#eef3f7"); txt(L+6,y+21,"Type",{size:10,bold:true,fill:"#1a2a3a"});
  m.cols.forEach((c,i)=>{ const cx=L+typeW+i*colW+colW/2; txt(cx,y+14,c.t,{size:9,bold:true,fill:"#1a2a3a",anchor:"middle"}); txt(cx,y+28,c.s,{size:9,fill:"#7a8690",anchor:"middle"}); });
  y+=hH;
  const rH=24;
  m.rows.forEach((r,ri)=>{ const tot=r.label==="Total"; rect(L,y,R-L,rH, tot?"#0e2a47":(ri%2?"#f7fafc":"#ffffff"));
    txt(L+6,y+16,r.label,{size:11,bold:true,fill:tot?"#fff":"#222"});
    r.cells.forEach((v,i)=>{ const cx=L+typeW+i*colW+colW-6; txt(cx,y+16,v,{size:11,bold:tot,fill:tot?"#fff":"#222",anchor:"end"}); }); y+=rH; });
  y+=16;
  const m2w=340;
  rect(L,y,m2w,22,"#0e2a47"); txt(L+8,y+15,"Monthly Production (current year)",{size:12,bold:true,fill:"#fff"}); y+=22;
  rect(L,y,m2w,22,"#eef3f7"); txt(L+6,y+15,"Month",{size:10,bold:true,fill:"#1a2a3a"}); txt(L+m2w-6,y+15,"Production (MT)",{size:10,bold:true,fill:"#1a2a3a",anchor:"end"}); y+=22;
  m.mrows.forEach((r,ri)=>{ rect(L,y,m2w,20, ri%2?"#f7fafc":"#ffffff"); txt(L+6,y+14,r.month,{size:10}); txt(L+m2w-6,y+14,r.qty,{size:10,anchor:"end"}); y+=20; });
  rect(L,y,m2w,22,"#0e2a47"); txt(L+6,y+15,"Total",{size:10,bold:true,fill:"#fff"}); txt(L+m2w-6,y+15,m.mtot,{size:10,bold:true,fill:"#fff",anchor:"end"}); y+=22;
  y+=18; txt(L,y,"Computer-generated report — may contain mistakes.  Bharat Steel (Chennai) Pvt. Ltd. · info@bharatsteels.in",{size:8,fill:"#999"}); y+=10;
  const H=y+6;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><rect width="${W}" height="${H}" fill="#ffffff"/>${els.join("")}</svg>`;
}
async function buildWheelsProdPng(){
  let Resvg; try{ Resvg=(await import("@resvg/resvg-js")).Resvg; }catch(e){ throw new Error("resvg-js not available on server: "+e.message); }
  const ff=prodFontFiles();
  const svg=prodSvg(await prodModel());
  const r=new Resvg(svg, { font:{ fontFiles:ff, loadSystemFonts: ff.length===0, defaultFontFamily:"DejaVu Sans" }, background:"#ffffff" });
  return Buffer.from(r.render().asPng());
}
async function buildWheelsProdPdf(){
  let PDFDocument; try{ PDFDocument=(await import("pdfkit")).default; }catch(e){ throw new Error("pdfkit not available on server: "+e.message); }
  const m=await prodModel();
  const doc=new PDFDocument({size:"A4",margin:28}); const chunks=[]; doc.on("data",c=>chunks.push(c)); const done=new Promise(r=>doc.on("end",r));
  const L=28,R=567,W=R-L; let y=28;
  doc.rect(L,y,W,44).fill("#0e2a47");
  doc.fillColor("#fff").font("Helvetica-Bold").fontSize(15).text("Wheels India — Production Report",L+10,y+9);
  doc.font("Helvetica").fontSize(8).fillColor("#cfe0ef").text("Bharat Steel (Chennai) Pvt. Ltd.  ·  Plate = MACHINE01 (MDS) · Strip = MACHINE15 (MDS) · Qty in MT",L+10,y+28,{width:W-20});
  y+=52; doc.fillColor("#445").font("Helvetica").fontSize(9).text("As on: "+m.asof,L,y); y+=16;
  const typeW=84, nC=m.cols.length, colW=(W-typeW)/nC;
  doc.rect(L,y,W,18).fill("#0e2a47"); doc.fillColor("#fff").font("Helvetica-Bold").fontSize(10).text("Production — Plate / Strip",L+8,y+5,{lineBreak:false}); y+=18;
  const hH=30; doc.rect(L,y,W,hH).fill("#eef3f7"); doc.fillColor("#1a2a3a").font("Helvetica-Bold").fontSize(8).text("Type",L+5,y+11,{lineBreak:false});
  m.cols.forEach((c,i)=>{ const cx=L+typeW+i*colW; doc.font("Helvetica-Bold").fontSize(8).fillColor("#1a2a3a").text(c.t,cx+2,y+5,{width:colW-4,align:"center",lineBreak:false}); doc.font("Helvetica").fontSize(7).fillColor("#7a8690").text(c.s,cx+2,y+18,{width:colW-4,align:"center",lineBreak:false}); });
  y+=hH;
  const rH=20;
  m.rows.forEach((r,ri)=>{ const tot=r.label==="Total"; if(tot){doc.rect(L,y,W,rH).fill("#0e2a47");} else if(ri%2){doc.rect(L,y,W,rH).fill("#f7fafc");}
    doc.font(tot?"Helvetica-Bold":"Helvetica").fontSize(10).fillColor(tot?"#fff":"#222").text(r.label,L+5,y+6,{lineBreak:false});
    r.cells.forEach((v,i)=>{ const cx=L+typeW+i*colW; doc.font(tot?"Helvetica-Bold":"Helvetica").fontSize(9).fillColor(tot?"#fff":"#222").text(v,cx+2,y+6,{width:colW-6,align:"right",lineBreak:false}); }); y+=rH; });
  y+=16;
  const m2w=300;
  doc.rect(L,y,m2w,18).fill("#0e2a47"); doc.fillColor("#fff").font("Helvetica-Bold").fontSize(10).text("Monthly Production (current year)",L+8,y+5,{lineBreak:false}); y+=18;
  doc.rect(L,y,m2w,18).fill("#eef3f7"); doc.fillColor("#1a2a3a").font("Helvetica-Bold").fontSize(8).text("Month",L+6,y+5,{lineBreak:false}); doc.text("Production (MT)",L,y+5,{width:m2w-6,align:"right",lineBreak:false}); y+=18;
  m.mrows.forEach((r,ri)=>{ if(ri%2){doc.rect(L,y,m2w,16).fill("#f7fafc");} doc.font("Helvetica").fontSize(9).fillColor("#222").text(r.month,L+6,y+4,{lineBreak:false}); doc.text(r.qty,L,y+4,{width:m2w-6,align:"right",lineBreak:false}); y+=16; });
  doc.rect(L,y,m2w,18).fill("#0e2a47"); doc.fillColor("#fff").font("Helvetica-Bold").fontSize(9).text("Total",L+6,y+5,{lineBreak:false}); doc.text(m.mtot,L,y+5,{width:m2w-6,align:"right",lineBreak:false}); y+=24;
  doc.fillColor("#999").font("Helvetica-Oblique").fontSize(7).text("Computer-generated report — may contain mistakes. Bharat Steel (Chennai) Pvt. Ltd. · info@bharatsteels.in",L,y,{width:W});
  doc.end(); await done; return Buffer.concat(chunks);
}
async function sendWheelsMail(to, cc, html, subject, attach){
  const tenant=process.env.TENANT_ID, cid=process.env.CLIENT_ID, sec=process.env.CLIENT_SECRET;
  const sender=process.env.WHEELS_SENDER || "pdqc@bharatsteels.in";
  const tok = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
    method:"POST", headers:{"Content-Type":"application/x-www-form-urlencoded"},
    body:new URLSearchParams({client_id:cid,client_secret:sec,scope:"https://graph.microsoft.com/.default",grant_type:"client_credentials"})
  }).then(r=>r.json());
  if(!tok.access_token) throw new Error("token: "+(tok.error_description||tok.error||"failed"));
  const message={ subject, body:{contentType:"HTML",content:html},
    toRecipients:(to||[]).map(a=>({emailAddress:{address:a}})),
    ccRecipients:(cc||[]).map(a=>({emailAddress:{address:a}})) };
  if(attach && attach.b64){
    const a={ "@odata.type":"#microsoft.graph.fileAttachment", name:attach.name,
      contentType:(attach.contentType||"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), contentBytes:attach.b64 };
    if(attach.cid){ a.contentId=attach.cid; a.isInline=true; }
    message.attachments=[a];
  }
  const msg={ message, saveToSentItems:true };
  const r = await fetch(`https://graph.microsoft.com/v1.0/users/${sender}/sendMail`, {
    method:"POST", headers:{Authorization:"Bearer "+tok.access_token,"Content-Type":"application/json"}, body:JSON.stringify(msg) });
  if(r.status!==202) throw new Error("sendMail HTTP "+r.status+" "+(await r.text()).slice(0,200));
}

// _router.js
async function handler5(req, res) {
  const url = req.url || "";
  let route = decodeURIComponent((/[?&]route=([^&]*)/.exec(url) || [])[1] || req.query && req.query.route || "");
  if (!route || route === "index") route = url.split("?")[0].replace(/\/+$/, "").replace(/^\/api\/?/, "");
  route = route.replace(/^\/+|\/+$/g, "");
  if (route === "dispatch") return handler(req, res);
  if (route === "sync") return handler2(req, res);
  if (route === "wheels") return handlerWheels(req, res);
  res.setHeader("Cache-Control", "no-store");
  res.status(404).json({ error: "unknown route: " + route });
}
export {
  handler5 as default
};
