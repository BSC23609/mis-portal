// NMDC backend for MIS portal — sync (Excel->Neon) + dispatch (read). Transporter/WATI/arrive disabled (routes removed; WATI env vars unset).
// original single-file build (dispatch, sync, tform, arrive) — build 2026-09-21g (consignment lines, Ref key, SO master)

// lib/db.js
import { neon } from "@neondatabase/serverless";
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
async function handler(req, res) {
  if (cors(req, res)) return;
  if (!process.env.DASH_PIN) return json(res, 503, { error: "DASH_PIN not set" });
  if (req.headers["x-dash-pin"] !== process.env.DASH_PIN) return json(res, 401, { error: "pin" });
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
async function send(mobile, template, parameters) {
  const base = process.env.WATI_API_URL, key = process.env.WATI_API_KEY;
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

// _router.js
async function handler5(req, res) {
  const url = req.url || "";
  let route = decodeURIComponent((/[?&]route=([^&]*)/.exec(url) || [])[1] || req.query && req.query.route || "");
  if (!route || route === "index") route = url.split("?")[0].replace(/\/+$/, "").replace(/^\/api\/?/, "");
  route = route.replace(/^\/+|\/+$/g, "");
  if (route === "dispatch") return handler(req, res);
  if (route === "sync") return handler2(req, res);
  res.setHeader("Cache-Control", "no-store");
  res.status(404).json({ error: "unknown route: " + route });
}
export {
  handler5 as default
};
