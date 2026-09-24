# MIS Portal — Vercel deployment (Stage 2)

One Vercel app, three modules, at **mis.bharatsteels.in**:
- **Stock** — 4 screens (Daily Stock Report, Coil T×W×Grade, Sheet list, Coil list). Reads CSVs pushed by PowerShell. No backend.
- **Dispatch** — weighbridge vehicle-movement dashboard. Reads `dispatch/data.csv` pushed by PowerShell. No backend.
- **NMDC** — dispatch dashboard. Reads Neon via `/api/dispatch`; Neon is filled from the OneDrive `tbl_Dispatch` workbook by `/api/sync` (Graph). PIN-gated.

Transporter links, WATI and the /t/:token page have been removed. Sync still runs (it is independent of WATI).

## 1. Create repo + Vercel project
1. New GitHub repo `mis-portal`, push all these files.
2. vercel.com → New Project → import `mis-portal`. Framework preset: **Other**. Deploy.
3. It gets a `*.vercel.app` URL — test everything there first (domain comes last).

## 2. Environment variables (Vercel → Settings → Environment Variables)
Needed for the NMDC module only. Stock & Dispatch need none.

| Key | Value |
|---|---|
| `DATABASE_URL` | the Neon connection string (rotate the password first) |
| `DASH_PIN` | dashboard PIN for the NMDC page |
| `CRON_SECRET` | random string; cron-job.org sends it as `Authorization: Bearer …` |
| `CLIENT_ID` | Azure app client id |
| `CLIENT_SECRET` | Azure app client secret (new one) |
| `TENANT_ID` | Azure tenant id |
| `TABLE_NAME` | `tbl_Dispatch` |
| `SHARE_LINK` | the OneDrive share link to the workbook |
| `ALLOWED_ORIGINS` | `https://mis.bharatsteels.in` |

Do NOT set `WATI_API_URL` / `WATI_API_KEY` — leaving them unset keeps WATI disabled by design.

## 3. Schedule the NMDC sync (Excel → Neon)
cron-job.org → URL `https://mis.bharatsteels.in/api/sync`, every 5 min, GET,
header `Authorization: Bearer <CRON_SECRET>`. (Until DNS cuts over, use the `*.vercel.app` URL.)
A healthy run returns JSON with `imported`/`updated`/`adopted` counts.

## 4. Point PowerShell at the new repo
`refresh_stock.ps1` currently pushes to `bsc-stock-data`. Change `$RepoDir` to the new
`mis-portal` checkout and adjust the file paths:
- `coils.csv`, `plates.csv`, `mis_stock.csv`, `mis_sales.csv`, `mis_purchase.csv` → `stock/`
- the dispatch/weighbridge job that writes `data.csv` → `dispatch/data.csv`
Vercel auto-deploys on every push, so the dashboards refresh the same as before.

## 5. Cut over the domain (last, reversible)
1. Confirm the `*.vercel.app` site works end-to-end.
2. GitHub `bsc-stock-data` → Settings → Pages → remove custom domain `mis.bharatsteels.in`.
3. Vercel project → Settings → Domains → add `mis.bharatsteels.in`. Follow its DNS instruction
   (CNAME to `cname.vercel-dns.com`, replacing the old GitHub Pages record).
4. `bsc-stock-data` stays intact as rollback — re-add the domain there to revert.

## Security note
Rotate the Neon password, DASH_PIN and CRON_SECRET now — they were shared in chat during setup.
