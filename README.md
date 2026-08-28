# ColdLine WMS — Multi-Warehouse Inventory & Stock Management System

A Vite + React app, ready to deploy on Netlify.

## What's inside
- `src/App.jsx` — full warehouse system: dashboard, items, stock in/out,
  ledger, suppliers, payments, branches, reports, daily closing, settings,
  activity log — now with **multiple independent warehouses** and a
  **login screen**.
- Data is saved in the browser's `localStorage`.

## Demo accounts (shown on the login screen too)
- `owner / owner123` — Platform Owner, sees & manages all warehouses
- `ayesha / ayesha123` — Cold Storage Warehouse owner
- `bilal / bilal123` — North Bakery Hub owner
- `sana / sana123` — Accountant, Cold Storage Warehouse
- `guest / guest123` — Viewer, both warehouses

## Deploy to Netlify

### Option A — Drag & drop (fastest, no CLI needed)
1. On your own machine (with internet access), unzip this project.
2. Run:
   ```
   npm install
   npm run build
   ```
3. Go to https://app.netlify.com/drop and drag the generated `dist` folder
   onto the page. Netlify gives you a live URL immediately.

### Option B — Connect a Git repo (recommended for updates)
1. Push this folder to a GitHub/GitLab/Bitbucket repo.
2. In Netlify: **Add new site → Import an existing project** → pick the repo.
3. Build settings (already in `netlify.toml`): build command `npm run build`,
   publish directory `dist`.
4. Deploy — Netlify installs dependencies and builds automatically.

### Option C — Netlify CLI
```
npm install -g netlify-cli
npm install
npm run build
netlify deploy --prod --dir=dist
```

## Local development
```
npm install
npm run dev
```

## Honest limits
- Login is a **client-side demo only** — usernames/passwords live in the
  browser's storage, not verified by a real server. Don't use real
  passwords here; a production deployment needs proper backend auth.
- Stock valuation uses current purchase rate, not FIFO/weighted-average.
