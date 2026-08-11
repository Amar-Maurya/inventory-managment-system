# GroutLine — Inventory Management (Preview Build)

A responsive inventory management app for a tile & adhesive business.
Built with React + Vite. Currently running on in-memory mock data —
not yet connected to a real backend.

## What's included

- **Login** — UI only, accepts any email/password (demo mode)
- **Dashboard** — quick stats, low-stock alerts, Excel export/import, recently added items
- **Products** — full list with search, category filter, sort, low-stock filter,
  add/delete, and Excel export (respects active filters)
- **Excel import/export** — real `.xlsx` file download/upload via SheetJS,
  including a downloadable sample import template

## Run locally

```bash
npm install
npm run dev
```

Opens at `http://localhost:5173`

## Build for production

```bash
npm run build
```

Outputs static files to `dist/` — deployable to any static host.

## Deploy (free)

**Vercel** (recommended)
1. Push this folder to a GitHub repo
2. Go to vercel.com → New Project → Import your repo
3. Vercel auto-detects Vite — just click Deploy
4. You'll get a live `https://your-project.vercel.app` URL

**Netlify** — same flow: connect repo → deploy → live `.netlify.app` URL

## What's next (not yet built)

- Real authentication (Supabase Auth — swap logic lives in one place, not
  scattered through the UI)
- Live database instead of mock in-memory state (Supabase/Postgres, free tier)
- Product image upload (Supabase Storage)
- Role-based access (Admin / Staff / Viewer)

## Project structure

```
groutline/
├── index.html
├── package.json
├── vite.config.js
├── tailwind.config.js
├── postcss.config.js

    ├── main.jsx
    ├── index.css
    └── App.jsx        # login + dashboard + products, all tabs
```

`App.jsx` currently holds everything in one file for fast iteration during
the preview phase. Once backend work starts, this splits into
`pages/`, `components/`, `services/`, and `context/` as planned.
