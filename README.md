# GroutLine — Inventory Management (Preview Build)

A responsive inventory management app for a tile & adhesive business.
Built with React + Vite. Currently running on in-memory mock data —
not yet connected to a real backend.

## What's included

- **Login** — UI only, accepts any email/password (demo mode)
- **Dashboard** — quick stats, low-stock alerts, Excel export/import, recently added items
- **Products** — full list with search, category filter, sort, low-stock filter,
  add/edit/delete, inline quantity editing, and Excel export (respects active filters)
- **Excel import/export** — real `.xlsx` file download/upload via SheetJS,
  including a downloadable sample import template
- **Prices in INR (₹)**, Indian number formatting

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

## Deploy (free)

Push this folder to GitHub, then import it on vercel.com — Vite is auto-detected,
no config needed. You'll get a live `https://your-project.vercel.app` URL.

## What's next (not yet built)

- Real authentication (Supabase Auth)
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
└── src/
    ├── main.jsx
    ├── index.css
    └── App.jsx        # login + dashboard + products, all tabs
```
