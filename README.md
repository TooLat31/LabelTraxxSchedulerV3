# Label Traxx Scheduler

Shared production scheduler built with `React + Vite`.

## Local development

```bash
npm install
npm run dev
```

## Shared Supabase mode

This app can now run in two modes:

- `Local only`: browser storage only
- `Shared`: Supabase-backed state with live updates across devices

To turn on shared mode:

1. Create a Supabase project
2. Run [`supabase/schema.sql`](./supabase/schema.sql)
3. Add the env vars from [`.env.example`](./.env.example)
4. Deploy to Vercel

Full setup notes are in [`SUPABASE_SETUP.md`](./SUPABASE_SETUP.md).

## Current shared features

- Shared scheduler data across devices
- Shared requests, shipment groups, and pull paper requests
- Shared user list and roles
- Live updates through Supabase Realtime
- Local per-device login session persistence
