# Shared Vercel + Supabase Setup

This app now supports a shared Supabase-backed state store with live updates.

## 1. Create the Supabase project

1. Create a new Supabase project.
2. Open the SQL editor.
3. Run [`supabase/schema.sql`](./supabase/schema.sql).
4. In `Database` -> `Replication` or `Realtime`, make sure `public.app_state` is enabled for realtime.

## 2. Add Vercel environment variables

Add these in Vercel project settings:

- `VITE_SUPABASE_URL`
- `VITE_SUPABASE_ANON_KEY`

Use the values from your Supabase project API settings.

For local development, copy `.env.example` to `.env.local` and fill in the same values.

## 3. Deploy

Push the repo to GitHub and let Vercel redeploy, or run a fresh Vercel deploy after adding the env vars.

## 4. How it works

- Shared data in Supabase:
  - jobs
  - assignments
  - requests
  - pull paper requests
  - shipment groups
  - user list / passwords / roles
  - shared week start
- Local-only per device:
  - which user is currently logged in

## 5. Live refresh

When Supabase is configured:

- schedule changes save to the shared row automatically
- new requests save automatically
- other open devices receive live updates through Supabase Realtime
- the header shows sync status like `Connecting`, `Saving`, `Live sync`, or `Sync error`

## 6. Security note

The current SQL policies allow browser clients with the anon key to read and write the shared scheduler row so the app works immediately on Vercel.

That is fine for an internal/private deployment, but if you want stronger security later, the next step would be moving the app from the current custom login system to real Supabase Auth plus tighter row-level policies.
