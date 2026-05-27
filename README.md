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

## Resend inbound requests

This project now includes a Vercel Function at [`api/resend-inbound.js`](./api/resend-inbound.js) that can turn inbound Resend emails into open requests.

Setup:

1. Add these Vercel env vars:
   - `RESEND_API_KEY`
   - `RESEND_WEBHOOK_SECRET`
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_ROLE_KEY`
2. In Resend, create an `email.received` webhook that points to:
   - `https://label-traxx-scheduler-v3.vercel.app/api/resend-inbound`
3. Forward your Outlook/shared mailbox to your Resend receiving address.

Behavior:

- The webhook pulls the full inbound email body from Resend
- Attachments are copied into the shared Supabase storage bucket
- A new `Email Request` is created in `Open Requests`
- If the mailbox alias matches a username, the request is auto-assigned to that user

## Current shared features

- Shared scheduler data across devices
- Shared requests, shipment groups, and pull paper requests
- Shared user list and roles
- Live updates through Supabase Realtime
- Local per-device login session persistence
