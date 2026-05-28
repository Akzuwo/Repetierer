# HWM Issue Handler

Small Cloudflare Worker that receives desktop app error reports via `POST /report` and stores each report as one JSON file in Cloudflare R2.

## Setup

Create the R2 bucket:

```sh
npx wrangler r2 bucket create hwm-issue-reports
```

Install dependencies:

```sh
npm install
```

Optionally set an API key. When this secret exists, requests must send `Authorization: Bearer <key>`.

```sh
npx wrangler secret put REPORT_API_KEY
```

Run locally:

```sh
npm run dev
```

Deploy:

```sh
npm run deploy
```

## Example Request

```sh
curl -X POST "http://localhost:8787/report" \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer your-api-key" \
  --data '{
    "appVersion": "6.2.0",
    "windowsUser": "timow",
    "comment": "Beim Excel neu laden kommt ein Fehler.",
    "log": "...",
    "clientTimestamp": "2026-05-28T12:40:00.000Z"
  }'
```

Successful reports are stored as `reports/<id>.json` in the `hwm-issue-reports` bucket.
