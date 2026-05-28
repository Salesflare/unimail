# Unimail viewer

A small read-only email viewer for poking at the `@salesflare/unimail` library
during development. It serves a tiny Hapi backend that calls the library and a
vanilla-JS frontend laid out like a typical 3-pane email client (folders on the
left, message list in the middle, message detail on the right).

The viewer supports Gmail, Office 365 and Unipile. Nylas is not wired up.

## Setup

From the repo root:

```bash
npm install
cp example/.env.example example/.env
# edit example/.env and fill in the connector credentials you have
npm run example
```

Then open <http://127.0.0.1:3000>.

You only need to fill in the credentials for the connectors you want to try.
Connectors with missing env vars are silently skipped at startup, and the
viewer's connector dropdown only lists the configured ones.

## How it works

- The browser holds your per-account auth (`access_token`, `refresh_token`,
  `expiration_date`) in `localStorage`. Use the form at the top to paste them
  and hit `Save`.
- Every API request to the local server sends two custom headers:
  - `X-Unimail-Connector`: `gmail` | `office365` | `unipile`
  - `X-Unimail-Auth`: base64-encoded JSON of the auth object above plus a
    per-session `id` (random UUID).
- When the library refreshes the access token, the server returns the updated
  auth on a `X-Unimail-Auth-Updated` response header and the UI overwrites
  `localStorage` with it.

The library is otherwise called exactly as a normal consumer would call it
(`unimail.folders.list`, `unimail.messages.list`, `unimail.messages.get`).

## API surface

| Verb | Path                              | Library call                |
|------|-----------------------------------|-----------------------------|
| GET  | `/connectors`                     | (lists the configured ones) |
| GET  | `/api/folders`                    | `unimail.folders.list`      |
| GET  | `/api/messages?folder=&limit=&pageToken=` | `unimail.messages.list` |
| GET  | `/api/messages/{id}`              | `unimail.messages.get`      |

There is no file/attachment download endpoint by design. Attachments are shown
as metadata only in the detail pane.

## Provider quirks

- **Gmail**: clicking a folder filters via `in:<folder name>`. Any label works.
- **Unipile**: clicking a folder filters by `role`/name. Standard roles like
  `inbox` and `sent` work; custom folders work by name.
- **Office 365**: only the `sent` and `inbox` (no filter) folders are supported
  by `listMessages`. For any other Office 365 folder, the viewer shows a banner
  and falls back to listing without a folder filter.

## Files

- `server.js` - Hapi bootstrap, header parsing, static file serving.
- `unimail.js` - builds a single `Unimail` instance and registers the
  connectors based on env vars; exports the instance and connector refs.
- `routes/folders.js`, `routes/messages.js` - thin wrappers around the library.
- `public/` - the static frontend (`index.html`, `app.js`, `styles.css`).
