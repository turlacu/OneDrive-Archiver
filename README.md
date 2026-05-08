# OneDrive Archiver

OneDrive Archiver signs in to Microsoft OneDrive, lets you browse your drive, and archives selected files or folders. It supports two target modes:

- **Browser target**: downloads to a folder selected in the browser. This uses the browser File System Access API and is best for local desktop use.
- **Server target**: downloads on the server into a configured path. This is intended for UNRAID, Coolify, or similar self-hosted deployments where the browser is only a remote control.

## Features

- OneDrive OAuth sign-in with server-side session storage.
- OneDrive folder browsing, search, select all, and deselect.
- Browser-side archive downloads with chunking, retry, resume state, verification, conflict handling, dry run, repair, and incremental mode.
- Server-side archive jobs for start, dry run, repair, cancel, and reconnectable progress polling.
- Server-side downloads are constrained to `SERVER_DOWNLOAD_ROOT`.

## Microsoft Azure Setup

1. Open [Azure App Registrations](https://portal.azure.com/) and create an app.
2. Set supported account types to include personal Microsoft accounts, or personal plus organizational accounts.
3. Add a **Web** redirect URI.
4. Create a client secret and copy the secret **Value**, not the Secret ID.
5. Add delegated Microsoft Graph permissions:
   - `User.Read`
   - `Files.Read.All`
   - `offline_access`

Redirect URI examples:

```txt
Local development:
http://localhost:3000/api/callback

Coolify or public server:
https://onedrivearchiver.example.com/api/callback
```

Keep both redirect URIs in Azure if you use both local development and a deployed server.

## Local Development

Prerequisites:

- Node.js
- Microsoft Edge or Chrome for browser-target downloads

Setup:

```bash
npm install
cp .env.example .env.local
```

Edit `.env.local`:

```env
APP_URL=http://localhost:3000
SESSION_SECRET=replace-with-a-long-random-string
MICROSOFT_CLIENT_ID=your-app-client-id
MICROSOFT_CLIENT_SECRET=your-client-secret-value
```

Run:

```bash
npm run dev
```

Open:

```txt
http://localhost:3000
```

For browser-target downloads, choose `Browser` in the Archive Target panel, select a local folder, select OneDrive items, and start the archive.

Optional local server-target testing:

```env
SERVER_DOWNLOAD_ROOT=/tmp/onedrive-archive
```

Then choose `Server` in the Archive Target panel.

## Coolify Deployment

Set these environment variables in Coolify:

```env
NODE_ENV=production
APP_URL=https://onedrivearchiver.example.com
SESSION_SECRET=replace-with-a-long-random-string
MICROSOFT_CLIENT_ID=your-app-client-id
MICROSOFT_CLIENT_SECRET=your-client-secret-value
```

In Azure, add the matching redirect URI:

```txt
https://onedrivearchiver.example.com/api/callback
```

Coolify should expose the app on port `3000` inside the container. The public URL must match `APP_URL`.

To enable server-side downloads in Coolify, mount a persistent volume into the container and set:

```env
SERVER_DOWNLOAD_ROOT=/downloads/onedrive-archive
```

The app will only write inside `SERVER_DOWNLOAD_ROOT`.

## UNRAID Deployment

Use server-target mode when you want downloads to land on the UNRAID server rather than the browser PC.

Recommended container path:

```txt
/downloads/onedrive-archive
```

Example UNRAID volume mapping:

```txt
Host path:      /mnt/user/Backups/OneDrive
Container path: /downloads/onedrive-archive
Access mode:    Read/Write
```

Environment variables:

```env
NODE_ENV=production
APP_URL=https://your-unraid-app-domain.example.com
SESSION_SECRET=replace-with-a-long-random-string
MICROSOFT_CLIENT_ID=your-app-client-id
MICROSOFT_CLIENT_SECRET=your-client-secret-value
SERVER_DOWNLOAD_ROOT=/downloads/onedrive-archive
```

Azure redirect URI:

```txt
https://your-unraid-app-domain.example.com/api/callback
```

Usage:

1. Open the app in your browser.
2. Sign in to Microsoft.
3. Select `Server` in Archive Target.
4. Select OneDrive files or folders.
5. Click `Start Archive`, `Dry Run`, or `Repair`.
6. You may close the browser after the server job starts. The server continues the job while the container remains running.

Server jobs are in memory in this version. They continue after the browser closes, but they do not survive a container/server restart.

## Target Modes

### Browser Target

Downloads are written by the browser to a folder you choose. This mode requires a Chromium-based browser because it uses the File System Access API.

Available actions:

- `Start Archive`
- `Dry Run`
- `Incremental`
- `Repair`
- retry failed files
- clear stale partial files

### Server Target

Downloads are written by the Node server into `SERVER_DOWNLOAD_ROOT`.

Available actions:

- `Start Archive`
- `Dry Run`
- `Repair`
- cancel running server job
- reconnect to active server job progress after page reload

Server-side `Incremental` is not enabled in this version.

## Validation

Run checks:

```bash
npm run lint
npm run build
npm test
```

## Notes

- `APP_URL` must not have a trailing slash.
- `SESSION_SECRET` must be a strong unique value in production.
- `MICROSOFT_CLIENT_SECRET` must be the secret **Value** from Azure.
- Public deployments should use HTTPS so secure session cookies work correctly.
- Existing `.syncpoint-manifests` files created by older builds can be deleted from archive folders. Current builds store repair metadata in browser storage instead.
