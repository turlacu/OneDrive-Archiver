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
APP_DATA_DIR=/tmp/onedrive-archiver-config
ALLOWED_USERS=your-microsoft-email@example.com
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
APP_DATA_DIR=/config
ALLOWED_USERS=you@example.com
```

The app will only write archived files inside `SERVER_DOWNLOAD_ROOT`. When `SERVER_DOWNLOAD_ROOT` is set, server archive mode is enabled and `ALLOWED_USERS` is required. If you leave `SERVER_DOWNLOAD_ROOT` unset, the app runs in browser-only mode and downloads stay on the user's local PC.

## UNRAID Deployment

Unraid installs are server archive deployments. Downloads land on the Unraid server, jobs are persisted under `/config`, and only Microsoft accounts listed in `ALLOWED_USERS` can use the app.

### Requirements

- Unraid Docker service enabled.
- A Microsoft Azure app registration with a Web redirect URI.
- A public HTTPS URL for the app, usually through a reverse proxy. Microsoft only allows `http://` redirect URIs for localhost development.
- A long random `SESSION_SECRET`.
- An `ALLOWED_USERS` value with the Microsoft email addresses allowed to use this server.

Recommended image:

```txt
ghcr.io/turlacu/onedrive-archiver:latest
```

Container port:

```txt
3000
```

Recommended archive path inside the container:

```txt
/downloads/onedrive-archive
```

Recommended app data path inside the container:

```txt
/config
```

### Install with the bundled Unraid template

Use this option for the easiest install. The template creates all required container settings for the user: port, archive path, OAuth variables, permissions, and timezone.

If you only type `ghcr.io/turlacu/onedrive-archiver:latest` into Unraid's `Repository` field, Unraid creates a plain Docker container form and will not know which settings to show. Import the XML template first if you want the settings to appear automatically in the container editor.

The repository includes a Docker template at:

```txt
unraid/onedrive-archiver.xml
```

To install it from the Unraid terminal:

```bash
mkdir -p /boot/config/plugins/dockerMan/templates-user
curl -L -o /boot/config/plugins/dockerMan/templates-user/my-onedrive-archiver.xml \
  https://raw.githubusercontent.com/turlacu/OneDrive-Archiver/main/unraid/onedrive-archiver.xml
```

Then create the container:

1. In the Unraid WebGUI, open the Docker tab.
2. Click `Add Container`.
3. Select `OneDrive-Archiver` from the template dropdown.
4. Fill in `APP_URL`, `SESSION_SECRET`, `MICROSOFT_CLIENT_ID`, `MICROSOFT_CLIENT_SECRET`, and `ALLOWED_USERS`.
5. Adjust the host archive path if needed.
6. Click `Apply`.

The template provides these visible settings:

```txt
Web Interface Port
Archive Download Path
App Data Path
APP_URL
SESSION_SECRET
MICROSOFT_CLIENT_ID
MICROSOFT_CLIENT_SECRET
SERVER_DOWNLOAD_ROOT
APP_DATA_DIR
ALLOWED_USERS
```

Advanced settings are also included:

```txt
NODE_ENV
PORT
PUID
PGID
UMASK
TZ
```

### Install with Add Container

If you do not use the XML template, create a container manually from the Docker tab and add the settings below yourself.

Main fields:

```txt
Name:       OneDrive-Archiver
Repository: ghcr.io/turlacu/onedrive-archiver:latest
Network:    bridge
WebUI:      http://[IP]:[PORT:3000]/
```

Port mapping:

```txt
Container port: 3000
Host port:      3000
Protocol:       TCP
```

You can change the host port if `3000` is already in use. Leave the container port as `3000` unless you also set the `PORT` environment variable.

Volume mapping:

```txt
Host path:      /mnt/user/Backups/OneDrive
Container path: /downloads/onedrive-archive
Access mode:    Read/Write
```

App data mapping:

```txt
Host path:      /mnt/user/appdata/onedrive-archiver
Container path: /config
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
APP_DATA_DIR=/config
ALLOWED_USERS=you@example.com,other@example.com
PUID=99
PGID=100
UMASK=022
TZ=Etc/UTC
```

Azure redirect URI:

```txt
https://your-unraid-app-domain.example.com/api/callback
```

`APP_URL` must exactly match the external URL used in your browser, without a trailing slash. The Azure redirect URI must be that URL plus `/api/callback`.

### File permissions

The container supports Unraid-style file ownership variables:

```env
PUID=99
PGID=100
UMASK=022
```

The defaults match the standard Unraid `nobody:users` ownership. Change them only if your share requires different permissions.

### Reverse proxy notes

For normal LAN installs, expose the app through an HTTPS reverse proxy and set `APP_URL` to the external HTTPS address. The app trusts one proxy hop and sets secure cookies when `APP_URL` starts with `https://`.

Direct `http://tower:3000` access is useful for checking whether the container starts, but Microsoft OAuth callbacks for non-localhost HTTP URLs are not valid production redirect URIs.

### Updating

When a new image is published, use the Docker tab to check for updates or force update the container. Your archived files stay on the mapped Unraid share.

### Build locally

If the published image is not available yet, build it from the repository root:

```bash
docker build -t onedrive-archiver:local .
```

Then use `onedrive-archiver:local` as the repository/image value in the Unraid container form.

Usage:

1. Open the app in your browser.
2. Sign in to Microsoft.
3. Select `Server` in Archive Target.
4. Select OneDrive files or folders.
5. Click `Start Archive`, `Dry Run`, `Incremental`, or `Repair`.
6. You may close the browser after the server job starts. The server continues the job while the container remains running.

Server jobs, logs, and incremental delta tokens are persisted in `APP_DATA_DIR`. If the container restarts while a job is active, the job is marked interrupted and resumes when the owning user signs in again.

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

Each allowed user writes under:

```txt
SERVER_DOWNLOAD_ROOT/users/<user-email>/
```

The folder name is sanitized for filesystem safety.

Available actions:

- `Start Archive`
- `Dry Run`
- `Repair`
- `Incremental`
- cancel running server job
- reconnect to active server job progress after page reload
- reconnect to interrupted jobs after signing in again

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
