# OneDrive Archiver

OneDrive Archiver signs in to Microsoft OneDrive, lets you browse your drive, and downloads selected files or folders to a local folder.

## Run Locally

**Prerequisites:** Node.js and Microsoft Edge or Chrome. Local folder downloads use the browser File System Access API.

## Microsoft Setup

1. Open Azure App Registrations and create an app.
2. Set supported account types to include personal Microsoft accounts.
3. Add a Web redirect URI: `http://localhost:3000/api/callback`
4. Create a client secret and copy the secret **Value**.
5. Add delegated Microsoft Graph permissions for `User.Read` and `Files.Read.All`.

## Run Locally

1. Install dependencies:
   `npm install`
2. Copy [.env.example](.env.example) to `.env.local` and configure the Microsoft OAuth values.
3. Run the app:
   `npm run dev`
4. Open `http://localhost:3000`, sign in, choose a local folder, select OneDrive files or folders, and start the download.
