import express from 'express';
import { createServer as createViteServer } from 'vite';
import path from 'path';
import { fileURLToPath } from 'url';
import { pathToFileURL } from 'url';
import * as msal from '@azure/msal-node';
import dotenv from 'dotenv';
import cookieParser from 'cookie-parser';
import session from 'express-session';
import { randomBytes, timingSafeEqual } from 'crypto';
import { ServerDownloadManager } from './src/server/serverDownloadEngine.ts';

dotenv.config({ path: '.env.local' });
dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

let APP_URL = process.env.APP_URL || 'http://localhost:3000';
// Remove trailing slash if present
if (APP_URL.endsWith('/')) {
    APP_URL = APP_URL.slice(0, -1);
}

const REDIRECT_URI = `${APP_URL}/api/callback`;

const CLIENT_ID = process.env.MICROSOFT_CLIENT_ID;
const CLIENT_SECRET = process.env.MICROSOFT_CLIENT_SECRET;
const FALLBACK_SESSION_SECRET = 'syncpoint-archiver-dev-secret';
const SESSION_SECRET = process.env.SESSION_SECRET || FALLBACK_SESSION_SECRET;
const isDevelopment = process.env.NODE_ENV !== 'production';
const isSecureAppUrl = APP_URL.startsWith('https://');
const appOrigin = new URL(APP_URL).origin;
const SERVER_DOWNLOAD_ROOT = process.env.SERVER_DOWNLOAD_ROOT;
const serverDownloads = new ServerDownloadManager(SERVER_DOWNLOAD_ROOT);
let msalClient: msal.ConfidentialClientApplication | null = null;

if (!isDevelopment) {
    const missing = [
        !process.env.APP_URL ? 'APP_URL' : '',
        (!process.env.SESSION_SECRET || SESSION_SECRET === FALLBACK_SESSION_SECRET) ? 'SESSION_SECRET' : '',
        !CLIENT_ID ? 'MICROSOFT_CLIENT_ID' : '',
        !CLIENT_SECRET ? 'MICROSOFT_CLIENT_SECRET' : '',
    ].filter(Boolean);
    if (missing.length > 0) {
        throw new Error(`Missing required production environment variable${missing.length === 1 ? '' : 's'}: ${missing.join(', ')}`);
    }
}

function isUsableAccessToken(value: unknown) {
    if (typeof value !== 'string') return false;
    const token = value.trim();
    return token.length > 20
        && !token.includes(' ')
        && token !== 'undefined'
        && token !== 'null'
        && !token.startsWith('M.');
}

function createMsalClient() {
    if (!CLIENT_ID || !CLIENT_SECRET) {
        throw new Error('MICROSOFT_CLIENT_ID and MICROSOFT_CLIENT_SECRET must be configured');
    }

    const msalConfig: msal.Configuration = {
        auth: {
            clientId: CLIENT_ID,
            authority: 'https://login.microsoftonline.com/common',
            clientSecret: CLIENT_SECRET,
        }
    };

    if (!msalClient) {
        msalClient = new msal.ConfidentialClientApplication(msalConfig);
    }
    return msalClient;
}

function generateOAuthState() {
    return randomBytes(32).toString('base64url');
}

function safeEqual(a: string, b: string) {
    const aBuffer = Buffer.from(a);
    const bBuffer = Buffer.from(b);
    return aBuffer.length === bBuffer.length && timingSafeEqual(aBuffer, bBuffer);
}

export function validateOAuthState(expected: unknown, actual: unknown) {
    return typeof expected === 'string'
        && typeof actual === 'string'
        && expected.length > 0
        && safeEqual(expected, actual);
}

async function refreshSessionToken(req: express.Request) {
    const session = req.session as any;
    if (!session.account) return session.accessToken;

    const expiresAt = Number(session.expiresAt || 0);
    if (session.accessToken && expiresAt > Date.now() + 5 * 60 * 1000) {
        return session.accessToken;
    }

    const cca = createMsalClient();
    const response = await cca.acquireTokenSilent({
        account: session.account,
        scopes: ['User.Read', 'Files.Read.All', 'offline_access'],
    });
    session.accessToken = response.accessToken;
    session.expiresAt = response.expiresOn?.getTime() || Date.now() + 45 * 60 * 1000;
    if (response.account) session.account = response.account;
    await new Promise<void>((resolve, reject) => req.session.save(error => error ? reject(error) : resolve()));
    return session.accessToken;
}

async function startServer() {
    const app = express();
    const PORT = Number(process.env.PORT || 3000);

    app.use(express.json({ limit: '1mb' }));
    app.use(cookieParser());
    app.set('trust proxy', 1);
    app.use(session({
        secret: SESSION_SECRET,
        resave: false,
        saveUninitialized: false,
        cookie: {
            secure: isSecureAppUrl,
            sameSite: isSecureAppUrl ? 'none' : 'lax',
            httpOnly: true,
            maxAge: 24 * 60 * 60 * 1000 // 24 hours
        }
    }));

    app.get('/api/health', (_req, res) => {
        res.json({ ok: true });
    });

    // Logging middleware. Do not log callback query strings because they contain OAuth codes.
    app.use((req, res, next) => {
        if (req.path.includes('/api/') || req.path.includes('/callback')) {
            console.log(`[AUTH-DEBUG] ${new Date().toISOString()} - ${req.method} ${req.path}`);
        }
        next();
    });

    // API Routes
    app.get('/api/auth/status', (req, res) => {
        const session = (req.session as any);
        if (session.accessToken) {
            if (!isUsableAccessToken(session.accessToken)) {
                console.warn('[AUTH-DEBUG] Invalid session token found; clearing session');
                req.session.destroy(() => {
                    res.clearCookie('connect.sid');
                    res.json({ authenticated: false });
                });
                return;
            }
            refreshSessionToken(req)
                .then(accessToken => {
                    console.log('[AUTH-DEBUG] Session token found for status check');
                    res.json({ authenticated: true, accessToken });
                })
                .catch(error => {
                    console.error('[AUTH-DEBUG] Token refresh failed:', error);
                    req.session.destroy(() => {
                        res.clearCookie('connect.sid');
                        res.json({ authenticated: false });
                    });
                });
        } else {
            res.json({ authenticated: false });
        }
    });

    app.get('/api/auth/token', async (req, res) => {
        try {
            const accessToken = await refreshSessionToken(req);
            if (!isUsableAccessToken(accessToken)) {
                return res.status(401).json({ error: 'No valid Microsoft access token in session' });
            }
            res.json({ accessToken });
        } catch (error) {
            console.error('[AUTH-DEBUG] Token endpoint refresh failed:', error);
            req.session.destroy(() => {
                res.clearCookie('connect.sid');
                res.status(401).json({ error: 'Microsoft session expired. Please sign in again.' });
            });
        }
    });

    app.post('/api/auth/logout', (req, res) => {
        req.session.destroy(() => {
            res.clearCookie('connect.sid');
            res.json({ ok: true });
        });
    });

    function requireServerDownloadRoot(res: express.Response) {
        if (!SERVER_DOWNLOAD_ROOT) {
            res.status(400).json({ error: 'SERVER_DOWNLOAD_ROOT is not configured on this server.' });
            return false;
        }
        return true;
    }

    async function requireAccessToken(req: express.Request, res: express.Response) {
        try {
            const accessToken = await refreshSessionToken(req);
            if (!isUsableAccessToken(accessToken)) {
                res.status(401).json({ error: 'Sign in to Microsoft before starting server-side downloads.' });
                return null;
            }
            return accessToken;
        } catch {
            res.status(401).json({ error: 'Microsoft session expired. Sign in again.' });
            return null;
        }
    }

    app.get('/api/server-jobs/config', (_req, res) => {
        res.json({
            configured: Boolean(SERVER_DOWNLOAD_ROOT),
            targetRoot: SERVER_DOWNLOAD_ROOT || null,
        });
    });

    app.get('/api/server-jobs', (_req, res) => {
        res.json({ jobs: serverDownloads.listJobs() });
    });

    app.get('/api/server-jobs/:id', (req, res) => {
        const job = serverDownloads.getJob(req.params.id);
        if (!job) return res.status(404).json({ error: 'Server job not found.' });
        res.json({ job });
    });

    app.post('/api/server-jobs/:id/cancel', (req, res) => {
        const cancelled = serverDownloads.cancel(req.params.id);
        if (!cancelled) return res.status(404).json({ error: 'Server job not found.' });
        res.json({ ok: true });
    });

    async function startServerJob(req: express.Request, res: express.Response, mode: 'start' | 'dry-run' | 'repair') {
        if (!requireServerDownloadRoot(res)) return;
        const accessToken = await requireAccessToken(req, res);
        if (!accessToken) return;
        const selections = Array.isArray(req.body?.selections) ? req.body.selections : [];
        if (selections.length === 0) {
            return res.status(400).json({ error: 'Select OneDrive files or folders before starting a server job.' });
        }
        const getAccessToken = async () => refreshSessionToken(req);
        try {
            const job = serverDownloads.start(mode, selections, req.body?.settings || {}, getAccessToken);
            res.json({ job });
        } catch (error) {
            res.status(400).json({ error: error instanceof Error ? error.message : String(error) });
        }
    }

    app.post('/api/server-jobs/start', (req, res) => startServerJob(req, res, 'start'));
    app.post('/api/server-jobs/dry-run', (req, res) => startServerJob(req, res, 'dry-run'));
    app.post('/api/server-jobs/repair', (req, res) => startServerJob(req, res, 'repair'));

    app.get('/api/auth/url', async (req, res) => {
        console.log('[AUTH-DEBUG] Generating auth URL');
        if (!CLIENT_ID || !CLIENT_SECRET) {
            return res.status(500).json({ error: 'MICROSOFT_CLIENT_ID and MICROSOFT_CLIENT_SECRET are not configured. Copy .env.example to .env.local and fill in your Azure app registration values.' });
        }

        const state = generateOAuthState();
        (req.session as any).oauthState = state;

        const authCodeUrlParameters = {
            scopes: ['User.Read', 'Files.Read.All', 'offline_access'],
            redirectUri: REDIRECT_URI,
            prompt: 'select_account',
            state,
        };

        try {
            const cca = createMsalClient();
            const url = await cca.getAuthCodeUrl(authCodeUrlParameters);
            console.log('[AUTH-DEBUG] Success generating URL');
            res.json({ url });
        } catch (error) {
            console.error('[AUTH-DEBUG] Auth URL Error:', error);
            res.status(500).json({ error: 'Failed to generate auth URL' });
        }
    });

    app.get('/api/callback', async (req, res) => {
        console.log('[AUTH-DEBUG] Callback hit');
        const { code, error, error_description, state } = req.query;

        if (error) {
            console.error('[AUTH-DEBUG] Auth Error in callback:', error, error_description);
            return res.status(400).send(`Authentication Error: ${error_description}`);
        }

        if (!code) {
            console.warn('[AUTH-DEBUG] No code in callback query');
            return res.status(400).send('No code provided in redirect');
        }

        const expectedState = (req.session as any).oauthState;
        delete (req.session as any).oauthState;
        if (!validateOAuthState(expectedState, Array.isArray(state) ? state[0] : state)) {
            console.warn('[AUTH-DEBUG] Rejected callback with invalid OAuth state');
            return res.status(400).send('Invalid authentication state. Please start sign in again.');
        }

        const tokenRequest = {
            code: code as string,
            scopes: ['User.Read', 'Files.Read.All', 'offline_access'],
            redirectUri: REDIRECT_URI,
        };

        try {
            console.log('[AUTH-DEBUG] Exchanging code for token...');
            const cca = createMsalClient();
            const response = await cca.acquireTokenByCode(tokenRequest);
            const accessToken = response.accessToken;
            if (!isUsableAccessToken(accessToken)) {
                console.error('[AUTH-DEBUG] MSAL returned an unusable access token');
                return res.status(500).send('Authentication failed: Microsoft did not return a usable Graph access token.');
            }
            console.log('[AUTH-DEBUG] Token acquired successfully');

            // Save to session for polling fallback
            (req.session as any).accessToken = accessToken;
            (req.session as any).expiresAt = response.expiresOn?.getTime() || Date.now() + 45 * 60 * 1000;
            (req.session as any).account = response.account;
            req.session.save((sessionError) => {
                if (sessionError) {
                    console.error('[AUTH-DEBUG] Session save error:', sessionError);
                }

                // Send success message to parent window and close popup
                res.send(`
                <html>
                    <body style="font-family: sans-serif; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; margin: 0; background-color: #F0F9FF;">
                        <div style="background: white; padding: 2rem; border-radius: 12px; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1); border: 1px solid #BAE6FD; text-align: center;">
                            <div style="color: #0284C7; font-size: 3rem; margin-bottom: 1rem;">✓</div>
                            <h2 style="color: #0369A1; margin: 0 0 0.5rem 0;">Auth Success</h2>
                            <p id="status-text" style="color: #64748B; margin-bottom: 1.5rem;">Authenticating with OneDrive Archiver...</p>
                            <div style="width: 24px; height: 24px; border: 3px solid #E0F2FE; border-top-color: #0284C7; border-radius: 50%; display: inline-block; animation: spin 1s linear infinite;"></div>
                        </div>
                        <style>
                            @keyframes spin { to { transform: rotate(360deg); } }
                        </style>
                        <script>
                            console.log('Finalizing authentication...');

                            // 1. Notify the main window to poll the server session.
                            try {
                                const channel = new BroadcastChannel('onedrive_auth_channel');
                                channel.postMessage({ type: 'OAUTH_AUTH_SUCCESS' });
                                console.log('BroadcastChannel sent');
                            } catch (e) {
                                console.error('BroadcastChannel failed:', e);
                            }

                            // 2. Try postMessage to opener using the configured app origin.
                            if (window.opener) {
                                try {
                                    window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS' }, ${JSON.stringify(appOrigin)});
                                    console.log('postMessage sent');
                                } catch (e) {
                                    console.error('postMessage failed:', e);
                                }
                            }

                            // Update UI to show completion
                            document.getElementById('status-text').innerText = 'Authenticated! Closing...';

                            // Close the popup after a short delay
                            setTimeout(() => {
                                try {
                                    window.close();
                                } catch (e) {}
                                
                                // Fallback if close fails
                                document.body.innerHTML = '<div style="padding:40px; text-align:center; font-family:sans-serif;">' +
                                    '<h2 style="color:#0369A1;">Success!</h2>' +
                                    '<p>You have been signed in. You can close this window now.</p>' +
                                    '<button onclick="window.close()" style="padding:10px 20px; background:#0284C7; color:white; border:none; border-radius:5px; cursor:pointer;">Close Window</button>' +
                                '</div>';
                            }, 1500);
                        </script>
                    </body>
                </html>
                `);
            });
        } catch (error: any) {
            console.error('[AUTH-DEBUG] Exchange error:', error);
            let errorMessage = 'Authentication failed. Please check your server logs.';
            
            if (error.errorMessage && error.errorMessage.includes('AADSTS7000215')) {
                errorMessage = 'Invalid Client Secret. Ensure you used the "Value" (not the "ID") from the Azure Portal Secrets section.';
            }

            res.status(500).send(`
                <html>
                    <body style="font-family: sans-serif; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; background-color: #FEF2F2;">
                        <div style="background: white; padding: 2rem; border-radius: 8px; border: 1px solid #FECACA; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);">
                            <h2 style="color: #991B1B; margin-top: 0;">Configuration Error</h2>
                            <p style="color: #4B5563;">${errorMessage}</p>
                            <button onclick="window.close()" style="margin-top: 1rem; padding: 0.5rem 1rem; background: #EF4444; color: white; border: none; rounded: 4px; cursor: pointer;">Close Window</button>
                        </div>
                    </body>
                </html>
            `);
        }
    });

    // Vite middleware
    if (process.env.NODE_ENV !== 'production') {
        const vite = await createViteServer({
            server: { middlewareMode: true },
            appType: 'spa',
        });
        app.use(vite.middlewares);
    } else {
        const distPath = path.join(process.cwd(), 'dist');
        app.use(express.static(distPath));
        app.get('*', (req, res) => {
            res.sendFile(path.join(distPath, 'index.html'));
        });
    }

    app.listen(PORT, '0.0.0.0', () => {
        console.log(`Server running at http://0.0.0.0:${PORT}`);
        console.log(`[CONFIG] APP_URL: ${APP_URL}`);
        console.log(`[CONFIG] REDIRECT_URI: ${REDIRECT_URI}`);
    });
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
    startServer();
}
