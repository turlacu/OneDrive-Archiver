import express from 'express';
import { createServer as createViteServer } from 'vite';
import path from 'path';
import { fileURLToPath } from 'url';
import * as msal from '@azure/msal-node';
import dotenv from 'dotenv';
import cookieParser from 'cookie-parser';
import session from 'express-session';

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

const msalConfig: msal.Configuration = {
    auth: {
        clientId: CLIENT_ID || '',
        authority: 'https://login.microsoftonline.com/common',
        clientSecret: CLIENT_SECRET,
    }
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

async function startServer() {
    const app = express();
    const PORT = 3000;

    app.use(cookieParser());
    app.use(session({
        secret: 'onedrive-sync-secret',
        resave: false,
        saveUninitialized: false,
        cookie: {
            secure: true,
            sameSite: 'none',
            httpOnly: true,
            maxAge: 24 * 60 * 60 * 1000 // 24 hours
        }
    }));

    // Logging middleware
    app.use((req, res, next) => {
        if (req.url.includes('/api/') || req.url.includes('/callback')) {
            console.log(`[AUTH-DEBUG] ${new Date().toISOString()} - ${req.method} ${req.url}`);
        }
        next();
    });

    // API Routes
    app.get('/api/auth/status', (req, res) => {
        const session = (req.session as any);
        if (session.accessToken) {
            console.log('[AUTH-DEBUG] Session token found for status check');
            res.json({ authenticated: true, accessToken: session.accessToken });
        } else {
            res.json({ authenticated: false });
        }
    });

    app.get('/api/auth/url', async (req, res) => {
        console.log('[AUTH-DEBUG] Generating auth URL for client:', CLIENT_ID);
        if (!CLIENT_ID || CLIENT_ID === '') {
            return res.status(500).json({ error: 'MICROSOFT_CLIENT_ID not configured in Secrets' });
        }

        const authCodeUrlParameters = {
            scopes: ['user.read', 'Files.Read.All', 'offline_access'],
            redirectUri: REDIRECT_URI,
        };

        try {
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
        const { code, error, error_description } = req.query;

        if (error) {
            console.error('[AUTH-DEBUG] Auth Error in callback:', error, error_description);
            return res.status(400).send(`Authentication Error: ${error_description}`);
        }

        if (!code) {
            console.warn('[AUTH-DEBUG] No code in callback query');
            return res.status(400).send('No code provided in redirect');
        }

        const tokenRequest = {
            code: code as string,
            scopes: ['user.read', 'Files.Read.All'],
            redirectUri: REDIRECT_URI,
        };

        try {
            console.log('[AUTH-DEBUG] Exchanging code for token...');
            const response = await cca.acquireTokenByCode(tokenRequest);
            const accessToken = response.accessToken;
            console.log('[AUTH-DEBUG] Token acquired successfully');

            // Save to session for polling fallback
            (req.session as any).accessToken = accessToken;
            req.session.save();
            
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
                            const token = '${accessToken}';
                            
                            console.log('Finalizing authentication...');
                            
                            // 1. Set localStorage (triggers 'storage' event in main window)
                            localStorage.setItem('ms_token', token);
                            localStorage.setItem('ms_token_signal', token);
                            
                            // Set a cookies as a final fallback for some environments
                            document.cookie = "ms_token=" + token + "; path=/; max-age=3600; SameSite=Lax";

                            // 2. Try BroadcastChannel
                            try {
                                const channel = new BroadcastChannel('onedrive_auth_channel');
                                channel.postMessage({ type: 'OAUTH_AUTH_SUCCESS', accessToken: token });
                                console.log('BroadcastChannel sent');
                            } catch (e) {
                                console.error('BroadcastChannel failed:', e);
                            }

                            // 3. Try postMessage to opener
                            if (window.opener) {
                                try {
                                    window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS', accessToken: token }, '*');
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

startServer();
