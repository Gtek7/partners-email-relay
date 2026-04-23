const express = require('express');
const axios = require('axios');

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const { MS_CLIENT_ID, MS_CLIENT_SECRET, MS_REDIRECT_URI, BASE_URL, N8N_WEBHOOK_URL, PORT = 3000 } = process.env;

let state = { accessToken: null, refreshToken: null, tokenExpiry: null, subscriptionId: null, subscriptionExpiry: null };

app.get('/auth/login', (req, res) => {
  const params = new URLSearchParams({ client_id: MS_CLIENT_ID, response_type: 'code', redirect_uri: MS_REDIRECT_URI, scope: 'openid offline_access Mail.Read', prompt: 'login' });
  res.redirect('https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?' + params);
});

app.get('/auth/callback', async (req, res) => {
  const { code, error } = req.query;
  if (error) return res.status(400).send('Auth error: ' + error);
  if (!code) return res.status(400).send('Missing code');
  try {
    const tokenRes = await axios.post('https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
      new URLSearchParams({ client_id: MS_CLIENT_ID, client_secret: MS_CLIENT_SECRET, code, redirect_uri: MS_REDIRECT_URI, grant_type: 'authorization_code', scope: 'openid offline_access Mail.Read' }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } });
    storeTokens(tokenRes.data);
    await registerSubscription();
    res.send('<h2>Authorization successful!</h2><p>Email webhook active.</p><p>Subscription ID: ' + state.subscriptionId + '</p><p>Expires: ' + state.subscriptionExpiry + '</p>');
  } catch (err) {
    console.error('Auth callback error:', err.response?.data || err.message);
    res.status(500).send('Authorization failed: ' + JSON.stringify(err.response?.data || err.message));
  }
});

app.get('/graph/notify', (req, res) => {
  const { validationToken } = req.query;
  if (validationToken) { console.log('Graph subscription validated'); res.setHeader('Content-Type', 'text/plain'); return res.send(validationToken); }
  res.status(400).send('Missing validationToken');
});

app.post('/graph/notify', async (req, res) => {
  res.status(202).send();
  const notifications = req.body?.value || [];
  for (const notification of notifications) {
    if (notification.lifecycleEvent) {
      console.log('Lifecycle event:', notification.lifecycleEvent);
      if (notification.lifecycleEvent === 'subscriptionRemoved') {
        try { await ensureValidToken(); await registerSubscription(); } catch (e) { console.error('Re-register failed:', e.message); }
      }
      continue;
    }
    const msgId = notification.resourceData?.id;
    if (!msgId) continue;
    try {
      await ensureValidToken();
      const msgRes = await axios.get('https://graph.microsoft.com/v1.0/me/messages/' + msgId + '?$select=id,subject,bodyPreview,body,from,conversationId,receivedDateTime',
        { headers: { Authorization: 'Bearer ' + state.accessToken } });
      const msg = msgRes.data;
      console.log('New email:', msg.subject, 'from:', msg.from?.emailAddress?.address);
      await axios.post(N8N_WEBHOOK_URL, { id: msg.id, subject: msg.subject, bodyPreview: msg.bodyPreview, conversationId: msg.conversationId, receivedDateTime: msg.receivedDateTime, body: { content: msg.body?.content }, sender: msg.from });
      console.log('Forwarded to n8n:', msg.subject);
    } catch (err) { console.error('Error processing notification:', err.response?.data || err.message); }
  }
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok', authorized: !!state.accessToken, tokenExpiry: state.tokenExpiry, subscriptionId: state.subscriptionId, subscriptionExpiry: state.subscriptionExpiry });
});

function storeTokens(data) {
  state.accessToken = data.access_token;
  if (data.refresh_token) state.refreshToken = data.refresh_token;
  state.tokenExpiry = new Date(Date.now() + (data.expires_in - 300) * 1000);
}

async function ensureValidToken() {
  if (!state.tokenExpiry || Date.now() >= state.tokenExpiry.getTime()) {
    if (!state.refreshToken) throw new Error('No refresh token - re-authorize at /auth/login');
    const res = await axios.post('https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
      new URLSearchParams({ client_id: MS_CLIENT_ID, client_secret: MS_CLIENT_SECRET, refresh_token: state.refreshToken, grant_type: 'refresh_token', scope: 'openid offline_access Mail.Read' }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } });
    storeTokens(res.data);
    console.log('Token refreshed, expires:', state.tokenExpiry);
  }
}

async function registerSubscription() {
  const expirationDateTime = new Date(Date.now() + 4000 * 60 * 1000).toISOString();
  const res = await axios.post('https://graph.microsoft.com/v1.0/subscriptions',
    { changeType: 'created', notificationUrl: BASE_URL + '/graph/notify', resource: "me/mailFolders('inbox')/messages", expirationDateTime, clientState: 'partners-email-relay' },
    { headers: { Authorization: 'Bearer ' + state.accessToken, 'Content-Type': 'application/json' } });
  state.subscriptionId = res.data.id;
  state.subscriptionExpiry = new Date(res.data.expirationDateTime);
  console.log('Subscription registered:', state.subscriptionId, 'expires:', state.subscriptionExpiry);
}

async function renewSubscription() {
  const newExpiry = new Date(Date.now() + 4000 * 60 * 1000).toISOString();
  try {
    await ensureValidToken();
    await axios.patch('https://graph.microsoft.com/v1.0/subscriptions/' + state.subscriptionId, { expirationDateTime: newExpiry },
      { headers: { Authorization: 'Bearer ' + state.accessToken, 'Content-Type': 'application/json' } });
    state.subscriptionExpiry = new Date(newExpiry);
    console.log('Subscription renewed until:', state.subscriptionExpiry);
  } catch (err) {
    console.error('Renewal failed, re-registering:', err.response?.data || err.message);
    await registerSubscription();
  }
}

setInterval(async () => {
  if (!state.subscriptionId || !state.subscriptionExpiry) return;
  const hoursLeft = (state.subscriptionExpiry.getTime() - Date.now()) / (1000 * 60 * 60);
  console.log('Subscription check: ' + hoursLeft.toFixed(1) + 'h remaining');
  if (hoursLeft < 12) { console.log('Renewing subscription...'); await renewSubscription(); }
}, 60 * 60 * 1000);

app.listen(PORT, () => { console.log('Partners Email Relay on port ' + PORT); console.log('Authorize at: ' + BASE_URL + '/auth/login'); });