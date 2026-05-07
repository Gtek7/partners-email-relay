const express = require('express');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const Stripe = require('stripe');

const app = express();

// Stripe webhook needs raw body — must come before express.json()
app.use('/stripe/webhook', express.raw({ type: 'application/json' }));

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const {
  MS_CLIENT_ID,
  MS_CLIENT_SECRET,
  MS_REDIRECT_URI,
  BASE_URL,
  N8N_WEBHOOK_URL,
  STRIPE_SECRET_KEY,
  STRIPE_WEBHOOK_SECRET,
  NO_SHOW_FEE_CENTS = '5250',   // Default $52.50 CAD ($50 + 5% GST)
  PORT = 3000
} = process.env;

const stripe = Stripe(STRIPE_SECRET_KEY);

// In-memory store: phone → { customerId, paymentMethodId, appointmentDateTime, serviceType }
// This persists between requests but resets on server restart.
// For production consider a database.
const paymentStore = {};

// Token + subscription state (in-memory; survives between requests, lost on restart)
let state = {
  accessToken: null,
  refreshToken: null,
  tokenExpiry: null,       // Date
  subscriptionId: null,
  subscriptionExpiry: null // Date
};

// ─── Auth ──────────────────────────────────────────────────────────────────

// Visit this URL once to authorize the app with gtektest@outlook.com
app.get('/auth/login', (req, res) => {
  const params = new URLSearchParams({
    client_id: MS_CLIENT_ID,
    response_type: 'code',
    redirect_uri: MS_REDIRECT_URI,
    scope: 'openid offline_access Mail.Read',
    prompt: 'login'
  });
  res.redirect(`https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?${params}`);
});

// Microsoft redirects here after login
app.get('/auth/callback', async (req, res) => {
  const { code, error } = req.query;
  if (error) return res.status(400).send(`Auth error: ${error}`);
  if (!code) return res.status(400).send('Missing code');

  try {
    const tokenRes = await axios.post(
      'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
      new URLSearchParams({
        client_id: MS_CLIENT_ID,
        client_secret: MS_CLIENT_SECRET,
        code,
        redirect_uri: MS_REDIRECT_URI,
        grant_type: 'authorization_code',
        scope: 'openid offline_access Mail.Read'
      }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    storeTokens(tokenRes.data);

    // Register the Graph subscription
    await registerSubscription();

    res.send(`
      <h2>✅ Authorization successful!</h2>
      <p>Email webhook is now active. Microsoft will push new emails from <strong>gtektest@outlook.com</strong> directly to n8n.</p>
      <p>Subscription ID: ${state.subscriptionId}</p>
      <p>Expires: ${state.subscriptionExpiry}</p>
      <p>You can close this window.</p>
    `);
  } catch (err) {
    console.error('Auth callback error:', err.response?.data || err.message);
    res.status(500).send('Authorization failed: ' + JSON.stringify(err.response?.data || err.message));
  }
});

// ─── Microsoft Graph Webhook ───────────────────────────────────────────────

// Microsoft validates the notification URL by sending a GET with validationToken
app.get('/graph/notify', (req, res) => {
  const { validationToken } = req.query;
  if (validationToken) {
    console.log('Graph subscription validated');
    res.setHeader('Content-Type', 'text/plain');
    return res.send(validationToken);
  }
  res.status(400).send('Missing validationToken');
});

// Microsoft POSTs here when a new email arrives
app.post('/graph/notify', async (req, res) => {
  // MUST respond 202 within 3 seconds or Microsoft will retry
  res.status(202).send();

  const notifications = req.body?.value || [];

  for (const notification of notifications) {
    // Skip lifecycle notifications (e.g. subscriptionRemoved)
    if (notification.lifecycleEvent) {
      console.log('Lifecycle event:', notification.lifecycleEvent);
      if (notification.lifecycleEvent === 'subscriptionRemoved') {
        // Re-register
        try {
          await ensureValidToken();
          await registerSubscription();
        } catch (e) {
          console.error('Failed to re-register subscription:', e.message);
        }
      }
      continue;
    }

    const msgId = notification.resourceData?.id;
    if (!msgId) continue;

    try {
      await ensureValidToken();

      // Fetch full message details from Graph
      const msgRes = await axios.get(
        `https://graph.microsoft.com/v1.0/me/messages/${msgId}?$select=id,subject,bodyPreview,body,from,conversationId,receivedDateTime`,
        { headers: { Authorization: `Bearer ${state.accessToken}` } }
      );

      const msg = msgRes.data;
      console.log('New email received:', msg.subject, 'from:', msg.from?.emailAddress?.address);

      // Forward to n8n webhook (same shape as what Power Automate was sending)
      await axios.post(N8N_WEBHOOK_URL, {
        id: msg.id,
        subject: msg.subject,
        bodyPreview: msg.bodyPreview,
        conversationId: msg.conversationId,
        receivedDateTime: msg.receivedDateTime,
        body: { content: msg.body?.content },
        sender: msg.from
      });

      console.log('Forwarded to n8n:', msg.subject);
    } catch (err) {
      console.error('Error processing notification:', err.response?.data || err.message);
    }
  }
});

// ─── Stripe ───────────────────────────────────────────────────────────────

/**
 * POST /stripe/create-setup
 * Called by Vapi after a booking is confirmed.
 * Creates a Stripe Checkout session in "setup" mode (saves card, no charge).
 * Returns { url } — Vapi sends this to the caller via SMS.
 *
 * Body: { customerName, customerEmail, customerPhone, appointmentDateTime, serviceType }
 */
app.post('/stripe/create-setup', async (req, res) => {
  const { customerName, customerEmail, customerPhone, appointmentDateTime, serviceType } = req.body;

  if (!customerPhone) {
    return res.status(400).json({ error: 'customerPhone is required' });
  }

  try {
    // Create or retrieve a Stripe customer keyed by phone
    let customer;
    const existing = await stripe.customers.search({
      query: `phone:'${customerPhone}'`,
      limit: 1
    });

    if (existing.data.length > 0) {
      customer = existing.data[0];
    } else {
      customer = await stripe.customers.create({
        name: customerName || 'Guest',
        email: customerEmail || undefined,
        phone: customerPhone,
        metadata: { source: 'haya_high_receptionist' }
      });
    }

    // Create a Checkout Session in setup mode
    const session = await stripe.checkout.sessions.create({
      mode: 'setup',
      currency: 'cad',
      customer: customer.id,
      payment_method_types: ['card'],
      success_url: `${BASE_URL}/stripe/success?session_id={CHECKOUT_SESSION_ID}`,
      cancel_url: `${BASE_URL}/stripe/cancel`,
      metadata: {
        customerPhone,
        customerName: customerName || '',
        appointmentDateTime: appointmentDateTime || '',
        serviceType: serviceType || ''
      }
    });

    console.log(`Stripe setup session created for ${customerPhone}: ${session.id}`);
    res.json({ url: session.url, sessionId: session.id });

  } catch (err) {
    console.error('Stripe create-setup error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

/**
 * POST /stripe/charge
 * Charges the saved card for a no-show or late cancellation.
 * Body: { customerPhone, reason, appointmentDateTime }
 */
app.post('/stripe/charge', async (req, res) => {
  const { customerPhone, reason = 'No-show / late cancellation fee', appointmentDateTime } = req.body;

  if (!customerPhone) {
    return res.status(400).json({ error: 'customerPhone is required' });
  }

  const stored = paymentStore[customerPhone];
  if (!stored) {
    return res.status(404).json({ error: 'No saved payment method found for this phone number' });
  }

  try {
    const { customerId, paymentMethodId } = stored;
    const feeCents = parseInt(NO_SHOW_FEE_CENTS, 10);

    const paymentIntent = await stripe.paymentIntents.create({
      amount: feeCents,
      currency: 'cad',
      customer: customerId,
      payment_method: paymentMethodId,
      confirm: true,
      off_session: true,
      description: `${reason}${appointmentDateTime ? ' — ' + appointmentDateTime : ''}`,
      metadata: { customerPhone, appointmentDateTime: appointmentDateTime || '' }
    });

    console.log(`Charged ${feeCents / 100} CAD for ${customerPhone}: ${paymentIntent.id}`);
    res.json({
      success: true,
      chargeId: paymentIntent.id,
      amount: feeCents / 100,
      currency: 'cad'
    });

  } catch (err) {
    console.error('Stripe charge error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

/**
 * POST /stripe/webhook
 * Stripe sends events here. We listen for checkout.session.completed
 * to save the customer's payment method after they fill in their card.
 */
app.post('/stripe/webhook', (req, res) => {
  const sig = req.headers['stripe-signature'];
  let event;

  try {
    event = stripe.webhooks.constructEvent(req.body, sig, STRIPE_WEBHOOK_SECRET);
  } catch (err) {
    console.error('Stripe webhook signature error:', err.message);
    return res.status(400).send(`Webhook Error: ${err.message}`);
  }

  if (event.type === 'checkout.session.completed') {
    const session = event.data.object;

    if (session.mode === 'setup' && session.setup_intent) {
      handleSetupComplete(session).catch(err =>
        console.error('Error handling setup complete:', err.message)
      );
    }
  }

  res.json({ received: true });
});

async function handleSetupComplete(session) {
  const setupIntent = await stripe.setupIntents.retrieve(session.setup_intent);
  const paymentMethodId = setupIntent.payment_method;
  const customerPhone = session.metadata?.customerPhone;

  if (!customerPhone || !paymentMethodId) {
    console.warn('handleSetupComplete: missing phone or paymentMethod');
    return;
  }

  // Attach payment method to customer
  await stripe.paymentMethods.attach(paymentMethodId, { customer: session.customer });

  // Set as default
  await stripe.customers.update(session.customer, {
    invoice_settings: { default_payment_method: paymentMethodId }
  });

  // Store in memory
  paymentStore[customerPhone] = {
    customerId: session.customer,
    paymentMethodId,
    appointmentDateTime: session.metadata?.appointmentDateTime || '',
    serviceType: session.metadata?.serviceType || '',
    savedAt: new Date().toISOString()
  };

  console.log(`Card saved for ${customerPhone}: customer=${session.customer}, pm=${paymentMethodId}`);
}

// Success / cancel pages shown after Stripe checkout
app.get('/stripe/success', (req, res) => {
  res.send(`
    <html><body style="font-family:sans-serif;text-align:center;padding:60px">
      <h2>✅ Card saved successfully!</h2>
      <p>Your appointment at Haya-High Massage is confirmed.</p>
      <p>Your card will only be charged if you cancel within 24 hours or miss your appointment.</p>
    </body></html>
  `);
});

app.get('/stripe/cancel', (req, res) => {
  res.send(`
    <html><body style="font-family:sans-serif;text-align:center;padding:60px">
      <h2>Payment setup cancelled</h2>
      <p>Your appointment has been held, but please call us back to complete card verification.</p>
    </body></html>
  `);
});

// ─── Health Check ──────────────────────────────────────────────────────────

app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    authorized: !!state.accessToken,
    tokenExpiry: state.tokenExpiry,
    subscriptionId: state.subscriptionId,
    subscriptionExpiry: state.subscriptionExpiry
  });
});

// ─── Helpers ───────────────────────────────────────────────────────────────

function storeTokens(data) {
  state.accessToken = data.access_token;
  if (data.refresh_token) state.refreshToken = data.refresh_token;
  // Access tokens expire in 3600s; refresh 5 min early
  state.tokenExpiry = new Date(Date.now() + (data.expires_in - 300) * 1000);
}

async function ensureValidToken() {
  if (!state.tokenExpiry || Date.now() >= state.tokenExpiry.getTime()) {
    if (!state.refreshToken) throw new Error('No refresh token — re-authorize at /auth/login');
    const res = await axios.post(
      'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
      new URLSearchParams({
        client_id: MS_CLIENT_ID,
        client_secret: MS_CLIENT_SECRET,
        refresh_token: state.refreshToken,
        grant_type: 'refresh_token',
        scope: 'openid offline_access Mail.Read'
      }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );
    storeTokens(res.data);
    console.log('Token refreshed, expires:', state.tokenExpiry);
  }
}

async function registerSubscription() {
  // Personal account subscriptions max out at ~4230 minutes (~3 days)
  const expirationDateTime = new Date(Date.now() + 4000 * 60 * 1000).toISOString();

  const res = await axios.post(
    'https://graph.microsoft.com/v1.0/subscriptions',
    {
      changeType: 'created',
      notificationUrl: `${BASE_URL}/graph/notify`,
      resource: "me/mailFolders('inbox')/messages",
      expirationDateTime,
      clientState: 'partners-email-relay'
    },
    {
      headers: {
        Authorization: `Bearer ${state.accessToken}`,
        'Content-Type': 'application/json'
      }
    }
  );

  state.subscriptionId = res.data.id;
  state.subscriptionExpiry = new Date(res.data.expirationDateTime);
  console.log('Subscription registered:', state.subscriptionId, 'expires:', state.subscriptionExpiry);
}

async function renewSubscription() {
  const newExpiry = new Date(Date.now() + 4000 * 60 * 1000).toISOString();
  try {
    await ensureValidToken();
    await axios.patch(
      `https://graph.microsoft.com/v1.0/subscriptions/${state.subscriptionId}`,
      { expirationDateTime: newExpiry },
      {
        headers: {
          Authorization: `Bearer ${state.accessToken}`,
          'Content-Type': 'application/json'
        }
      }
    );
    state.subscriptionExpiry = new Date(newExpiry);
    console.log('Subscription renewed until:', state.subscriptionExpiry);
  } catch (err) {
    console.error('Renewal failed, re-registering:', err.response?.data || err.message);
    await registerSubscription();
  }
}

// Check every hour; renew subscription if it expires within 12 hours
setInterval(async () => {
  if (!state.subscriptionId || !state.subscriptionExpiry) return;
  const hoursLeft = (state.subscriptionExpiry.getTime() - Date.now()) / (1000 * 60 * 60);
  console.log(`Subscription check: ${hoursLeft.toFixed(1)}h remaining`);
  if (hoursLeft < 12) {
    console.log('Renewing subscription...');
    await renewSubscription();
  }
}, 60 * 60 * 1000);

app.listen(PORT, () => {
  console.log(`Partners Email Relay running on port ${PORT}`);
  console.log(`Authorize at: ${BASE_URL}/auth/login`);
});
