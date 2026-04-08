# Stripe Billing Setup — What You Need To Do

Everything in the codebase is already wired. These are the manual steps you need to complete before billing goes live.

---

## Step 1 — Create or log into your Stripe account

Go to [https://dashboard.stripe.com](https://dashboard.stripe.com) and sign in (or create a free account).

> Use **Test mode** (toggle in the top-left) while you are testing. Switch to **Live mode** only when you are ready to charge real users.

---

## Step 2 — Get your Stripe Secret Key

1. In the Stripe dashboard, go to **Developers → API keys**.
2. Copy the **Secret key** (starts with `sk_test_...` for test mode, `sk_live_...` for live).

> Keep this safe — treat it like a password. Never commit it to code.

---

## Step 3 — Create the $5/month Pro price

Run this one-time script from your project root (replace the key with yours):

```bash
STRIPE_SECRET_KEY=sk_test_... node scripts/stripe-seed.mjs
```

The script will print something like:

```
Created price: price_1Abc123XYZ
Set STRIPE_PRO_PRICE_ID=price_1Abc123XYZ
```

Copy that price ID — you will need it in Step 5.

> If you already created the price manually in the Stripe dashboard, the script will find it and print the existing ID instead.

---

## Step 4 — Configure the Stripe Customer Portal

This powers the "Manage subscription" button so users can cancel or update their payment method.

1. In Stripe, go to **Settings → Billing → Customer portal**.
2. Click **Activate portal**.
3. Under **Functionality**, enable at minimum:
   - Cancel subscriptions
   - Update payment methods
4. Save.

---

## Step 5 — Set environment variables on Railway

In your Railway project, go to your API server service → **Variables** and add:

| Variable | Value |
|---|---|
| `STRIPE_SECRET_KEY` | Your secret key from Step 2 |
| `STRIPE_PRO_PRICE_ID` | The price ID from Step 3 |
| `STRIPE_WEBHOOK_SECRET` | *(set this in Step 6 below)* |

---

## Step 6 — Register the Stripe webhook

Stripe needs to call your API whenever a subscription changes (payment success, cancellation, etc.).

1. In Stripe, go to **Developers → Webhooks → Add endpoint**.
2. Set the **Endpoint URL** to your Railway API URL + `/api/stripe/webhook`:
   ```
   https://your-api.railway.app/api/stripe/webhook
   ```
3. Under **Events to listen to**, select:
   - `checkout.session.completed`
   - `customer.subscription.deleted`
   - `customer.subscription.updated`
4. Click **Add endpoint**.
5. On the next screen, click **Reveal** next to **Signing secret**.
6. Copy that value and add it to Railway as `STRIPE_WEBHOOK_SECRET`.

---

## Step 7 — Redeploy on Railway

After adding all three environment variables, trigger a redeploy of your API server on Railway so the new values take effect.

---

## Step 8 — Test the full flow

Use Stripe's test card **`4242 4242 4242 4242`** (any future expiry, any CVC):

1. Open your app and sign in.
2. Click **Upgrade** in the top navigation or the user dropdown.
3. Complete checkout with the test card.
4. After redirect back, confirm the header shows **Pro** instead of **Free**.
5. Click **Manage subscription** in the dropdown to confirm the portal opens.
6. Cancel the subscription in the portal and confirm the badge returns to **Free**.

---

## Going live

When you are ready for real users:

1. Switch Stripe dashboard to **Live mode**.
2. Repeat Steps 2–3 and 6 using your **live** keys (they start with `sk_live_` and `whsec_`).
3. Update the Railway environment variables with the live values.
4. Redeploy.
