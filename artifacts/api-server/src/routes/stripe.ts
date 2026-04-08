import { Router, type IRouter, type Request, type Response } from "express";
import Stripe from "stripe";
import { db, usersTable } from "@workspace/db";
import { eq } from "drizzle-orm";
import { requireAuth } from "./user-auth";

const router: IRouter = Router();

function getStripe(): Stripe | null {
  const key = process.env["STRIPE_SECRET_KEY"];
  if (!key) return null;
  return new Stripe(key, { apiVersion: "2025-04-30.basil" });
}

function getFrontendUrl(req: Request): string {
  const corsOrigin = process.env["CORS_ORIGIN"];
  if (corsOrigin) return corsOrigin.split(",")[0].trim();
  return `${req.protocol}://${req.get("host")}`;
}

router.post("/stripe/checkout", requireAuth, async (req: Request, res: Response): Promise<void> => {
  const stripe = getStripe();
  if (!stripe) {
    res.status(503).json({ error: "Stripe is not configured" });
    return;
  }

  const priceId = process.env["STRIPE_PRO_PRICE_ID"];
  if (!priceId) {
    res.status(503).json({ error: "Stripe Pro price is not configured (STRIPE_PRO_PRICE_ID)" });
    return;
  }

  const user = req.user!;
  const frontendUrl = getFrontendUrl(req);

  try {
    let customerId = user.stripeCustomerId;

    if (!customerId) {
      const customer = await stripe.customers.create({
        email: user.email ?? undefined,
        name: user.name,
        metadata: { userId: user.id },
      });
      customerId = customer.id;
      await db.update(usersTable)
        .set({ stripeCustomerId: customerId, updatedAt: new Date() })
        .where(eq(usersTable.id, user.id));
    }

    const session = await stripe.checkout.sessions.create({
      customer: customerId,
      payment_method_types: ["card"],
      line_items: [{ price: priceId, quantity: 1 }],
      mode: "subscription",
      success_url: `${frontendUrl}/upgrade/success`,
      cancel_url: `${frontendUrl}/upgrade/cancelled`,
    });

    res.json({ url: session.url });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[stripe/checkout]", message);
    res.status(500).json({ error: message });
  }
});

router.post("/stripe/portal", requireAuth, async (req: Request, res: Response): Promise<void> => {
  const stripe = getStripe();
  if (!stripe) {
    res.status(503).json({ error: "Stripe is not configured" });
    return;
  }

  const user = req.user!;
  const frontendUrl = getFrontendUrl(req);

  if (!user.stripeCustomerId) {
    res.status(400).json({ error: "No Stripe customer on file" });
    return;
  }

  try {
    const session = await stripe.billingPortal.sessions.create({
      customer: user.stripeCustomerId,
      return_url: frontendUrl,
    });
    res.json({ url: session.url });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[stripe/portal]", message);
    res.status(500).json({ error: message });
  }
});

export async function stripeWebhookHandler(req: Request, res: Response): Promise<void> {
  const stripe = getStripe();
  if (!stripe) {
    res.status(503).json({ error: "Stripe is not configured" });
    return;
  }

  const webhookSecret = process.env["STRIPE_WEBHOOK_SECRET"];
  const signature = req.headers["stripe-signature"];

  if (!signature) {
    res.status(400).json({ error: "Missing stripe-signature header" });
    return;
  }

  if (!webhookSecret) {
    console.error("[stripe/webhook] STRIPE_WEBHOOK_SECRET is not configured — rejecting request");
    res.status(503).json({ error: "Webhook secret not configured on server" });
    return;
  }

  let event: Stripe.Event;

  try {
    const sig = Array.isArray(signature) ? signature[0] : signature;
    event = stripe.webhooks.constructEvent(req.body as Buffer, sig, webhookSecret);
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[stripe/webhook] signature verification failed:", message);
    res.status(400).json({ error: "Webhook signature verification failed" });
    return;
  }

  try {
    switch (event.type) {
      case "checkout.session.completed": {
        const session = event.data.object as Stripe.Checkout.Session;
        if (session.mode === "subscription" && session.customer && session.subscription) {
          const customerId = typeof session.customer === "string" ? session.customer : session.customer.id;
          const subscriptionId = typeof session.subscription === "string" ? session.subscription : session.subscription.id;
          await db.update(usersTable)
            .set({ plan: "pro", stripeSubscriptionId: subscriptionId, updatedAt: new Date() })
            .where(eq(usersTable.stripeCustomerId, customerId));
        }
        break;
      }
      case "customer.subscription.deleted": {
        const sub = event.data.object as Stripe.Subscription;
        const customerId = typeof sub.customer === "string" ? sub.customer : sub.customer.id;
        await db.update(usersTable)
          .set({ plan: "free", stripeSubscriptionId: null, updatedAt: new Date() })
          .where(eq(usersTable.stripeCustomerId, customerId));
        break;
      }
      case "customer.subscription.updated": {
        const sub = event.data.object as Stripe.Subscription;
        const customerId = typeof sub.customer === "string" ? sub.customer : sub.customer.id;
        const isActive = sub.status === "active" || sub.status === "trialing";
        await db.update(usersTable)
          .set({
            plan: isActive ? "pro" : "free",
            stripeSubscriptionId: isActive ? sub.id : null,
            updatedAt: new Date(),
          })
          .where(eq(usersTable.stripeCustomerId, customerId));
        break;
      }
      default:
        break;
    }
  } catch (err: unknown) {
    console.error("[stripe/webhook] handler error:", err);
    res.status(500).json({ error: "Internal error processing webhook" });
    return;
  }

  res.json({ received: true });
}

export default router;
