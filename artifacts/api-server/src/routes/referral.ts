import { Router, type IRouter, type Request, type Response } from "express";
import crypto from "crypto";
import Stripe from "stripe";
import { db, usersTable, referralRewardsTable } from "@workspace/db";
import { eq, and } from "drizzle-orm";
import { requireAuth } from "./user-auth";

const router: IRouter = Router();

function generateReferralCode(): string {
  return crypto.randomBytes(5).toString("hex").toUpperCase();
}

async function ensureReferralCode(userId: string): Promise<string> {
  const [user] = await db
    .select({ referralCode: usersTable.referralCode })
    .from(usersTable)
    .where(eq(usersTable.id, userId))
    .limit(1);

  if (user?.referralCode) return user.referralCode;

  let code: string = "";
  let attempts = 0;
  while (true) {
    code = generateReferralCode();
    try {
      await db
        .update(usersTable)
        .set({ referralCode: code, updatedAt: new Date() })
        .where(eq(usersTable.id, userId));
      break;
    } catch {
      attempts++;
      if (attempts >= 10) throw new Error("Failed to generate unique referral code");
    }
  }
  return code;
}

router.get("/referral/info", requireAuth, async (req: Request, res: Response): Promise<void> => {
  try {
    const userId = (req.user as import("@workspace/db").User).id;
    const code = await ensureReferralCode(userId);

    const appUrl = process.env["CORS_ORIGIN"]?.split(",")[0]?.trim() ?? "";
    const referralLink = `${appUrl}?ref=${code}`;

    const rewards = await db
      .select({
        id: referralRewardsTable.id,
        status: referralRewardsTable.status,
      })
      .from(referralRewardsTable)
      .where(eq(referralRewardsTable.referrerId, userId));

    const totalReferred = await db
      .select({ id: usersTable.id })
      .from(usersTable)
      .where(eq(usersTable.referredBy, code))
      .then((rows) => rows.length);

    const totalConverted = rewards.length;
    const totalRewarded = rewards.filter((r) => r.status === "applied").length;

    res.json({
      referralCode: code,
      referralLink,
      totalReferred,
      totalConverted,
      totalRewarded,
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: "Failed to get referral info: " + message });
  }
});

router.post(
  "/referral/stripe-webhook",
  async (req: Request, res: Response): Promise<void> => {
    const webhookSecret = process.env["STRIPE_WEBHOOK_SECRET"];
    const stripeSecretKey = process.env["STRIPE_SECRET_KEY"];
    const stripeSignature = req.headers["stripe-signature"] as string | undefined;
    const isProduction = process.env["NODE_ENV"] === "production";

    if (!stripeSecretKey) {
      res.status(400).json({ error: "Stripe not configured" });
      return;
    }

    const stripeClient = new Stripe(stripeSecretKey);

    let event: Stripe.Event;

    if (webhookSecret || isProduction) {
      if (!webhookSecret) {
        res.status(500).json({ error: "STRIPE_WEBHOOK_SECRET must be set in production" });
        return;
      }
      if (!stripeSignature) {
        res.status(400).json({ error: "Missing Stripe webhook signature" });
        return;
      }
      try {
        event = stripeClient.webhooks.constructEvent(
          req.body as Buffer,
          stripeSignature,
          webhookSecret,
        );
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err);
        res.status(400).json({ error: "Webhook signature verification failed: " + message });
        return;
      }
    } else {
      try {
        event = JSON.parse((req.body as Buffer).toString()) as Stripe.Event;
      } catch {
        res.status(400).json({ error: "Invalid webhook payload" });
        return;
      }
    }

    try {
      await handleStripeEvent(stripeClient, event);
    } catch (err) {
      console.error("[referral] Webhook event processing failed:", err);
      res.status(500).json({ error: "Webhook processing failed, retry will be attempted" });
      return;
    }

    res.json({ received: true });
  },
);

async function resolveUserByStripeCustomer(
  stripeClient: Stripe,
  customerId: string,
): Promise<(typeof usersTable.$inferSelect) | null> {
  const [existing] = await db
    .select()
    .from(usersTable)
    .where(eq(usersTable.stripeCustomerId, customerId))
    .limit(1);
  if (existing) return existing;

  const customer = await stripeClient.customers.retrieve(customerId);
  if (customer.deleted) return null;

  const appUserId = customer.metadata?.["app_user_id"];
  if (appUserId) {
    const [byMeta] = await db
      .select()
      .from(usersTable)
      .where(eq(usersTable.id, appUserId))
      .limit(1);
    if (byMeta) {
      await db
        .update(usersTable)
        .set({ stripeCustomerId: customerId, updatedAt: new Date() })
        .where(eq(usersTable.id, byMeta.id));
      return { ...byMeta, stripeCustomerId: customerId };
    }
  }

  const email = customer.email;
  if (email) {
    const [byEmail] = await db
      .select()
      .from(usersTable)
      .where(eq(usersTable.email, email))
      .limit(1);
    if (byEmail) {
      await db
        .update(usersTable)
        .set({ stripeCustomerId: customerId, updatedAt: new Date() })
        .where(eq(usersTable.id, byEmail.id));
      return { ...byEmail, stripeCustomerId: customerId };
    }
  }

  return null;
}

async function applyStripeBalanceCredit(
  stripeClient: Stripe,
  stripeCustomerId: string,
  stripeSubscriptionId: string,
  count: number,
  description: string,
  idempotencyKey: string,
): Promise<void> {
  const referrerSub = await stripeClient.subscriptions.retrieve(stripeSubscriptionId);
  const priceId = referrerSub.items.data[0]?.price.id;
  if (!priceId) throw new Error("Referrer subscription has no price");

  const price = await stripeClient.prices.retrieve(priceId);
  const unitAmount = (price.unit_amount ?? 500) * count;
  await stripeClient.customers.createBalanceTransaction(
    stripeCustomerId,
    {
      amount: -unitAmount,
      currency: price.currency,
      description,
    },
    { idempotencyKey },
  );
  console.log(
    `[referral] Applied ${count} month(s) credit ($${unitAmount / 100} ${price.currency}): ${description}`,
  );
}

async function processReward(
  stripeClient: Stripe,
  rewardId: string,
): Promise<void> {
  const [reward] = await db
    .select({
      id: referralRewardsTable.id,
      referrerId: referralRewardsTable.referrerId,
      status: referralRewardsTable.status,
    })
    .from(referralRewardsTable)
    .where(eq(referralRewardsTable.id, rewardId))
    .limit(1);

  if (!reward || reward.status === "applied") return;

  const [referrer] = await db
    .select({
      id: usersTable.id,
      stripeCustomerId: usersTable.stripeCustomerId,
      stripeSubscriptionId: usersTable.stripeSubscriptionId,
    })
    .from(usersTable)
    .where(eq(usersTable.id, reward.referrerId))
    .limit(1);

  if (!referrer?.stripeCustomerId || !referrer?.stripeSubscriptionId) {
    console.log(`[referral] Reward ${rewardId}: referrer not yet on Pro; stays pending`);
    return;
  }

  await applyStripeBalanceCredit(
    stripeClient,
    referrer.stripeCustomerId,
    referrer.stripeSubscriptionId,
    1,
    "Referral reward: one free month of Pro",
    `referral-reward-${rewardId}`,
  );

  await db
    .update(referralRewardsTable)
    .set({ status: "applied", appliedAt: new Date(), updatedAt: new Date() })
    .where(eq(referralRewardsTable.id, rewardId));
}

async function redeemPendingRewards(
  stripeClient: Stripe,
  referrerId: string,
): Promise<void> {
  const pendingRewards = await db
    .select({ id: referralRewardsTable.id })
    .from(referralRewardsTable)
    .where(
      and(
        eq(referralRewardsTable.referrerId, referrerId),
        eq(referralRewardsTable.status, "pending"),
      ),
    );

  for (const reward of pendingRewards) {
    try {
      await processReward(stripeClient, reward.id);
    } catch (err) {
      console.error(`[referral] Failed to process pending reward ${reward.id}:`, err);
      throw err;
    }
  }
}

async function attributeReferral(
  referrerId: string,
  referredUserId: string,
): Promise<string | null> {
  try {
    const [row] = await db
      .insert(referralRewardsTable)
      .values({
        referrerId,
        referredUserId,
        status: "pending",
      })
      .onConflictDoNothing({ target: referralRewardsTable.referredUserId })
      .returning({ id: referralRewardsTable.id });
    return row?.id ?? null;
  } catch (err) {
    console.error("[referral] Failed to insert referral_rewards row:", err);
    return null;
  }
}

async function handleStripeEvent(stripeClient: Stripe, event: Stripe.Event): Promise<void> {
  if (event.type === "checkout.session.completed") {
    const session = event.data.object as Stripe.Checkout.Session;
    const customerId =
      typeof session.customer === "string" ? session.customer : session.customer?.id;
    const appUserId = session.metadata?.["app_user_id"] ?? session.client_reference_id;

    if (customerId && appUserId) {
      await db
        .update(usersTable)
        .set({ stripeCustomerId: customerId, updatedAt: new Date() })
        .where(eq(usersTable.id, appUserId));
    }
    return;
  }

  if (
    event.type === "customer.subscription.created" ||
    event.type === "customer.subscription.updated"
  ) {
    const subscription = event.data.object as Stripe.Subscription;
    const customerId =
      typeof subscription.customer === "string"
        ? subscription.customer
        : subscription.customer.id;
    const subscriptionId = subscription.id;
    const status = subscription.status;

    if (status !== "active") return;

    const user = await resolveUserByStripeCustomer(stripeClient, customerId);
    if (!user) {
      console.warn(`[referral] No user found for Stripe customer ${customerId}`);
      return;
    }

    const isNewActivation = !user.stripeSubscriptionId;

    if (isNewActivation && user.referredBy) {
      const [referrer] = await db
        .select({ id: usersTable.id })
        .from(usersTable)
        .where(eq(usersTable.referralCode, user.referredBy))
        .limit(1);

      if (referrer && referrer.id !== user.id) {
        const rewardId = await attributeReferral(referrer.id, user.id);
        if (rewardId) {
          await db
            .update(usersTable)
            .set({ stripeCustomerId: customerId, stripeSubscriptionId: subscriptionId, updatedAt: new Date() })
            .where(eq(usersTable.id, user.id));

          await processReward(stripeClient, rewardId);
          return;
        }
      }
    }

    await db
      .update(usersTable)
      .set({ stripeCustomerId: customerId, stripeSubscriptionId: subscriptionId, updatedAt: new Date() })
      .where(eq(usersTable.id, user.id));

    if (isNewActivation) {
      await redeemPendingRewards(stripeClient, user.id);
    }
  }
}

export default router;
