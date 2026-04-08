#!/usr/bin/env node
/**
 * Stripe seed script — creates the $5/month Pro plan price if it does not already exist.
 *
 * Usage:
 *   STRIPE_SECRET_KEY=sk_... node scripts/stripe-seed.mjs
 *
 * On success it prints the price ID. Set that value as STRIPE_PRO_PRICE_ID in Railway.
 */
import Stripe from "stripe";

const key = process.env.STRIPE_SECRET_KEY;
if (!key) {
  console.error("Error: STRIPE_SECRET_KEY environment variable is not set.");
  process.exit(1);
}

const stripe = new Stripe(key, { apiVersion: "2025-01-27.acacia" });

async function main() {
  const existing = await stripe.prices.search({
    query: 'product_data.name:"Budgify Pro" AND active:"true" AND currency:"usd"',
    limit: 10,
  });

  const match = existing.data.find(
    (p) =>
      p.recurring?.interval === "month" &&
      p.unit_amount === 500
  );

  if (match) {
    console.log("Price already exists:", match.id);
    console.log("Set STRIPE_PRO_PRICE_ID=" + match.id);
    return;
  }

  const product = await stripe.products.create({
    name: "Budgify Pro",
    description: "Unlimited spending categories, unlimited savings goals, Google Sheets sync, and more.",
  });

  const price = await stripe.prices.create({
    product: product.id,
    unit_amount: 500,
    currency: "usd",
    recurring: { interval: "month" },
    nickname: "Budgify Pro Monthly",
  });

  console.log("Created price:", price.id);
  console.log("Set STRIPE_PRO_PRICE_ID=" + price.id);
}

main().catch((err) => {
  console.error(err.message ?? err);
  process.exit(1);
});
