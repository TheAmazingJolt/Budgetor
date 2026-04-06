import { pgTable, text, timestamp, uuid, uniqueIndex, bigint, jsonb, index } from "drizzle-orm/pg-core";
import { sql } from "drizzle-orm";

export const referralRewardsTable = pgTable("referral_rewards", {
  id: uuid("id").defaultRandom().primaryKey(),
  referrerId: uuid("referrer_id").notNull().references(() => usersTable.id, { onDelete: "cascade" }),
  referredUserId: uuid("referred_user_id").notNull().unique().references(() => usersTable.id, { onDelete: "cascade" }),
  status: text("status").notNull().$type<"pending" | "applied">().default("pending"),
  appliedAt: timestamp("applied_at"),
  createdAt: timestamp("created_at").defaultNow().notNull(),
  updatedAt: timestamp("updated_at").defaultNow().notNull(),
}, (table) => [
  index("referral_rewards_referrer_idx").on(table.referrerId),
  index("referral_rewards_status_idx").on(table.status),
]);

export const usersTable = pgTable("users", {
  id: uuid("id").defaultRandom().primaryKey(),
  email: text("email"),
  name: text("name").notNull(),
  avatarUrl: text("avatar_url"),
  provider: text("provider").notNull().$type<"google" | "apple" | "guest" | "microsoft" | "email">(),
  providerId: text("provider_id"),
  passwordHash: text("password_hash"),
  passwordResetToken: text("password_reset_token"),
  passwordResetExpires: bigint("password_reset_expires", { mode: "number" }),
  googleAccessToken: text("google_access_token"),
  googleRefreshToken: text("google_refresh_token"),
  googleTokenExpiry: bigint("google_token_expiry", { mode: "number" }),
  microsoftAccessToken: text("microsoft_access_token"),
  microsoftRefreshToken: text("microsoft_refresh_token"),
  microsoftTokenExpiry: bigint("microsoft_token_expiry", { mode: "number" }),
  debts: jsonb("debts").default([]),
  bills: jsonb("bills").default([]),
  preferences: jsonb("preferences").default({}),
  referralCode: text("referral_code"),
  referredBy: text("referred_by"),
  stripeCustomerId: text("stripe_customer_id"),
  stripeSubscriptionId: text("stripe_subscription_id"),
  proExpiresAt: timestamp("pro_expires_at"),
  referralCreditsOwed: text("referral_credits_owed").default("0"),
  referralRewardGrantedAt: timestamp("referral_reward_granted_at"),
  createdAt: timestamp("created_at").defaultNow().notNull(),
  updatedAt: timestamp("updated_at").defaultNow().notNull(),
}, (table) => [
  uniqueIndex("users_provider_provider_id_unique")
    .on(table.provider, table.providerId)
    .where(sql`provider_id IS NOT NULL`),
  uniqueIndex("users_referral_code_unique").on(table.referralCode).where(sql`referral_code IS NOT NULL`),
]);

export type InsertUser = {
  name: string;
  email?: string | null;
  avatarUrl?: string | null;
  provider: string;
  providerId?: string | null;
};

export type User = typeof usersTable.$inferSelect;
