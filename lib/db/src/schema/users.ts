import { pgTable, text, timestamp, uuid, uniqueIndex } from "drizzle-orm/pg-core";
import { sql } from "drizzle-orm";

export const usersTable = pgTable("users", {
  id: uuid("id").defaultRandom().primaryKey(),
  email: text("email"),
  name: text("name").notNull(),
  avatarUrl: text("avatar_url"),
  provider: text("provider").notNull().$type<"google" | "apple" | "guest">(),
  providerId: text("provider_id"),
  createdAt: timestamp("created_at").defaultNow().notNull(),
  updatedAt: timestamp("updated_at").defaultNow().notNull(),
}, (table) => [
  uniqueIndex("users_provider_provider_id_unique")
    .on(table.provider, table.providerId)
    .where(sql`provider_id IS NOT NULL`),
]);

export type InsertUser = {
  name: string;
  email?: string | null;
  avatarUrl?: string | null;
  provider: string;
  providerId?: string | null;
};

export type User = typeof usersTable.$inferSelect;
