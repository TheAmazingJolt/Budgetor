import { pgTable, text, timestamp, uuid, jsonb } from "drizzle-orm/pg-core";
import { usersTable } from "./users";

export const savedBudgetsTable = pgTable("saved_budgets", {
  id: uuid("id").defaultRandom().primaryKey(),
  userId: uuid("user_id").notNull().references(() => usersTable.id, { onDelete: "cascade" }),
  name: text("name").notNull(),
  bills: jsonb("bills").notNull(),
  settings: jsonb("settings").notNull(),
  debts: jsonb("debts").default([]),
  createdAt: timestamp("created_at").defaultNow().notNull(),
  updatedAt: timestamp("updated_at").defaultNow().notNull(),
});

export type InsertSavedBudget = {
  userId: string;
  name: string;
  bills: unknown;
  settings: unknown;
  debts?: unknown;
};

export type SavedBudget = typeof savedBudgetsTable.$inferSelect;
