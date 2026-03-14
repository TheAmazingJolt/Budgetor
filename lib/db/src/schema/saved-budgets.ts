import { pgTable, text, timestamp, uuid, jsonb } from "drizzle-orm/pg-core";
import { createInsertSchema } from "drizzle-zod";
import { usersTable } from "./users";

export const savedBudgetsTable = pgTable("saved_budgets", {
  id: uuid("id").defaultRandom().primaryKey(),
  userId: uuid("user_id").notNull().references(() => usersTable.id, { onDelete: "cascade" }),
  name: text("name").notNull(),
  bills: jsonb("bills").notNull(),
  settings: jsonb("settings").notNull(),
  createdAt: timestamp("created_at").defaultNow().notNull(),
  updatedAt: timestamp("updated_at").defaultNow().notNull(),
});

export const insertSavedBudgetSchema = createInsertSchema(savedBudgetsTable).omit({ id: true, createdAt: true, updatedAt: true });

export type InsertSavedBudget = {
  userId: string;
  name: string;
  bills: unknown;
  settings: unknown;
};

export type SavedBudget = typeof savedBudgetsTable.$inferSelect;
