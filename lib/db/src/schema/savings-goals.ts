import { pgTable, text, timestamp, uuid, numeric, boolean } from "drizzle-orm/pg-core";
import { usersTable } from "./users";
import { savedBudgetsTable } from "./saved-budgets";

export const savingsGoalsTable = pgTable("savings_goals", {
  id: uuid("id").defaultRandom().primaryKey(),
  userId: uuid("user_id").notNull().references(() => usersTable.id, { onDelete: "cascade" }),
  budgetId: uuid("budget_id").notNull().references(() => savedBudgetsTable.id, { onDelete: "cascade" }),
  name: text("name").notNull(),
  targetAmount: numeric("target_amount", { precision: 12, scale: 2 }).notNull(),
  targetDate: text("target_date").notNull(),
  note: text("note"),
  includeInBudget: boolean("include_in_budget").notNull().default(false),
  createdAt: timestamp("created_at").defaultNow().notNull(),
  updatedAt: timestamp("updated_at").defaultNow().notNull(),
});

export type SavingsGoal = typeof savingsGoalsTable.$inferSelect;
export type InsertSavingsGoal = {
  userId: string;
  budgetId: string;
  name: string;
  targetAmount: string;
  targetDate: string;
  note?: string | null;
  includeInBudget?: boolean;
};
