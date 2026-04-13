import { pgTable, text, timestamp, uuid, numeric, boolean } from "drizzle-orm/pg-core";
import { usersTable } from "./users";
import { savedBudgetsTable } from "./saved-budgets";

export const savingsContributionsTable = pgTable("savings_contributions", {
  id: uuid("id").defaultRandom().primaryKey(),
  userId: uuid("user_id").notNull().references(() => usersTable.id, { onDelete: "cascade" }),
  budgetId: uuid("budget_id").notNull().references(() => savedBudgetsTable.id, { onDelete: "cascade" }),
  billName: text("bill_name").notNull(),
  amount: numeric("amount", { precision: 12, scale: 2 }).notNull(),
  date: text("date").notNull(),
  note: text("note"),
  isExtra: boolean("is_extra").default(false).notNull(),
  createdAt: timestamp("created_at").defaultNow().notNull(),
});

export type SavingsContribution = typeof savingsContributionsTable.$inferSelect;
export type InsertSavingsContribution = {
  userId: string;
  budgetId: string;
  billName: string;
  amount: string;
  date: string;
  note?: string | null;
  isExtra?: boolean;
};
