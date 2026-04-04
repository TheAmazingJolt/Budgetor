import { pgTable, text, timestamp, uuid } from "drizzle-orm/pg-core";
import { usersTable } from "./users";

export const bugReportsTable = pgTable("bug_reports", {
  id: uuid("id").defaultRandom().primaryKey(),
  userId: uuid("user_id").references(() => usersTable.id, { onDelete: "set null" }),
  description: text("description").notNull(),
  errorMessage: text("error_message"),
  errorStack: text("error_stack"),
  pageUrl: text("page_url").notNull(),
  userAgent: text("user_agent").notNull(),
  appVersion: text("app_version").notNull(),
  createdAt: timestamp("created_at").defaultNow().notNull(),
});

export type BugReport = typeof bugReportsTable.$inferSelect;
export type InsertBugReport = {
  userId?: string | null;
  description: string;
  errorMessage?: string | null;
  errorStack?: string | null;
  pageUrl: string;
  userAgent: string;
  appVersion: string;
};
