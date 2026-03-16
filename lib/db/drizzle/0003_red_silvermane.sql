ALTER TABLE "users" ADD COLUMN "debts" jsonb DEFAULT '[]'::jsonb;--> statement-breakpoint
ALTER TABLE "users" ADD COLUMN "preferences" jsonb DEFAULT '{}'::jsonb;