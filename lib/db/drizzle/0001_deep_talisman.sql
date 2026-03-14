ALTER TABLE "users" ADD COLUMN "google_access_token" text;--> statement-breakpoint
ALTER TABLE "users" ADD COLUMN "google_refresh_token" text;--> statement-breakpoint
ALTER TABLE "users" ADD COLUMN "google_token_expiry" bigint;--> statement-breakpoint
ALTER TABLE "users" ADD COLUMN "microsoft_access_token" text;--> statement-breakpoint
ALTER TABLE "users" ADD COLUMN "microsoft_refresh_token" text;--> statement-breakpoint
ALTER TABLE "users" ADD COLUMN "microsoft_token_expiry" bigint;