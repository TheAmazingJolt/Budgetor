import { customFetch } from "./custom-fetch";

export const stripeCheckout = async (options?: RequestInit): Promise<{ url: string }> => {
  return customFetch<{ url: string }>("/api/stripe/checkout", {
    ...options,
    method: "POST",
  });
};

export const stripePortal = async (options?: RequestInit): Promise<{ url: string }> => {
  return customFetch<{ url: string }>("/api/stripe/portal", {
    ...options,
    method: "POST",
  });
};

export const stripeVerifySession = async (body: { sessionId: string }, options?: RequestInit): Promise<{ upgraded: boolean }> => {
  return customFetch<{ upgraded: boolean }>("/api/stripe/verify-session", {
    ...options,
    method: "POST",
    headers: { "Content-Type": "application/json", ...(options?.headers ?? {}) },
    body: JSON.stringify(body),
  });
};
