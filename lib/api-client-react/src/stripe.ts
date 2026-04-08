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
