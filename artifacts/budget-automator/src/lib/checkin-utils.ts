const API_BASE = ((import.meta.env.VITE_API_BASE_URL as string | undefined) ?? "").replace(/\/+$/, "");

export async function apiFetch<T>(path: string, init?: RequestInit): Promise<T> {
  const headers = new Headers(init?.headers);
  const token = typeof localStorage !== "undefined" ? localStorage.getItem("auth_token") : null;
  if (token && !headers.has("authorization")) headers.set("authorization", `Bearer ${token}`);
  const res = await fetch(`${API_BASE}${path}`, { credentials: "include", ...init, headers });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error((err as any).error ?? `HTTP ${res.status}`);
  }
  return res.json();
}

function checkinDismissKey(budgetId: string, weekLabel: string) {
  return `checkin_dismissed_${budgetId}_${weekLabel}`;
}

export function isDismissed(budgetId: string, weekLabel: string): boolean {
  try {
    return sessionStorage.getItem(checkinDismissKey(budgetId, weekLabel)) === "1";
  } catch {
    return false;
  }
}

export function setDismissed(budgetId: string, weekLabel: string) {
  try {
    sessionStorage.setItem(checkinDismissKey(budgetId, weekLabel), "1");
  } catch {}
}

function paydayDismissKey(budgetId: string, weekLabel: string) {
  return `payday_dismissed_${budgetId}_${weekLabel}`;
}

export function isPaydayDismissed(budgetId: string, weekLabel: string): boolean {
  try {
    return localStorage.getItem(paydayDismissKey(budgetId, weekLabel)) === "1";
  } catch {
    return false;
  }
}

export function setPaydayDismissed(budgetId: string, weekLabel: string) {
  try {
    localStorage.setItem(paydayDismissKey(budgetId, weekLabel), "1");
  } catch {}
}
