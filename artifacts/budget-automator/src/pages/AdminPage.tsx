import { useState } from "react";
import { DollarSign } from "lucide-react";

const apiBase = ((import.meta.env.VITE_API_BASE_URL as string | undefined) ?? "").replace(/\/+$/, "");

export function AdminPage() {
  const [secret, setSecret] = useState("");
  const [email, setEmail] = useState("");
  const [plan, setPlan] = useState<"pro" | "free">("pro");
  const [status, setStatus] = useState<{ ok: boolean; message: string } | null>(null);
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setStatus(null);
    try {
      const res = await fetch(`${apiBase}/api/admin/set-plan`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-admin-secret": secret,
        },
        body: JSON.stringify({ email, plan }),
      });
      const data = await res.json() as { ok?: boolean; error?: string; user?: { email: string; plan: string } };
      if (!res.ok) {
        setStatus({ ok: false, message: data.error ?? `Error ${res.status}` });
      } else {
        setStatus({ ok: true, message: `${data.user?.email ?? email} is now on the ${data.user?.plan ?? plan} plan.` });
      }
    } catch (err) {
      setStatus({ ok: false, message: err instanceof Error ? err.message : "Network error" });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex flex-col items-center justify-center bg-gradient-to-br from-slate-50 via-white to-emerald-50 p-6">
      <div className="w-14 h-14 rounded-2xl bg-gradient-to-br from-primary to-emerald-600 flex items-center justify-center mb-6 shadow-lg shadow-primary/20">
        <DollarSign className="w-7 h-7 text-white" />
      </div>
      <h1 className="text-xl font-bold text-foreground mb-1">Admin</h1>
      <p className="text-sm text-muted-foreground mb-6">Set a user's plan directly</p>

      <form onSubmit={handleSubmit} className="w-full max-w-sm flex flex-col gap-3">
        <input
          type="password"
          placeholder="Admin secret"
          value={secret}
          onChange={(e) => setSecret(e.target.value)}
          required
          className="w-full px-4 py-2.5 rounded-xl border border-slate-200 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500"
        />
        <input
          type="email"
          placeholder="User email"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          required
          className="w-full px-4 py-2.5 rounded-xl border border-slate-200 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500"
        />
        <select
          value={plan}
          onChange={(e) => setPlan(e.target.value as "pro" | "free")}
          className="w-full px-4 py-2.5 rounded-xl border border-slate-200 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500"
        >
          <option value="pro">Pro</option>
          <option value="free">Free</option>
        </select>

        <button
          type="submit"
          disabled={loading}
          className="w-full px-4 py-2.5 rounded-xl bg-emerald-600 text-white text-sm font-medium hover:bg-emerald-700 transition-colors disabled:opacity-50"
        >
          {loading ? "Saving…" : "Set plan"}
        </button>
      </form>

      {status && (
        <p className={`mt-4 text-sm font-medium ${status.ok ? "text-emerald-600" : "text-red-500"}`}>
          {status.message}
        </p>
      )}
    </div>
  );
}
