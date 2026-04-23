import { useEffect, useState, Component } from "react";
import type { ReactNode } from "react";
import { QueryClient, QueryClientProvider, useQueryClient } from "@tanstack/react-query";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import { useToast } from "@/hooks/use-toast";
import { BudgetWizard } from "@/pages/BudgetWizard";
import { SignInPage } from "@/pages/SignInPage";
import { AdminPage } from "@/pages/AdminPage";
import { PrivacyPage } from "@/pages/PrivacyPage";
import { TermsPage } from "@/pages/TermsPage";
import { BugReportDialog } from "@/components/BugReportDialog";
import type { AuthUser } from "@workspace/api-client-react";
import {
  useAuthMe,
  useAuthProviders,
  useAuthGuestLogin,
  authLoginGoogle,
  getAuthMeQueryKey,
  getAuthProvidersQueryKey,
  getMicrosoftAuthStatusQueryKey,
  authEmailLogin,
  authEmailRegister,
  authForgotPassword,
  authResetPassword,
  authClaimAccount,
  stripeCheckout,
  stripeVerifySession,
} from "@workspace/api-client-react";
import { DollarSign, Loader2, RefreshCw, Bug } from "lucide-react";

interface ErrorBoundaryState {
  error: Error | null;
  bugReportOpen: boolean;
}

class AppErrorBoundary extends Component<{ children: ReactNode }, ErrorBoundaryState> {
  constructor(props: { children: ReactNode }) {
    super(props);
    this.state = { error: null, bugReportOpen: false };
  }

  static getDerivedStateFromError(error: Error) {
    return { error };
  }

  handleReset = async () => {
    try {
      const keysToRemove: string[] = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key !== "auth_token") keysToRemove.push(key);
      }
      keysToRemove.forEach(k => localStorage.removeItem(k));
    } catch { }
    try {
      if ("caches" in window) {
        const keys = await caches.keys();
        await Promise.all(keys.map(k => caches.delete(k)));
      }
    } catch { }
    const url = new URL(window.location.href);
    url.searchParams.set("_r", Date.now().toString());
    window.location.replace(url.toString());
  };

  render() {
    if (this.state.error) {
      const error = this.state.error;
      return (
        <QueryClientProvider client={new QueryClient()}>
          <div className="min-h-screen flex flex-col items-center justify-center bg-gradient-to-br from-slate-50 via-white to-emerald-50 p-6">
            <div className="w-14 h-14 rounded-2xl bg-gradient-to-br from-primary to-emerald-600 flex items-center justify-center mb-6 shadow-lg shadow-primary/20">
              <DollarSign className="w-7 h-7 text-white" />
            </div>
            <h1 className="text-xl font-bold text-foreground mb-2">Something went wrong</h1>
            <p className="text-muted-foreground text-sm text-center mb-6 max-w-xs">
              Cached app data may be out of date. Clearing it and reloading usually fixes this.
            </p>
            <div className="flex flex-col sm:flex-row items-center gap-3">
              <button
                onClick={this.handleReset}
                className="flex items-center gap-2 px-5 py-2.5 rounded-xl bg-emerald-600 text-white text-sm font-medium hover:bg-emerald-700 transition-colors"
              >
                <RefreshCw className="w-4 h-4" /> Clear cache &amp; reload
              </button>
              <button
                onClick={() => this.setState({ bugReportOpen: true })}
                className="flex items-center gap-2 px-5 py-2.5 rounded-xl border border-slate-200 bg-white text-slate-700 text-sm font-medium hover:bg-slate-50 transition-colors"
              >
                <Bug className="w-4 h-4 text-amber-500" /> Report this issue
              </button>
            </div>
            <details className="mt-6 max-w-sm w-full">
              <summary className="text-xs text-muted-foreground cursor-pointer">Error details</summary>
              <pre className="mt-2 text-xs text-red-600 bg-red-50 rounded-lg p-3 overflow-auto whitespace-pre-wrap break-all">
                {error.message}
              </pre>
            </details>
          </div>

          <BugReportDialog
            open={this.state.bugReportOpen}
            onOpenChange={(open) => this.setState({ bugReportOpen: open })}
            prefillDescription={`I encountered an error: ${error.message}`}
            prefillErrorMessage={error.message}
            prefillErrorStack={error.stack ?? null}
          />
          <Toaster />
        </QueryClientProvider>
      );
    }
    return this.props.children;
  }
}

const queryClient = new QueryClient();

function SplashScreen() {
  return (
    <div className="min-h-screen flex flex-col items-center justify-center bg-gradient-to-br from-slate-50 via-white to-emerald-50">
      <div className="w-14 h-14 rounded-2xl bg-gradient-to-br from-primary to-emerald-600 flex items-center justify-center mb-4 shadow-lg shadow-primary/20">
        <DollarSign className="w-7 h-7 text-white" />
      </div>
      <Loader2 className="w-5 h-5 animate-spin text-muted-foreground" />
    </div>
  );
}

function AppRouting() {
  const qc = useQueryClient();
  const { toast } = useToast();

  const authQuery = useAuthMe({ query: { queryKey: getAuthMeQueryKey(), retry: false, staleTime: 30000 } });
  const providersQuery = useAuthProviders({ query: { queryKey: getAuthProvidersQueryKey(), retry: false, staleTime: 60000 } });
  const guestLoginMutation = useAuthGuestLogin();

  const currentUser: AuthUser | null = authQuery.data?.user ?? null;
  const isSignedIn = !!currentUser;
  const isGuest = currentUser?.provider === "guest";
  const googleLoginAvailable = providersQuery.data?.google ?? false;
  const appleLoginAvailable = providersQuery.data?.apple ?? false;

  const initialRefCode = new URLSearchParams(window.location.search).get("ref");
  const [referralReady, setReferralReady] = useState<boolean>(!initialRefCode);

  const resetToken = new URLSearchParams(window.location.search).get("reset_token")
    ?? new URLSearchParams(window.location.hash.slice(1)).get("reset_token");
  const claimEmail = new URLSearchParams(window.location.search).get("claim_email")
    ?? new URLSearchParams(window.location.hash.slice(1)).get("claim_email");
  const claimToken = new URLSearchParams(window.location.search).get("claim_token")
    ?? new URLSearchParams(window.location.hash.slice(1)).get("claim_token");

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const refCode = params.get("ref");
    if (!refCode) {
      setReferralReady(true);
      return;
    }
    params.delete("ref");
    const newSearch = params.toString();
    const newUrl = window.location.pathname + (newSearch ? "?" + newSearch : "") + window.location.hash;
    window.history.replaceState({}, "", newUrl);

    const apiBase = ((import.meta.env.VITE_API_BASE_URL as string | undefined) ?? "").replace(/\/+$/, "");
    fetch(`${apiBase}/api/auth/referral-code`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      credentials: "include",
      body: JSON.stringify({ code: refCode }),
    })
      .catch(() => {})
      .finally(() => setReferralReady(true));
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const authCode = params.get("auth_code");
    if (!authCode) return;
    params.delete("auth_code");
    const newSearch = params.toString();
    const newUrl = window.location.pathname + (newSearch ? "?" + newSearch : "") + window.location.hash;
    window.history.replaceState({}, "", newUrl);
    const apiBase = ((import.meta.env.VITE_API_BASE_URL as string | undefined) ?? "").replace(/\/+$/, "");
    fetch(`${apiBase}/api/auth/exchange`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      credentials: "include",
      body: JSON.stringify({ code: authCode }),
    })
      .then((r) => {
        if (!r.ok) throw new Error(`Exchange failed: ${r.status}`);
        return r.json();
      })
      .then((data: { user?: AuthUser; token?: string }) => {
        if (data.token) {
          localStorage.setItem("auth_token", data.token);
        }
        if (data.user) {
          qc.setQueryData(getAuthMeQueryKey(), { user: data.user });
          qc.invalidateQueries({ queryKey: ["/api/auth/google/status"] });
          qc.invalidateQueries({ queryKey: getMicrosoftAuthStatusQueryKey() });
          const intent = localStorage.getItem("upgrade_intent");
          if (intent === "pro") {
            localStorage.removeItem("upgrade_intent");
            stripeCheckout()
              .then(({ url }) => { window.location.href = url; })
              .catch(() => {});
          }
        }
      })
      .catch(() => {
        const failParams = new URLSearchParams(window.location.search);
        failParams.delete("auth_code");
        const failSearch = failParams.toString();
        const failUrl = window.location.pathname + (failSearch ? "?" + failSearch : "") + window.location.hash;
        window.history.replaceState({}, "", failUrl);
        toast({
          title: "Sign-in failed — please try again.",
          variant: "destructive",
        });
      });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const errorParam = params.get("error");
    if (errorParam === "link_only" || errorParam === "link_required") {
      params.delete("error");
      const newSearch = params.toString();
      const newUrl = window.location.pathname + (newSearch ? "?" + newSearch : "") + window.location.hash;
      window.history.replaceState({}, "", newUrl);
      toast({
        title: "Sign in required",
        description: "Please sign in with your email and password first, then connect your account in Settings.",
        variant: "destructive",
      });
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleGoogleLogin = async () => {
    try {
      const currentUrl = window.location.href;
      const result = await authLoginGoogle({ redirect: currentUrl });
      window.location.href = result.url;
    } catch (err) {
      toast({
        title: "Failed to start Google login",
        description: err instanceof Error ? err.message : "Unknown error",
        variant: "destructive",
      });
    }
  };

  const handleEmailLogin = async (email: string, password: string) => {
    const data = await authEmailLogin({ email, password });
    if (data.token) {
      localStorage.setItem("auth_token", data.token);
    }
    if (data.user) {
      qc.setQueryData(getAuthMeQueryKey(), { user: data.user });
      qc.invalidateQueries({ queryKey: ["/api/auth/google/status"] });
      qc.invalidateQueries({ queryKey: getMicrosoftAuthStatusQueryKey() });
    }
  };

  const handleEmailRegister = async (name: string, email: string, password: string) => {
    const data = await authEmailRegister({ name, email, password });
    if (data.token) {
      localStorage.setItem("auth_token", data.token);
    }
    if (data.user) {
      qc.setQueryData(getAuthMeQueryKey(), { user: data.user });
      qc.invalidateQueries({ queryKey: ["/api/auth/google/status"] });
      qc.invalidateQueries({ queryKey: getMicrosoftAuthStatusQueryKey() });
    }
  };

  const handleForgotPassword = async (email: string) => {
    const data = await authForgotPassword({ email });
    return data;
  };

  const handleResetPassword = async (token: string, password: string) => {
    const data = await authResetPassword({ token, password });
    if (data.token) {
      localStorage.setItem("auth_token", data.token);
    }
    if (data.user) {
      qc.setQueryData(getAuthMeQueryKey(), { user: data.user });
      qc.invalidateQueries({ queryKey: ["/api/auth/google/status"] });
      qc.invalidateQueries({ queryKey: getMicrosoftAuthStatusQueryKey() });
      try {
        const params = new URLSearchParams(window.location.search);
        params.delete("reset_token");
        const newSearch = params.toString();
        const hashParams = new URLSearchParams(window.location.hash.slice(1));
        hashParams.delete("reset_token");
        const newHash = hashParams.toString() ? "#" + hashParams.toString() : "";
        const cleanPath = window.location.pathname.replace(/^\/\/+/, "/");
        const newUrl = cleanPath + (newSearch ? "?" + newSearch : "") + newHash;
        window.history.replaceState({}, "", newUrl);
      } catch {
        // URL cleanup is cosmetic — ignore if blocked
      }
    }
  };

  const handleClaimAccount = async (_email: string, password: string) => {
    const data = await authClaimAccount({
      password,
      ...(claimToken ? { claimToken } : {}),
    });
    if (data.token) {
      localStorage.setItem("auth_token", data.token);
    }
    if (data.user) {
      qc.setQueryData(getAuthMeQueryKey(), { user: data.user });
      qc.invalidateQueries({ queryKey: ["/api/auth/google/status"] });
      qc.invalidateQueries({ queryKey: getMicrosoftAuthStatusQueryKey() });
      const params = new URLSearchParams(window.location.search);
      params.delete("claim_email");
      params.delete("claim_token");
      const newSearch = params.toString();
      const newUrl = window.location.pathname + (newSearch ? "?" + newSearch : "") + window.location.hash;
      window.history.replaceState({}, "", newUrl);
    }
  };

  const handleGuestLogin = () => {
    guestLoginMutation.mutate(undefined, {
      onSuccess: (data) => {
        if (data.token) {
          localStorage.setItem("auth_token", data.token);
        }
        qc.invalidateQueries({ queryKey: getAuthMeQueryKey() });
        toast({ title: "Signed in as guest" });
      },
      onError: (err) => {
        toast({
          title: "Could not sign in as guest",
          description: err instanceof Error ? err.message : "Something went wrong. Please try again.",
          variant: "destructive",
        });
      },
    });
  };

  useEffect(() => {
    const path = window.location.pathname;
    if (path === "/upgrade/success") {
      window.history.replaceState({}, "", "/");
      const params = new URLSearchParams(window.location.search);
      const sessionId = params.get("session_id");
      if (sessionId) {
        // Verify the session directly with Stripe so plan is updated immediately,
        // regardless of webhook delivery timing.
        stripeVerifySession({ sessionId })
          .then(() => queryClient.invalidateQueries({ queryKey: getAuthMeQueryKey() }))
          .catch(() => {})
          .finally(() => {
            toast({ title: "Welcome to Pro!", description: "Your subscription is active. Enjoy all Pro features." });
          });
      } else {
        // Fallback: just refetch to pick up any webhook-applied plan change
        queryClient.invalidateQueries({ queryKey: getAuthMeQueryKey() });
        toast({ title: "Welcome to Pro!", description: "Your subscription is active. Enjoy all Pro features." });
      }
    } else if (path === "/upgrade/cancelled") {
      window.history.replaceState({}, "", "/");
      toast({ title: "Upgrade cancelled", description: "Your plan has not changed. Upgrade any time from your account." });
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  if (window.location.pathname === "/admin") {
    return <AdminPage />;
  }

  if (window.location.pathname === "/privacy") {
    return <PrivacyPage />;
  }

  if (window.location.pathname === "/terms") {
    return <TermsPage />;
  }

  if (authQuery.isLoading || providersQuery.isLoading || !referralReady) {
    return <SplashScreen />;
  }

  if (!isSignedIn || resetToken || claimEmail || claimToken) {
    return (
      <SignInPage
        googleLoginAvailable={googleLoginAvailable}
        appleLoginAvailable={appleLoginAvailable}
        onGoogleLogin={handleGoogleLogin}
        onAppleLogin={() => {}}
        onGuestLogin={handleGuestLogin}
        onEmailLogin={handleEmailLogin}
        onEmailRegister={handleEmailRegister}
        onForgotPassword={handleForgotPassword}
        onResetPassword={handleResetPassword}
        onClaimAccount={handleClaimAccount}
        resetToken={resetToken}
        claimEmail={claimEmail}
        isLoggingIn={guestLoginMutation.isPending}
      />
    );
  }

  return (
    <div className="min-h-screen bg-background text-foreground">
      <BudgetWizard
        currentUser={currentUser}
        isSignedIn={isSignedIn}
        isGuest={isGuest}
        googleLoginAvailable={googleLoginAvailable}
        appleLoginAvailable={appleLoginAvailable}
      />
    </div>
  );
}

function App() {
  return (
    <AppErrorBoundary>
      <QueryClientProvider client={queryClient}>
        <TooltipProvider>
          <AppRouting />
          <Toaster />
        </TooltipProvider>
      </QueryClientProvider>
    </AppErrorBoundary>
  );
}

export default App;
