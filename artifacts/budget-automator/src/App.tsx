import { useEffect, Component } from "react";
import type { ReactNode } from "react";
import { QueryClient, QueryClientProvider, useQueryClient } from "@tanstack/react-query";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import { useToast } from "@/hooks/use-toast";
import { BudgetWizard } from "@/pages/BudgetWizard";
import { SignInPage } from "@/pages/SignInPage";
import type { AuthUser } from "@workspace/api-client-react";
import {
  useAuthMe,
  useAuthProviders,
  useAuthGuestLogin,
  authLoginGoogle,
  authLoginApple,
  getAuthMeQueryKey,
  getAuthProvidersQueryKey,
  stripeCheckout,
} from "@workspace/api-client-react";
import { DollarSign, Loader2, RefreshCw } from "lucide-react";

class AppErrorBoundary extends Component<{ children: ReactNode }, { error: Error | null }> {
  constructor(props: { children: ReactNode }) {
    super(props);
    this.state = { error: null };
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
      return (
        <div className="min-h-screen flex flex-col items-center justify-center bg-gradient-to-br from-slate-50 via-white to-emerald-50 p-6">
          <div className="w-14 h-14 rounded-2xl bg-gradient-to-br from-primary to-emerald-600 flex items-center justify-center mb-6 shadow-lg shadow-primary/20">
            <DollarSign className="w-7 h-7 text-white" />
          </div>
          <h1 className="text-xl font-bold text-foreground mb-2">Something went wrong</h1>
          <p className="text-muted-foreground text-sm text-center mb-6 max-w-xs">
            Cached app data may be out of date. Clearing it and reloading usually fixes this.
          </p>
          <button
            onClick={this.handleReset}
            className="flex items-center gap-2 px-5 py-2.5 rounded-xl bg-emerald-600 text-white text-sm font-medium hover:bg-emerald-700 transition-colors"
          >
            <RefreshCw className="w-4 h-4" /> Clear cache &amp; reload
          </button>
          <details className="mt-6 max-w-sm w-full">
            <summary className="text-xs text-muted-foreground cursor-pointer">Error details</summary>
            <pre className="mt-2 text-xs text-red-600 bg-red-50 rounded-lg p-3 overflow-auto whitespace-pre-wrap break-all">
              {this.state.error.message}
            </pre>
          </details>
        </div>
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

  const handleAppleLogin = async () => {
    try {
      const currentUrl = window.location.href;
      const result = await authLoginApple({ redirect: currentUrl });
      window.location.href = result.url;
    } catch (err) {
      toast({
        title: "Failed to start Apple login",
        description: err instanceof Error ? err.message : "Unknown error",
        variant: "destructive",
      });
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
      toast({ title: "Welcome to Pro!", description: "Your subscription is active. Enjoy all Pro features." });
    } else if (path === "/upgrade/cancelled") {
      window.history.replaceState({}, "", "/");
      toast({ title: "Upgrade cancelled", description: "Your plan has not changed. Upgrade any time from your account." });
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  if (authQuery.isLoading || providersQuery.isLoading) {
    return <SplashScreen />;
  }

  if (!isSignedIn) {
    return (
      <SignInPage
        googleLoginAvailable={googleLoginAvailable}
        appleLoginAvailable={appleLoginAvailable}
        onGoogleLogin={handleGoogleLogin}
        onAppleLogin={handleAppleLogin}
        onGuestLogin={handleGuestLogin}
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
