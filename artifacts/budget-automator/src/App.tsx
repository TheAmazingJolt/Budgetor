import { useEffect } from "react";
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
} from "@workspace/api-client-react";
import { DollarSign, Loader2 } from "lucide-react";

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
      .then((r) => r.json())
      .then((data: { user?: AuthUser; token?: string }) => {
        if (data.token) {
          localStorage.setItem("auth_token", data.token);
        }
        if (data.user) {
          qc.setQueryData(getAuthMeQueryKey(), { user: data.user });
          qc.invalidateQueries({ queryKey: ["/api/auth/google/status"] });
        }
      })
      .catch(() => {
        authQuery.refetch();
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
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <AppRouting />
        <Toaster />
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;
