import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  DollarSign,
  UserPlus,
  Loader2,
  Zap,
  Sheet,
  TrendingDown,
  PiggyBank,
  Check,
  Star,
  ChevronDown,
  Eye,
  EyeOff,
  Mail,
  Lock,
  User as UserIcon,
  ArrowRight,
} from "lucide-react";

interface SignInPageProps {
  googleLoginAvailable: boolean;
  appleLoginAvailable: boolean;
  microsoftLoginAvailable?: boolean;
  onGoogleLogin: () => void;
  onAppleLogin: () => void;
  onMicrosoftLogin?: () => void;
  onGuestLogin: () => void;
  onEmailLogin: (email: string, password: string) => Promise<void>;
  onEmailRegister: (name: string, email: string, password: string) => Promise<void>;
  onForgotPassword: (email: string) => Promise<{ resetUrl?: string } | void>;
  onResetPassword?: (token: string, password: string) => Promise<void>;
  onClaimAccount?: (email: string, password: string) => Promise<void>;
  resetToken?: string | null;
  claimEmail?: string | null;
  isLoggingIn: boolean;
}

function GoogleIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 24 24">
      <path d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92a5.06 5.06 0 0 1-2.2 3.32v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.1z" fill="#4285F4" />
      <path d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" fill="#34A853" />
      <path d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z" fill="#FBBC05" />
      <path d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z" fill="#EA4335" />
    </svg>
  );
}

type AuthView = "signin" | "register" | "forgot" | "forgot_sent" | "reset" | "claim";

function PasswordInput({ value, onChange, placeholder, id }: { value: string; onChange: (v: string) => void; placeholder?: string; id?: string }) {
  const [show, setShow] = useState(false);
  return (
    <div className="relative">
      <Lock className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground pointer-events-none" />
      <Input
        id={id}
        type={show ? "text" : "password"}
        value={value}
        onChange={e => onChange(e.target.value)}
        placeholder={placeholder ?? "Password"}
        className="pl-9 pr-10"
        autoComplete={id?.includes("confirm") ? "new-password" : id?.includes("new") ? "new-password" : "current-password"}
      />
      <button
        type="button"
        onClick={() => setShow(s => !s)}
        className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground"
        tabIndex={-1}
      >
        {show ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
      </button>
    </div>
  );
}

function AuthPanel({
  googleLoginAvailable,
  onGoogleLogin,
  onGuestLogin,
  onEmailLogin,
  onEmailRegister,
  onForgotPassword,
  onResetPassword,
  onClaimAccount,
  resetToken,
  claimEmail,
  isLoggingIn,
  instanceId = "hero",
}: {
  googleLoginAvailable: boolean;
  onGoogleLogin: () => void;
  onGuestLogin: () => void;
  onEmailLogin: (email: string, password: string) => Promise<void>;
  onEmailRegister: (name: string, email: string, password: string) => Promise<void>;
  onForgotPassword: (email: string) => Promise<{ resetUrl?: string } | void>;
  onResetPassword?: (token: string, password: string) => Promise<void>;
  onClaimAccount?: (email: string, password: string) => Promise<void>;
  resetToken?: string | null;
  claimEmail?: string | null;
  isLoggingIn: boolean;
  instanceId?: string;
}) {
  const initialView: AuthView = claimEmail ? "claim" : resetToken ? "reset" : "signin";
  const [view, setView] = useState<AuthView>(initialView);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [clickedGoogle, setClickedGoogle] = useState(false);

  const [signinEmail, setSigninEmail] = useState("");
  const [signinPassword, setSigninPassword] = useState("");

  const [regName, setRegName] = useState("");
  const [regEmail, setRegEmail] = useState("");
  const [regPassword, setRegPassword] = useState("");

  const [forgotEmail, setForgotEmail] = useState("");
  const [devResetUrl, setDevResetUrl] = useState<string | null>(null);

  const [resetPassword, setResetPassword] = useState("");
  const [resetConfirm, setResetConfirm] = useState("");

  const [claimPassword, setClaimPassword] = useState("");
  const [claimConfirm, setClaimConfirm] = useState("");

  const busy = loading || isLoggingIn;

  const handleSignin = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setLoading(true);
    try {
      await onEmailLogin(signinEmail.trim(), signinPassword);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Sign in failed. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  const handleRegister = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    if (regPassword.length < 8) {
      setError("Password must be at least 8 characters");
      return;
    }
    setLoading(true);
    try {
      await onEmailRegister(regName.trim(), regEmail.trim(), regPassword);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Registration failed. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  const handleForgot = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setLoading(true);
    try {
      const result = await onForgotPassword(forgotEmail.trim());
      if (result?.resetUrl) setDevResetUrl(result.resetUrl);
      setView("forgot_sent");
    } catch {
      setView("forgot_sent");
    } finally {
      setLoading(false);
    }
  };

  const handleReset = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    if (resetPassword.length < 8) {
      setError("Password must be at least 8 characters");
      return;
    }
    if (resetPassword !== resetConfirm) {
      setError("Passwords don't match");
      return;
    }
    if (!onResetPassword || !resetToken) return;
    setLoading(true);
    try {
      await onResetPassword(resetToken, resetPassword);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Reset failed. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  const handleClaim = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    if (claimPassword.length < 8) {
      setError("Password must be at least 8 characters");
      return;
    }
    if (claimPassword !== claimConfirm) {
      setError("Passwords don't match");
      return;
    }
    if (!onClaimAccount || !claimEmail) return;
    setLoading(true);
    try {
      await onClaimAccount(claimEmail, claimPassword);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to set password. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  const handleGoogle = () => {
    setClickedGoogle(true);
    onGoogleLogin();
  };

  const switchView = (v: AuthView) => {
    setView(v);
    setError(null);
  };

  return (
    <div className="w-full max-w-sm">
      {view === "reset" && (
        <form onSubmit={handleReset} className="flex flex-col gap-4">
          <div className="text-center mb-2">
            <h2 className="text-xl font-bold">Set new password</h2>
            <p className="text-muted-foreground text-sm mt-1">Choose a new password for your account.</p>
          </div>
          {error && <p className="text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2">{error}</p>}
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="reset-pass">New password</Label>
            <PasswordInput id="reset-pass-new" value={resetPassword} onChange={setResetPassword} placeholder="At least 8 characters" />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="reset-confirm">Confirm password</Label>
            <PasswordInput id="reset-confirm" value={resetConfirm} onChange={setResetConfirm} placeholder="Re-enter password" />
          </div>
          <Button type="submit" className="w-full h-11 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white" disabled={busy}>
            {loading ? <Loader2 className="w-4 h-4 animate-spin" /> : "Set password & sign in"}
          </Button>
        </form>
      )}

      {view === "claim" && claimEmail && (
        <form onSubmit={handleClaim} className="flex flex-col gap-4">
          <div className="text-center mb-2">
            <h2 className="text-xl font-bold">Set a password</h2>
            <p className="text-muted-foreground text-sm mt-1">
              We found your account for <strong>{claimEmail}</strong>. Set a password to sign in with email going forward.
            </p>
          </div>
          {error && <p className="text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2">{error}</p>}
          <div className="flex flex-col gap-1.5">
            <Label>New password</Label>
            <PasswordInput value={claimPassword} onChange={setClaimPassword} placeholder="At least 8 characters" />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Confirm password</Label>
            <PasswordInput value={claimConfirm} onChange={setClaimConfirm} placeholder="Re-enter password" />
          </div>
          <Button type="submit" className="w-full h-11 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white" disabled={busy}>
            {loading ? <Loader2 className="w-4 h-4 animate-spin" /> : "Set password & sign in"}
          </Button>
          <button type="button" onClick={() => switchView("signin")} className="text-sm text-muted-foreground hover:text-foreground text-center transition-colors">
            Back to sign in
          </button>
        </form>
      )}

      {view === "forgot" && (
        <form onSubmit={handleForgot} className="flex flex-col gap-4">
          <div className="text-center mb-2">
            <h2 className="text-xl font-bold">Forgot password?</h2>
            <p className="text-muted-foreground text-sm mt-1">Enter your email and we'll send a reset link.</p>
          </div>
          {error && <p className="text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2">{error}</p>}
          <div className="flex flex-col gap-1.5">
            <Label htmlFor={`${instanceId}-forgot-email`}>Email address</Label>
            <div className="relative">
              <Mail className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground pointer-events-none" />
              <Input id={`${instanceId}-forgot-email`} type="email" value={forgotEmail} onChange={e => setForgotEmail(e.target.value)} placeholder="you@example.com" className="pl-9" autoComplete="email" />
            </div>
          </div>
          <Button type="submit" className="w-full h-11 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white" disabled={busy || !forgotEmail.includes("@")}>
            {loading ? <Loader2 className="w-4 h-4 animate-spin" /> : "Send reset link"}
          </Button>
          <button type="button" onClick={() => switchView("signin")} className="text-sm text-muted-foreground hover:text-foreground text-center transition-colors">
            Back to sign in
          </button>
        </form>
      )}

      {view === "forgot_sent" && (
        <div className="flex flex-col items-center gap-4 text-center">
          <div className="w-12 h-12 rounded-full bg-emerald-100 flex items-center justify-center">
            <Check className="w-6 h-6 text-emerald-600" />
          </div>
          <div>
            <h2 className="text-xl font-bold mb-1">Check your email</h2>
            <p className="text-muted-foreground text-sm">If an account exists for {forgotEmail || "that address"}, we've sent a reset link. It expires in 1 hour.</p>
          </div>
          {devResetUrl && (
            <div className="w-full text-left bg-amber-50 border border-amber-200 rounded-xl p-3">
              <p className="text-xs font-semibold text-amber-700 mb-1">Dev mode — reset link:</p>
              <a href={devResetUrl} className="text-xs text-amber-800 break-all underline underline-offset-2">{devResetUrl}</a>
            </div>
          )}
          <button type="button" onClick={() => switchView("signin")} className="text-sm text-muted-foreground hover:text-foreground transition-colors">
            Back to sign in
          </button>
        </div>
      )}

      {(view === "signin" || view === "register") && (
        <>
          <div className="flex rounded-xl border border-border/60 p-1 mb-5 bg-muted/30">
            <button
              type="button"
              onClick={() => switchView("signin")}
              className={`flex-1 py-2 text-sm font-medium rounded-lg transition-all ${view === "signin" ? "bg-white shadow-sm text-foreground" : "text-muted-foreground hover:text-foreground"}`}
            >
              Sign in
            </button>
            <button
              type="button"
              onClick={() => switchView("register")}
              className={`flex-1 py-2 text-sm font-medium rounded-lg transition-all ${view === "register" ? "bg-white shadow-sm text-foreground" : "text-muted-foreground hover:text-foreground"}`}
            >
              Create account
            </button>
          </div>

          {error && <p className="text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2 mb-4">{error}</p>}

          {view === "signin" && (
            <form onSubmit={handleSignin} className="flex flex-col gap-4">
              <div className="flex flex-col gap-1.5">
                <Label htmlFor={`${instanceId}-signin-email`}>Email</Label>
                <div className="relative">
                  <Mail className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground pointer-events-none" />
                  <Input id={`${instanceId}-signin-email`} type="email" value={signinEmail} onChange={e => setSigninEmail(e.target.value)} placeholder="you@example.com" className="pl-9" autoComplete="email" />
                </div>
              </div>
              <div className="flex flex-col gap-1.5">
                <div className="flex items-center justify-between">
                  <Label htmlFor={`${instanceId}-signin-password`}>Password</Label>
                  <button type="button" onClick={() => switchView("forgot")} className="text-xs text-muted-foreground hover:text-foreground transition-colors">
                    Forgot password?
                  </button>
                </div>
                <PasswordInput id={`${instanceId}-signin-password`} value={signinPassword} onChange={setSigninPassword} />
              </div>
              <Button type="submit" className="w-full h-11 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white" disabled={busy}>
                {loading ? <Loader2 className="w-4 h-4 animate-spin" /> : "Sign in"}
              </Button>
            </form>
          )}

          {view === "register" && (
            <form onSubmit={handleRegister} className="flex flex-col gap-4">
              <div className="flex flex-col gap-1.5">
                <Label htmlFor={`${instanceId}-reg-name`}>Your name</Label>
                <div className="relative">
                  <UserIcon className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground pointer-events-none" />
                  <Input id={`${instanceId}-reg-name`} type="text" value={regName} onChange={e => setRegName(e.target.value)} placeholder="Alex Johnson" className="pl-9" autoComplete="name" />
                </div>
              </div>
              <div className="flex flex-col gap-1.5">
                <Label htmlFor={`${instanceId}-reg-email`}>Email</Label>
                <div className="relative">
                  <Mail className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground pointer-events-none" />
                  <Input id={`${instanceId}-reg-email`} type="email" value={regEmail} onChange={e => setRegEmail(e.target.value)} placeholder="you@example.com" className="pl-9" autoComplete="email" />
                </div>
              </div>
              <div className="flex flex-col gap-1.5">
                <Label htmlFor={`${instanceId}-reg-password`}>Password</Label>
                <PasswordInput id={`${instanceId}-reg-password-new`} value={regPassword} onChange={setRegPassword} placeholder="At least 8 characters" />
              </div>
              <Button type="submit" className="w-full h-11 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white" disabled={busy}>
                {loading ? <Loader2 className="w-4 h-4 animate-spin" /> : "Create account"}
              </Button>
            </form>
          )}

          <div className="flex flex-col gap-3 mt-4">
            <div className="relative">
              <div className="absolute inset-0 flex items-center">
                <div className="w-full border-t border-border/50" />
              </div>
              <div className="relative flex justify-center text-xs">
                <span className="bg-white px-3 text-muted-foreground">or</span>
              </div>
            </div>

            <Button
              type="button"
              size="lg"
              variant="ghost"
              className="w-full h-11 rounded-xl gap-2 text-muted-foreground hover:text-foreground border border-border/30"
              onClick={onGuestLogin}
              disabled={busy}
            >
              {clickedGoogle === false && isLoggingIn ? (
                <Loader2 className="w-4 h-4 animate-spin" />
              ) : (
                <UserPlus className="w-4 h-4" />
              )}
              Try for free — no account needed
            </Button>

            <p className="text-xs text-center text-muted-foreground/70 leading-relaxed">
              Had a Budgify account via Google or Apple?{" "}
              <button type="button" onClick={() => switchView("forgot")} className="underline underline-offset-2 hover:text-foreground transition-colors">
                Use forgot password
              </button>{" "}
              with your email to set a password.
            </p>

            {googleLoginAvailable && (
              <button
                type="button"
                onClick={handleGoogle}
                disabled={busy}
                className="text-xs text-center text-muted-foreground hover:text-foreground transition-colors flex items-center justify-center gap-1.5"
              >
                {clickedGoogle && isLoggingIn ? (
                  <Loader2 className="w-3 h-3 animate-spin" />
                ) : (
                  <GoogleIcon className="w-3.5 h-3.5" />
                )}
                Had a Google account? Sign in here
              </button>
            )}
          </div>
        </>
      )}
    </div>
  );
}

const features = [
  {
    icon: Zap,
    title: "Weekly Budget Automation",
    description:
      "Set your budget once and let Budgify do the rest. Automatic weekly resets keep your spending on track without any manual work.",
    color: "from-amber-400 to-orange-500",
  },
  {
    icon: Sheet,
    title: "Google Sheets Sync",
    description:
      "Every transaction, category, and summary syncs to your own Google Sheet in real time — your data, always accessible and exportable.",
    color: "from-emerald-400 to-teal-500",
  },
  {
    icon: TrendingDown,
    title: "Debt Tracking",
    description:
      "Visualize your debt paydown journey. Track balances, minimum payments, and watch the total shrink over time with clear progress indicators.",
    color: "from-rose-400 to-pink-500",
  },
  {
    icon: PiggyBank,
    title: "Savings Goals",
    description:
      "Name your goals, set targets, and see how close you are. Budgify helps you put money toward what matters — vacation, emergency fund, or a new car.",
    color: "from-violet-400 to-purple-500",
  },
];

const testimonials = [
  {
    name: "Alex M.",
    role: "Freelance designer",
    quote:
      "I finally know where my money goes each week. The Google Sheets sync means my partner and I are always on the same page.",
    rating: 5,
  },
  {
    name: "Jordan K.",
    role: "Software engineer",
    quote:
      "Paid off $12,000 in credit card debt last year using the debt tracker. Seeing the number shrink kept me motivated.",
    rating: 5,
  },
  {
    name: "Sam R.",
    role: "Teacher",
    quote:
      "The weekly reset is genius. It made budgeting feel doable instead of overwhelming. I've stuck with it for 8 months now.",
    rating: 5,
  },
];

export function SignInPage({
  googleLoginAvailable,
  onGoogleLogin,
  onAppleLogin: _onAppleLogin,
  onGuestLogin,
  onEmailLogin,
  onEmailRegister,
  onForgotPassword,
  onResetPassword,
  onClaimAccount,
  resetToken,
  claimEmail,
  isLoggingIn,
}: SignInPageProps) {

  const scrollToSignIn = () => {
    document.getElementById("cta-section")?.scrollIntoView({ behavior: "smooth" });
  };

  return (
    <div className="min-h-screen bg-white text-foreground overflow-x-hidden">
      <nav className="sticky top-0 z-50 bg-white/90 backdrop-blur-md border-b border-border/40 px-4 sm:px-8 py-3 flex items-center justify-between">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-primary to-emerald-600 flex items-center justify-center shadow-sm">
            <DollarSign className="w-4 h-4 text-white" />
          </div>
          <span className="text-lg font-bold tracking-tight">Budgify</span>
        </div>
        <Button
          size="sm"
          className="rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-sm px-4 h-9 shadow-sm"
          onClick={scrollToSignIn}
        >
          Get started free
        </Button>
      </nav>

      <section className="relative px-4 sm:px-8 pt-16 pb-20 sm:pt-24 sm:pb-28 flex flex-col items-center text-center overflow-hidden">
        <div
          className="absolute inset-0 -z-10 pointer-events-none"
          aria-hidden="true"
          style={{
            background:
              "radial-gradient(ellipse 80% 60% at 50% 0%, rgba(16,185,129,0.10) 0%, transparent 70%)",
          }}
        />

        <div className="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-emerald-50 border border-emerald-100 text-emerald-700 text-xs font-medium mb-6">
          <Zap className="w-3.5 h-3.5" />
          Automated weekly budgets — no spreadsheet skills needed
        </div>

        <h1 className="text-4xl sm:text-5xl md:text-6xl font-extrabold tracking-tight leading-[1.1] mb-5 max-w-3xl">
          Your money,{" "}
          <span className="bg-gradient-to-r from-emerald-500 to-teal-400 bg-clip-text text-transparent">
            finally under control
          </span>
        </h1>

        <p className="text-lg sm:text-xl text-muted-foreground max-w-xl mb-10 leading-relaxed">
          Budgify automates your weekly budget, syncs everything to Google Sheets, and keeps your savings goals and debt paydown on track — all in one place.
        </p>

        <div className="flex flex-col items-center gap-4 w-full">
          <AuthPanel
            googleLoginAvailable={googleLoginAvailable}
            onGoogleLogin={onGoogleLogin}
            onGuestLogin={onGuestLogin}
            onEmailLogin={onEmailLogin}
            onEmailRegister={onEmailRegister}
            onForgotPassword={onForgotPassword}
            onResetPassword={onResetPassword}
            onClaimAccount={onClaimAccount}
            resetToken={resetToken}
            claimEmail={claimEmail}
            isLoggingIn={isLoggingIn}
          />
          <p className="text-xs text-muted-foreground/70">
            Free forever · No credit card required
          </p>
        </div>

        <button
          onClick={scrollToSignIn}
          className="mt-12 flex flex-col items-center gap-1 text-muted-foreground/50 text-xs hover:text-muted-foreground transition-colors"
          aria-label="Scroll down"
        >
          <span>Learn more</span>
          <ChevronDown className="w-4 h-4 animate-bounce" />
        </button>
      </section>

      <section className="px-4 sm:px-8 py-20 sm:py-28 max-w-5xl mx-auto">
        <div className="text-center mb-14">
          <h2 className="text-3xl sm:text-4xl font-bold tracking-tight mb-4">
            Everything you need to budget smarter
          </h2>
          <p className="text-muted-foreground text-lg max-w-xl mx-auto">
            Four powerful features that work together so you spend less time worrying and more time living.
          </p>
        </div>

        <div className="grid sm:grid-cols-2 gap-6">
          {features.map((feature) => {
            const Icon = feature.icon;
            return (
              <div
                key={feature.title}
                className="group p-6 rounded-2xl border border-border/50 bg-white hover:border-emerald-200 hover:shadow-lg hover:shadow-emerald-50/50 transition-all duration-200"
              >
                <div className={`w-11 h-11 rounded-xl bg-gradient-to-br ${feature.color} flex items-center justify-center mb-4 shadow-sm`}>
                  <Icon className="w-5 h-5 text-white" />
                </div>
                <h3 className="text-lg font-semibold mb-2">{feature.title}</h3>
                <p className="text-muted-foreground text-sm leading-relaxed">{feature.description}</p>
              </div>
            );
          })}
        </div>
      </section>

      <section className="bg-slate-50 border-y border-border/30 px-4 sm:px-8 py-20 sm:py-24">
        <div className="max-w-5xl mx-auto">
          <div className="text-center mb-12">
            <h2 className="text-3xl sm:text-4xl font-bold tracking-tight mb-3">
              Real people, real results
            </h2>
            <p className="text-muted-foreground text-lg">
              Join thousands who have taken control of their finances with Budgify.
            </p>
          </div>

          <div className="grid sm:grid-cols-3 gap-6">
            {testimonials.map((t) => (
              <div key={t.name} className="bg-white rounded-2xl p-6 border border-border/50 shadow-sm flex flex-col gap-4">
                <div className="flex gap-0.5">
                  {Array.from({ length: t.rating }).map((_, i) => (
                    <Star key={i} className="w-4 h-4 fill-amber-400 text-amber-400" />
                  ))}
                </div>
                <p className="text-sm text-foreground/80 leading-relaxed flex-1">"{t.quote}"</p>
                <div>
                  <div className="font-semibold text-sm">{t.name}</div>
                  <div className="text-xs text-muted-foreground">{t.role}</div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </section>

      <section
        id="cta-section"
        className="bg-gradient-to-br from-emerald-600 to-teal-600 px-4 sm:px-8 py-20 sm:py-28 flex flex-col items-center text-center text-white"
      >
        <div className="w-14 h-14 rounded-2xl bg-white/20 flex items-center justify-center mb-6">
          <DollarSign className="w-7 h-7 text-white" />
        </div>

        <h2 className="text-3xl sm:text-4xl font-bold mb-4 max-w-xl">
          Ready to take control of your budget?
        </h2>
        <p className="text-emerald-100 text-lg mb-10 max-w-md">
          Join thousands of people who have made their weekly budget automatic with Budgify. Start free, today.
        </p>

        <div className="flex flex-col items-center gap-3 w-full max-w-sm">
          <div className="w-full bg-white/10 backdrop-blur-sm border border-white/20 rounded-2xl p-5">
            <AuthPanel
              googleLoginAvailable={googleLoginAvailable}
              onGoogleLogin={onGoogleLogin}
              onGuestLogin={onGuestLogin}
              onEmailLogin={onEmailLogin}
              onEmailRegister={onEmailRegister}
              onForgotPassword={onForgotPassword}
              onResetPassword={onResetPassword}
              onClaimAccount={onClaimAccount}
              resetToken={resetToken}
              claimEmail={claimEmail}
              isLoggingIn={isLoggingIn}
              instanceId="cta"
            />
          </div>
          <p className="text-xs text-emerald-200 mt-1">
            Free forever · No credit card required
          </p>
        </div>
      </section>

      <footer className="border-t border-border/30 bg-white px-4 sm:px-8 py-8 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-muted-foreground">
        <div className="flex items-center gap-2">
          <div className="w-5 h-5 rounded bg-gradient-to-br from-primary to-emerald-600 flex items-center justify-center">
            <DollarSign className="w-3 h-3 text-white" />
          </div>
          <span className="font-semibold text-foreground">Budgify</span>
        </div>
        <span>© {new Date().getFullYear()} Budgify. All rights reserved.</span>
        <div className="flex gap-4">
          <a href="#" className="hover:text-foreground transition-colors">Privacy</a>
          <a href="#" className="hover:text-foreground transition-colors">Terms</a>
        </div>
      </footer>
    </div>
  );
}
