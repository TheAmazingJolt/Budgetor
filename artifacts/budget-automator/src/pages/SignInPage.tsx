import { useState } from "react";
import { Button } from "@/components/ui/button";
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
  ArrowRight,
  ChevronDown,
} from "lucide-react";

interface SignInPageProps {
  googleLoginAvailable: boolean;
  appleLoginAvailable: boolean;
  onGoogleLogin: () => void;
  onAppleLogin: () => void;
  onGuestLogin: () => void;
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

function AppleIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="currentColor">
      <path d="M17.05 20.28c-.98.95-2.05.88-3.08.4-1.09-.5-2.08-.48-3.24 0-1.44.62-2.2.44-3.06-.4C2.79 15.25 3.51 7.59 9.05 7.31c1.35.07 2.29.74 3.08.8 1.18-.24 2.31-.93 3.57-.84 1.51.12 2.65.72 3.4 1.8-3.12 1.87-2.38 5.98.48 7.13-.57 1.5-1.31 2.99-2.54 4.09zM12.03 7.25c-.15-2.23 1.66-4.07 3.74-4.25.29 2.58-2.34 4.5-3.74 4.25z" />
    </svg>
  );
}

function AuthButtons({
  googleLoginAvailable,
  appleLoginAvailable,
  onGoogleLogin,
  onAppleLogin,
  onGuestLogin,
  isLoggingIn,
  clickedProvider,
  setClickedProvider,
  variant = "default",
}: {
  googleLoginAvailable: boolean;
  appleLoginAvailable: boolean;
  onGoogleLogin: () => void;
  onAppleLogin: () => void;
  onGuestLogin: () => void;
  isLoggingIn: boolean;
  clickedProvider: string | null;
  setClickedProvider: (v: string | null) => void;
  variant?: "default" | "compact";
}) {
  const handleGoogle = () => { setClickedProvider("google"); onGoogleLogin(); };
  const handleApple = () => { setClickedProvider("apple"); onAppleLogin(); };
  const handleGuest = () => { setClickedProvider("guest"); onGuestLogin(); };

  const isCompact = variant === "compact";

  return (
    <div className={`flex flex-col gap-3 ${isCompact ? "w-full max-w-xs" : "w-full max-w-sm"}`}>
      {googleLoginAvailable && (
        <Button
          size="lg"
          className="w-full h-12 rounded-xl gap-2 bg-white text-foreground border border-border/60 hover:bg-slate-50 shadow-sm"
          variant="outline"
          onClick={handleGoogle}
          disabled={isLoggingIn}
        >
          {clickedProvider === "google" && isLoggingIn ? (
            <Loader2 className="w-5 h-5 animate-spin" />
          ) : (
            <GoogleIcon className="w-5 h-5" />
          )}
          Sign in with Google
        </Button>
      )}

      {appleLoginAvailable && (
        <Button
          size="lg"
          className="w-full h-12 rounded-xl gap-2 bg-black text-white hover:bg-black/90 shadow-sm"
          onClick={handleApple}
          disabled={isLoggingIn}
        >
          {clickedProvider === "apple" && isLoggingIn ? (
            <Loader2 className="w-5 h-5 animate-spin" />
          ) : (
            <AppleIcon className="w-5 h-5" />
          )}
          Sign in with Apple
        </Button>
      )}

      {(googleLoginAvailable || appleLoginAvailable) && (
        <div className="relative my-1">
          <div className="absolute inset-0 flex items-center">
            <div className="w-full border-t border-border/50" />
          </div>
          <div className="relative flex justify-center text-xs">
            <span className="bg-white px-3 text-muted-foreground">or</span>
          </div>
        </div>
      )}

      <Button
        size="lg"
        variant="ghost"
        className="w-full h-12 rounded-xl gap-2 text-muted-foreground hover:text-foreground border border-border/30"
        onClick={handleGuest}
        disabled={isLoggingIn}
      >
        {clickedProvider === "guest" && isLoggingIn ? (
          <Loader2 className="w-5 h-5 animate-spin" />
        ) : (
          <UserPlus className="w-5 h-5" />
        )}
        Try for free — no account needed
      </Button>
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

const freeFeatures = [
  "Weekly budget automation",
  "Up to 5 spending categories",
  "Basic debt tracker",
  "1 savings goal",
  "Guest mode — no sign-up required",
];

const proFeatures = [
  "Everything in Free",
  "Unlimited spending categories",
  "Google Sheets sync",
  "Unlimited savings goals",
  "Advanced analytics & charts",
  "Priority support",
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
  appleLoginAvailable,
  onGoogleLogin,
  onAppleLogin,
  onGuestLogin,
  isLoggingIn,
}: SignInPageProps) {
  const [clickedProvider, setClickedProvider] = useState<string | null>(null);

  const scrollToSignIn = () => {
    document.getElementById("cta-section")?.scrollIntoView({ behavior: "smooth" });
  };

  return (
    <div className="min-h-screen bg-white text-foreground overflow-x-hidden">
      {/* Nav */}
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

      {/* Hero */}
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
          <AuthButtons
            googleLoginAvailable={googleLoginAvailable}
            appleLoginAvailable={appleLoginAvailable}
            onGoogleLogin={onGoogleLogin}
            onAppleLogin={onAppleLogin}
            onGuestLogin={onGuestLogin}
            isLoggingIn={isLoggingIn}
            clickedProvider={clickedProvider}
            setClickedProvider={setClickedProvider}
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

      {/* Stats bar */}
      <section className="bg-slate-50 border-y border-border/30 px-4 sm:px-8 py-8">
        <div className="max-w-4xl mx-auto grid grid-cols-2 sm:grid-cols-4 gap-6 text-center">
          {[
            { value: "10,000+", label: "Active budgeters" },
            { value: "$2.4M", label: "Debt tracked & paid off" },
            { value: "4.8★", label: "Average rating" },
            { value: "5 min", label: "Setup time" },
          ].map((stat) => (
            <div key={stat.label}>
              <div className="text-2xl sm:text-3xl font-extrabold text-foreground">{stat.value}</div>
              <div className="text-xs sm:text-sm text-muted-foreground mt-1">{stat.label}</div>
            </div>
          ))}
        </div>
      </section>

      {/* Features */}
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

      {/* Social proof */}
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

      {/* Pricing */}
      <section className="px-4 sm:px-8 py-20 sm:py-28 max-w-4xl mx-auto">
        <div className="text-center mb-14">
          <h2 className="text-3xl sm:text-4xl font-bold tracking-tight mb-4">
            Simple, transparent pricing
          </h2>
          <p className="text-muted-foreground text-lg">
            Start free. Upgrade when you're ready.
          </p>
        </div>

        <div className="grid sm:grid-cols-2 gap-6 items-start">
          {/* Free */}
          <div className="rounded-2xl border border-border/60 bg-white p-7 flex flex-col gap-5">
            <div>
              <div className="text-sm font-medium text-muted-foreground mb-1">Free</div>
              <div className="flex items-end gap-1">
                <span className="text-4xl font-extrabold">$0</span>
                <span className="text-muted-foreground text-sm mb-1.5">/month</span>
              </div>
              <p className="text-muted-foreground text-sm mt-2">Everything you need to get started — forever free.</p>
            </div>
            <ul className="flex flex-col gap-2.5">
              {freeFeatures.map((f) => (
                <li key={f} className="flex items-center gap-2.5 text-sm">
                  <Check className="w-4 h-4 text-emerald-500 flex-shrink-0" />
                  {f}
                </li>
              ))}
            </ul>
            <Button
              variant="outline"
              className="w-full rounded-xl h-11 mt-auto"
              onClick={scrollToSignIn}
            >
              Get started free
            </Button>
          </div>

          {/* Pro */}
          <div className="rounded-2xl border-2 border-emerald-500 bg-gradient-to-b from-emerald-50 to-white p-7 flex flex-col gap-5 relative shadow-lg shadow-emerald-100">
            <div className="absolute -top-3.5 left-1/2 -translate-x-1/2">
              <span className="bg-emerald-500 text-white text-xs font-semibold px-3 py-1 rounded-full shadow-sm">
                Most popular
              </span>
            </div>
            <div>
              <div className="text-sm font-medium text-emerald-700 mb-1">Pro</div>
              <div className="flex items-end gap-1">
                <span className="text-4xl font-extrabold">$5</span>
                <span className="text-muted-foreground text-sm mb-1.5">/month</span>
              </div>
              <p className="text-muted-foreground text-sm mt-2">For serious budgeters who want the full toolkit.</p>
            </div>
            <ul className="flex flex-col gap-2.5">
              {proFeatures.map((f) => (
                <li key={f} className="flex items-center gap-2.5 text-sm">
                  <Check className="w-4 h-4 text-emerald-500 flex-shrink-0" />
                  {f}
                </li>
              ))}
            </ul>
            <Button
              className="w-full rounded-xl h-11 mt-auto bg-gradient-to-r from-emerald-600 to-teal-600 text-white hover:from-emerald-700 hover:to-teal-700 gap-2"
              onClick={() => {
                localStorage.setItem("upgrade_intent", "pro");
                if (googleLoginAvailable) {
                  setClickedProvider("google");
                  onGoogleLogin();
                } else if (appleLoginAvailable) {
                  setClickedProvider("apple");
                  onAppleLogin();
                } else {
                  scrollToSignIn();
                }
              }}
              disabled={isLoggingIn}
            >
              {isLoggingIn && (clickedProvider === "google" || clickedProvider === "apple") ? (
                <Loader2 className="w-4 h-4 animate-spin" />
              ) : (
                <Star className="w-4 h-4" />
              )}{" "}
              Get started — $5/mo
            </Button>
          </div>
        </div>
      </section>

      {/* CTA footer */}
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
          {googleLoginAvailable && (
            <Button
              size="lg"
              className="w-full h-12 rounded-xl gap-2 bg-white text-foreground border border-white/60 hover:bg-slate-50 shadow-sm"
              variant="outline"
              onClick={() => { setClickedProvider("google"); onGoogleLogin(); }}
              disabled={isLoggingIn}
            >
              {clickedProvider === "google" && isLoggingIn ? (
                <Loader2 className="w-5 h-5 animate-spin" />
              ) : (
                <GoogleIcon className="w-5 h-5" />
              )}
              Sign in with Google
            </Button>
          )}

          {appleLoginAvailable && (
            <Button
              size="lg"
              className="w-full h-12 rounded-xl gap-2 bg-black text-white hover:bg-black/90 shadow-sm"
              onClick={() => { setClickedProvider("apple"); onAppleLogin(); }}
              disabled={isLoggingIn}
            >
              {clickedProvider === "apple" && isLoggingIn ? (
                <Loader2 className="w-5 h-5 animate-spin" />
              ) : (
                <AppleIcon className="w-5 h-5" />
              )}
              Sign in with Apple
            </Button>
          )}

          <Button
            size="lg"
            variant="ghost"
            className="w-full h-12 rounded-xl gap-2 text-white hover:text-white hover:bg-white/10 border border-white/30"
            onClick={() => { setClickedProvider("guest"); onGuestLogin(); }}
            disabled={isLoggingIn}
          >
            {clickedProvider === "guest" && isLoggingIn ? (
              <Loader2 className="w-5 h-5 animate-spin" />
            ) : (
              <ArrowRight className="w-5 h-5" />
            )}
            Try for free — no account needed
          </Button>

          <p className="text-xs text-emerald-200 mt-1">
            Free forever · No credit card required
          </p>
        </div>
      </section>

      {/* Footer */}
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
