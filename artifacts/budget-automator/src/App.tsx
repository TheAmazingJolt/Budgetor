import { Switch, Route, Router as WouterRouter, useLocation } from "wouter";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import { motion } from "framer-motion";
import { Calculator, Settings } from "lucide-react";

import { Setup } from "@/pages/setup";
import { Budgets } from "@/pages/budgets";
import NotFound from "@/pages/not-found";

const queryClient = new QueryClient();

function Navigation() {
  const [location, setLocation] = useLocation();

  const navItems = [
    { path: "/", label: "Setup & Bills", icon: Settings },
    { path: "/budgets", label: "My Budgets", icon: Calculator },
  ];

  return (
    <header className="sticky top-0 z-50 w-full backdrop-blur-xl bg-white/70 border-b border-border/50 shadow-sm">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-20 flex items-center justify-between">
        
        <div className="flex items-center gap-3">
          <div className="bg-gradient-to-br from-primary to-emerald-600 p-2.5 rounded-xl shadow-lg shadow-primary/30">
            <Calculator className="w-6 h-6 text-white" />
          </div>
          <div>
            <h1 className="font-display text-2xl font-bold text-foreground leading-none tracking-tight">Budgex</h1>
            <p className="text-xs text-emerald-600 font-medium tracking-widest uppercase">Automator</p>
          </div>
        </div>

        <nav className="flex gap-2 bg-slate-100/80 p-1.5 rounded-2xl border border-border/60">
          {navItems.map((item) => {
            const isActive = location === item.path;
            const Icon = item.icon;
            
            return (
              <button
                key={item.path}
                onClick={() => setLocation(item.path)}
                className={`
                  relative flex items-center gap-2 px-5 py-2.5 rounded-xl text-sm font-semibold transition-colors duration-200
                  ${isActive ? 'text-primary-foreground' : 'text-muted-foreground hover:text-foreground hover:bg-black/5'}
                `}
              >
                {isActive && (
                  <motion.div
                    layoutId="nav-pill"
                    className="absolute inset-0 bg-primary rounded-xl shadow-sm"
                    transition={{ type: "spring", stiffness: 350, damping: 30 }}
                  />
                )}
                <span className="relative z-10 flex items-center gap-2">
                  <Icon className="w-4 h-4" />
                  {item.label}
                </span>
              </button>
            );
          })}
        </nav>
      </div>
    </header>
  );
}

function Router() {
  return (
    <div className="relative min-h-screen pb-24">
      {/* Abstract Background Elements */}
      <div className="fixed inset-0 overflow-hidden pointer-events-none -z-10">
        <div className="absolute -top-[20%] -left-[10%] w-[50%] h-[50%] rounded-full bg-emerald-100/40 blur-[120px]" />
        <div className="absolute top-[40%] -right-[10%] w-[40%] h-[60%] rounded-full bg-teal-100/30 blur-[100px]" />
        {/* We requested an image but also added css gradients for fallback/layering. Using the requested image: */}
        <img 
          src={`${import.meta.env.BASE_URL}images/hero-bg.png`} 
          alt="" 
          className="absolute inset-0 w-full h-full object-cover opacity-10 mix-blend-overlay"
        />
      </div>

      <Navigation />
      
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-10">
        <Switch>
          <Route path="/" component={Setup} />
          <Route path="/budgets" component={Budgets} />
          <Route component={NotFound} />
        </Switch>
      </main>
    </div>
  );
}

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <WouterRouter base={import.meta.env.BASE_URL.replace(/\/$/, "")}>
          <Router />
        </WouterRouter>
        <Toaster />
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;
