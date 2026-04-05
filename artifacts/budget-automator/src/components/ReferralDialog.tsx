import { useState } from "react";
import { Copy, Check, Users, Gift, Star, ArrowRight } from "lucide-react";
import {
  Dialog,
  DialogContent,
  DialogTitle,
  DialogDescription,
} from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { useToast } from "@/hooks/use-toast";

interface ReferralInfo {
  referralCode: string;
  referralLink: string;
  totalReferred: number;
  totalConverted: number;
  totalRewarded: number;
}

interface ReferralDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  referralInfo: ReferralInfo | null;
  isLoading: boolean;
}

const HOW_IT_WORKS = [
  { label: "Share your unique referral link with a friend" },
  { label: "They sign up using your link" },
  { label: "When they upgrade to Pro, you earn one free month" },
];

const STAT_CONFIGS = [
  {
    key: "totalReferred" as const,
    label: "Referred",
    Icon: Users,
    bg: "bg-emerald-50",
    border: "border-emerald-100",
    iconBg: "bg-emerald-100",
    numColor: "text-emerald-700",
    labelColor: "text-emerald-600",
    iconColor: "text-emerald-500",
  },
  {
    key: "totalConverted" as const,
    label: "Converted",
    Icon: ArrowRight,
    bg: "bg-amber-50",
    border: "border-amber-100",
    iconBg: "bg-amber-100",
    numColor: "text-amber-700",
    labelColor: "text-amber-600",
    iconColor: "text-amber-500",
  },
  {
    key: "totalRewarded" as const,
    label: "Rewarded",
    Icon: Star,
    bg: "bg-violet-50",
    border: "border-violet-100",
    iconBg: "bg-violet-100",
    numColor: "text-violet-700",
    labelColor: "text-violet-600",
    iconColor: "text-violet-500",
  },
];

export function ReferralDialog({ open, onOpenChange, referralInfo, isLoading }: ReferralDialogProps) {
  const { toast } = useToast();
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    if (!referralInfo?.referralLink) return;
    try {
      await navigator.clipboard.writeText(referralInfo.referralLink);
      setCopied(true);
      toast({ title: "Copied!", description: "Referral link copied to clipboard." });
      setTimeout(() => setCopied(false), 2000);
    } catch {
      toast({ title: "Could not copy", description: "Please copy the link manually.", variant: "destructive" });
    }
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-[calc(100vw-2rem)] sm:max-w-md overflow-y-auto max-h-[90vh] p-0">
        <DialogTitle className="sr-only">Refer a Friend</DialogTitle>
        <DialogDescription className="sr-only">
          Share your referral link. When a friend signs up and upgrades to Pro, you earn one free month of Pro.
        </DialogDescription>
        <div className="relative overflow-hidden rounded-t-lg bg-gradient-to-br from-emerald-500 via-emerald-600 to-teal-700 px-6 pt-8 pb-10">
          <div className="absolute inset-0 opacity-10"
            style={{
              backgroundImage:
                "radial-gradient(circle at 20% 50%, white 1px, transparent 1px), radial-gradient(circle at 80% 20%, white 1px, transparent 1px)",
              backgroundSize: "40px 40px",
            }}
          />
          <div className="relative flex flex-col items-center text-center gap-3">
            <div className="flex items-center justify-center w-14 h-14 rounded-full bg-white/20 backdrop-blur-sm shadow-inner">
              <Gift className="w-7 h-7 text-white" />
            </div>
            <div>
              <h2 className="text-xl font-bold text-white leading-tight">Refer a Friend</h2>
              <p className="mt-1 text-emerald-100 text-sm leading-relaxed max-w-xs">
                Share your link and earn{" "}
                <span className="font-semibold text-white">1 free month of Pro</span>{" "}
                for every friend who upgrades.
              </p>
            </div>
          </div>
        </div>

        <div className="space-y-5 px-6 py-5">
          {isLoading ? (
            <div className="flex items-center justify-center h-20">
              <div className="w-5 h-5 border-2 border-emerald-500 border-t-transparent rounded-full animate-spin" />
            </div>
          ) : referralInfo ? (
            <>
              <div className="space-y-2">
                <label className="text-xs font-semibold text-muted-foreground uppercase tracking-wider block">
                  Your referral link
                </label>
                <div className="flex flex-col gap-2">
                  <div className="w-full px-3 py-2.5 rounded-lg bg-slate-50 border border-border/60 text-sm font-mono truncate text-muted-foreground">
                    {referralInfo.referralLink}
                  </div>
                  <Button
                    className="w-full gap-2 bg-emerald-600 hover:bg-emerald-700 text-white font-semibold"
                    onClick={handleCopy}
                  >
                    {copied ? (
                      <Check className="w-4 h-4" />
                    ) : (
                      <Copy className="w-4 h-4" />
                    )}
                    {copied ? "Link Copied!" : "Copy Referral Link"}
                  </Button>
                </div>
              </div>

              <div className="grid grid-cols-3 gap-1.5 sm:gap-2.5">
                {STAT_CONFIGS.map(({ key, label, Icon, bg, border, iconBg, numColor, labelColor, iconColor }) => (
                  <div
                    key={key}
                    className={`rounded-xl ${bg} ${border} border p-2 sm:p-3.5 flex flex-col items-center gap-1.5 sm:gap-2 min-w-0`}
                  >
                    <div className={`flex items-center justify-center w-7 h-7 sm:w-8 sm:h-8 rounded-full ${iconBg} flex-shrink-0`}>
                      <Icon className={`w-3.5 h-3.5 sm:w-4 sm:h-4 ${iconColor}`} />
                    </div>
                    <div className={`text-2xl sm:text-3xl font-extrabold leading-none ${numColor}`}>
                      {referralInfo[key]}
                    </div>
                    <div className={`text-[10px] sm:text-xs font-medium ${labelColor} text-center leading-tight`}>{label}</div>
                  </div>
                ))}
              </div>

              <div className="rounded-xl bg-slate-50 border border-border/50 p-4 space-y-3">
                <p className="font-semibold text-foreground text-xs uppercase tracking-wider">How it works</p>
                <ol className="space-y-2.5">
                  {HOW_IT_WORKS.map((step, i) => (
                    <li key={i} className="flex items-start gap-3">
                      <span className="flex-shrink-0 flex items-center justify-center w-6 h-6 rounded-full bg-emerald-100 text-emerald-700 text-xs font-bold leading-none">
                        {i + 1}
                      </span>
                      <span className="text-xs text-muted-foreground leading-relaxed pt-0.5">{step.label}</span>
                    </li>
                  ))}
                </ol>
              </div>
            </>
          ) : (
            <p className="text-sm text-muted-foreground text-center py-4">
              Could not load referral info. Please try again.
            </p>
          )}
        </div>
      </DialogContent>
    </Dialog>
  );
}
