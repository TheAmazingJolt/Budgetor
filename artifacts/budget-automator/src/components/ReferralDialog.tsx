import { useState } from "react";
import { Copy, Check, Users, Gift } from "lucide-react";
import {
  Dialog,
  DialogContent,
  DialogHeader,
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
      <DialogContent className="max-w-[calc(100vw-2rem)] sm:max-w-md overflow-hidden">
        <DialogHeader>
          <DialogTitle className="flex items-center gap-2">
            <Gift className="w-5 h-5 text-emerald-500" />
            Refer a friend
          </DialogTitle>
          <DialogDescription>
            Share your referral link. When a friend signs up and upgrades to Pro, you earn one free month of Pro.
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-5 py-2">
          {isLoading ? (
            <div className="flex items-center justify-center h-20">
              <div className="w-5 h-5 border-2 border-emerald-500 border-t-transparent rounded-full animate-spin" />
            </div>
          ) : referralInfo ? (
            <>
              <div>
                <label className="text-xs font-medium text-muted-foreground uppercase tracking-wide mb-1.5 block">
                  Your referral link
                </label>
                <div className="flex items-center gap-2">
                  <div className="flex-1 min-w-0 px-3 py-2 rounded-lg bg-slate-50 border border-border/60 text-sm font-mono truncate text-muted-foreground">
                    {referralInfo.referralLink}
                  </div>
                  <Button
                    size="sm"
                    variant="outline"
                    className="shrink-0 gap-1.5"
                    onClick={handleCopy}
                  >
                    {copied ? (
                      <Check className="w-3.5 h-3.5 text-emerald-500" />
                    ) : (
                      <Copy className="w-3.5 h-3.5" />
                    )}
                    {copied ? "Copied" : "Copy"}
                  </Button>
                </div>
              </div>

              <div className="grid grid-cols-3 gap-3 overflow-hidden">
                <div className="rounded-xl bg-emerald-50 border border-emerald-100 p-4 text-center">
                  <div className="text-2xl font-bold text-emerald-700">{referralInfo.totalReferred}</div>
                  <div className="text-xs text-emerald-600 mt-1 flex items-center justify-center gap-1">
                    <Users className="w-3 h-3" />
                    Referred
                  </div>
                </div>
                <div className="rounded-xl bg-amber-50 border border-amber-100 p-4 text-center">
                  <div className="text-2xl font-bold text-amber-700">{referralInfo.totalConverted}</div>
                  <div className="text-xs text-amber-600 mt-1 flex items-center justify-center gap-1">
                    <Gift className="w-3 h-3" />
                    Converted
                  </div>
                </div>
                <div className="rounded-xl bg-violet-50 border border-violet-100 p-4 text-center">
                  <div className="text-2xl font-bold text-violet-700">{referralInfo.totalRewarded}</div>
                  <div className="text-xs text-violet-600 mt-1 flex items-center justify-center gap-1">
                    <Check className="w-3 h-3" />
                    Rewarded
                  </div>
                </div>
              </div>

              <div className="rounded-xl bg-slate-50 border border-border/50 p-4 text-sm text-muted-foreground space-y-1.5">
                <p className="font-medium text-foreground text-xs uppercase tracking-wide">How it works</p>
                <ol className="list-decimal list-inside space-y-1 text-xs">
                  <li>Share your link with friends</li>
                  <li>They sign up using your link</li>
                  <li>When they upgrade to Pro, you get one free month</li>
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
