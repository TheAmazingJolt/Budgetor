import { useState, useCallback, useEffect } from "react";
import { Bug, Loader2, Send } from "lucide-react";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { useCreateBugReport } from "@workspace/api-client-react";
import { toast } from "@/hooks/use-toast";

const APP_VERSION = import.meta.env.VITE_APP_VERSION ?? "dev";

interface BugReportDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  prefillDescription?: string;
  prefillErrorMessage?: string | null;
  prefillErrorStack?: string | null;
  userName?: string | null;
  userEmail?: string | null;
}

export function BugReportDialog({
  open,
  onOpenChange,
  prefillDescription = "",
  prefillErrorMessage,
  prefillErrorStack,
  userName,
  userEmail,
}: BugReportDialogProps) {
  const [description, setDescription] = useState(prefillDescription);

  useEffect(() => {
    if (open) {
      setDescription(prefillDescription);
    }
  }, [open, prefillDescription]);

  const mutation = useCreateBugReport();

  const handleOpenChange = useCallback(
    (next: boolean) => {
      if (!mutation.isPending) {
        onOpenChange(next);
      }
    },
    [mutation.isPending, onOpenChange],
  );

  const handleSubmit = useCallback(async () => {
    const trimmed = description.trim();
    if (!trimmed) return;

    try {
      await mutation.mutateAsync({
        data: {
          description: trimmed,
          errorMessage: prefillErrorMessage ?? null,
          errorStack: prefillErrorStack ?? null,
          pageUrl: window.location.href,
          userAgent: navigator.userAgent,
          appVersion: APP_VERSION,
        },
      });
      toast({
        title: "Bug report submitted",
        description: "Thank you! We'll look into this.",
      });
      onOpenChange(false);
    } catch {
      toast({
        title: "Failed to submit report",
        description: "Your report was not lost. Please try again.",
        variant: "destructive",
      });
    }
  }, [description, mutation, onOpenChange, prefillErrorMessage, prefillErrorStack]);

  const hasError = !!(prefillErrorMessage || prefillErrorStack);

  return (
    <Dialog open={open} onOpenChange={handleOpenChange}>
      <DialogContent className="sm:max-w-md">
        <DialogHeader>
          <DialogTitle className="flex items-center gap-2">
            <Bug className="w-4 h-4 text-amber-500" />
            Report a Bug
          </DialogTitle>
          <DialogDescription>
            Describe what happened and we'll look into it.
            {(userName || userEmail) && (
              <span className="block mt-1 text-xs text-muted-foreground">
                Submitting as {userName ?? userEmail}
              </span>
            )}
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-4 py-2">
          <div className="space-y-1.5">
            <Label htmlFor="bug-description">What went wrong?</Label>
            <Textarea
              id="bug-description"
              placeholder="Describe the issue as clearly as possible..."
              className="min-h-[120px] resize-none"
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              disabled={mutation.isPending}
            />
          </div>

          {hasError && (
            <div className="rounded-lg bg-red-50 border border-red-200 p-3 space-y-1">
              <p className="text-xs font-medium text-red-700">Error details (included automatically)</p>
              {prefillErrorMessage && (
                <p className="text-xs text-red-600 font-mono break-all">{prefillErrorMessage}</p>
              )}
            </div>
          )}
        </div>

        <DialogFooter>
          <Button
            variant="outline"
            onClick={() => handleOpenChange(false)}
            disabled={mutation.isPending}
          >
            Cancel
          </Button>
          <Button
            onClick={handleSubmit}
            disabled={mutation.isPending || description.trim().length === 0}
            className="gap-2"
          >
            {mutation.isPending ? (
              <>
                <Loader2 className="w-3.5 h-3.5 animate-spin" />
                Submitting…
              </>
            ) : (
              <>
                <Send className="w-3.5 h-3.5" />
                Submit Report
              </>
            )}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
