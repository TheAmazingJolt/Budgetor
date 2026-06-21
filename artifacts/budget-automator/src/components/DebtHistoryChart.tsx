import { LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer } from "recharts";

interface Snapshot {
  date: string;
  balance: number;
  note?: string;
}

export function DebtHistoryChart({
  snapshots,
  currentBalance,
}: {
  snapshots: Snapshot[];
  currentBalance: number;
}) {
  const sorted = [...snapshots].sort((a, b) => a.date.localeCompare(b.date));
  const startBalance = sorted[0]?.balance ?? currentBalance;
  const totalPaid = Math.max(0, startBalance - currentBalance);
  const startDate = sorted[0]?.date ?? "";
  const fmtUSD = (v: number) =>
    v.toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });

  if (sorted.length < 2) {
    return (
      <p className="text-xs text-muted-foreground py-2 text-center">
        Log a payment to start tracking your progress.
      </p>
    );
  }

  return (
    <div className="space-y-2">
      <div style={{ height: 110 }}>
        <ResponsiveContainer width="100%" height="100%">
          <LineChart data={sorted} margin={{ top: 4, right: 8, left: 0, bottom: 0 }}>
            <XAxis
              dataKey="date"
              fontSize={10}
              tickFormatter={(d) => {
                const dt = new Date(d + "T00:00:00");
                return `${dt.getMonth() + 1}/${dt.getDate()}`;
              }}
              interval="preserveStartEnd"
            />
            <YAxis
              fontSize={10}
              tickFormatter={(v) => `$${Math.round(v / 1000)}k`}
              width={42}
              domain={["auto", "auto"]}
            />
            <Tooltip
              formatter={(v: number) => [fmtUSD(v), "Balance"]}
              labelFormatter={(d) =>
                new Date(d + "T00:00:00").toLocaleDateString("en-US", {
                  month: "short",
                  day: "numeric",
                  year: "numeric",
                })
              }
            />
            <Line
              type="monotone"
              dataKey="balance"
              stroke="#8b5cf6"
              strokeWidth={2}
              dot={{ r: 3, fill: "#8b5cf6" }}
              isAnimationActive={false}
            />
          </LineChart>
        </ResponsiveContainer>
      </div>
      {totalPaid > 0 && (
        <p className="text-xs text-emerald-600 font-medium">
          {fmtUSD(totalPaid)} paid down since{" "}
          {new Date(startDate + "T00:00:00").toLocaleDateString("en-US", {
            month: "short",
            year: "numeric",
          })}
        </p>
      )}
    </div>
  );
}
