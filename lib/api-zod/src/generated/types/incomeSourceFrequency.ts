export type IncomeSourceFrequency =
  (typeof IncomeSourceFrequency)[keyof typeof IncomeSourceFrequency];

export const IncomeSourceFrequency = {
  weekly: "weekly",
  biweekly: "biweekly",
  monthly: "monthly",
  variable: "variable",
} as const;
