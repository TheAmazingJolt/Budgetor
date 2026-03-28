import { Router, type IRouter } from "express";
import healthRouter from "./health";
import budgetRouter from "./budget";
import googleAuthRouter from "./google-auth";
import sheetsRouter from "./sheets";
import microsoftAuthRouter from "./microsoft-auth";
import excelRouter from "./excel";
import userAuthRouter, { attachUser } from "./user-auth";
import savedBudgetsRouter from "./saved-budgets";
import savingsContributionsRouter from "./savings-contributions";
import weeklyCheckinsRouter from "./weekly-checkins";

const router: IRouter = Router();

router.use(attachUser);
router.use(healthRouter);
router.use(budgetRouter);
router.use(googleAuthRouter);
router.use(sheetsRouter);
router.use(microsoftAuthRouter);
router.use(excelRouter);
router.use(userAuthRouter);
router.use(savedBudgetsRouter);
router.use(savingsContributionsRouter);
router.use(weeklyCheckinsRouter);

export default router;
