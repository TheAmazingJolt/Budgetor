import { Router, type IRouter } from "express";
import healthRouter from "./health";
import budgetRouter from "./budget";
import googleAuthRouter from "./google-auth";
import sheetsRouter from "./sheets";
import userAuthRouter, { attachUser } from "./user-auth";
import savedBudgetsRouter from "./saved-budgets";

const router: IRouter = Router();

router.use(attachUser);
router.use(healthRouter);
router.use(budgetRouter);
router.use(googleAuthRouter);
router.use(sheetsRouter);
router.use(userAuthRouter);
router.use(savedBudgetsRouter);

export default router;
