import { Router, type IRouter } from "express";
import healthRouter from "./health";
import budgetRouter from "./budget";
import googleAuthRouter from "./google-auth";
import sheetsRouter from "./sheets";

const router: IRouter = Router();

router.use(healthRouter);
router.use(budgetRouter);
router.use(googleAuthRouter);
router.use(sheetsRouter);

export default router;
