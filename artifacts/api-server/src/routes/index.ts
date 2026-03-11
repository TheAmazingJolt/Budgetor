import { Router, type IRouter } from "express";
import healthRouter from "./health";
import budgetRouter from "./budget";

const router: IRouter = Router();

router.use(healthRouter);
router.use(budgetRouter);

export default router;
