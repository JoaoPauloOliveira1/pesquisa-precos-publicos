import { Router } from "express";
import { obterStatusGeral } from "../repositories/sinapiRepository.js";
import { verificarReferenciaDisponivel } from "../services/syncService.js";

export const statusRouter = Router();

statusRouter.get("/status", async (_req, res, next) => {
  try {
    const [statusBanco, statusFonte] = await Promise.all([
      obterStatusGeral(),
      verificarReferenciaDisponivel(),
    ]);

    res.json({
      ok: true,
      banco: statusBanco,
      fonte: statusFonte,
    });
  } catch (error) {
    next(error);
  }
});
