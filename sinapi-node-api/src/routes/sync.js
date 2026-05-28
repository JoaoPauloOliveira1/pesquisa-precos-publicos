import { Router } from "express";
import { validarSyncToken } from "../lib/auth.js";
import { iniciarSincronizacaoManual } from "../services/syncService.js";

export const syncRouter = Router();

syncRouter.post("/sync/manual", validarSyncToken, async (_req, res, next) => {
  try {
    const resultado = await iniciarSincronizacaoManual();
    res.status(202).json(resultado);
  } catch (error) {
    next(error);
  }
});
