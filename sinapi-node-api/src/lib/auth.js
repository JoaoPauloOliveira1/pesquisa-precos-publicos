import { env } from "../config/env.js";

export function validarApiKey(req, res, next) {
  const chave = req.header("X-API-KEY");

  if (!env.apiKey || chave !== env.apiKey) {
    return res.status(401).json({
      erro: "Chave de API inválida.",
    });
  }

  next();
}

export function validarSyncToken(req, res, next) {
  const token = req.header("X-SYNC-TOKEN");

  if (!env.syncToken || token !== env.syncToken) {
    return res.status(401).json({
      erro: "Token de sincronização inválido.",
    });
  }

  next();
}
