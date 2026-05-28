import express from "express";
import { healthRouter } from "./routes/health.js";
import { statusRouter } from "./routes/status.js";
import { insumosRouter } from "./routes/insumos.js";
import { composicoesRouter } from "./routes/composicoes.js";
import { syncRouter } from "./routes/sync.js";
import { validarApiKey } from "./lib/auth.js";

export function criarApp() {
  const app = express();

  app.use(express.json({ limit: "2mb" }));

  app.use(healthRouter);
  app.use(statusRouter);
  app.use(validarApiKey);
  app.use(insumosRouter);
  app.use(composicoesRouter);
  app.use(syncRouter);

  app.use((req, res) => {
    res.status(404).json({
      erro: "Rota não encontrada.",
      rota: req.originalUrl,
    });
  });

  app.use((error, _req, res, _next) => {
    console.error("Erro na API SINAPI:", error);
    res.status(500).json({
      erro: "Erro interno na API SINAPI.",
      detalhe: error.message,
    });
  });

  return app;
}
