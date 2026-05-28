import { Router } from "express";
import { buscarComposicoes } from "../repositories/sinapiRepository.js";

export const composicoesRouter = Router();

composicoesRouter.get("/composicoes", async (req, res, next) => {
  try {
    const limit = Number.parseInt(req.query.limit || "50", 10);
    const skip = Number.parseInt(req.query.skip || "0", 10);

    const itens = await buscarComposicoes({
      q: req.query.q || "",
      uf: req.query.uf || "",
      dataReferencia: req.query.data_referencia || "",
      regime: req.query.regime || "",
      limit: Number.isFinite(limit) ? Math.min(limit, 100) : 50,
      skip: Number.isFinite(skip) ? Math.max(skip, 0) : 0,
    });

    res.json(itens);
  } catch (error) {
    next(error);
  }
});
