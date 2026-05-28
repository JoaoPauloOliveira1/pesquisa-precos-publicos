import { consultarUmaLinha, consultarVariasLinhas } from "../db/index.js";
import { normalizarDataReferencia, normalizarRegime, normalizarTexto } from "../lib/normalize.js";

function montarContexto({ uf, dataReferencia, regime }) {
  return {
    uf: String(uf || "PE").trim().toUpperCase(),
    dataReferencia: normalizarDataReferencia(dataReferencia),
    regime: normalizarRegime(regime),
  };
}

export async function obterStatusGeral() {
  const referencia = await consultarUmaLinha(
    `select data_referencia, status, atualizado_em
       from sinapi_referencias
      order by data_referencia desc
      limit 1`
  );

  const ultimoSync = await consultarUmaLinha(
    `select tipo, status, referencia_encontrada, iniciado_em, finalizado_em, mensagem
       from sinapi_sync_runs
      order by iniciado_em desc
      limit 1`
  );

  const insumos = await consultarUmaLinha(`select count(*)::int as total from sinapi_insumos`);
  const composicoes = await consultarUmaLinha(`select count(*)::int as total from sinapi_composicoes`);

  return {
    referenciaAtual: referencia,
    ultimoSync,
    totais: {
      insumos: insumos?.total || 0,
      composicoes: composicoes?.total || 0,
    },
  };
}

export async function buscarInsumos({ q, uf, dataReferencia, regime, limit = 50, skip = 0 }) {
  const contexto = montarContexto({ uf, dataReferencia, regime });
  const termo = normalizarTexto(q);

  if (!termo || termo.length < 3) return [];

  return consultarVariasLinhas(
    `select
        codigo,
        descricao,
        unidade,
        uf,
        data_referencia,
        regime,
        preco_mediano
      from sinapi_insumos
      where uf = $1
        and data_referencia = $2
        and regime = $3
        and descricao_normalizada % $4
      order by similarity(descricao_normalizada, $4) desc, descricao asc
      offset $5
      limit $6`,
    [contexto.uf, contexto.dataReferencia, contexto.regime, termo, skip, limit]
  );
}

export async function buscarComposicoes({ q, uf, dataReferencia, regime, limit = 50, skip = 0 }) {
  const contexto = montarContexto({ uf, dataReferencia, regime });
  const termo = normalizarTexto(q);

  if (!termo || termo.length < 3) return [];

  return consultarVariasLinhas(
    `select
        codigo,
        descricao,
        unidade,
        uf,
        data_referencia,
        regime,
        custo_total
      from sinapi_composicoes
      where uf = $1
        and data_referencia = $2
        and regime = $3
        and descricao_normalizada % $4
      order by similarity(descricao_normalizada, $4) desc, descricao asc
      offset $5
      limit $6`,
    [contexto.uf, contexto.dataReferencia, contexto.regime, termo, skip, limit]
  );
}
