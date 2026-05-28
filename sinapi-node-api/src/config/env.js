import "dotenv/config";

function numeroInteiro(valor, fallback) {
  const numero = Number.parseInt(valor, 10);
  return Number.isFinite(numero) ? numero : fallback;
}

export const env = {
  port: numeroInteiro(process.env.PORT, 3100),
  host: process.env.HOST || "0.0.0.0",
  nodeEnv: process.env.NODE_ENV || "development",
  databaseUrl: process.env.DATABASE_URL || "",
  apiKey: process.env.SINAPI_API_KEY || "",
  syncToken: process.env.SINAPI_SYNC_TOKEN || "",
  caixaBaseUrl:
    process.env.SINAPI_CAIXA_BASE_URL ||
    "https://www.caixa.gov.br/poder-publico/modernizacao-gestao/sinapi/Paginas/default.aspx",
  publicationCheckStartDay: numeroInteiro(process.env.SINAPI_PUBLICATION_CHECK_START_DAY, 5),
  publicationCheckEndDay: numeroInteiro(process.env.SINAPI_PUBLICATION_CHECK_END_DAY, 15),
  defaultUf: (process.env.SINAPI_DEFAULT_UF || "PE").toUpperCase(),
  defaultRegime: process.env.SINAPI_DEFAULT_REGIME || "NAO_DESONERADO",
};

export function assertEnv() {
  const faltando = [];

  if (!env.databaseUrl) faltando.push("DATABASE_URL");
  if (!env.apiKey) faltando.push("SINAPI_API_KEY");

  if (faltando.length) {
    throw new Error(`Variáveis obrigatórias ausentes: ${faltando.join(", ")}`);
  }
}
