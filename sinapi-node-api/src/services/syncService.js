import { env } from "../config/env.js";

export function janelaEsperadaDePublicacao(data = new Date()) {
  const dia = data.getDate();
  return dia >= env.publicationCheckStartDay && dia <= env.publicationCheckEndDay;
}

export async function verificarReferenciaDisponivel() {
  return {
    fonte: env.caixaBaseUrl,
    janelaEsperada: janelaEsperadaDePublicacao(),
    referenciaDetectada: null,
    observacao:
      "Sincronizador ainda não implementado. Próximo passo: ler a página oficial da CAIXA, localizar a referência mensal e baixar o ZIP/XLSX oficial.",
  };
}

export async function iniciarSincronizacaoManual() {
  const diagnostico = await verificarReferenciaDisponivel();

  return {
    ok: true,
    iniciado: false,
    ...diagnostico,
  };
}
