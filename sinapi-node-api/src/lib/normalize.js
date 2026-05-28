export function normalizarTexto(valor = "") {
  return String(valor || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/([a-zA-Z])(\d)/g, "$1 $2")
    .replace(/(\d)([a-zA-Z])/g, "$1 $2")
    .replace(/[^\w\s/.-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

export function normalizarDataReferencia(valor = "") {
  const texto = String(valor || "").trim();
  if (/^\d{4}-\d{2}$/.test(texto)) return texto;
  return "";
}

export function normalizarRegime(valor = "") {
  const texto = String(valor || "NAO_DESONERADO").trim().toUpperCase();
  return texto === "DESONERADO" ? "DESONERADO" : "NAO_DESONERADO";
}
