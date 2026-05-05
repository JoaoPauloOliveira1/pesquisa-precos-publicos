const express = require("express");
const cors = require("cors");
const PDFDocument = require("pdfkit");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");
const XLSX = require("xlsx");
const sevenZip = require("7zip-bin");

const envPath = path.join(__dirname, ".env");
if (fs.existsSync(envPath)) {
  for (const line of fs.readFileSync(envPath, "utf8").split(/\r?\n/)) {
    const match = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*)\s*$/i);
    if (!match || process.env[match[1]]) continue;
    process.env[match[1]] = match[2].replace(/^["']|["']$/g, "");
  }
}

const app = express();
const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || "0.0.0.0";

app.use(cors());
app.use(express.json({ limit: "2mb" }));
app.use(express.static("public"));

const BASE = "https://dadosabertos.compras.gov.br";
const PNCP_BASE = "https://pncp.gov.br";
const CNBS_BASE = "https://cnbs.estaleiro.serpro.gov.br/cnbs-api";
const DNIT_SICRO_NORDESTE_URL =
  "https://www.gov.br/dnit/pt-br/assuntos/planejamento-e-pesquisa/custos-referenciais/sistemas-de-custos/sicro/relatorios/relatorios-sicro/nordeste/nordeste";
const ORSE_BASE_DADOS_URL = "https://orse-portal.cehop.se.gov.br/base-de-dados/";
const ORSE_INSUMOS_URL = "https://orse.cehop.se.gov.br/insumosargumento.asp";
const PE_INTEGRADO_ARP_BASE = "https://www.peintegrado.pe.gov.br/Portal/Pages/AtasRegistroPreco.aspx";
const SINAPI_API_URL = process.env.SINAPI_API_URL || "";
const SINAPI_API_KEY = process.env.SINAPI_API_KEY || "";
const DEBUG_LOGS = process.env.DEBUG_LOGS === "1";
const IS_RENDER = process.env.RENDER === "true";
const CURL_BIN = process.platform === "win32" ? "curl.exe" : "curl";
const AUTO_UPDATE_BASES =
  process.env.AUTO_UPDATE_BASES === "1" ||
  (process.env.AUTO_UPDATE_BASES !== "0" && !IS_RENDER && process.env.NODE_ENV !== "production");
const SINAPI_HABILITADO = process.env.ENABLE_SINAPI !== "0" && Boolean(SINAPI_API_URL);
const DATA_DIR = path.join(__dirname, "data");
const BASES_DIR = path.join(DATA_DIR, "bases");
const SICRO_DOWNLOAD_DIR = path.join(BASES_DIR, "sicro-downloads");
const SICRO_EXTRACT_DIR = path.join(BASES_DIR, "sicro-extraidos");
const SICRO_LITE_JSON_PATH = path.join(BASES_DIR, "sicro-lite.json");
const SICRO_JSON_PATH = path.join(BASES_DIR, "sicro.json");
const SICRO_MANIFEST_PATH = path.join(SICRO_DOWNLOAD_DIR, "manifest.json");
const ORSE_DOWNLOAD_DIR = path.join(BASES_DIR, "orse-downloads");
const ORSE_EXTRACT_DIR = path.join(BASES_DIR, "orse-extraidos");
const ORSE_JSON_PATH = path.join(BASES_DIR, "orse.json");
const ORSE_MANIFEST_PATH = path.join(ORSE_DOWNLOAD_DIR, "manifest.json");
const CACHE_PATH = path.join(DATA_DIR, "cache-consultas.json");
const CACHE_TTL_MS = 1000 * 60 * 60 * 24 * 7;
const BUSCA_ITENS_FAILURE_TTL_MS = 1000 * 60 * 2;
const SICRO_CHECK_INTERVAL_MS = 1000 * 60 * 60 * 24;
const ORSE_CHECK_INTERVAL_MS = 1000 * 60 * 60 * 24;

let cachePdmMaterial = null;
let cacheServico = null;
const cachePncpPorIdCompra = new Map();
const cacheDetalheContratacaoPncp = new Map();
const cacheResultadoItemPncp = new Map();
const cacheItensContratacaoPncp = new Map();
const cacheFalhasBuscaItens = new Map();
let cacheConsultas = carregarCacheConsultas();

function debugLog(...args) {
  if (DEBUG_LOGS) console.log(...args);
}

function debugWarn(...args) {
  if (DEBUG_LOGS) console.warn(...args);
}

function limitarTimeout(restanteMs, maximoMs) {
  return Math.max(1000, Math.min(maximoMs, restanteMs));
}

function carregarCacheConsultas() {
  try {
    if (!fs.existsSync(CACHE_PATH)) return { itens: {}, precos: {} };
    const dados = JSON.parse(fs.readFileSync(CACHE_PATH, "utf8"));
    return {
      itens: dados.itens && typeof dados.itens === "object" ? dados.itens : {},
      precos: dados.precos && typeof dados.precos === "object" ? dados.precos : {},
    };
  } catch (error) {
    debugWarn("Não foi possível carregar cache local:", error.message);
    return { itens: {}, precos: {} };
  }
}

function salvarCacheConsultas() {
  try {
    fs.mkdirSync(DATA_DIR, { recursive: true });
    fs.writeFileSync(CACHE_PATH, JSON.stringify(cacheConsultas, null, 2));
  } catch (error) {
    debugWarn("Não foi possível salvar cache local:", error.message);
  }
}

function chaveCache(...partes) {
  return partes.map((parte) => normalizarTexto(parte)).join("|");
}

function obterCache(secao, chave) {
  const entrada = cacheConsultas[secao]?.[chave];
  if (!entrada) return null;

  return {
    ...entrada,
    idadeMs: Date.now() - Date.parse(entrada.atualizadoEm || 0),
  };
}

function gravarCache(secao, chave, payload) {
  cacheConsultas[secao][chave] = {
    atualizadoEm: new Date().toISOString(),
    payload,
  };
  salvarCacheConsultas();
}

function carregarManifestSicro() {
  try {
    if (!fs.existsSync(SICRO_MANIFEST_PATH)) return {};
    return JSON.parse(fs.readFileSync(SICRO_MANIFEST_PATH, "utf8"));
  } catch (error) {
    debugWarn("Não foi possível carregar manifesto SICRO:", error.message);
    return {};
  }
}

function salvarManifestSicro(manifest) {
  fs.mkdirSync(SICRO_DOWNLOAD_DIR, { recursive: true });
  fs.writeFileSync(SICRO_MANIFEST_PATH, JSON.stringify(manifest, null, 2));
}

function carregarManifestOrse() {
  try {
    if (!fs.existsSync(ORSE_MANIFEST_PATH)) return {};
    return JSON.parse(fs.readFileSync(ORSE_MANIFEST_PATH, "utf8"));
  } catch (error) {
    debugWarn("Não foi possível carregar manifesto ORSE:", error.message);
    return {};
  }
}

function salvarManifestOrse(manifest) {
  fs.mkdirSync(ORSE_DOWNLOAD_DIR, { recursive: true });
  fs.writeFileSync(ORSE_MANIFEST_PATH, JSON.stringify(manifest, null, 2));
}

function limparErroProcesso(mensagem) {
  const texto = String(mensagem || "").trim();
  const linhas = texto
    .split(/\r?\n/)
    .map((linha) => linha.trim())
    .filter(Boolean)
    .filter((linha) => !linha.includes("% Total") && !linha.includes("Dload") && !/^\d+\s+\d+/.test(linha));

  const limpo = linhas.join(" ").replace(/\s+/g, " ").trim();
  return limpo || texto.slice(0, 300) || "Falha ao executar comando externo.";
}

const SINONIMOS = {
  "caixa de passagem": ["caixa passagem", "caixa inspecao", "caixa eletrica", "caixa derivacao"],
  "cimento": ["cimento", "cimento portland"],
  "ar condicionado": ["ar condicionado", "condicionador ar", "aparelho ar condicionado"],
  "cabo de cobre nu": [
    "cabo cobre nu",
    "cobre nu",
    "condutor cobre nu",
    "fio cobre nu",
    "cabo para raios cobre nu",
  ],
  "eletroduto": ["eletroduto", "conduite", "conduíte"],
};

const STOP_WORDS = new Set(["a", "o", "as", "os", "de", "da", "do", "das", "dos", "e", "em", "para"]);

const CLASSES_MATERIAL_RELACIONADAS = [
  {
    codigoClasse: 6145,
    termos: ["cabo", "cabinho", "fio", "condutor", "eletrico", "cobre"],
  },
];

const FONTES_LOCAIS = {
  sicro: {
    nome: "SICRO",
    arquivos: ["sicro-lite.json", "sicro-lite.csv", "sicro.json", "sicro.csv"],
    origem: "SICRO",
  },
  orse: {
    nome: "ORSE-SE",
    arquivos: ["orse.json", "orse.csv"],
    origem: "ORSE-SE",
  },
};

const UF_NORDESTE_SICRO = {
  AL: "ALAGOAS",
  BA: "BAHIA",
  CE: "CEARÁ",
  MA: "MARANHÃO",
  PB: "PARAÍBA",
  PE: "PERNAMBUCO",
  PI: "PIAUÍ",
  RN: "RIO GRANDE DO NORTE",
  SE: "SERGIPE",
};

const FONTES_EXTERNAS = {
  peintegrado: {
    nome: "Pernambuco Integrado",
    origem: "Pernambuco Integrado",
    link: PE_INTEGRADO_ARP_BASE,
  },
};

function normalizarTexto(texto) {
  return String(texto || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\w\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function gerarTermosBusca(termo) {
  const base = normalizarTexto(termo);
  const extras = SINONIMOS[base] || [];
  return [base, ...extras.map(normalizarTexto)].filter(Boolean);
}

function palavrasRelevantes(termoBusca) {
  return normalizarTexto(termoBusca)
    .split(" ")
    .filter((p) => p && !STOP_WORDS.has(p));
}

function normalizarValor(valor) {
  if (valor === null || valor === undefined || valor === "") return 0;
  let texto = String(valor).trim().replace(/\s/g, "");

  if (texto.includes(",") && texto.includes(".")) {
    const ultimaVirgula = texto.lastIndexOf(",");
    const ultimoPonto = texto.lastIndexOf(".");
    if (ultimaVirgula > ultimoPonto) {
      texto = texto.replace(/\./g, "").replace(",", ".");
    } else {
      texto = texto.replace(/,/g, "");
    }
  } else if (texto.includes(",")) {
    texto = texto.replace(",", ".");
  }

  return Number(texto);
}

function parseCsvSimples(texto) {
  const linhas = String(texto || "").split(/\r?\n/).filter((linha) => linha.trim());
  if (linhas.length < 2) return [];

  const separador = linhas[0].includes(";") ? ";" : ",";
  const headers = linhas[0].split(separador).map((h) => normalizarTexto(h));

  return linhas.slice(1).map((linha) => {
    const valores = linha.split(separador);
    return headers.reduce((acc, header, index) => {
      acc[header] = valores[index] || "";
      return acc;
    }, {});
  });
}

function parseCsvComDelimitador(texto) {
  const linhas = String(texto || "").split(/\r?\n/).filter((linha) => linha.trim());
  if (!linhas.length) return [];

  const separador = linhas[0].includes(";") ? ";" : ",";
  return linhas.map((linha) => linha.split(separador).map((valor) => valor.trim()));
}

function primeiroCampo(obj, nomes) {
  const entradasNormalizadas = Object.entries(obj).map(([chave, valor]) => [normalizarTexto(chave), valor]);

  for (const nome of nomes) {
    const valor = obj[nome] ?? obj[normalizarTexto(nome)];
    if (valor !== undefined && valor !== null && String(valor).trim() !== "") return valor;

    const nomeNormalizado = normalizarTexto(nome);
    const encontrado = entradasNormalizadas.find(([chave, valorEntrada]) =>
      chave.includes(nomeNormalizado) &&
      valorEntrada !== undefined &&
      valorEntrada !== null &&
      String(valorEntrada).trim() !== ""
    );

    if (encontrado) return encontrado[1];
  }

  return "";
}

function normalizarRegistroBaseLocal(item, fonte) {
  return {
    origem: fonte.origem,
    tipoSinapi: primeiroCampo(item, ["tipo", "categoria", "classe"]) || "",
    codigo: primeiroCampo(item, ["codigo", "código", "cod", "item", "referencia"]),
    descricao: primeiroCampo(item, ["descricao", "descrição", "nome", "servico", "serviço", "insumo"]),
    unidade: primeiroCampo(item, ["unidade", "un", "und"]),
    precoUnitario: primeiroCampo(item, ["preco", "preço", "valor", "custo", "custo_unitario", "custo unitario"]),
    dataReferencia: primeiroCampo(item, ["data", "referencia", "referência", "mes", "mês"]),
    uf: primeiroCampo(item, ["uf", "estado"]),
    regime: primeiroCampo(item, ["regime"]),
    link: primeiroCampo(item, ["link", "url"]),
  };
}

function detectarLinhaCabecalho(linhas) {
  let melhor = { index: -1, score: 0 };

  linhas.slice(0, 40).forEach((linha, index) => {
    const texto = linha.map(normalizarTexto);
    const score =
      (texto.some((celula) => celula.includes("codigo") || celula.includes("cod")) ? 2 : 0) +
      (texto.some((celula) => celula.includes("descricao") || celula.includes("discriminacao")) ? 3 : 0) +
      (texto.some((celula) => celula.includes("unidade") || celula === "un") ? 1 : 0) +
      (texto.some((celula) => celula.includes("preco") || celula.includes("custo") || celula.includes("valor")) ? 2 : 0);

    if (score > melhor.score) melhor = { index, score };
  });

  return melhor.score >= 4 ? melhor.index : -1;
}

function normalizarTabelaPlanilha(linhas, contexto) {
  const headerIndex = detectarLinhaCabecalho(linhas);
  if (headerIndex === -1) return [];

  const headers = linhas[headerIndex].map((header) => normalizarTexto(header));
  const registros = [];
  const fonte = FONTES_LOCAIS[contexto.fonteId] || FONTES_LOCAIS.sicro;

  for (const linha of linhas.slice(headerIndex + 1)) {
    const obj = {};
    headers.forEach((header, index) => {
      if (header) obj[header] = linha[index] || "";
    });

    const registro = normalizarRegistroBaseLocal(
      {
        ...obj,
        uf: contexto.uf,
        referencia: contexto.dataReferencia,
        tipo: contexto.tipo,
      },
      fonte
    );

    const valor = normalizarValor(registro.precoUnitario);
    if (registro.codigo && registro.descricao && !Number.isNaN(valor) && valor > 0) {
      registros.push({
        ...registro,
        arquivo: contexto.arquivo,
        planilha: contexto.planilha,
        estado: contexto.estado,
      });
    }
  }

  return registros;
}

function carregarBaseLocal(fonteId) {
  const fonte = FONTES_LOCAIS[fonteId];
  if (!fonte) return null;

  for (const arquivo of fonte.arquivos) {
    const caminho = path.join(BASES_DIR, arquivo);
    if (!fs.existsSync(caminho)) continue;

    const conteudo = fs.readFileSync(caminho, "utf8");
    const bruto = arquivo.endsWith(".json") ? JSON.parse(conteudo) : parseCsvSimples(conteudo);
    const lista = Array.isArray(bruto) ? bruto : bruto.items || bruto.registros || [];

    return lista.map((item) => normalizarRegistroBaseLocal(item, fonte));
  }

  return null;
}

function executarArquivo(comando, args) {
  return new Promise((resolve, reject) => {
    execFile(comando, args, { windowsHide: true }, (error, stdout, stderr) => {
      if (error) {
        reject(new Error(limparErroProcesso(stderr || stdout || error.message)));
        return;
      }

      resolve({ stdout, stderr });
    });
  });
}

async function consultarTextoComCurl(url, timeoutSegundos = 60) {
  const { stdout } = await executarArquivo(CURL_BIN, [
    "-sS",
    "-L",
    "-A",
    "Mozilla/5.0",
    "--connect-timeout",
    "20",
    "--max-time",
    String(timeoutSegundos),
    url.toString(),
  ]);

  if (!stdout || !stdout.trim()) {
    throw new Error("curl não retornou conteúdo.");
  }

  return stdout;
}

async function baixarArquivoComCurl(url, destino, timeoutSegundos = 240) {
  fs.mkdirSync(path.dirname(destino), { recursive: true });
  await executarArquivo(CURL_BIN, [
    "-sS",
    "--http1.1",
    "-L",
    "-A",
    "Mozilla/5.0",
    "-e",
    ORSE_BASE_DADOS_URL,
    "--connect-timeout",
    "20",
    "--max-time",
    String(timeoutSegundos),
    "-o",
    destino,
    url.toString(),
  ]);

  if (!fs.existsSync(destino) || fs.statSync(destino).size === 0) {
    throw new Error("Download não gerou arquivo.");
  }

  return fs.statSync(destino).size;
}

async function baixarArquivoComCurlResumivel(url, destino, tamanhoEsperado = 0, tentativas = 8) {
  fs.mkdirSync(path.dirname(destino), { recursive: true });
  let ultimoErro = "";

  for (let tentativa = 1; tentativa <= tentativas; tentativa++) {
    try {
      await executarArquivo(CURL_BIN, [
        "-sS",
        "--http1.1",
        "-L",
        "-C",
        "-",
        "-A",
        "Mozilla/5.0",
        "-e",
        ORSE_BASE_DADOS_URL,
        "--connect-timeout",
        "20",
        "--max-time",
        "180",
        "-o",
        destino,
        url.toString(),
      ]);
    } catch (error) {
      ultimoErro = limparErroProcesso(error.message);
    }

    const bytes = fs.existsSync(destino) ? fs.statSync(destino).size : 0;
    if (bytes > 0 && (!tamanhoEsperado || bytes >= tamanhoEsperado)) return bytes;

    debugWarn(`Download ORSE incompleto na tentativa ${tentativa}: ${bytes}/${tamanhoEsperado || "?"} bytes.`);
  }

  const bytes = fs.existsSync(destino) ? fs.statSync(destino).size : 0;
  throw new Error(
    `Download ORSE incompleto após ${tentativas} tentativa(s): ${bytes}/${tamanhoEsperado || "?"} bytes. ${ultimoErro}`.trim()
  );
}

async function obterCabecalhosHttp(url, timeoutSegundos = 60) {
  const { stdout } = await executarArquivo(CURL_BIN, [
    "-sS",
    "-I",
    "-L",
    "-A",
    "Mozilla/5.0",
    "--connect-timeout",
    "20",
    "--max-time",
    String(timeoutSegundos),
    url.toString(),
  ]);

  return stdout;
}

function obterContentLength(cabecalhos) {
  const matches = [...String(cabecalhos || "").matchAll(/content-length:\s*(\d+)/gi)];
  if (!matches.length) return 0;
  return Number(matches[matches.length - 1][1]) || 0;
}

function listarArquivosRecursivo(dir) {
  if (!fs.existsSync(dir)) return [];

  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((item) => {
    const caminho = path.join(dir, item.name);
    return item.isDirectory() ? listarArquivosRecursivo(caminho) : [caminho];
  });
}

async function extrairArquivoSicro(arquivo, destino) {
  fs.mkdirSync(destino, { recursive: true });
  await executarArquivo(sevenZip.path7za, ["x", "-y", `-o${destino}`, arquivo]);
  return listarArquivosRecursivo(destino);
}

async function tentarExtrairArquivo(arquivo, destino) {
  try {
    fs.mkdirSync(destino, { recursive: true });
    await executarArquivo(sevenZip.path7za, ["x", "-y", `-o${destino}`, arquivo]);
    return {
      ok: true,
      arquivos: listarArquivosRecursivo(destino),
      erro: "",
    };
  } catch (error) {
    return {
      ok: false,
      arquivos: [],
      erro: error.message,
    };
  }
}

function lerRegistrosDeArquivoTabela(arquivo, contexto) {
  const ext = path.extname(arquivo).toLowerCase();

  if ([".xlsx", ".xls", ".xlsm"].includes(ext)) {
    const workbook = XLSX.readFile(arquivo, { cellDates: false });
    return workbook.SheetNames.flatMap((nomePlanilha) => {
      const linhas = XLSX.utils.sheet_to_json(workbook.Sheets[nomePlanilha], {
        header: 1,
        defval: "",
        raw: false,
      });
      return normalizarTabelaPlanilha(linhas, { ...contexto, planilha: nomePlanilha });
    });
  }

  if (ext === ".csv") {
    const linhas = parseCsvComDelimitador(fs.readFileSync(arquivo, "utf8"));
    return normalizarTabelaPlanilha(linhas, { ...contexto, planilha: path.basename(arquivo) });
  }

  return [];
}

function lerRegistrosDeArquivoSicro(arquivo, contexto) {
  return lerRegistrosDeArquivoTabela(arquivo, { ...contexto, fonteId: "sicro" });
}

async function importarSicroDownloads() {
  const manifest = carregarManifestSicro();
  const registros = [];
  const resumo = [];

  for (const [uf, info] of Object.entries(manifest)) {
    if (!info?.caminho || !fs.existsSync(info.caminho)) continue;

    const destino = path.join(SICRO_EXTRACT_DIR, uf.toLowerCase(), info.ano || "sem-ano", path.basename(info.caminho, path.extname(info.caminho)));
    let arquivosExtraidos = listarArquivosRecursivo(destino);

    if (!arquivosExtraidos.length) {
      arquivosExtraidos = await extrairArquivoSicro(info.caminho, destino);
    }

    const arquivosDados = arquivosExtraidos.filter((arquivo) => {
      const ext = path.extname(arquivo).toLowerCase();
      const nome = normalizarTexto(path.basename(arquivo));
      return (
        [".xlsx", ".xls", ".xlsm", ".csv"].includes(ext) &&
        nome.includes("relatorio sintetico") &&
        !nome.includes("origem de precos")
      );
    });

    const antes = registros.length;
    for (const arquivo of arquivosDados) {
      registros.push(...lerRegistrosDeArquivoSicro(arquivo, {
        uf,
        estado: info.estado,
        dataReferencia: info.dataReferencia,
        tipo: "",
        arquivo: path.basename(arquivo),
      }));
    }

    resumo.push({
      uf,
      arquivo: info.nomeArquivo,
      arquivosDados: arquivosDados.length,
      registros: registros.length - antes,
    });
  }

  fs.mkdirSync(BASES_DIR, { recursive: true });
  fs.writeFileSync(SICRO_JSON_PATH, JSON.stringify(registros, null, 2));

  return {
    atualizadoEm: new Date().toISOString(),
    arquivo: SICRO_JSON_PATH,
    totalRegistros: registros.length,
    resumo,
  };
}

function compactarRegistrosSicro(registros, uf = "") {
  const ufFiltro = String(uf || "").trim().toUpperCase();
  const mapa = new Map();

  for (const item of Array.isArray(registros) ? registros : []) {
    const ufItem = String(item.uf || "").trim().toUpperCase();
    if (ufFiltro && ufItem && ufItem !== ufFiltro) continue;

    const precoUnitario = normalizarValor(item.precoUnitario);
    if (!item.codigo || !item.descricao || Number.isNaN(precoUnitario) || precoUnitario <= 0) continue;

    const registro = {
      origem: "SICRO",
      tipoSinapi: item.tipoSinapi || item.tipo || "",
      codigo: item.codigo,
      descricao: item.descricao,
      unidade: item.unidade || "",
      precoUnitario,
      dataReferencia: item.dataReferencia || "",
      uf: ufItem || item.uf || "",
      referencia: item.referencia || "",
      arquivo: item.arquivo || "",
      planilha: item.planilha || "",
      estado: item.estado || "",
    };

    const chave = [
      registro.uf,
      registro.dataReferencia,
      registro.codigo,
      normalizarTexto(registro.descricao),
      registro.precoUnitario,
    ].join("|");

    if (!mapa.has(chave)) {
      mapa.set(chave, registro);
    }
  }

  return Array.from(mapa.values()).sort((a, b) => {
    const data = String(b.dataReferencia || "").localeCompare(String(a.dataReferencia || ""));
    if (data !== 0) return data;
    return String(a.codigo || "").localeCompare(String(b.codigo || ""));
  });
}

function calcularEstatisticas(lista) {
  const valores = lista
    .map((x) => normalizarValor(x.precoUnitario))
    .filter((v) => !Number.isNaN(v) && v > 0)
    .sort((a, b) => a - b);

  if (!valores.length) {
    return {
      quantidade: 0,
      menor: 0,
      maior: 0,
      media: 0,
      mediana: 0,
    };
  }

  const soma = valores.reduce((acc, v) => acc + v, 0);
  const meio = Math.floor(valores.length / 2);

  const mediana =
    valores.length % 2 === 0
      ? (valores[meio - 1] + valores[meio]) / 2
      : valores[meio];

  return {
    quantidade: valores.length,
    menor: valores[0],
    maior: valores[valores.length - 1],
    media: soma / valores.length,
    mediana,
  };
}

async function consultarJson(url, timeoutMs = 15000) {
  debugLog("Consultando:", url.toString());

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const resposta = await fetch(url.toString(), {
      headers: {
        accept: "application/json",
        "user-agent": "Mozilla/5.0",
      },
      signal: controller.signal,
    });

    const texto = await resposta.text();

    if (!resposta.ok) {
      throw new Error(`Status ${resposta.status}: ${texto}`);
    }

    try {
      return JSON.parse(texto);
    } catch (error) {
      throw new Error("A API retornou um conteúdo que não é JSON.");
    }
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error("A API pública demorou demais para responder. Tente novamente em alguns segundos.");
    }

    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

function timestampRegistro(registro) {
  const data =
    registro.pncp?.item?.dataResultadoHomologacao ||
    registro.dataResultado ||
    registro.dataReferencia ||
    registro.dataCompra ||
    registro.pncp?.contratacao?.dataDivulgacaoPncp ||
    "";

  const timestamp = Date.parse(data);
  return Number.isNaN(timestamp) ? 0 : timestamp;
}

function normalizarPeriodoMeses(valor) {
  const meses = Number(valor || 6);
  if (!Number.isFinite(meses) || meses <= 0) return 6;
  return Math.min(Math.floor(meses), 6);
}

function normalizarLimiteItens(valor) {
  const limite = Number(valor || 100);
  if (!Number.isFinite(limite) || limite <= 0) return 100;
  return Math.min(Math.floor(limite), 200);
}

function mensagemErroApiPublica(error) {
  const mensagem = String(error?.message || error || "");

  if (mensagem.includes("Could not open JPA EntityManager")) {
    return "O Compras.gov está instável agora para este item. Tente novamente em alguns segundos ou use outro item do catálogo.";
  }

  if (mensagem.includes("demorou demais")) {
    return "A API pública demorou para responder. Tente novamente em alguns segundos.";
  }

  return mensagem;
}

function apiCatalogoInstavel(error) {
  const mensagem = String(error?.message || error || "");
  return (
    mensagem.includes("Could not open JPA EntityManager") ||
    mensagem.includes("demorou demais") ||
    mensagem.includes("Status 400") ||
    mensagem.includes("fetch failed")
  );
}

function obterFalhaBuscaItens(chave) {
  const falha = cacheFalhasBuscaItens.get(chave);
  if (!falha) return null;

  if (Date.now() - falha.em > BUSCA_ITENS_FAILURE_TTL_MS) {
    cacheFalhasBuscaItens.delete(chave);
    return null;
  }

  return falha;
}

function gravarFalhaBuscaItens(chave, error) {
  cacheFalhasBuscaItens.set(chave, {
    em: Date.now(),
    mensagem: mensagemErroApiPublica(error),
  });
}

function filtrarRegistrosPorPeriodo(registros, meses) {
  const limiteMeses = normalizarPeriodoMeses(meses);
  const dataMinima = new Date();
  dataMinima.setMonth(dataMinima.getMonth() - limiteMeses);
  dataMinima.setHours(0, 0, 0, 0);

  return registros.filter((registro) => {
    const timestamp = timestampRegistro(registro);
    return timestamp && timestamp >= dataMinima.getTime();
  });
}

function dataReferenciaPadrao() {
  const data = new Date();
  data.setMonth(data.getMonth() - 1);
  const ano = data.getFullYear();
  const mes = String(data.getMonth() + 1).padStart(2, "0");
  return `${ano}-${mes}`;
}

function normalizarDataReferenciaSinapi(valor) {
  const texto = String(valor || "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) return texto.slice(0, 7);
  return texto || dataReferenciaPadrao();
}

function resumirTextoBuscaSinapi(texto) {
  const normalizado = normalizarTexto(texto);
  const palavras = normalizado
    .split(" ")
    .filter((palavra) => palavra.length >= 3 && !STOP_WORDS.has(palavra));

  return palavras.slice(0, 6).join(" ") || normalizado;
}

async function consultarPncpJson(url, timeoutMs = 15000) {
  debugLog("Consultando PNCP:", url.toString());

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const resposta = await fetch(url.toString(), {
      headers: {
        accept: "application/json",
        "user-agent": "Mozilla/5.0",
      },
      signal: controller.signal,
    });

    const texto = await resposta.text();

    if (!resposta.ok) {
      throw new Error(`Status ${resposta.status}: ${texto}`);
    }

    return texto ? JSON.parse(texto) : null;
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error("O PNCP demorou demais para responder. Tente novamente em alguns segundos.");
    }

    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

async function consultarCnbsJson(url, timeoutMs = 15000) {
  debugLog("Consultando CNBS:", url.toString());

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const resposta = await fetch(url.toString(), {
      headers: {
        accept: "application/json",
        "user-agent": "Mozilla/5.0",
      },
      signal: controller.signal,
    });

    const texto = await resposta.text();

    if (!resposta.ok) {
      throw new Error(`Status ${resposta.status}: ${texto}`);
    }

    return texto ? JSON.parse(texto) : null;
  } finally {
    clearTimeout(timeout);
  }
}

async function consultarSinapiJson(url, timeoutMs = 12000) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  const headers = {
    accept: "application/json",
    "user-agent": "Mozilla/5.0",
  };

  if (SINAPI_API_KEY) {
    headers["X-API-KEY"] = SINAPI_API_KEY;
  }

  try {
    const resposta = await fetch(url.toString(), {
      headers,
      signal: controller.signal,
    });

    const texto = await resposta.text();

    if (!resposta.ok) {
      throw new Error(`Status ${resposta.status}: ${texto}`);
    }

    return texto ? JSON.parse(texto) : null;
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error("A API SINAPI demorou demais para responder.");
    }

    if (error.message === "fetch failed") {
      throw new Error(
        `Não consegui conectar na API SINAPI em ${SINAPI_API_URL}. Verifique se o Docker/autoSINAPI_API está rodando na porta configurada.`
      );
    }

    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

async function consultarTextoPublico(url, timeoutMs = 20000) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const resposta = await fetch(url.toString(), {
      headers: {
        accept: "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "user-agent": "Mozilla/5.0",
      },
      signal: controller.signal,
    });

    const texto = await resposta.text();
    if (!resposta.ok) throw new Error(`Status ${resposta.status}: ${texto.slice(0, 200)}`);
    return texto;
  } finally {
    clearTimeout(timeout);
  }
}

async function consultarJsonPostPublico(url, payload, timeoutMs = 30000) {
  debugLog("Consultando POST público:", url.toString());

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const resposta = await fetch(url.toString(), {
      method: "POST",
      headers: {
        accept: "application/json",
        "content-type": "application/json; charset=utf-8",
        "user-agent": "Mozilla/5.0",
      },
      body: JSON.stringify(payload || {}),
      signal: controller.signal,
    });

    const texto = await resposta.text();

    if (!resposta.ok) {
      throw new Error(`Status ${resposta.status}: ${texto.slice(0, 300)}`);
    }

    try {
      return texto ? JSON.parse(texto) : null;
    } catch (error) {
      throw new Error("A fonte pública retornou um conteúdo que não é JSON.");
    }
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error("A fonte pública demorou demais para responder.");
    }

    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

async function baixarArquivoPublico(url, destino, timeoutMs = 120000) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const resposta = await fetch(url.toString(), {
      headers: {
        "user-agent": "Mozilla/5.0",
      },
      signal: controller.signal,
    });

    if (!resposta.ok) {
      const texto = await resposta.text().catch(() => "");
      throw new Error(`Status ${resposta.status}: ${texto.slice(0, 200)}`);
    }

    const buffer = Buffer.from(await resposta.arrayBuffer());
    fs.mkdirSync(path.dirname(destino), { recursive: true });
    fs.writeFileSync(destino, buffer);
    return buffer.length;
  } finally {
    clearTimeout(timeout);
  }
}

function textoContemTodasPalavras(textoBase, termoBusca) {
  const palavras = palavrasRelevantes(termoBusca);
  const texto = ` ${normalizarTexto(textoBase)} `;

  return palavras.every((p) => texto.includes(` ${p} `));
}

function textoContemAlgumaPalavra(textoBase, palavras) {
  const texto = ` ${normalizarTexto(textoBase)} `;
  return palavras.some((p) => texto.includes(` ${normalizarTexto(p)} `));
}

function pontuarTexto(textoBase, termosBusca) {
  const texto = normalizarTexto(textoBase);
  let pontuacao = 0;

  for (const termoBusca of termosBusca) {
    const termo = normalizarTexto(termoBusca);
    const palavras = palavrasRelevantes(termo);

    if (!palavras.length) continue;
    if (texto === termo) pontuacao += 100;
    if (texto.startsWith(termo)) pontuacao += 50;
    if (texto.includes(termo)) pontuacao += 25;

    pontuacao += palavras.filter((palavra) => texto.includes(palavra)).length * 5;
  }

  return pontuacao;
}

function pontuarResultadoBusca(textoBase, termosBusca) {
  return pontuarTexto(textoBase, termosBusca.slice(0, 1)) * 10 + pontuarTexto(textoBase, termosBusca);
}

function primeiraPalavraRelevante(termo) {
  return palavrasRelevantes(termo)[0] || "";
}

function textoComecaComPalavra(textoBase, palavra) {
  if (!palavra) return true;
  return normalizarTexto(textoBase).startsWith(normalizarTexto(palavra));
}

function ordenarItensPorAderencia(itens, termosBusca, montarTexto) {
  const primeiraPalavra = primeiraPalavraRelevante(termosBusca[0]);

  return [...itens].sort((a, b) => {
    const textoA = montarTexto(a);
    const textoB = montarTexto(b);
    const aComeca = textoComecaComPalavra(textoA, primeiraPalavra) ? 1 : 0;
    const bComeca = textoComecaComPalavra(textoB, primeiraPalavra) ? 1 : 0;

    if (aComeca !== bComeca) return bComeca - aComeca;

    return pontuarResultadoBusca(textoB, termosBusca) - pontuarResultadoBusca(textoA, termosBusca);
  });
}

function criarLinkPncp(registro) {
  const linkExistente =
    registro.linkPncp ||
    registro.linkPNCP ||
    registro.urlPncp ||
    registro.urlPNCP ||
    registro.linkCompra ||
    registro.urlCompra;

  if (linkExistente) return linkExistente;

  const cnpj =
    registro.cnpjOrgao ||
    registro.cnpjOrgaoEntidade ||
    registro.cnpjUnidadeOrgao ||
    registro.orgaoCnpj ||
    registro.cnpj;

  const ano =
    registro.anoCompra ||
    registro.anoContratacao ||
    registro.ano ||
    registro.dataCompra?.slice?.(0, 4);

  const sequencial =
    registro.sequencialCompra ||
    registro.numeroSequencialCompra ||
    registro.sequencialContratacao ||
    registro.numeroControlePNCP?.split?.("-")?.pop?.();

  if (cnpj && ano && sequencial) {
    const cnpjLimpo = String(cnpj).replace(/\D/g, "");
    return `https://pncp.gov.br/app/editais/${cnpjLimpo}/${ano}/${sequencial}`;
  }

  return "";
}

function parseIdCompra(idCompra) {
  const texto = String(idCompra || "").replace(/\D/g, "");

  if (texto.length < 17) return null;

  return {
    uasg: texto.slice(0, 6),
    modalidade: texto.slice(6, 8),
    numeroCompra: texto.slice(8, -4).replace(/^0+/, "") || "0",
    numeroCompraOriginal: texto.slice(8, -4),
    anoCompra: texto.slice(-4),
  };
}

function obterPrimeiroValor(registro, campos) {
  for (const campo of campos) {
    const valor = registro[campo];

    if (valor !== null && valor !== undefined && valor !== "") {
      return valor;
    }
  }

  return "";
}

function criarLinkPncpPorNumeroControle(numeroControle) {
  if (!numeroControle) return "";

  const partes = String(numeroControle).match(/^(\d{14})-\d-(\d+)\/(\d{4})/);

  if (!partes) return "";

  const [, cnpj, sequencial, ano] = partes;
  return `https://pncp.gov.br/app/editais/${cnpj}/${ano}/${sequencial}`;
}

function linkPncpValido(link) {
  return String(link || "").startsWith(`${PNCP_BASE}/app/editais/`);
}

function parseNumeroControlePncp(numeroControle) {
  if (!numeroControle) return null;

  const partes = String(numeroControle).match(/^(\d{14})-\d-(\d+)\/(\d{4})/);

  if (!partes) return null;

  return {
    cnpj: partes[1],
    sequencial: String(Number(partes[2])),
    ano: partes[3],
  };
}

function extrairIdentificacaoLicitacao(registro) {
  const numeroControlePncp = obterPrimeiroValor(registro, [
    "numeroControlePNCP",
    "numeroControlePNCPCompra",
    "idContratacaoPNCP",
    "numeroControlePncp",
    "idPncp",
  ]);

  const linkPncp =
    criarLinkPncpPorNumeroControle(numeroControlePncp) ||
    criarLinkPncp(registro);

  const codigoCompra = obterPrimeiroValor(registro, [
    "idCompra",
    "numeroCompra",
    "compra",
    "codigoCompra",
  ]);

  return {
    numeroControlePncp,
    codigoCompra,
    processo: obterPrimeiroValor(registro, ["processo", "numeroProcesso"]),
    modalidade: obterPrimeiroValor(registro, ["modalidadeNome", "modalidade"]),
    linkPncp,
    identificador: numeroControlePncp || codigoCompra || "",
  };
}

function resumirContratacaoPncp(meta, detalhe = null) {
  if (!meta && !detalhe) return null;

  const numeroControle =
    detalhe?.numeroControlePNCP ||
    meta?.numero_controle_pncp ||
    meta?.numeroControlePNCP ||
    "";

  const partes = parseNumeroControlePncp(numeroControle);
  const cnpj = meta?.orgao_cnpj || detalhe?.orgaoEntidade?.cnpj || partes?.cnpj || "";
  const ano = meta?.ano || detalhe?.anoCompra || partes?.ano || "";
  const sequencial =
    meta?.numero_sequencial ||
    detalhe?.sequencialCompra ||
    partes?.sequencial ||
    "";

  return {
    numeroEdital: detalhe?.numeroCompra || meta?.title?.replace(/^Edital n[ºo]\s*/i, "") || "",
    numeroControlePncp: numeroControle,
    linkPncp: cnpj && ano && sequencial ? `${PNCP_BASE}/app/editais/${cnpj}/${ano}/${sequencial}` : "",
    local: [detalhe?.unidadeOrgao?.municipioNome || meta?.municipio_nome, detalhe?.unidadeOrgao?.ufSigla || meta?.uf]
      .filter(Boolean)
      .join("/"),
    orgao: detalhe?.orgaoEntidade?.razaoSocial || meta?.orgao_nome || "",
    unidadeCompradora:
      detalhe?.unidadeOrgao?.codigoUnidade && detalhe?.unidadeOrgao?.nomeUnidade
        ? `${detalhe.unidadeOrgao.codigoUnidade} - ${detalhe.unidadeOrgao.nomeUnidade}`
        : meta?.unidade_codigo && meta?.unidade_nome
          ? `${meta.unidade_codigo} - ${meta.unidade_nome}`
          : "",
    modalidade: detalhe?.modalidadeNome || meta?.modalidade_licitacao_nome || "",
    amparoLegal: detalhe?.amparoLegal?.nome || "",
    tipo: detalhe?.tipoInstrumentoConvocatorioNome || meta?.tipo_nome || "",
    modoDisputa: detalhe?.modoDisputaNome || "",
    registroPreco:
      detalhe?.srp !== undefined
        ? detalhe.srp
          ? "Sim"
          : "Não"
        : "",
    fonteOrcamentaria: detalhe?.fontesOrcamentarias?.[0]?.nome || meta?.fonte_orcamentaria_nome || "",
    dataDivulgacaoPncp: detalhe?.dataPublicacaoPncp || meta?.data_publicacao_pncp || "",
    situacao: detalhe?.situacaoCompraNome || meta?.situacao_nome || "",
    dataInicioPropostas: detalhe?.dataAberturaProposta || meta?.data_inicio_vigencia || "",
    dataFimPropostas: detalhe?.dataEncerramentoProposta || meta?.data_fim_vigencia || "",
    objeto: detalhe?.objetoCompra || meta?.description || "",
    informacaoComplementar: detalhe?.informacaoComplementar || "",
    valorTotalEstimado: detalhe?.valorTotalEstimado ?? "",
    valorTotalHomologado: detalhe?.valorTotalHomologado ?? "",
  };
}

function resumirItemPncp(item, resultado) {
  if (!item && !resultado) return null;

  return {
    numeroItem: item?.numeroItem || resultado?.numeroItem || "",
    descricao: item?.descricao || "",
    criterioJulgamento: item?.criterioJulgamentoNome || "",
    situacaoItem: item?.situacaoCompraItemNome || "",
    tipo: item?.materialOuServicoNome || "",
    categoria: item?.itemCategoriaNome || "",
    quantidade: item?.quantidade ?? "",
    unidadeMedida: item?.unidadeMedida || "",
    valorUnitarioEstimado: item?.valorUnitarioEstimado ?? "",
    valorTotalEstimado: item?.valorTotal ?? "",
    dataResultadoHomologacao: resultado?.dataResultado || "",
    quantidadeHomologada: resultado?.quantidadeHomologada ?? "",
    valorUnitarioHomologado: resultado?.valorUnitarioHomologado ?? "",
    valorTotalHomologado: resultado?.valorTotalHomologado ?? "",
    fornecedor: resultado?.nomeRazaoSocialFornecedor || "",
    niFornecedor: resultado?.niFornecedor || "",
    situacaoResultado: resultado?.situacaoCompraItemResultadoNome || "",
  };
}

async function buscarMetaPncpPorIdCompra(idCompra) {
  if (!idCompra) return null;
  if (cachePncpPorIdCompra.has(idCompra)) return cachePncpPorIdCompra.get(idCompra);

  const infoCompra = parseIdCompra(idCompra);
  const url = new URL(PNCP_BASE + "/api/search/");
  url.searchParams.set("q", idCompra);
  url.searchParams.set("tipos_documento", "edital");
  url.searchParams.set("ordenacao", "-data");
  url.searchParams.set("pagina", "1");
  url.searchParams.set("tam_pagina", "10");
  url.searchParams.set("status", "todos");

  try {
    const dados = await consultarPncpJson(url);
    const itens = dados?.items || [];
    const encontrado =
      itens.find((item) => {
        const mesmoUasg = !infoCompra?.uasg || String(item.unidade_codigo || "") === infoCompra.uasg;
        const mesmoAno = !infoCompra?.anoCompra || String(item.ano || "") === infoCompra.anoCompra;
        const titulo = normalizarTexto(item.title || "");
        const mesmoNumero =
          !infoCompra?.numeroCompra ||
          titulo.includes(infoCompra.numeroCompra) ||
          titulo.includes(infoCompra.numeroCompraOriginal);

        return mesmoUasg && mesmoAno && mesmoNumero;
      }) ||
      itens[0] ||
      null;

    cachePncpPorIdCompra.set(idCompra, encontrado);
    return encontrado;
  } catch (error) {
    debugWarn(`Não foi possível localizar contratação PNCP para ${idCompra}:`, error.message);
    cachePncpPorIdCompra.set(idCompra, null);
    return null;
  }
}

async function buscarDetalheContratacaoPncp(meta, timeoutMs = 15000) {
  if (!meta?.orgao_cnpj || !meta?.ano || !meta?.numero_sequencial) return null;

  const chave = `${meta.orgao_cnpj}/${meta.ano}/${meta.numero_sequencial}`;
  if (cacheDetalheContratacaoPncp.has(chave)) return cacheDetalheContratacaoPncp.get(chave);

  try {
    const url = new URL(`${PNCP_BASE}/api/consulta/v1/orgaos/${meta.orgao_cnpj}/compras/${meta.ano}/${meta.numero_sequencial}`);
    const detalhe = await consultarPncpJson(url, timeoutMs);
    cacheDetalheContratacaoPncp.set(chave, detalhe);
    return detalhe;
  } catch (error) {
    debugWarn(`Não foi possível detalhar contratação PNCP ${chave}:`, error.message);
    return null;
  }
}

async function buscarItemResultadoPncp(meta, numeroItem, nomeFornecedor = "") {
  if (!meta?.orgao_cnpj || !meta?.ano || !meta?.numero_sequencial || !numeroItem) return null;

  const chave = `${meta.orgao_cnpj}/${meta.ano}/${meta.numero_sequencial}/${numeroItem}`;
  if (cacheResultadoItemPncp.has(chave)) return cacheResultadoItemPncp.get(chave);

  try {
    const baseItem = `${PNCP_BASE}/api/pncp/v1/orgaos/${meta.orgao_cnpj}/compras/${meta.ano}/${meta.numero_sequencial}/itens/${numeroItem}`;
    const [item, resultados] = await Promise.all([
      consultarPncpJson(new URL(baseItem)),
      consultarPncpJson(new URL(`${baseItem}/resultados`)),
    ]);

    const fornecedorNormalizado = normalizarTexto(nomeFornecedor);
    const resultado =
      (Array.isArray(resultados) ? resultados : []).find((r) =>
        fornecedorNormalizado && normalizarTexto(r.nomeRazaoSocialFornecedor).includes(fornecedorNormalizado)
      ) ||
      (Array.isArray(resultados) ? resultados : []).find((r) => r.ordemClassificacaoSrp === 1) ||
      (Array.isArray(resultados) ? resultados : [])[0] ||
      null;

    const resumo = resumirItemPncp(item, resultado);
    cacheResultadoItemPncp.set(chave, resumo);
    return resumo;
  } catch (error) {
    debugWarn(`Não foi possível detalhar item PNCP ${chave}:`, error.message);
    return null;
  }
}

async function buscarItensContratacaoPncp(meta, timeoutMs = 15000) {
  if (!meta?.orgao_cnpj || !meta?.ano || !meta?.numero_sequencial) return [];

  const chave = `${meta.orgao_cnpj}/${meta.ano}/${meta.numero_sequencial}`;
  if (cacheItensContratacaoPncp.has(chave)) return cacheItensContratacaoPncp.get(chave);

  try {
    const url = new URL(`${PNCP_BASE}/api/pncp/v1/orgaos/${meta.orgao_cnpj}/compras/${meta.ano}/${meta.numero_sequencial}/itens`);
    const itens = await consultarPncpJson(url, timeoutMs);
    const lista = Array.isArray(itens) ? itens : [];
    cacheItensContratacaoPncp.set(chave, lista);
    return lista;
  } catch (error) {
    debugWarn(`Não foi possível listar itens PNCP ${chave}:`, error.message);
    return [];
  }
}

async function buscarResultadoHomologadoPncp(meta, item, timeoutMs = 15000) {
  if (!meta?.orgao_cnpj || !meta?.ano || !meta?.numero_sequencial || !item?.numeroItem) return null;

  try {
    const url = new URL(
      `${PNCP_BASE}/api/pncp/v1/orgaos/${meta.orgao_cnpj}/compras/${meta.ano}/${meta.numero_sequencial}/itens/${item.numeroItem}/resultados`
    );
    const resultados = await consultarPncpJson(url, timeoutMs);
    const lista = Array.isArray(resultados) ? resultados : [];

    return (
      lista
        .filter((resultado) => {
          const valor = normalizarValor(resultado.valorUnitarioHomologado);
          return !Number.isNaN(valor) && valor > 0;
        })
        .sort((a, b) => {
          const ordemA = Number(a.ordemClassificacaoSrp || 9999);
          const ordemB = Number(b.ordemClassificacaoSrp || 9999);
          return ordemA - ordemB;
        })[0] || null
    );
  } catch (error) {
    debugWarn(`Não foi possível buscar resultado PNCP do item ${item.numeroItem}:`, error.message);
    return null;
  }
}

function montarRegistroPncpDireto({ meta, detalhe, item, resultado, codigo }) {
  const contratacao = resumirContratacaoPncp(meta, detalhe);
  const resumoItem = resumirItemPncp(item, resultado);
  const precoUnitario = resultado?.valorUnitarioHomologado ?? item?.valorUnitarioEstimado ?? 0;

  return {
    origem: "PNCP",
    codigo: codigo || item?.catalogoCodigoItem || item?.numeroItem || "",
    descricao: item?.descricao || "",
    orgao: contratacao?.orgao || meta?.orgao_nome || "",
    nomeUasg: contratacao?.unidadeCompradora || "",
    municipio: detalhe?.unidadeOrgao?.municipioNome || meta?.municipio_nome || "",
    uf: detalhe?.unidadeOrgao?.ufSigla || meta?.uf || "",
    nomeFornecedor: resultado?.nomeRazaoSocialFornecedor || "",
    quantidade: resultado?.quantidadeHomologada ?? item?.quantidade ?? "",
    unidade: item?.unidadeMedida || "",
    precoUnitario,
    dataResultado: resultado?.dataResultado || item?.dataAtualizacao || contratacao?.dataDivulgacaoPncp || "",
    numeroItemCompra: item?.numeroItem || "",
    licitacao: {
      numeroControlePncp: contratacao?.numeroControlePncp || "",
      codigoCompra: contratacao?.numeroEdital || "",
      modalidade: contratacao?.modalidade || "",
      linkPncp: contratacao?.linkPncp || "",
      identificador: contratacao?.numeroControlePncp || contratacao?.numeroEdital || "",
    },
    pncp: {
      encontrado: true,
      contratacao,
      item: resumoItem,
    },
    linkPncpGerado: contratacao?.linkPncp || "",
  };
}

async function buscarPrecosPncpDireto({ termo, codigo, tipo, limite = 150 }) {
  const termoBusca = String(termo || codigo || "").trim();
  if (!termoBusca || tipo === "material-pdm") return [];

  const prazoFinal = Date.now() + 55000;
  const termosBusca = gerarTermosBusca(termoBusca);
  const url = new URL(PNCP_BASE + "/api/search/");
  url.searchParams.set("q", termoBusca);
  url.searchParams.set("tipos_documento", "edital");
  url.searchParams.set("ordenacao", "-data");
  url.searchParams.set("pagina", "1");
  url.searchParams.set("tam_pagina", "20");
  url.searchParams.set("status", "todos");

  const dados = await consultarPncpJson(url, 10000);
  const metas = (dados?.items || [])
    .filter((meta) => meta?.orgao_cnpj && meta?.ano && meta?.numero_sequencial)
    .filter((meta) => meta.tem_resultado !== false)
    .sort((a, b) => Number(Boolean(b.tem_resultado)) - Number(Boolean(a.tem_resultado)))
    .slice(0, 6);

  const registros = [];

  for (const meta of metas) {
    const restante = prazoFinal - Date.now();
    if (registros.length >= limite || restante <= 3000) break;

    const [detalhe, itens] = await Promise.all([
      buscarDetalheContratacaoPncp(meta, limitarTimeout(restante, 6000)),
      buscarItensContratacaoPncp(meta, limitarTimeout(restante, 6000)),
    ]);

    const itensRelacionados = itens
      .map((item) => ({
        item,
        pontuacao: pontuarResultadoBusca(
          `${item.descricao || ""} ${item.catalogoCodigoItem || ""} ${item.patrimonio || ""}`,
          termosBusca
        ),
      }))
      .filter(({ item, pontuacao }) => {
        const mesmoCodigo = codigo && String(item.catalogoCodigoItem || item.patrimonio || "") === String(codigo);
        return item.temResultado && (pontuacao > 0 || mesmoCodigo);
      })
      .sort((a, b) => b.pontuacao - a.pontuacao)
      .slice(0, 3)
      .map(({ item }) => item);

    const registrosDaContratacao = await Promise.all(itensRelacionados.map(async (item) => {
      const resultado = await buscarResultadoHomologadoPncp(meta, item, limitarTimeout(prazoFinal - Date.now(), 5000));
      if (!resultado) return null;

      const registro = montarRegistroPncpDireto({ meta, detalhe, item, resultado, codigo });
      const valor = normalizarValor(registro.precoUnitario);
      if (!Number.isNaN(valor) && valor > 0 && linkPncpValido(registro.linkPncpGerado)) {
        return registro;
      }

      return null;
    }));

    registros.push(...registrosDaContratacao.filter(Boolean));
  }

  return registros.sort((a, b) => timestampRegistro(b) - timestampRegistro(a));
}

async function enriquecerRegistroComPncp(registro) {
  const idCompra = registro.idCompra || registro.licitacao?.codigoCompra;
  const meta = await buscarMetaPncpPorIdCompra(idCompra);
  const detalhe = await buscarDetalheContratacaoPncp(meta);
  const contratacao = resumirContratacaoPncp(meta, detalhe);
  const item = await buscarItemResultadoPncp(meta, registro.numeroItemCompra, registro.nomeFornecedor);

  return {
    ...registro,
    pncp: {
      encontrado: Boolean(meta),
      contratacao,
      item,
    },
    linkPncpGerado: contratacao?.linkPncp || registro.linkPncpGerado || "",
  };
}

async function enriquecerRegistrosComPncp(registros, limite = 150) {
  const enriquecidos = [];
  const tamanhoLote = 5;

  for (let i = 0; i < registros.length; i += tamanhoLote) {
    if (i >= limite) {
      enriquecidos.push(...registros.slice(i));
      continue;
    }

    const lote = registros.slice(i, Math.min(i + tamanhoLote, limite));
    const loteEnriquecido = await Promise.all(lote.map(enriquecerRegistroComPncp));
    enriquecidos.push(...loteEnriquecido);
  }

  return enriquecidos;
}

async function consultarPrecosSinapi({ termo, uf, dataReferencia, regime }) {
  if (!SINAPI_HABILITADO) {
    return {
      configurado: Boolean(SINAPI_API_URL),
      desativado: true,
      erro: SINAPI_API_URL
        ? "SINAPI desativado por configuração. Defina ENABLE_SINAPI=1 para habilitar."
        : "SINAPI sem API configurada neste ambiente.",
      registros: [],
    };
  }

  const termoBusca = resumirTextoBuscaSinapi(termo);
  if (!termoBusca || termoBusca.length < 3) {
    return {
      configurado: true,
      erro: "",
      registros: [],
    };
  }

  const base = SINAPI_API_URL.replace(/\/+$/, "");
  const parametros = {
    q: termoBusca,
    uf: String(uf || "PB").toUpperCase(),
    data_referencia: normalizarDataReferenciaSinapi(dataReferencia),
    regime: regime || "NAO_DESONERADO",
    skip: "0",
    limit: "50",
  };

  try {
    const [insumosRaw, composicoesRaw] = await Promise.all([
      (async () => {
        const url = new URL(`${base}/insumos/`);
        Object.entries(parametros).forEach(([chave, valor]) => url.searchParams.set(chave, valor));
        return consultarSinapiJson(url);
      })(),
      (async () => {
        const url = new URL(`${base}/composicoes/`);
        Object.entries(parametros).forEach(([chave, valor]) => url.searchParams.set(chave, valor));
        return consultarSinapiJson(url);
      })(),
    ]);

    const insumos = Array.isArray(insumosRaw) ? insumosRaw : [];
    const composicoes = Array.isArray(composicoesRaw) ? composicoesRaw : [];

    const registros = [
      ...insumos.map((item) => ({
        origem: "SINAPI",
        tipoSinapi: "Insumo",
        codigo: item.codigo,
        descricao: item.descricao,
        unidade: item.unidade,
        precoUnitario: item.preco_mediano,
        dataReferencia: parametros.data_referencia,
        uf: parametros.uf,
        regime: parametros.regime,
      })),
      ...composicoes.map((item) => ({
        origem: "SINAPI",
        tipoSinapi: "Composição",
        codigo: item.codigo,
        descricao: item.descricao,
        unidade: item.unidade,
        precoUnitario: item.custo_total,
        dataReferencia: parametros.data_referencia,
        uf: parametros.uf,
        regime: parametros.regime,
      })),
    ].filter((item) => {
      const valor = normalizarValor(item.precoUnitario);
      return !Number.isNaN(valor) && valor > 0;
    });

    return {
      configurado: true,
      erro: "",
      termoBusca,
      registros: registros.sort((a, b) => normalizarValor(a.precoUnitario) - normalizarValor(b.precoUnitario)),
    };
  } catch (error) {
    return {
      configurado: true,
      erro: error.message,
      registros: [],
    };
  }
}

async function consultarBaseLocal({ fonteId, termo, uf }) {
  const fonte = FONTES_LOCAIS[fonteId];
  const registros = carregarBaseLocal(fonteId);

  if (!fonte || !registros) {
    return {
      id: fonteId,
      nome: fonte?.nome || fonteId,
      configurado: false,
      erro: `Base local não encontrada. Coloque ${fonte?.arquivos?.join(" ou ")} em data/bases para ativar.`,
      registros: [],
    };
  }

  const termosBusca = gerarTermosBusca(termo);
  const ufBusca = normalizarTexto(uf || "");

  const encontrados = registros
    .filter((registro) => {
      const texto = `${registro.codigo || ""} ${registro.descricao || ""}`;
      const correspondeTexto = termosBusca.some((termoBusca) => textoContemTodasPalavras(texto, termoBusca));
      const correspondeUf = !ufBusca || !registro.uf || normalizarTexto(registro.uf) === ufBusca;
      const valor = normalizarValor(registro.precoUnitario);
      return correspondeTexto && correspondeUf && !Number.isNaN(valor) && valor > 0;
    })
    .slice(0, 100);

  return {
    id: fonteId,
    nome: fonte.nome,
    configurado: true,
    erro: "",
    registros: encontrados,
  };
}

function consultarFonteExternaPendente(fonteId) {
  const fonte = FONTES_EXTERNAS[fonteId];

  return {
    id: fonteId,
    nome: fonte?.nome || fonteId,
    configurado: false,
    erro: "Fonte adicionada ao fluxo, mas ainda sem endpoint público configurado para consulta automática.",
    link: fonte?.link || "",
    registros: [],
  };
}

function dataIsoHoje() {
  return new Date().toISOString().slice(0, 10);
}

function dataIsoMesesAtras(meses) {
  const data = new Date();
  data.setMonth(data.getMonth() - normalizarPeriodoMeses(meses));
  return data.toISOString().slice(0, 10);
}

function dataAspNetParaIso(valor) {
  const match = String(valor || "").match(/\/Date\((-?\d+)\)\//);
  if (!match) return "";

  const data = new Date(Number(match[1]));
  if (Number.isNaN(data.getTime())) return "";
  return data.toISOString().slice(0, 10);
}

async function consultarPeIntegradoJson(metodo, payload, timeoutMs = 45000) {
  return consultarJsonPostPublico(new URL(`${PE_INTEGRADO_ARP_BASE}/${metodo}`), payload, timeoutMs);
}

async function buscarAtasPeIntegrado({ termo, periodoMeses, somenteObjeto = true, limiteAtas = 80 }) {
  const quantidadePorPagina = 20;
  const atas = [];
  const paginasMaximas = Math.ceil(limiteAtas / quantidadePorPagina);

  for (let pagina = 0; pagina < paginasMaximas; pagina += 1) {
    const arpFiltroDTO = {
      sProcesso: "",
      sARP: "",
      nOrigem: "999",
      nTipoAta: "999",
      nCdSituacao: "5",
      sObjeto: somenteObjeto ? String(termo || "").trim() : "",
      tDtInicio: dataIsoMesesAtras(periodoMeses),
      tDtEncerramento: dataIsoHoje(),
      exportar: "N",
      nPaginaDe: pagina * quantidadePorPagina,
      sCampoOrderBy: "TDTINICIO",
      sDirecaoOrderBy: "DESC",
    };

    const dados = await consultarPeIntegradoJson("PesquisarARPs", {
      arpFiltroDTO,
      quantidadePorPagina,
    });

    const lista = Array.isArray(dados?.d?.Dados) ? dados.d.Dados : [];
    atas.push(...lista);

    if (!lista.length || atas.length >= limiteAtas || atas.length >= Number(dados?.d?.TotalRegistros || 0)) {
      break;
    }
  }

  return atas.slice(0, limiteAtas);
}

async function buscarItensAtaPeIntegrado(nCdRegistroPreco) {
  if (!nCdRegistroPreco) return [];

  const dados = await consultarPeIntegradoJson("ListarDetalhesItens", { nCdRegistroPreco }, 45000);
  return Array.isArray(dados?.d) ? dados.d : [];
}

function normalizarRegistroPeIntegrado(ata, item, termo) {
  const quantidade = normalizarValor(item.dQtRegistrada || item.dQtQuantidadeTempo || item.dQtSaldoConsumo);
  const precoUnitario = normalizarValor(item.dVlUnitario);
  const dataInicio = dataAspNetParaIso(ata.tDtInicio);
  const dataFim = dataAspNetParaIso(ata.tDtEncerramento);
  const ataNumero = ata.sNrRegistroPreco || "";
  const processo = ata.sNrProcesso || "";

  return {
    origem: "Pernambuco Integrado",
    tipoSinapi: "Ata de Registro de Preço",
    codigo: item.nCdItem || item.nSequencial || "",
    numeroItem: item.nSequencial || "",
    descricao: item.sDsProduto || ata.sDsTitulo || termo,
    unidade: item.sSgUnidadeMedida || "",
    quantidade,
    precoUnitario,
    precoTotal: normalizarValor(item.dVlTotal),
    fornecedor: item.sNmEmpresa || "",
    orgao: ata.sNmOrgaoGestor || "",
    uf: "PE",
    dataReferencia: dataInicio,
    dataResultado: dataInicio,
    vigenciaFim: dataFim,
    ata: ataNumero,
    processo,
    referencia: [ataNumero, processo].filter(Boolean).join(" | "),
    link: PE_INTEGRADO_ARP_BASE,
    licitacao: {
      codigoCompra: processo,
      numeroControlePncp: "",
      modalidade: "Ata de Registro de Preço",
      linkPncp: PE_INTEGRADO_ARP_BASE,
      identificador: ataNumero || processo,
    },
    peIntegrado: {
      nCdRegistroPreco: ata.nCdRegistroPreco,
      ata: ataNumero,
      processo,
      titulo: ata.sDsTitulo || "",
      orgaoGestor: ata.sNmOrgaoGestor || "",
      situacao: ata.sDsSituacao || "",
      inicioVigencia: dataInicio,
      fimVigencia: dataFim,
      permiteAdesao: ata.bFlPermiteAdesao == 1 ? "Sim" : "Não",
    },
  };
}

async function consultarPeIntegrado({ termo, periodoMeses }) {
  const termoBusca = String(termo || "").trim();
  const fonte = FONTES_EXTERNAS.peintegrado;

  if (!termoBusca || termoBusca.length < 3) {
    return {
      id: "peintegrado",
      nome: fonte.nome,
      configurado: true,
      erro: "",
      link: fonte.link,
      registros: [],
    };
  }

  try {
    let atas = await buscarAtasPeIntegrado({
      termo: termoBusca,
      periodoMeses,
      somenteObjeto: true,
      limiteAtas: 80,
    });
    let modoBusca = "atas cujo título/objeto contém o termo pesquisado";

    if (!atas.length) {
      atas = await buscarAtasPeIntegrado({
        termo: termoBusca,
        periodoMeses,
        somenteObjeto: false,
        limiteAtas: 120,
      });
      modoBusca = "atas vigentes recentes, filtrando item por item";
    }

    const termosBusca = gerarTermosBusca(termoBusca);
    const registros = [];

    for (const ata of atas) {
      if (registros.length >= 100) break;

      let itens = [];
      try {
        itens = await buscarItensAtaPeIntegrado(ata.nCdRegistroPreco);
      } catch (error) {
        debugWarn(`Não foi possível abrir itens da ata ${ata.sNrRegistroPreco}:`, error.message);
        continue;
      }

      for (const item of itens) {
        const descricao = item.sDsProduto || "";
        const corresponde = termosBusca.some((termoBuscaItem) => textoContemTodasPalavras(descricao, termoBuscaItem));
        const valor = normalizarValor(item.dVlUnitario);

        if (!corresponde || Number.isNaN(valor) || valor <= 0) continue;

        registros.push(normalizarRegistroPeIntegrado(ata, item, termoBusca));
        if (registros.length >= 100) break;
      }
    }

    return {
      id: "peintegrado",
      nome: fonte.nome,
      configurado: true,
      erro: `Consulta em ${modoBusca}. Atas analisadas: ${atas.length}.`,
      link: fonte.link,
      registros: registros.sort((a, b) => timestampRegistro(b) - timestampRegistro(a)),
    };
  } catch (error) {
    return {
      id: "peintegrado",
      nome: fonte.nome,
      configurado: true,
      erro: `PE Integrado indisponível agora: ${error.message}`,
      link: fonte.link,
      registros: [],
    };
  }
}

function mesSicroParaNumero(nome) {
  const meses = {
    janeiro: "01",
    fevereiro: "02",
    marco: "03",
    março: "03",
    abril: "04",
    maio: "05",
    junho: "06",
    julho: "07",
    agosto: "08",
    setembro: "09",
    outubro: "10",
    novembro: "11",
    dezembro: "12",
  };

  return meses[normalizarTexto(nome)] || "00";
}

function mesOrseParaNumero(nome) {
  const meses = {
    jan: "01",
    janeiro: "01",
    fev: "02",
    fevereiro: "02",
    mar: "03",
    marco: "03",
    março: "03",
    abr: "04",
    abril: "04",
    mai: "05",
    maio: "05",
    jun: "06",
    junho: "06",
    jul: "07",
    julho: "07",
    ago: "08",
    agosto: "08",
    set: "09",
    setembro: "09",
    out: "10",
    outubro: "10",
    nov: "11",
    novembro: "11",
    dez: "12",
    dezembro: "12",
  };

  return meses[normalizarTexto(nome)] || "00";
}

function decodificarHtmlBasico(texto) {
  return String(texto || "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&#8211;/g, "-")
    .replace(/&ccedil;/gi, "ç")
    .replace(/&atilde;/gi, "ã")
    .replace(/&aacute;/gi, "á")
    .replace(/&eacute;/gi, "é")
    .replace(/&iacute;/gi, "í")
    .replace(/&oacute;/gi, "ó")
    .replace(/&uacute;/gi, "ú")
    .replace(/&acirc;/gi, "â")
    .replace(/&ecirc;/gi, "ê")
    .replace(/&ocirc;/gi, "ô")
    .replace(/&agrave;/gi, "à")
    .replace(/&ordm;/gi, "º")
    .replace(/&raquo;/gi, "»")
    .replace(/&laquo;/gi, "«")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&#039;/g, "'");
}

function limparHtml(texto) {
  return decodificarHtmlBasico(String(texto || "").replace(/<[^>]+>/g, " ")).replace(/\s+/g, " ").trim();
}

function dataReferenciaPeriodoOrse(periodo) {
  const match = String(periodo || "").match(/^(\d{4})-(\d{1,2})-/);
  if (!match) return "";
  return `${match[1]}-${String(match[2]).padStart(2, "0")}`;
}

function criarArquivoTemporario(nome) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  return path.join(DATA_DIR, `${nome}-${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2)}.tmp`);
}

function excluirArquivoSilenciosamente(caminho) {
  try {
    if (caminho && fs.existsSync(caminho)) fs.unlinkSync(caminho);
  } catch (error) {
    debugWarn("Não foi possível apagar arquivo temporário:", error.message);
  }
}

async function consultarOrseAspComCurl({ pagina = 1, body = null, timeoutSegundos = 90, cookieFile = "" } = {}) {
  const url = new URL(ORSE_INSUMOS_URL);
  if (body) url.searchParams.set("tarefa", "consultar");
  if (pagina > 1) url.searchParams.set("page", String(pagina));

  const args = [
    "-sS",
    "--http1.1",
    "-L",
    "-A",
    "Mozilla/5.0",
    "-e",
    ORSE_INSUMOS_URL,
    "-H",
    "Origin: https://orse.cehop.se.gov.br",
    "-H",
    "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "--connect-timeout",
    "20",
    "--max-time",
    String(timeoutSegundos),
  ];

  if (cookieFile) {
    args.push("-c", cookieFile, "-b", cookieFile);
  }

  if (body) {
    args.push(
      "-H",
      "Content-Type: application/x-www-form-urlencoded",
      "--data",
      body
    );
  }

  args.push(url.toString());

  const { stdout } = await executarArquivo(CURL_BIN, args);
  if (!stdout || !stdout.trim()) {
    throw new Error("curl não retornou HTML da consulta ORSE.");
  }

  return stdout;
}

async function consultarOrseAsp({ pagina = 1, body = null, timeoutMs = 90000, cookieFile = "" } = {}) {
  const url = new URL(ORSE_INSUMOS_URL);
  if (body) url.searchParams.set("tarefa", "consultar");
  if (pagina > 1) url.searchParams.set("page", String(pagina));

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const resposta = await fetch(url.toString(), {
      method: body ? "POST" : "GET",
      headers: {
        accept: "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "content-type": "application/x-www-form-urlencoded",
        origin: "https://orse.cehop.se.gov.br",
        referer: ORSE_INSUMOS_URL,
        "user-agent": "Mozilla/5.0",
      },
      body,
      signal: controller.signal,
    });

    const buffer = Buffer.from(await resposta.arrayBuffer());
    const texto = buffer.toString("latin1");
    if (!resposta.ok) throw new Error(`Status ${resposta.status}: ${texto.slice(0, 200)}`);
    if (!texto.trim()) throw new Error("HTML vazio na consulta ORSE.");
    return texto;
  } catch (error) {
    debugWarn("Fetch da consulta ORSE falhou, tentando curl:", error.message);
    return consultarOrseAspComCurl({
      pagina,
      body,
      timeoutSegundos: Math.max(30, Math.ceil(timeoutMs / 1000)),
      cookieFile,
    });
  } finally {
    clearTimeout(timeout);
  }
}

async function obterPeriodoAtualOrseInsumos(cookieFile = "") {
  const html = await consultarOrseAsp({ cookieFile });
  const match =
    html.match(/<select[^>]+name="sltPeriodo"[\s\S]*?<option\s+value="([^"]+)"[^>]*selected[^>]*>([^<]+)<\/option>/i) ||
    html.match(/<select[^>]+name="sltPeriodo"[\s\S]*?<option\s+value="([^"]+)"[^>]*>([^<]+)<\/option>/i);

  if (!match) {
    throw new Error("Não consegui identificar o período mais recente da pesquisa de insumos ORSE.");
  }

  return {
    valor: limparHtml(match[1]),
    rotulo: limparHtml(match[2]),
    dataReferencia: dataReferenciaPeriodoOrse(match[1]),
  };
}

function extrairRegistrosOrseInsumos(html, { termo, periodo }) {
  const registros = [];
  const rowRegex = /<tr>\s*<td[^>]*class="CorpoTabela"[^>]*>([\s\S]*?)<\/td>\s*<td[^>]*class="CorpoTabela"[^>]*>([\s\S]*?)<\/td>\s*<td[^>]*class="CorpoTabela"[^>]*>([\s\S]*?)<\/td>\s*<td[^>]*class="CorpoTabela"[^>]*>([\s\S]*?)<\/td>\s*<\/tr>/gi;
  let match;

  while ((match = rowRegex.exec(html))) {
    const codigo = limparHtml(match[1]);
    const descricao = limparHtml(match[2]);
    const unidade = limparHtml(match[3]);
    const precoUnitario = limparHtml(match[4]);
    const valor = normalizarValor(precoUnitario);

    if (!codigo || !descricao || Number.isNaN(valor) || valor <= 0) continue;
    if (termo && !textoContemTodasPalavras(`${codigo} ${descricao}`, termo)) continue;

    registros.push({
      origem: "ORSE-SE",
      tipoSinapi: "Insumo",
      codigo,
      descricao,
      unidade,
      precoUnitario,
      dataReferencia: periodo.dataReferencia,
      periodo: periodo.rotulo,
      uf: "SE",
      regime: "",
      link: ORSE_INSUMOS_URL,
    });
  }

  return registros;
}

function totalPaginasOrse(html) {
  const match = String(html || "").match(/Página\s+\d+\s+de\s+(\d+)/i);
  return match ? Number(match[1]) || 1 : 1;
}

async function consultarOrseInsumosOnline({ termo, limite = 100 }) {
  const cookieFile = criarArquivoTemporario("orse-cookie");

  try {
    const periodo = await obterPeriodoAtualOrseInsumos(cookieFile);
    const body = new URLSearchParams({
      sltFOnte: "ORSE",
      sltPeriodo: periodo.valor,
      sltGrupoInsumo: "0",
      rdbCriterio: "2",
      txtDescricao: termo || "",
      Submit: "Consultar",
    }).toString();

    const registros = [];
    const primeiraPagina = await consultarOrseAsp({ pagina: 1, body, cookieFile });
    registros.push(...extrairRegistrosOrseInsumos(primeiraPagina, { termo, periodo }));

    const paginas = Math.min(totalPaginasOrse(primeiraPagina), Math.ceil(limite / 12) + 1, 12);
    for (let pagina = 2; pagina <= paginas && registros.length < limite; pagina++) {
      const html = await consultarOrseAsp({ pagina, body, cookieFile });
      const itens = extrairRegistrosOrseInsumos(html, { termo, periodo });
      if (!itens.length) break;
      registros.push(...itens);
    }

    const unicos = new Map();
    for (const registro of registros) {
      if (!unicos.has(registro.codigo)) unicos.set(registro.codigo, registro);
    }

    return {
      periodo,
      registros: Array.from(unicos.values()).slice(0, limite),
    };
  } finally {
    excluirArquivoSilenciosamente(cookieFile);
  }
}

async function obterReferenciaOrseSe() {
  let html = "";

  try {
    html = await consultarTextoPublico(new URL(ORSE_BASE_DADOS_URL), 45000);
  } catch (error) {
    debugWarn("Fetch falhou ao ler ORSE, tentando curl:", error.message);
    html = await consultarTextoComCurl(new URL(ORSE_BASE_DADOS_URL), 90);
  }

  const rowRegex = /<tr[^>]*>\s*<td[^>]*>\s*<center>\s*(20\d{2})\s*<\/center>\s*<\/td>\s*<td[^>]*>\s*<center>\s*([^<]+)\s*<\/center>\s*<\/td>[\s\S]*?<a\s+href="([^"]+\.ORSE)"[^>]*>\s*<b>\s*([^<]+\.ORSE)/i;
  const match = html.match(rowRegex);

  if (!match) {
    throw new Error("Não consegui identificar a base ORSE mais recente na página oficial.");
  }

  const ano = limparHtml(match[1]);
  const mes = limparHtml(match[2]);
  const downloadUrlOriginal = new URL(decodificarHtmlBasico(match[3]), ORSE_BASE_DADOS_URL).toString();
  const downloadUrl = downloadUrlOriginal.replace(
    /^http:\/\/orse\.cehop\.se\.gov\.br\/downloads\//i,
    "https://orse.cehop.se.gov.br/downloads/"
  );
  const nomeArquivo = limparHtml(match[4]);
  const mesNumero = mesOrseParaNumero(mes);

  return {
    uf: "SE",
    estado: "SERGIPE",
    ano,
    mes,
    dataReferencia: mesNumero === "00" ? ano : `${ano}-${mesNumero}`,
    pagina: ORSE_BASE_DADOS_URL,
    fonte: ORSE_BASE_DADOS_URL,
    downloadUrl,
    nomeArquivo,
  };
}

function descreverArquivoBaixado(arquivo) {
  if (!arquivo || !fs.existsSync(arquivo)) return { tipo: "ausente", assinatura: "" };
  const buffer = fs.readFileSync(arquivo);
  const assinatura = buffer.subarray(0, 16).toString("hex").toUpperCase();
  const textoInicial = buffer.subarray(0, 120).toString("utf8");

  if (textoInicial.trim().startsWith("<")) return { tipo: "html", assinatura };
  if (assinatura.startsWith("504B0304")) return { tipo: "zip", assinatura };
  if (assinatura.startsWith("377ABCAF271C")) return { tipo: "7z", assinatura };
  if (textoInicial.includes("SQLite format")) return { tipo: "sqlite", assinatura };
  if (assinatura.startsWith("D0CF11E0")) return { tipo: "ole-office", assinatura };
  return { tipo: "desconhecido", assinatura };
}

async function atualizarOrseSe({ forcar = false } = {}) {
  const manifest = carregarManifestOrse();
  const info = await obterReferenciaOrseSe();
  const atual = manifest.SE;
  const jaBaixado =
    atual?.downloadUrl === info.downloadUrl &&
    atual?.nomeArquivo === info.nomeArquivo &&
    atual?.caminho &&
    fs.existsSync(atual.caminho);

  if (jaBaixado && !forcar) {
    return {
      atualizadoEm: new Date().toISOString(),
      resultados: [{ uf: "SE", status: "atual", ...info, caminho: atual.caminho }],
    };
  }

  const destino = path.join(ORSE_DOWNLOAD_DIR, info.ano, info.nomeArquivo);
  const cabecalhos = await obterCabecalhosHttp(new URL(info.downloadUrl), 60).catch(() => "");
  const tamanhoEsperado = obterContentLength(cabecalhos);
  const arquivoExistenteCompleto =
    fs.existsSync(destino) &&
    fs.statSync(destino).size > 0 &&
    (!tamanhoEsperado || fs.statSync(destino).size >= tamanhoEsperado);

  if (arquivoExistenteCompleto && !forcar) {
    const descricaoArquivo = descreverArquivoBaixado(destino);
    manifest.SE = {
      ...info,
      caminho: destino,
      bytes: fs.statSync(destino).size,
      ...descricaoArquivo,
      baixadoEm: atual?.baixadoEm || new Date().toISOString(),
    };
    salvarManifestOrse(manifest);

    return {
      atualizadoEm: new Date().toISOString(),
      resultados: [{ uf: "SE", status: "atual", bytes: fs.statSync(destino).size, ...descricaoArquivo, ...info, caminho: destino }],
    };
  }

  let bytes = 0;

  try {
    bytes = await baixarArquivoPublico(new URL(info.downloadUrl), destino, 240000);
  } catch (error) {
    debugWarn("Fetch falhou ao baixar ORSE, tentando curl:", error.message);
    try {
      bytes = await baixarArquivoComCurlResumivel(new URL(info.downloadUrl), destino, tamanhoEsperado, 10);
    } catch (curlError) {
      const erro = limparErroProcesso(curlError.message);
      manifest.SE = {
        ...info,
        caminho: destino,
        status: "erro_download",
        erro,
        tentadoEm: new Date().toISOString(),
      };
      salvarManifestOrse(manifest);
      throw new Error(`Servidor da CEHOP fechou o download da base ${info.nomeArquivo}: ${erro}`);
    }
  }

  const descricaoArquivo = descreverArquivoBaixado(destino);

  if (descricaoArquivo.tipo === "html") {
    throw new Error("A CEHOP retornou uma página HTML no lugar do arquivo .ORSE. Tente novamente mais tarde.");
  }

  manifest.SE = {
    ...info,
    caminho: destino,
    bytes,
    ...descricaoArquivo,
    baixadoEm: new Date().toISOString(),
  };
  salvarManifestOrse(manifest);

  return {
    atualizadoEm: new Date().toISOString(),
    resultados: [{ uf: "SE", status: "baixado", bytes, ...descricaoArquivo, ...info, caminho: destino }],
  };
}

async function importarOrseDownloads() {
  const manifest = carregarManifestOrse();
  const info = manifest.SE;
  const registros = [];
  const resumo = [];

  if (!info?.caminho || !fs.existsSync(info.caminho)) {
    fs.mkdirSync(BASES_DIR, { recursive: true });
    fs.writeFileSync(ORSE_JSON_PATH, JSON.stringify(registros, null, 2));
    return {
      atualizadoEm: new Date().toISOString(),
      arquivo: ORSE_JSON_PATH,
      totalRegistros: 0,
      resumo: [{ uf: "SE", status: "sem_download", registros: 0 }],
    };
  }

  const destino = path.join(ORSE_EXTRACT_DIR, info.ano || "sem-ano", path.basename(info.caminho, path.extname(info.caminho)));
  let arquivos = listarArquivosRecursivo(destino);
  let extracao = { ok: true, erro: "" };

  if (!arquivos.length) {
    extracao = await tentarExtrairArquivo(info.caminho, destino);
    arquivos = extracao.ok ? extracao.arquivos : [info.caminho];
  }

  const arquivosDados = arquivos.filter((arquivo) => {
    const ext = path.extname(arquivo).toLowerCase();
    return [".xlsx", ".xls", ".xlsm", ".csv"].includes(ext);
  });

  for (const arquivo of arquivosDados) {
    registros.push(...lerRegistrosDeArquivoTabela(arquivo, {
      fonteId: "orse",
      uf: "SE",
      estado: "SERGIPE",
      dataReferencia: info.dataReferencia,
      tipo: "",
      arquivo: path.basename(arquivo),
    }));
  }

  fs.mkdirSync(BASES_DIR, { recursive: true });
  fs.writeFileSync(ORSE_JSON_PATH, JSON.stringify(registros, null, 2));

  resumo.push({
    uf: "SE",
    arquivo: info.nomeArquivo,
    tipoArquivo: info.tipo || descreverArquivoBaixado(info.caminho).tipo,
    extracao: extracao.ok ? "ok" : "nao_suportada",
    erroExtracao: extracao.erro || "",
    arquivosDados: arquivosDados.length,
    registros: registros.length,
  });

  return {
    atualizadoEm: new Date().toISOString(),
    arquivo: ORSE_JSON_PATH,
    totalRegistros: registros.length,
    resumo,
  };
}

async function obterReferenciaSicroNordeste(uf = "PE") {
  const ufNormalizada = String(uf || "PE").toUpperCase();
  const nomeEstado = UF_NORDESTE_SICRO[ufNormalizada] || UF_NORDESTE_SICRO.PE;
  const html = await consultarTextoPublico(new URL(DNIT_SICRO_NORDESTE_URL), 30000);
  const texto = html.replace(/\s+/g, " ");
  const estadoIndex = texto.indexOf(nomeEstado);

  if (estadoIndex === -1) {
    throw new Error(`Não encontrei ${nomeEstado} na página oficial SICRO Nordeste.`);
  }

  const trecho = texto.slice(estadoIndex, estadoIndex + 5000);
  const match = trecho.match(/href="([^"]+)"[^>]*>\s*([^<]*(?:Janeiro|Fevereiro|Março|Marco|Abril|Maio|Junho|Julho|Agosto|Setembro|Outubro|Novembro|Dezembro)[^<]*)<\/a>/i);
  const anoMatch = trecho.match(/>\s*(20\d{2})\s*</);

  if (!match || !anoMatch) {
    throw new Error(`Não consegui identificar a referência mais recente do SICRO para ${nomeEstado}.`);
  }

  const url = new URL(match[1], DNIT_SICRO_NORDESTE_URL).toString();
  const mesNome = match[2].replace(/<[^>]+>/g, "").trim();
  const ano = anoMatch[1];

  return {
    uf: ufNormalizada,
    estado: nomeEstado,
    ano,
    mes: mesNome,
    dataReferencia: `${ano}-${mesSicroParaNumero(mesNome)}`,
    pagina: url,
    fonte: DNIT_SICRO_NORDESTE_URL,
  };
}

async function obterDownloadSicroNordeste(uf = "PE") {
  const referencia = await obterReferenciaSicroNordeste(uf);
  const html = await consultarTextoPublico(new URL(referencia.pagina), 30000);
  const match = html.match(/href="([^"]+\.7z)"/i);

  if (!match) {
    throw new Error(`Não encontrei arquivo .7z do SICRO para ${referencia.estado} ${referencia.mes}/${referencia.ano}.`);
  }

  const downloadUrl = new URL(match[1], referencia.pagina).toString();
  const nomeArquivo = path.basename(new URL(downloadUrl).pathname);

  return {
    ...referencia,
    downloadUrl,
    nomeArquivo,
  };
}

async function atualizarSicroNordeste({ forcar = false } = {}) {
  const manifest = carregarManifestSicro();
  const resultados = [];

  for (const uf of Object.keys(UF_NORDESTE_SICRO)) {
    try {
      const info = await obterDownloadSicroNordeste(uf);
      const atual = manifest[uf];
      const jaBaixado =
        atual?.downloadUrl === info.downloadUrl &&
        atual?.nomeArquivo === info.nomeArquivo &&
        atual?.caminho &&
        fs.existsSync(atual.caminho);

      if (jaBaixado && !forcar) {
        resultados.push({ uf, status: "atual", ...info, caminho: atual.caminho });
        continue;
      }

      const destino = path.join(SICRO_DOWNLOAD_DIR, uf.toLowerCase(), info.ano, info.nomeArquivo);
      const bytes = await baixarArquivoPublico(new URL(info.downloadUrl), destino);

      manifest[uf] = {
        ...info,
        caminho: destino,
        bytes,
        baixadoEm: new Date().toISOString(),
      };
      salvarManifestSicro(manifest);

      resultados.push({ uf, status: "baixado", bytes, ...info, caminho: destino });
    } catch (error) {
      resultados.push({
        uf,
        status: "erro",
        erro: error.message,
      });
    }
  }

  return {
    atualizadoEm: new Date().toISOString(),
    resultados,
  };
}

let atualizacaoSicroEmAndamento = null;
let ultimoResultadoSicro = null;
let statusAtualizacaoSicro = {
  emAndamento: false,
  iniciadoEm: "",
  finalizadoEm: "",
  mensagem: "Atualização SICRO ainda não executada nesta sessão.",
};
let atualizacaoOrseEmAndamento = null;
let ultimoResultadoOrse = null;
let statusAtualizacaoOrse = {
  emAndamento: false,
  iniciadoEm: "",
  finalizadoEm: "",
  mensagem: "Atualização ORSE-SE ainda não executada nesta sessão.",
};

function agendarAtualizacaoSicro() {
  setTimeout(() => {
    executarAtualizacaoSicroEmSegundoPlano();
  }, 3000);

  setInterval(() => {
    executarAtualizacaoSicroEmSegundoPlano();
  }, SICRO_CHECK_INTERVAL_MS);
}

function agendarAtualizacaoOrse() {
  setTimeout(() => {
    executarAtualizacaoOrseEmSegundoPlano();
  }, 9000);

  setInterval(() => {
    executarAtualizacaoOrseEmSegundoPlano();
  }, ORSE_CHECK_INTERVAL_MS);
}

function executarAtualizacaoSicroEmSegundoPlano(opcoes = {}) {
  if (atualizacaoSicroEmAndamento) return atualizacaoSicroEmAndamento;

  statusAtualizacaoSicro = {
    emAndamento: true,
    iniciadoEm: new Date().toISOString(),
    finalizadoEm: "",
    mensagem: "Verificando e baixando arquivos SICRO Nordeste em segundo plano.",
  };

  atualizacaoSicroEmAndamento = atualizarSicroNordeste(opcoes)
    .then((resultado) => {
      const baixados = resultado.resultados.filter((item) => item.status === "baixado").length;
      const erros = resultado.resultados.filter((item) => item.status === "erro").length;
      console.log(`SICRO Nordeste verificado: ${baixados} arquivo(s) baixado(s), ${erros} erro(s).`);
      if (baixados || !fs.existsSync(SICRO_JSON_PATH)) {
        return importarSicroDownloads()
          .then((importacao) => {
            console.log(`SICRO importado: ${importacao.totalRegistros} registro(s).`);
            return { ...resultado, importacao };
          })
          .catch((error) => {
            console.error("Erro ao importar SICRO:", error.message);
            return { ...resultado, importacao: { erro: error.message } };
          });
      }
      return resultado;
    })
    .then((resultado) => {
      ultimoResultadoSicro = resultado;
      statusAtualizacaoSicro = {
        emAndamento: false,
        iniciadoEm: statusAtualizacaoSicro.iniciadoEm,
        finalizadoEm: new Date().toISOString(),
        mensagem: "Atualização SICRO concluída.",
      };
      return resultado;
    })
    .catch((error) => {
      console.error("Erro ao atualizar SICRO Nordeste:", error.message);
      ultimoResultadoSicro = { erro: error.message, resultados: [] };
      statusAtualizacaoSicro = {
        emAndamento: false,
        iniciadoEm: statusAtualizacaoSicro.iniciadoEm,
        finalizadoEm: new Date().toISOString(),
        mensagem: `Erro ao atualizar SICRO: ${error.message}`,
      };
      return ultimoResultadoSicro;
    })
    .finally(() => {
      atualizacaoSicroEmAndamento = null;
    });

  return atualizacaoSicroEmAndamento;
}

function executarAtualizacaoOrseEmSegundoPlano(opcoes = {}) {
  if (atualizacaoOrseEmAndamento) return atualizacaoOrseEmAndamento;

  statusAtualizacaoOrse = {
    emAndamento: true,
    iniciadoEm: new Date().toISOString(),
    finalizadoEm: "",
    mensagem: "Verificando e baixando a base ORSE-SE em segundo plano.",
  };

  atualizacaoOrseEmAndamento = atualizarOrseSe(opcoes)
    .then((resultado) => {
      const baixados = resultado.resultados.filter((item) => item.status === "baixado").length;
      console.log(`ORSE-SE verificada: ${baixados} arquivo(s) baixado(s).`);
      if (baixados || !fs.existsSync(ORSE_JSON_PATH)) {
        return importarOrseDownloads()
          .then((importacao) => {
            console.log(`ORSE-SE importada: ${importacao.totalRegistros} registro(s).`);
            return { ...resultado, importacao };
          })
          .catch((error) => {
            console.error("Erro ao importar ORSE-SE:", error.message);
            return { ...resultado, importacao: { erro: error.message } };
          });
      }
      return resultado;
    })
    .then((resultado) => {
      ultimoResultadoOrse = resultado;
      statusAtualizacaoOrse = {
        emAndamento: false,
        iniciadoEm: statusAtualizacaoOrse.iniciadoEm,
        finalizadoEm: new Date().toISOString(),
        mensagem: "Atualização ORSE-SE concluída.",
      };
      return resultado;
    })
    .catch((error) => {
      const erro = limparErroProcesso(error.message);
      console.error("Erro ao atualizar ORSE-SE:", erro);
      ultimoResultadoOrse = { erro, resultados: [] };
      statusAtualizacaoOrse = {
        emAndamento: false,
        iniciadoEm: statusAtualizacaoOrse.iniciadoEm,
        finalizadoEm: new Date().toISOString(),
        mensagem: `Erro ao atualizar ORSE-SE: ${erro}`,
      };
      return ultimoResultadoOrse;
    })
    .finally(() => {
      atualizacaoOrseEmAndamento = null;
    });

  return atualizacaoOrseEmAndamento;
}

async function consultarSicroNordeste({ termo, uf }) {
  let referencia = null;

  try {
    referencia = await obterReferenciaSicroNordeste(uf);
  } catch (error) {
    debugWarn("Não foi possível consultar referência SICRO DNIT:", error.message);
  }

  const base = await consultarBaseLocal({ fonteId: "sicro", termo, uf });
  const erroReferencia = referencia
    ? `Referência oficial DNIT mais recente para ${referencia.estado}: ${referencia.mes}/${referencia.ano}.`
    : "Não consegui confirmar a referência mais recente no site oficial do DNIT agora.";

  return {
    ...base,
    nome: "SICRO",
    referenciaOficial: referencia,
    erro: base.configurado
      ? erroReferencia
      : `${base.erro} ${erroReferencia}`,
  };
}

async function consultarOrseSe({ termo }) {
  try {
    const online = await consultarOrseInsumosOnline({ termo, limite: 100 });
    return {
      id: "orse",
      nome: "ORSE-SE",
      configurado: true,
      erro: `Consulta online ORSE-SE. Período: ${online.periodo.rotulo}.`,
      referenciaOficial: {
        uf: "SE",
        estado: "SERGIPE",
        dataReferencia: online.periodo.dataReferencia,
        periodo: online.periodo.rotulo,
        fonte: ORSE_INSUMOS_URL,
      },
      registros: online.registros,
    };
  } catch (error) {
    debugWarn("Consulta online ORSE-SE falhou:", error.message);
  }

  let referencia = null;
  try {
    referencia = await obterReferenciaOrseSe();
  } catch (error) {
    debugWarn("Não foi possível consultar referência ORSE-SE:", error.message);
  }

  const base = await consultarBaseLocal({ fonteId: "orse", termo, uf: "SE" });
  const erroReferencia = referencia
    ? `Consulta online ORSE-SE indisponível. Usando base local, se existir. Referência de download mais recente: ${referencia.mes}/${referencia.ano}.`
    : "Consulta online ORSE-SE indisponível e não consegui confirmar a referência mais recente no site oficial da ORSE agora.";

  return {
    ...base,
    nome: "ORSE-SE",
    referenciaOficial: referencia,
    erro: base.configurado
      ? erroReferencia
      : `${base.erro} ${erroReferencia}`,
  };
}

function criarItemBuscaDiretaPncp(termo, tipo = "material") {
  const descricao = String(termo || "").trim();

  return {
    tipo: "pncp-direto",
    codigo: descricao,
    descricao: `Buscar direto no PNCP: ${descricao}`,
    codigoGrupo: "",
    grupo: "",
    codigoClasse: "",
    classe: "",
    codigoPdm: "",
    pdm: descricao,
    unidade: "",
    status: "Busca por termo",
    origemTipo: tipo,
    diretoPncp: true,
  };
}

async function buscarItensMaterialPorClassesRelacionadas(termosBusca) {
  const palavrasDaBusca = Array.from(new Set(termosBusca.flatMap(palavrasRelevantes)));
  const textoBusca = palavrasDaBusca.join(" ");
  const classes = CLASSES_MATERIAL_RELACIONADAS.filter((classe) =>
    classe.termos.some((termo) => textoContemTodasPalavras(textoBusca, termo))
  );

  if (!classes.length) return [];

  const encontrados = [];
  const tamanhoPagina = 500;
  const maxPaginasPorClasse = 8;

  for (const classe of classes) {
    for (let pagina = 1; pagina <= maxPaginasPorClasse && encontrados.length < 100; pagina++) {
      const url = new URL(BASE + "/modulo-material/4_consultarItemMaterial");
      url.searchParams.set("pagina", String(pagina));
      url.searchParams.set("tamanhoPagina", String(tamanhoPagina));
      url.searchParams.set("codigoClasse", String(classe.codigoClasse));

      let dados;
      try {
        dados = await consultarJson(url, 30000);
      } catch (error) {
        debugWarn(`Não foi possível consultar classe ${classe.codigoClasse} na página ${pagina}:`, error.message);
        break;
      }

      const resultado = dados.resultado || [];

      const itensFiltrados = resultado.filter((item) => {
        const texto = `${item.nomePdm || ""} ${item.descricaoItem || ""}`;

        return termosBusca.some((termoBusca) => textoContemTodasPalavras(texto, termoBusca));
      });

      encontrados.push(...itensFiltrados);

      if (!dados.paginasRestantes || dados.paginasRestantes <= 0) {
        break;
      }
    }
  }

  return encontrados;
}

async function buscarPdmsCnbsPorPalavra(termosBusca) {
  const mapa = new Map();

  for (const termoBusca of termosBusca) {
    const url = new URL(CNBS_BASE + "/material/v1/palavra");
    url.searchParams.set("palavra", termoBusca);
    url.searchParams.set("apenasAtivos", "nao");

    try {
      const dados = await consultarCnbsJson(url, 30000);
      const pdms = Array.isArray(dados) ? dados : [];

      pdms.forEach((pdm) => {
        const codigoPdm = pdm.codigoPdm || pdm.codigoPDM;
        if (!codigoPdm || mapa.has(String(codigoPdm))) return;

        mapa.set(String(codigoPdm), {
          codigoPdm,
          nomePdm: pdm.nomePdm || pdm.descricaoPDM || "",
          codigoClasse: pdm.codigoClasse || "",
          nomeClasse: pdm.nomeClasse || pdm.descricaoClasse || "",
          codigoGrupo: pdm.codigoGrupo || "",
          nomeGrupo: pdm.nomeGrupo || pdm.descricaoGrupo || "",
          statusPDM: pdm.statusPDM,
          textoBusca: normalizarTexto(
            `${pdm.nomePdm || ""} ${pdm.descricaoPDM || ""} ${pdm.nomeClasse || ""} ${pdm.descricaoClasse || ""} ${pdm.nomeGrupo || ""} ${pdm.descricaoGrupo || ""}`
          ),
        });
      });
    } catch (error) {
      debugWarn(`Não foi possível consultar CNBS material para "${termoBusca}":`, error.message);
    }
  }

  return Array.from(mapa.values()).sort(
    (a, b) => pontuarResultadoBusca(b.textoBusca, termosBusca) - pontuarResultadoBusca(a.textoBusca, termosBusca)
  );
}

async function buscarItensMaterialPorPdms(pdms, limite = 120) {
  const encontrados = [];
  const tamanhoLote = 2;
  let falhasSeguidas = 0;

  for (let inicio = 0; inicio < pdms.length && encontrados.length < limite && falhasSeguidas < 3; inicio += tamanhoLote) {
    const lote = pdms.slice(inicio, inicio + tamanhoLote).filter((pdm) => pdm.codigoPdm);

    const consultas = lote.map(async (pdm) => {
      const urlItensPdm = new URL(BASE + "/modulo-material/4_consultarItemMaterial");
      urlItensPdm.searchParams.set("pagina", "1");
      urlItensPdm.searchParams.set("tamanhoPagina", "100");
      urlItensPdm.searchParams.set("codigoPdm", pdm.codigoPdm);

      try {
        const dadosItensPdm = await consultarJson(urlItensPdm, 30000);
        return {
          ok: true,
          itens: dadosItensPdm.resultado || [],
        };
      } catch (error) {
        debugWarn(`Não foi possível consultar itens do PDM ${pdm.codigoPdm}:`, error.message);
        return {
          ok: false,
          itens: [],
          erro: error,
        };
      }
    });

    const resultados = await Promise.all(consultas);
    const loteComSucesso = resultados.some((resultado) => resultado.ok);
    falhasSeguidas = loteComSucesso ? 0 : falhasSeguidas + lote.length;

    for (const resultado of resultados) {
      encontrados.push(...resultado.itens);
      if (encontrados.length >= limite) break;
    }
  }

  return encontrados;
}

async function carregarCachePdmMaterial() {
  if (cachePdmMaterial) return cachePdmMaterial;

  debugLog("Carregando cache de PDMs de material. Aguarde...");

  const todos = [];
  const tamanhoPagina = 500;
  const maxPaginas = 400;

  for (let pagina = 1; pagina <= maxPaginas; pagina++) {
    const url = new URL(BASE + "/modulo-material/3_consultarPdmMaterial");
    url.searchParams.set("pagina", String(pagina));
    url.searchParams.set("tamanhoPagina", String(tamanhoPagina));

    const dados = await consultarJson(url, 30000);
    const resultado = dados.resultado || [];

    todos.push(...resultado);

    debugLog(`PDM material página ${pagina} | recebidos: ${resultado.length} | total cache: ${todos.length}`);

    if (!dados.paginasRestantes || dados.paginasRestantes <= 0) {
      break;
    }
  }

  cachePdmMaterial = todos.map((pdm) => ({
    ...pdm,
    textoBusca: normalizarTexto(
      `${pdm.nomePdm || ""} ${pdm.nomeClasse || ""} ${pdm.nomeGrupo || ""}`
    ),
  }));

  debugLog("Cache PDM material carregado:", cachePdmMaterial.length);

  return cachePdmMaterial;
}

async function carregarCacheServico() {
  if (cacheServico) return cacheServico;

  debugLog("Carregando cache de serviços. Aguarde...");

  const todos = [];
  const tamanhoPagina = 500;
  const maxPaginas = 300;

  for (let pagina = 1; pagina <= maxPaginas; pagina++) {
    const url = new URL(BASE + "/modulo-servico/6_consultarItemServico");
    url.searchParams.set("pagina", String(pagina));
    url.searchParams.set("tamanhoPagina", String(tamanhoPagina));

    const dados = await consultarJson(url, 30000);
    const resultado = dados.resultado || [];

    todos.push(...resultado);

    debugLog(`Serviços página ${pagina} | recebidos: ${resultado.length} | total cache: ${todos.length}`);

    if (!dados.paginasRestantes || dados.paginasRestantes <= 0) {
      break;
    }
  }

  cacheServico = todos.map((item) => ({
    ...item,
    textoBusca: normalizarTexto(
      `${item.nomeServico || ""} ${item.nomeGrupo || ""} ${item.nomeClasse || ""} ${item.nomeSubclasse || ""}`
    ),
  }));

  debugLog("Cache serviço carregado:", cacheServico.length);

  return cacheServico;
}

app.get("/api/buscar-itens", async (req, res) => {
  try {
    const termo = req.query.q;
    const tipo = req.query.tipo || "material";
    const permitirPncpDireto = req.query.pncpDireto !== "nao";
    const limiteItens = normalizarLimiteItens(req.query.limiteItens);

    if (!termo) {
      return res.status(400).json({
        erro: "Informe uma descrição para buscar.",
      });
    }

    const termosBusca = gerarTermosBusca(termo);
    const cacheKeyItens = chaveCache("itens", tipo, termo);
    const cacheItens = obterCache("itens", cacheKeyItens);
    const falhaRecente = obterFalhaBuscaItens(cacheKeyItens);

    if (falhaRecente) {
      if (cacheItens?.payload?.itens?.length) {
        const itensComFallback = permitirPncpDireto
          ? [...cacheItens.payload.itens, criarItemBuscaDiretaPncp(termo, tipo)]
          : cacheItens.payload.itens;

        return res.json({
          termo,
          tipo,
          total: itensComFallback.length,
          itens: itensComFallback,
          aviso: `O catálogo do Compras.gov oscilou há pouco (${falhaRecente.mensagem}). Mostrei dados locais e evitei repetir a chamada por alguns instantes.`,
          cache: true,
          cacheAtualizadoEm: cacheItens.atualizadoEm,
        });
      }

      return res.json({
        termo,
        tipo,
        total: permitirPncpDireto ? 1 : 0,
        itens: permitirPncpDireto ? [criarItemBuscaDiretaPncp(termo, tipo)] : [],
        aviso: `O catálogo do Compras.gov oscilou há pouco (${falhaRecente.mensagem}). Aguarde um pouco ou pesquise nas outras fontes ativas.`,
      });
    }

    if (tipo === "material") {
      let pdmsEncontrados = await buscarPdmsCnbsPorPalavra(termosBusca);

      if (!pdmsEncontrados.length) {
        const cache = await carregarCachePdmMaterial();

        pdmsEncontrados = ordenarItensPorAderencia(
          cache.filter((pdm) =>
            termosBusca.some((termoBusca) =>
              textoContemTodasPalavras(pdm.textoBusca, termoBusca)
            )
          ),
          termosBusca,
          (pdm) => pdm.nomePdm || pdm.textoBusca || ""
        );
      }

      pdmsEncontrados = ordenarItensPorAderencia(
        pdmsEncontrados,
        termosBusca,
        (pdm) => pdm.nomePdm || pdm.textoBusca || ""
      ).slice(0, 25);

      debugLog("PDMs encontrados:", pdmsEncontrados.length);

      if (req.query.rapido !== "nao" && pdmsEncontrados.length) {
        const itensPdm = pdmsEncontrados.slice(0, 20).map((pdm) => ({
          tipo: "material-pdm",
          codigo: pdm.codigoPdm,
          descricao: pdm.nomePdm || "PDM sem descrição",
          codigoGrupo: pdm.codigoGrupo || "",
          grupo: pdm.nomeGrupo || "",
          codigoClasse: pdm.codigoClasse || "",
          classe: pdm.nomeClasse || "",
          codigoPdm: pdm.codigoPdm || "",
          pdm: pdm.nomePdm || "",
          unidade: "",
          status: pdm.statusPDM,
          somenteSinapi: true,
        }));

        return res.json({
          termo,
          tipo,
          total: itensPdm.length,
          itens: itensPdm,
          aviso: "Resultado rápido por PDM. A consulta PNCP depende do código do item catálogo; o SINAPI usa a descrição do PDM.",
        });
      }

      let itensEncontrados = await buscarItensMaterialPorPdms(pdmsEncontrados, 150);

      if (itensEncontrados.length < 20) {
        const itensPorClasse = await buscarItensMaterialPorClassesRelacionadas(termosBusca);
        itensEncontrados.push(...itensPorClasse);
      }

      const mapa = new Map();

      itensEncontrados.forEach((item) => {
        if (!item.codigoItem) return;

        if (!mapa.has(item.codigoItem)) {
          mapa.set(item.codigoItem, item);
        }
      });

      const itens = ordenarItensPorAderencia(
          Array.from(mapa.values()),
          termosBusca,
          (item) => item.descricaoItem || item.nomePdm || ""
        )
        .slice(0, limiteItens)
        .map((item) => ({
        tipo: "material",
        codigo: item.codigoItem,
        descricao: item.descricaoItem || item.nomePdm || "Sem descrição",
        codigoGrupo: item.codigoGrupo || "",
        grupo: item.nomeGrupo || "",
        codigoClasse: item.codigoClasse || "",
        classe: item.nomeClasse || "",
        codigoPdm: item.codigoPdm || "",
        pdm: item.nomePdm || "",
        unidade:
          item.nomeUnidadeFornecimento ||
          item.siglaUnidadeFornecimento ||
          item.siglaUnidadeMedida ||
          item.nomeUnidadeMedida ||
          "",
        status: item.statusItem,
      }));

      if (!itens.length) {
        if (cacheItens?.payload?.itens?.length) {
          const itensComFallback = permitirPncpDireto
            ? [...cacheItens.payload.itens, criarItemBuscaDiretaPncp(termo, tipo)]
            : cacheItens.payload.itens;

          return res.json({
            termo,
            tipo,
            total: itensComFallback.length,
            itens: itensComFallback,
            aviso: "O catálogo não respondeu com itens novos. Mostrei os últimos itens salvos localmente.",
            cache: true,
            cacheAtualizadoEm: cacheItens.atualizadoEm,
          });
        }

        if (!permitirPncpDireto) {
          return res.json({
            termo,
            tipo,
            total: 0,
            itens: [],
            aviso: "O catálogo ainda não retornou itens.",
          });
        }

        const itemPncpDireto = criarItemBuscaDiretaPncp(termo, tipo);

        return res.json({
          termo,
          tipo,
          total: 1,
          itens: [itemPncpDireto],
          aviso: "O catálogo não retornou itens. A busca direta no PNCP foi liberada como último recurso.",
        });
      }

      gravarCache("itens", cacheKeyItens, { itens });

      return res.json({
        termo,
        tipo,
        total: itens.length,
        limiteItens,
        itens,
      });
    }

    if (tipo === "servico") {
      const cache = await carregarCacheServico();

      const encontrados = cache
        .filter((item) =>
          termosBusca.some((termoBusca) =>
            textoContemTodasPalavras(item.textoBusca, termoBusca)
          )
        )
        .slice(0, limiteItens);

      const itens = encontrados.map((item) => ({
        tipo: "servico",
        codigo: item.codigoServico,
        descricao: item.nomeServico,
        codigoGrupo: item.codigoGrupo || "",
        grupo: item.nomeGrupo || "",
        codigoClasse: item.codigoClasse || "",
        classe: item.nomeClasse || "",
        codigoSubclasse: item.codigoSubclasse || "",
        subclasse: item.nomeSubclasse || "",
        unidade: item.nomeUnidadeMedida || item.siglaUnidadeMedida || "",
      }));

      if (!itens.length) {
        if (cacheItens?.payload?.itens?.length) {
          const itensComFallback = permitirPncpDireto
            ? [...cacheItens.payload.itens, criarItemBuscaDiretaPncp(termo, tipo)]
            : cacheItens.payload.itens;

          return res.json({
            termo,
            tipo,
            total: itensComFallback.length,
            itens: itensComFallback,
            aviso: "O catálogo não respondeu com serviços novos. Mostrei os últimos serviços salvos localmente.",
            cache: true,
            cacheAtualizadoEm: cacheItens.atualizadoEm,
          });
        }

        if (!permitirPncpDireto) {
          return res.json({
            termo,
            tipo,
            total: 0,
            itens: [],
            aviso: "O catálogo ainda não retornou serviços.",
          });
        }

        const itemPncpDireto = criarItemBuscaDiretaPncp(termo, tipo);

        return res.json({
          termo,
          tipo,
          total: 1,
          itens: [itemPncpDireto],
          aviso: "O catálogo não retornou serviços. A busca direta no PNCP foi liberada como último recurso.",
        });
      }

      gravarCache("itens", cacheKeyItens, { itens });

      return res.json({
        termo,
        tipo,
        total: itens.length,
        limiteItens,
        itens,
      });
    }

    return res.status(400).json({
      erro: "Tipo inválido. Use material ou servico.",
    });
  } catch (error) {
    const cacheKeyItens = chaveCache("itens", req.query.tipo || "material", req.query.q || "");
    const cacheItens = obterCache("itens", cacheKeyItens);
    const mensagemAmigavel = mensagemErroApiPublica(error);

    if (apiCatalogoInstavel(error)) {
      gravarFalhaBuscaItens(cacheKeyItens, error);
    }

    if (DEBUG_LOGS) {
      console.error("ERRO DETALHADO NA BUSCA:", error);
    } else {
      console.warn("Catálogo Compras.gov indisponível:", mensagemAmigavel);
    }

    if (req.query.pncpDireto === "nao") {
      if (cacheItens?.payload?.itens?.length) {
        return res.json({
          termo: req.query.q || "",
          tipo: req.query.tipo || "material",
          total: cacheItens.payload.itens.length,
          itens: cacheItens.payload.itens,
          aviso: `A API oscilou: ${mensagemAmigavel}. Mostrei itens salvos localmente.`,
          cache: true,
          cacheAtualizadoEm: cacheItens.atualizadoEm,
        });
      }

      return res.json({
        termo: req.query.q || "",
        tipo: req.query.tipo || "material",
        total: 0,
        itens: [],
        aviso: `O catálogo ainda está indisponível: ${mensagemAmigavel}`,
      });
    }

    if (cacheItens?.payload?.itens?.length) {
      const itemPncpDireto = criarItemBuscaDiretaPncp(req.query.q || "", req.query.tipo || "material");
      return res.json({
        termo: req.query.q || "",
        tipo: req.query.tipo || "material",
        total: cacheItens.payload.itens.length + 1,
        itens: [...cacheItens.payload.itens, itemPncpDireto],
        aviso: `A API oscilou: ${mensagemAmigavel}. Mostrei itens salvos localmente e mantive PNCP direto como último recurso.`,
        cache: true,
        cacheAtualizadoEm: cacheItens.atualizadoEm,
      });
    }

    res.json({
      termo: req.query.q || "",
      tipo: req.query.tipo || "material",
      total: 1,
      itens: [criarItemBuscaDiretaPncp(req.query.q || "", req.query.tipo || "material")],
      aviso: `O catálogo ficou indisponível agora: ${mensagemAmigavel}. A busca direta no PNCP foi liberada como último recurso.`,
    });
  }
});

app.get("/api/precos", async (req, res) => {
  try {
    const codigo = req.query.codigo || req.query.termo || "";
    const tipo = req.query.tipo || "material";
    const tamanhoPagina = req.query.tamanhoPagina || 500;
    const termoSinapi = req.query.termo || "";
    const sinapiUf = req.query.sinapiUf || "PB";
    const sinapiDataReferencia = req.query.sinapiDataReferencia || dataReferenciaPadrao();
    const sinapiRegime = req.query.sinapiRegime || "NAO_DESONERADO";
    const incluirSinapi = req.query.incluirSinapi === "1";
    const incluirPncp = req.query.incluirPncp !== "0";
    const fontesAdicionais = String(req.query.fontesAdicionais || "")
      .split(",")
      .map((fonte) => fonte.trim().toLowerCase())
      .filter(Boolean);
    const periodoMeses = normalizarPeriodoMeses(req.query.periodoMeses);

    if (!codigo && incluirPncp) {
      return res.status(400).json({
        erro: "Informe o código do item para consultar o PNCP.",
      });
    }

    let resultado = [];
    let erroPncp = "";
    let usandoCachePrecos = false;
    let cacheAtualizadoEm = "";
    const cacheKeyPrecos = chaveCache("precos", tipo, codigo, termoSinapi, periodoMeses);
    const cachePrecos = obterCache("precos", cacheKeyPrecos);

    if (incluirPncp && tipo !== "material-pdm" && tipo !== "pncp-direto" && !resultado.length) {
      const endpoint =
        tipo === "servico"
          ? "/modulo-pesquisa-preco/3_consultarServico"
          : "/modulo-pesquisa-preco/1_consultarMaterial";

      const url = new URL(BASE + endpoint);

      url.searchParams.set("pagina", "1");
      url.searchParams.set("tamanhoPagina", tamanhoPagina);
      url.searchParams.set("codigoItemCatalogo", codigo);

      try {
        const dados = await consultarJson(url, 45000);
        const resultadoOriginal = (dados.resultado || []).filter((registro) => {
          const valorUnitario = normalizarValor(registro.precoUnitario);
          return !Number.isNaN(valorUnitario) && valorUnitario > 0;
        });

        const resultadoBase = resultadoOriginal.map((registro) => {
          const licitacao = extrairIdentificacaoLicitacao(registro);

          return {
            ...registro,
            licitacao,
            linkPncpGerado: licitacao.linkPncp,
          };
        });

        resultado = (await enriquecerRegistrosComPncp(resultadoBase, 50))
          .filter((registro) => linkPncpValido(registro.linkPncpGerado))
          .sort(
          (a, b) => timestampRegistro(b) - timestampRegistro(a)
        );

        if (resultado.length) {
          erroPncp = "";
        }

        debugLog("Campos do primeiro registro de preço:", Object.keys(resultado[0] || {}));
        debugLog("Primeiro registro de preço:", resultado[0]);
      } catch (error) {
        erroPncp = erroPncp || mensagemErroApiPublica(error);
        debugWarn("PNCP/Compras.gov indisponível:", error.message);
      }
    }

    if (incluirPncp && tipo === "pncp-direto") {
      try {
        resultado = await buscarPrecosPncpDireto({
          termo: termoSinapi || codigo,
          codigo,
          tipo,
          limite: 80,
        });

        if (resultado.length) {
          erroPncp = "";
        }
      } catch (error) {
        erroPncp = erroPncp || mensagemErroApiPublica(error);
        debugWarn("PNCP direto indisponível:", error.message);
      }
    }

    resultado = filtrarRegistrosPorPeriodo(resultado, periodoMeses)
      .sort((a, b) => timestampRegistro(b) - timestampRegistro(a));

    if (resultado.length) {
      gravarCache("precos", cacheKeyPrecos, { registros: resultado });
    } else if (incluirPncp && cachePrecos?.payload?.registros?.length) {
      resultado = cachePrecos.payload.registros;
      usandoCachePrecos = true;
      cacheAtualizadoEm = cachePrecos.atualizadoEm;
      erroPncp = erroPncp
        ? `A API pública oscilou (${erroPncp}). Mostrei os últimos preços salvos localmente.`
        : "A consulta atual não retornou preços. Mostrei os últimos preços salvos localmente.";
    }

    const registrosPncp = [];
    const sinapi = incluirSinapi
      ? await consultarPrecosSinapi({
          termo: termoSinapi,
          uf: sinapiUf,
          dataReferencia: sinapiDataReferencia,
          regime: sinapiRegime,
        })
      : {
          configurado: SINAPI_HABILITADO,
          desativado: true,
          erro: "",
          termoBusca: termoSinapi,
          registros: [],
        };
    const fontes = [];

    for (const fonteId of fontesAdicionais) {
      if (fonteId === "sinapi") continue;
      if (fonteId === "sicro") {
        fontes.push(await consultarSicroNordeste({
          termo: termoSinapi,
          uf: sinapiUf,
        }));
      } else if (fonteId === "orse") {
        fontes.push(await consultarOrseSe({
          termo: termoSinapi,
        }));
      } else if (fonteId === "peintegrado") {
        fontes.push(await consultarPeIntegrado({
          termo: termoSinapi,
          periodoMeses,
        }));
      } else if (FONTES_LOCAIS[fonteId]) {
        fontes.push(await consultarBaseLocal({
          fonteId,
          termo: termoSinapi,
          uf: sinapiUf,
        }));
      } else if (FONTES_EXTERNAS[fonteId]) {
        fontes.push(consultarFonteExternaPendente(fonteId));
      }
    }

    const estatisticas = calcularEstatisticas(resultado);

    res.json({
      codigo,
      tipo,
      incluirPncp,
      periodoMeses,
      totalRegistros: resultado.length,
      totalRegistrosPncp: registrosPncp.length,
      erroPncp,
      cache: usandoCachePrecos,
      cacheAtualizadoEm,
      estatisticas,
      registros: resultado,
      registrosPncp,
      sinapi,
      fontes,
    });
  } catch (error) {
    const codigo = req.query.codigo || "";
    const tipo = req.query.tipo || "material";
    const termoSinapi = req.query.termo || "";
    const periodoMeses = normalizarPeriodoMeses(req.query.periodoMeses);
    const cachePrecos = obterCache("precos", chaveCache("precos", tipo, codigo, termoSinapi, periodoMeses));

    if (cachePrecos?.payload?.registros?.length) {
      const registros = cachePrecos.payload.registros;
      return res.json({
        codigo,
        tipo,
        periodoMeses,
        totalRegistros: registros.length,
        totalRegistrosPncp: 0,
        erroPncp: `A consulta atual falhou (${mensagemErroApiPublica(error)}). Mostrei os últimos preços salvos localmente.`,
        cache: true,
        cacheAtualizadoEm: cachePrecos.atualizadoEm,
        estatisticas: calcularEstatisticas(registros),
        registros,
        registrosPncp: [],
        sinapi: {
          configurado: SINAPI_HABILITADO,
          erro: "",
          registros: [],
        },
        fontes: [],
      });
    }

    if (DEBUG_LOGS) {
      console.error("ERRO DETALHADO NA CONSULTA DE PREÇOS:", error);
    } else {
      console.error("Erro na consulta de preços:", error.message);
    }

    res.status(500).json({
      erro: "Erro ao consultar preços.",
      detalhe: error.message,
    });
  }
});

app.get("/api/precos-fonte", async (req, res) => {
  const fonteId = String(req.query.fonte || "").trim().toLowerCase();
  const termo = req.query.termo || req.query.codigo || "";
  const uf = req.query.sinapiUf || "PB";
  const periodoMeses = normalizarPeriodoMeses(req.query.periodoMeses);

  try {
    if (!fonteId) {
      return res.status(400).json({
        erro: "Informe a fonte a consultar.",
      });
    }

    if (fonteId === "sinapi") {
      return res.json(await consultarPrecosSinapi({
        termo,
        uf,
        dataReferencia: req.query.sinapiDataReferencia || dataReferenciaPadrao(),
        regime: req.query.sinapiRegime || "NAO_DESONERADO",
      }));
    }

    if (fonteId === "sicro") {
      return res.json(await consultarSicroNordeste({ termo, uf }));
    }

    if (fonteId === "orse") {
      return res.json(await consultarOrseSe({ termo }));
    }

    if (fonteId === "peintegrado") {
      return res.json(await consultarPeIntegrado({ termo, periodoMeses }));
    }

    if (FONTES_LOCAIS[fonteId]) {
      return res.json(await consultarBaseLocal({ fonteId, termo, uf }));
    }

    if (FONTES_EXTERNAS[fonteId]) {
      return res.json(consultarFonteExternaPendente(fonteId));
    }

    return res.status(400).json({
      erro: "Fonte inválida.",
    });
  } catch (error) {
    const fonte = FONTES_LOCAIS[fonteId] || FONTES_EXTERNAS[fonteId] || { nome: fonteId };
    return res.json({
      id: fonteId,
      nome: fonte.nome || fonteId,
      configurado: true,
      erro: error.message,
      registros: [],
    });
  }
});

function formatarMoedaPdf(valor) {
  const numero = normalizarValor(valor);
  if (Number.isNaN(numero)) return "-";
  return numero.toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
  });
}

function textoPdf(valor, fallback = "-") {
  return String(valor ?? "").trim() || fallback;
}

app.post("/api/relatorio-pdf", (req, res) => {
  const selecionados = Array.isArray(req.body?.selecionados) ? req.body.selecionados : [];

  if (!selecionados.length) {
    return res.status(400).json({
      erro: "Selecione ao menos um preço para gerar o relatório.",
    });
  }

  const doc = new PDFDocument({
    size: "A4",
    margin: 36,
    bufferPages: true,
  });

  const nomeArquivo = `relatorio-pesquisa-precos-${new Date().toISOString().slice(0, 10)}.pdf`;
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="${nomeArquivo}"`);
  doc.pipe(res);

  doc.font("Helvetica-Bold").fontSize(16).text("Relatório de Pesquisa de Preços", { align: "center" });
  doc.moveDown(0.4);
  doc.font("Helvetica").fontSize(9).text(`Gerado em ${new Date().toLocaleString("pt-BR")}`, { align: "center" });
  doc.moveDown(1);

  const resumoPorOrigem = selecionados.reduce((acc, item) => {
    const origem = item.origem || "Fonte não informada";
    acc[origem] = (acc[origem] || 0) + 1;
    return acc;
  }, {});

  doc.font("Helvetica-Bold").fontSize(11).text("Resumo");
  doc.font("Helvetica").fontSize(9).text(`Itens selecionados: ${selecionados.length}`);
  doc.text(
    Object.entries(resumoPorOrigem)
      .map(([origem, total]) => `${origem}: ${total}`)
      .join(" | ")
  );
  doc.moveDown(0.8);

  selecionados.forEach((item, index) => {
    if (doc.y > 690) doc.addPage();

    const x = doc.page.margins.left;
    const y = doc.y;
    const largura = doc.page.width - doc.page.margins.left - doc.page.margins.right;

    doc.roundedRect(x, y, largura, 112, 6).stroke("#d1d5db");
    doc.font("Helvetica-Bold").fontSize(10).fillColor("#111827").text(`${index + 1}. ${textoPdf(item.descricao)}`, x + 10, y + 10, {
      width: largura - 20,
      height: 28,
      ellipsis: true,
    });

    doc.font("Helvetica").fontSize(8).fillColor("#374151");
    doc.text(`Fonte: ${textoPdf(item.origem)}${item.tipoSinapi ? ` - ${item.tipoSinapi}` : ""}`, x + 10, y + 42, { width: 250 });
    doc.text(`Código: ${textoPdf(item.codigo)}`, x + 10, y + 56, { width: 160 });
    doc.text(`Unidade: ${textoPdf(item.unidade)}`, x + 180, y + 56, { width: 120 });
    doc.text(`Preço unitário: ${formatarMoedaPdf(item.precoUnitario)}`, x + 330, y + 56, { width: 170 });
    doc.text(`Item PNCP: ${item.numeroItem ? `Item nº ${item.numeroItem}` : "-"}`, x + 10, y + 70, { width: 160 });
    doc.text(`Quantidade: ${textoPdf(item.quantidade)}`, x + 180, y + 70, { width: 120 });
    doc.text(`Fornecedor/Órgão: ${textoPdf(item.fornecedor || item.orgao)}`, x + 330, y + 70, { width: 170 });
    doc.text(`Data: ${textoPdf(item.data)}`, x + 10, y + 84, { width: 150 });
    doc.text(`Licitação/Referência: ${textoPdf(item.licitacao || item.referencia)}`, x + 180, y + 84, { width: 320 });

    if (item.link) {
      doc.fillColor("#1d4ed8").text("Abrir referência", x + 10, y + 98, {
        link: item.link,
        underline: true,
      });
      doc.fillColor("#374151");
    }

    doc.y = y + 126;
  });

  const pages = doc.bufferedPageRange();
  for (let i = 0; i < pages.count; i++) {
    doc.switchToPage(i);
    doc.font("Helvetica").fontSize(8).fillColor("#6b7280").text(
      `Página ${i + 1} de ${pages.count}`,
      doc.page.margins.left,
      doc.page.height - 28,
      { align: "center" }
    );
  }

  doc.end();
});

app.get("/api/limpar-cache", (req, res) => {
  cachePdmMaterial = null;
  cacheServico = null;

  res.json({
    ok: true,
    mensagem: "Cache limpo com sucesso.",
  });
});

app.get("/api/sicro/atualizar", async (req, res) => {
  const forcar = req.query.forcar === "1";
  const aguardar = req.query.aguardar === "1";
  const tarefa = executarAtualizacaoSicroEmSegundoPlano({ forcar });

  if (aguardar) {
    const resultado = await tarefa;
    return res.json(resultado);
  }

  res.status(202).json({
    ok: true,
    emAndamento: true,
    mensagem: "Atualização SICRO iniciada em segundo plano. Acompanhe em /api/sicro/status.",
    statusUrl: "/api/sicro/status",
  });
});

app.get("/api/sicro/status", (req, res) => {
  res.json({
    emAndamento: Boolean(atualizacaoSicroEmAndamento),
    status: statusAtualizacaoSicro,
    ultimoResultado: ultimoResultadoSicro,
    manifest: carregarManifestSicro(),
    pasta: SICRO_DOWNLOAD_DIR,
    basePesquisavel: {
      arquivo: SICRO_JSON_PATH,
      existe: fs.existsSync(SICRO_JSON_PATH),
      tamanho: fs.existsSync(SICRO_JSON_PATH) ? fs.statSync(SICRO_JSON_PATH).size : 0,
    },
    baseLeve: {
      arquivo: SICRO_LITE_JSON_PATH,
      existe: fs.existsSync(SICRO_LITE_JSON_PATH),
      tamanho: fs.existsSync(SICRO_LITE_JSON_PATH) ? fs.statSync(SICRO_LITE_JSON_PATH).size : 0,
    },
    fonte: DNIT_SICRO_NORDESTE_URL,
  });
});

app.get("/api/sicro/importar", async (req, res) => {
  const resultado = await importarSicroDownloads();
  res.json(resultado);
});

app.get("/api/sicro/gerar-lite", (req, res) => {
  if (!fs.existsSync(SICRO_JSON_PATH)) {
    return res.status(404).json({
      erro: "Base SICRO principal não encontrada. Gere primeiro o sicro.json localmente.",
    });
  }

  const uf = String(req.query.uf || "").trim().toUpperCase();
  const bruto = JSON.parse(fs.readFileSync(SICRO_JSON_PATH, "utf8"));
  const registros = compactarRegistrosSicro(bruto, uf);

  fs.mkdirSync(BASES_DIR, { recursive: true });
  fs.writeFileSync(SICRO_LITE_JSON_PATH, JSON.stringify(registros, null, 2));

  res.json({
    ok: true,
    arquivo: SICRO_LITE_JSON_PATH,
    totalRegistros: registros.length,
    uf: uf || "TODAS",
    tamanho: fs.statSync(SICRO_LITE_JSON_PATH).size,
  });
});

app.get("/api/orse/atualizar", async (req, res) => {
  const forcar = req.query.forcar === "1";
  const aguardar = req.query.aguardar === "1";
  const tarefa = executarAtualizacaoOrseEmSegundoPlano({ forcar });

  if (aguardar) {
    const resultado = await tarefa;
    return res.json(resultado);
  }

  res.status(202).json({
    ok: true,
    emAndamento: true,
    mensagem: "Atualização ORSE-SE iniciada em segundo plano. Acompanhe em /api/orse/status.",
    statusUrl: "/api/orse/status",
  });
});

app.get("/api/orse/status", (req, res) => {
  res.json({
    emAndamento: Boolean(atualizacaoOrseEmAndamento),
    status: statusAtualizacaoOrse,
    ultimoResultado: ultimoResultadoOrse,
    manifest: carregarManifestOrse(),
    pasta: ORSE_DOWNLOAD_DIR,
    basePesquisavel: {
      arquivo: ORSE_JSON_PATH,
      existe: fs.existsSync(ORSE_JSON_PATH),
      tamanho: fs.existsSync(ORSE_JSON_PATH) ? fs.statSync(ORSE_JSON_PATH).size : 0,
    },
    fonte: ORSE_BASE_DADOS_URL,
  });
});

app.get("/api/orse/importar", async (req, res) => {
  const resultado = await importarOrseDownloads();
  res.json(resultado);
});

app.get("/api/health", (req, res) => {
  res.json({
    ok: true,
    ambiente: {
      render: IS_RENDER,
      nodeEnv: process.env.NODE_ENV || "",
      autoUpdateBases: AUTO_UPDATE_BASES,
      sinapiHabilitado: SINAPI_HABILITADO,
      sinapiConfigurado: Boolean(SINAPI_API_URL),
    },
    fontes: {
      pncp: true,
      sinapi: SINAPI_HABILITADO,
      sicro: fs.existsSync(SICRO_LITE_JSON_PATH) || fs.existsSync(SICRO_JSON_PATH),
      orse: true,
      peIntegrado: true,
    },
  });
});

app.listen(PORT, HOST, () => {
  console.log(`Sistema rodando em http://${HOST === "0.0.0.0" ? "localhost" : HOST}:${PORT}`);

  if (AUTO_UPDATE_BASES) {
    agendarAtualizacaoSicro();
    agendarAtualizacaoOrse();
  } else {
    console.log("Atualização automática de bases desativada. Use /api/sicro/atualizar ou /api/orse/atualizar manualmente.");
  }
});
