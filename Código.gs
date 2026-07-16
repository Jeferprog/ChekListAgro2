/**
 * SISTEMA DE CONSULTA DE CRÉDITO RURAL - CRESOL
 * Gerenciamento de Linhas de Crédito e Regras de Enquadramento
 * Desenvolvido com Google Apps Script
 */

// ==================== CONFIGURAÇÕES GLOBAIS ====================
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_LINHAS = SS.getSheetByName("Linhas") || SS.insertSheet("Linhas");
const SHEET_CONFIG = SS.getSheetByName("Configurações") || SS.insertSheet("Configurações");
const SHEET_HISTORICO = SS.getSheetByName("Histórico") || SS.insertSheet("Histórico");
const SHEET_BASE = SS.getSheetByName("Base") || SS.insertSheet("Base");
const SHEET_BASE_CREDITO = SS.getSheetByName("BaseCredito") || SS.insertSheet("BaseCredito");
const SHEET_CHECKLIST = SS.getSheetByName("ChecklistDocs") || SS.insertSheet("ChecklistDocs");

// ==================== INICIALIZAÇÃO DO SISTEMA ====================

function inicializarSistema() {
  inicializarSheetLinhas();
  inicializarSheetConfig();
  inicializarSheetHistorico();
  inicializarSheetBase();
  inicializarSheetCredito();
  inicializarSheetChecklist();
  Logger.log("✓ Sistema inicializado com sucesso");
}

function inicializarPlanilha() {
  inicializarSistema();
}

function inicializarSheetLinhas() {
  if (SHEET_LINHAS.getLastRow() === 0) {
    const headers = [
      "ID", "Nome Linha", "Órgão/Instituição", "Finalidade Principal",
      "Finalidades (tags)", "Enquadramento (Renda Min/Max)", "Taxa Mín (%)",
      "Taxa Máx (%)", "Prazo (meses)", "Carência (meses)", "Limite Min (R$)",
      "Limite Máx (R$)", "Requisitos", "Documentos Necessários",
      "Status (Ativa/Inativa)", "Data Atualização", "Observações",
      "Itens Financiáveis", "Culturas Financiadas", "Taxa (descrição)", "Tipo Pessoa"
    ];

    SHEET_LINHAS.appendRow(headers);
    SHEET_LINHAS.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#005c46");
    SHEET_LINHAS.getRange(1, 1, 1, headers.length).setFontColor("white");

    SHEET_LINHAS.appendRow([
      "L001",
      "PRONAF CUSTEIO AGRÍCOLA Faixa I",
      "BNDES / Cresol",
      "Custeio",
      "agricola,custeio",
      "Sem limite/R$ 500 mil",
      "3",
      "3",
      "36",
      "0",
      "0",
      "250000",
      "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições",
      "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda",
      "Ativa",
      new Date(),
      "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38%",
      "Itens de custeio relacionados à atividade agrícola",
      "MILHO (até 25 mil), SOJA, TRIGO, FEIJÃO",
      "3% a.a.",
      ""
    ]);
  }
}

function inicializarSheetConfig() {
  if (SHEET_CONFIG.getLastRow() === 0) {
    SHEET_CONFIG.appendRow(["Parâmetro", "Valor"]);
    SHEET_CONFIG.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#f58220").setFontColor("white");
    SHEET_CONFIG.appendRow(["Link Consulta Crédito (SICOR/CACR)", "https://www.bcb.gov.br/sicor/"]);
    SHEET_CONFIG.appendRow(["Link Pasta Base de Associados", ""]);
    SHEET_CONFIG.appendRow(["E-mails Administradores", "jeferson.deimling@cresolsicoper.com.br"]);
    SHEET_CONFIG.appendRow(["Limite Custeio PRONAF", "250000"]);
    SHEET_CONFIG.appendRow(["Limite Custeio PRONAMP", "1500000"]);
  }
}

function inicializarSheetHistorico() {
  if (SHEET_HISTORICO.getLastRow() === 0) {
    const headers = ["Data Hora", "Tipo Operação", "Finalidade", "Enquadramento", "Resultado", "Usuário"];
    SHEET_HISTORICO.appendRow(headers);
    SHEET_HISTORICO.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#005c46").setFontColor("white");
  }
}

function inicializarSheetBase() {
  if (SHEET_BASE.getLastRow() === 0) {
    const cabecalhos = ["nr_cpf_cnpj", "nr_conta_corrente", "nm_nome", "ds_pessoa_tipo", "vl_anual_fonte_renda_total"];
    SHEET_BASE.appendRow(cabecalhos);
    SHEET_BASE.getRange(1, 1, 1, cabecalhos.length).setFontWeight("bold").setBackground("#005c46").setFontColor("white");

    // vl_anual_fonte_renda_total guarda a renda MENSAL (apesar do nome)
    SHEET_BASE.appendRow(["12345678909", "12345-6", "JEFERSON OLIVEIRA DA SILVA", "Física", 29000]);
    SHEET_BASE.appendRow(["98765432100", "54321-0", "MARIA SOUZA REIS", "Física", 100000]);
  }
}

function inicializarSheetCredito() {
  if (SHEET_BASE_CREDITO.getLastRow() === 0) {
    const cabecalhos = ["nr_cpf_cnpj", "ano_safra", "produto", "atividade", "finalidade_recurso", "if_fin", "valor_financiado", "aliquota_proagro", "valor_tomado"];
    SHEET_BASE_CREDITO.appendRow(cabecalhos);
    SHEET_BASE_CREDITO.getRange(1, 1, 1, cabecalhos.length).setFontWeight("bold").setBackground("#005c46").setFontColor("white");

    SHEET_BASE_CREDITO.appendRow(["12345678909", "2025/2026", "PRONAF CUSTEIO AGRÍCOLA", "Milho", "Custeio", "Cresol", 45000, 3, 46350]);
  }
}

function inicializarSheetChecklist() {
  if (SHEET_CHECKLIST.getLastRow() === 0) {
    SHEET_CHECKLIST.appendRow(["Nome Linha", "Documentos"]);
    SHEET_CHECKLIST.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#005c46").setFontColor("white");
  }
}

// ==================== HANDLER WEB ====================

function doGet() {
  inicializarPlanilha();
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Cresol Crédito Rural")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function doPost(e) {
  try {
    const file = e.parameter.arquivo || e.parameter.file;
    if (!file) {
      return ContentService.createTextOutput(JSON.stringify({
        sucesso: false,
        erro: "Nenhum arquivo enviado"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const blob = e.parameter[file];
    if (!blob) {
      return ContentService.createTextOutput(JSON.stringify({
        sucesso: false,
        erro: "Arquivo não encontrado no envio"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const filename = blob.getName ? blob.getName() : (file || "arquivo");
    const content = blob.getDataAsString ? blob.getDataAsString() : String(blob);

    let result;
    if (filename.toLowerCase().endsWith(".docx")) {
      result = processarArquivoCresol({ nome: filename, conteudo: content });
    } else if (filename.toLowerCase().endsWith(".xlsx") || filename.toLowerCase().endsWith(".csv")) {
      result = processarArquivoCreditoBase({ nome: filename, conteudo: content });
    } else {
      result = { sucesso: false, erro: "Tipo de arquivo não suportado" };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      sucesso: false,
      erro: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================== CONFIGURAÇÕES (LEITURA & ESCRITA) ====================

function obterValorConfig(parametro) {
  if (!SHEET_CONFIG) return "";
  const dados = SHEET_CONFIG.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] === parametro) {
      return String(dados[i][1]);
    }
  }
  return "";
}

function salvarValorConfig(parametro, valor) {
  if (!SHEET_CONFIG) return;
  const dados = SHEET_CONFIG.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] === parametro) {
      SHEET_CONFIG.getRange(i + 1, 2).setValue(valor);
      return;
    }
  }
  SHEET_CONFIG.appendRow([parametro, valor]);
}

function obterLinkConsulta() {
  return obterValorConfig("Link Consulta Crédito (SICOR/CACR)") || "https://www.bcb.gov.br/sicor/";
}

function obterLinkBase() {
  return obterValorConfig("Link Pasta Base de Associados") || "";
}

function obterEmailsAdmin() {
  const emailsText = obterValorConfig("E-mails Administradores") || "";
  return emailsText
    .split(/[;,\n]+/)
    .map(function (e) { return e.trim().toLowerCase(); })
    .filter(Boolean);
}

// ==================== CONTROLE DE ACESSO (ADMIN) ====================

// Super administradores: sempre têm acesso, mesmo que não estejam na lista
// da aba Configurações (evita bloqueio acidental do responsável).
const SUPER_ADMINS = ["jeferson.deimling@cresolsicoper.com.br"];

function usuarioAtualEmail() {
  try {
    var e = Session.getActiveUser().getEmail();
    if (!e) e = Session.getEffectiveUser().getEmail();
    return (e || "").toLowerCase();
  } catch (err) {
    return "";
  }
}

/**
 * Indica se o usuário logado pode acessar as abas administrativas
 * (Linhas de Crédito e Configurações do Sistema).
 * - Super admins sempre podem.
 * - Se a lista de e-mails ainda não foi configurada, libera (config inicial).
 * - Caso contrário, só os e-mails da lista têm acesso.
 */
function usuarioEhAdmin() {
  const email = usuarioAtualEmail();
  if (email && SUPER_ADMINS.indexOf(email) !== -1) return true;
  const lista = obterEmailsAdmin();
  if (lista.length === 0) return true;
  return !!email && lista.indexOf(email) !== -1;
}

/** Retorna e-mail logado e se é admin — usado pelo frontend para exibir abas. */
function obterInfoUsuario() {
  return { email: usuarioAtualEmail(), admin: usuarioEhAdmin() };
}

function obterLimitesEnquadramento() {
  const pronaf = parseFloat(obterValorConfig("Limite Custeio PRONAF") || "250000");
  const pronamp = parseFloat(obterValorConfig("Limite Custeio PRONAMP") || "1500000");
  return { pronaf, pronamp };
}

function salvarLinkConsulta(valor) {
  salvarValorConfig("Link Consulta Crédito (SICOR/CACR)", valor);
  return { success: true };
}

function salvarLinkBase(valor) {
  salvarValorConfig("Link Pasta Base de Associados", valor);
  return { success: true };
}

function salvarEmailsAdmin(valor) {
  salvarValorConfig("E-mails Administradores", valor);
  return { success: true };
}

function salvarLimitesEnquadramento(pronaf, pronamp) {
  salvarValorConfig("Limite Custeio PRONAF", String(pronaf));
  salvarValorConfig("Limite Custeio PRONAMP", String(pronamp));
  return { success: true };
}

// ==================== MOTOR DE BUSCA E REGRAS ====================

function buscarLinhas(parametros) {
  try {
    const dados = SHEET_LINHAS.getDataRange().getValues();
    if (!dados || dados.length <= 1) return [];

    const headers = dados[0];
    const linhas = dados.slice(1);

    const resultado = linhas
      .filter(linha => {
        try {
          const statusIdx = headers.indexOf("Status (Ativa/Inativa)");
          if (statusIdx === -1 || linha[statusIdx] !== "Ativa") return false;

          const enquadramentoIdx = headers.indexOf("Enquadramento (Renda Min/Max)");
          if (enquadramentoIdx === -1) return false;

          if (parametros.enquadramento) {
            const grupoLinha = _grupoEnquadramentoLinha(linha[headers.indexOf("Nome Linha")], linha[enquadramentoIdx]);
            if (grupoLinha !== _grupoAssociado(parametros.enquadramento)) return false;
          }

          if (!validarRenda(parametros.renda, linha[enquadramentoIdx])) return false;

          const finalidadesIdx = headers.indexOf("Finalidades (tags)");
          if (finalidadesIdx === -1) return false;

          if (!validarFinalidade(parametros.finalidade, linha[finalidadesIdx])) return false;

          if (!validarProduto(parametros.produto, linha, headers)) return false;

          return true;
        } catch (e) {
          return false;
        }
      })
      .map(linha => {
        try {
          const culturasTxt = linha[headers.indexOf("Culturas Financiadas")] || "";
          let limiteMax = parseInt(linha[headers.indexOf("Limite Máx (R$)")]) || 0;
          let limiteMin = parseInt(linha[headers.indexOf("Limite Min (R$)")]) || 0;

          let capCulturaValor = 0, capCulturaTipo = "";
          const cap = _capCultura(parametros.produto, culturasTxt);
          if (cap) {
            if (cap.tipo === "max") {
              if (limiteMax === 0 || cap.valor < limiteMax) limiteMax = cap.valor;
              capCulturaValor = cap.valor; capCulturaTipo = "max";
            } else if (cap.tipo === "min") {
              if (cap.valor > limiteMin) limiteMin = cap.valor;
              capCulturaValor = cap.valor; capCulturaTipo = "min";
            }
          }

          return {
            id: linha[headers.indexOf("ID")] || "",
            nome: linha[headers.indexOf("Nome Linha")] || "Sem nome",
            orgao: linha[headers.indexOf("Órgão/Instituição")] || "",
            finalidade: linha[headers.indexOf("Finalidade Principal")] || "",
            taxaMin: parseFloat(linha[headers.indexOf("Taxa Mín (%)")]) || 0,
            taxaMax: parseFloat(linha[headers.indexOf("Taxa Máx (%)")]) || 0,
            taxaDescricao: linha[headers.indexOf("Taxa (descrição)")] || "",
            prazo: parseInt(linha[headers.indexOf("Prazo (meses)")]) || 0,
            carencia: parseInt(linha[headers.indexOf("Carência (meses)")]) || 0,
            limiteMin: limiteMin,
            limiteMax: limiteMax,
            requisitos: linha[headers.indexOf("Requisitos")] || "",
            documentos: linha[headers.indexOf("Documentos Necessários")] || "",
            observacoes: linha[headers.indexOf("Observações")] || "",
            itensFinanciaveis: linha[headers.indexOf("Itens Financiáveis")] || "",
            culturas: culturasTxt,
            limiteDisponivel: Math.max(0, limiteMax - (parametros.valorTomado || 0)),
            valorTomado: parametros.valorTomado || 0,
            capCulturaValor: capCulturaValor,
            capCulturaTipo: capCulturaTipo
          };
        } catch (e) {
          return null;
        }
      })
      .filter(item => item !== null)
      .filter(item => {
        if ((parametros.valorTomado || 0) > 0 && item.limiteMax > 0 && item.limiteDisponivel <= 0) {
          return false;
        }
        return true;
      });

    registrarConsulta(parametros, resultado.length);
    return resultado;
  } catch (e) {
    Logger.log("Erro em buscarLinhas: " + e);
    return [];
  }
}

function parseValorRenda(parte) {
  try {
    if (!parte) return null;
    const t = String(parte).toLowerCase().trim();

    const m = t.match(/([\d.,]+)/);
    if (!m) return null;

    let num = parseFloat(m[1].replace(/\s/g, "").replace(",", "."));
    if (isNaN(num)) return null;

    if (/milh|\bmi\b/.test(t)) {
      num = num * 1000000;
    } else if (/\bmil\b/.test(t)) {
      num = num * 1000;
    }
    return num;
  } catch (e) {
    return null;
  }
}

function validarRenda(renda, enquadramentoTexto) {
  try {
    if (!enquadramentoTexto || enquadramentoTexto === "") return true;
    if (typeof enquadramentoTexto !== "string") return true;

    const t = enquadramentoTexto.toLowerCase();

    if (t.includes("conforme")) return true;

    if (enquadramentoTexto.includes("/")) {
      const partes = enquadramentoTexto.split("/");
      const minTexto = partes[0].trim();
      const maxTexto = partes[1].trim();

      const min = minTexto.toLowerCase().includes("sem limite")
        ? 0 : (parseValorRenda(minTexto) || 0);
      const max = maxTexto.toLowerCase().includes("sem limite")
        ? Infinity : (parseValorRenda(maxTexto) || Infinity);

      return renda >= min && renda <= max;
    }

    if (t.includes("acima")) {
      const min = parseValorRenda(enquadramentoTexto);
      if (min === null) return true;
      return renda > min;
    }

    if (t.includes("até") || t.includes("ate")) {
      const max = parseValorRenda(enquadramentoTexto);
      if (max === null) return true;
      return renda <= max;
    }

    return true;
  } catch (e) {
    return true;
  }
}

function validarProduto(produtoBuscado, linha, headers) {
  try {
    if (!produtoBuscado || typeof produtoBuscado !== "string") return true;

    const normalizar = txt => String(txt || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[̀-ͯ]/g, "");

    const termo = normalizar(produtoBuscado.trim());
    if (termo === "") return true;

    const camposRelevantes = [
      "Nome Linha", "Finalidade Principal", "Finalidades (tags)",
      "Itens Financiáveis", "Culturas Financiadas", "Documentos Necessários", "Observações"
    ];
    const textoBusca = normalizar(camposRelevantes
      .map(c => {
        const idx = headers.indexOf(c);
        return idx === -1 ? "" : String(linha[idx] || "");
      })
      .join(" "));

    const palavras = termo.split(/\s+/).filter(p => p.length >= 3);
    if (palavras.length === 0) return true;

    return palavras.some(p => textoBusca.includes(p));
  } catch (e) {
    return true;
  }
}

function _normalizarTexto(s) {
  return String(s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
}

function _parseValorTexto(s) {
  const t = String(s || "").toLowerCase();
  const m = t.match(/([\d.,]+)/);
  if (!m) return null;
  let n = parseFloat(m[1].replace(/\./g, "").replace(",", "."));
  if (isNaN(n)) return null;
  if (/milh|\bmi\b/.test(t)) n *= 1000000;
  else if (/\bmil\b/.test(t)) n *= 1000;
  return n;
}

function _capCultura(produtoBuscado, culturasTexto) {
  try {
    if (!produtoBuscado || !culturasTexto) return null;
    const termo = _normalizarTexto(produtoBuscado.trim());
    if (termo.length < 3) return null;
    const texto = _normalizarTexto(culturasTexto);
    const termoEsc = termo.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

    let m = texto.match(new RegExp(termoEsc + "[^,;()]*\\(\\s*ate\\s*([^)]+)\\)"));
    if (m) { const v = _parseValorTexto(m[1]); if (v) return { tipo: "max", valor: v }; }

    m = texto.match(new RegExp(termoEsc + "[^,;()]*\\(\\s*acima\\s*de\\s*([^)]+)\\)"));
    if (m) { const v = _parseValorTexto(m[1]); if (v) return { tipo: "min", valor: v }; }

    return null;
  } catch (e) {
    return null;
  }
}

function _grupoEnquadramentoLinha(nome, enqTexto) {
  const n = String(nome || "").toUpperCase();
  if (n.indexOf("PRONAF") !== -1) return "pronaf";
  if (n.indexOf("PRONAMP") !== -1) return "pronamp";
  const e = String(enqTexto || "").toLowerCase();
  if (e.indexOf("sem limite") !== -1) return "pronaf";
  if (e.indexOf("500 mil") !== -1 && (e.indexOf("3.5") !== -1 || e.indexOf("3,5") !== -1)) return "pronamp";
  return "demais";
}

function _grupoAssociado(enquadramento) {
  if (enquadramento === "pronaf") return "pronaf";
  if (enquadramento === "pronamp") return "pronamp";
  return "demais";
}

function validarFinalidade(finalidadeBuscada, finalidadeLinha) {
  try {
    if (!finalidadeBuscada || !finalidadeLinha) return true;
    if (typeof finalidadeBuscada !== "string" || typeof finalidadeLinha !== "string") return true;

    const tags = finalidadeLinha.toLowerCase().split(",").map(t => t.trim()).filter(t => t);
    const buscaTermos = finalidadeBuscada.toLowerCase().split(",").map(t => t.trim()).filter(t => t);

    if (tags.length === 0 || buscaTermos.length === 0) return true;

    return buscaTermos.some(termo => tags.some(tag => tag.includes(termo) || termo.includes(tag)));
  } catch (e) {
    return true;
  }
}

function registrarConsulta(parametros, qtdResultados) {
  try {
    if (!SHEET_HISTORICO) return;

    const dataHora = new Date().toLocaleString('pt-BR');
    const finalidade = parametros.finalidade || "Não especificado";
    const enquadramento = parametros.enquadramento || "Não especificado";
    const resultado = `${qtdResultados} linha(s) encontrada(s)`;

    SHEET_HISTORICO.appendRow([
      dataHora,
      "Consulta Linha",
      finalidade,
      enquadramento,
      resultado,
      Session.getActiveUser().getEmail()
    ]);

    Logger.log("Consulta registrada: " + finalidade + " - " + resultado);
  } catch (e) {
    Logger.log("Erro ao registrar consulta: " + e.toString());
  }
}

function obterHistorico() {
  try {
    const dados = SHEET_HISTORICO.getDataRange().getValues();
    if (!dados || dados.length <= 1) return [];
    return dados.slice(1);
  } catch (e) {
    Logger.log("Erro ao obter histórico: " + e);
    return [];
  }
}

// ==================== LINHAS DE CRÉDITO ====================

function listarTodasAsLinhas() {
  if (!SHEET_LINHAS) return [];
  const dados = SHEET_LINHAS.getDataRange().getValues();
  if (dados.length <= 1) return [];

  const checklistLines = [];
  if (SHEET_CHECKLIST) {
    const ckDados = SHEET_CHECKLIST.getDataRange().getValues();
    for (var r = 1; r < ckDados.length; r++) {
      if (String(ckDados[r][1]).trim() !== "") {
        checklistLines.push(String(ckDados[r][0]));
      }
    }
  }

  const H = dados[0];
  const iTipoPessoa = H.indexOf("Tipo Pessoa");
  return dados.slice(1).map(function(linha) {
    const nome = String(linha[H.indexOf("Nome Linha")]);
    const documentos = String(linha[H.indexOf("Documentos Necessários")]);
    // "Tem checklist" = existe checklist IMPORTADO (aba ChecklistDocs)
    const temChecklist = checklistLines.indexOf(nome) !== -1;

    return {
      id: String(linha[H.indexOf("ID")]),
      nome: nome,
      orgao: String(linha[H.indexOf("Órgão/Instituição")]),
      finalidadePrincipal: String(linha[H.indexOf("Finalidade Principal")]),
      finalidades: String(linha[H.indexOf("Finalidades (tags)")]),
      enquadramento: String(linha[H.indexOf("Enquadramento (Renda Min/Max)")]),
      taxaMin: parseFloat(linha[H.indexOf("Taxa Mín (%)")]) || 0,
      taxaMax: parseFloat(linha[H.indexOf("Taxa Máx (%)")]) || 0,
      taxaDescricao: String(linha[H.indexOf("Taxa (descrição)")]),
      prazo: parseInt(linha[H.indexOf("Prazo (meses)")]) || 0,
      carencia: parseInt(linha[H.indexOf("Carência (meses)")]) || 0,
      limiteMin: parseFloat(linha[H.indexOf("Limite Min (R$)")]) || 0,
      limiteMax: parseFloat(linha[H.indexOf("Limite Máx (R$)")]) || 0,
      requisitos: String(linha[H.indexOf("Requisitos")]),
      documentos: documentos,
      status: String(linha[H.indexOf("Status (Ativa/Inativa)")]) === "Inativa" ? "Inativa" : "Ativa",
      observacoes: String(linha[H.indexOf("Observações")]),
      itensFinanciaveis: String(linha[H.indexOf("Itens Financiáveis")]),
      culturas: String(linha[H.indexOf("Culturas Financiadas")]),
      tipoPessoa: iTipoPessoa === -1 ? "" : String(linha[iTipoPessoa] || "").trim().toUpperCase(),
      temChecklist: temChecklist
    };
  });
}

function adicionarLinha(linha) {
  if (!SHEET_LINHAS) return { success: false, erro: "Aba Linhas não encontrada." };

  const dados = SHEET_LINHAS.getDataRange().getValues();
  let maxIdNum = 0;
  for (let i = 1; i < dados.length; i++) {
    const currentId = String(dados[i][0]);
    const num = parseInt(currentId.replace(/\D/g, ""));
    if (!isNaN(num) && num > maxIdNum) maxIdNum = num;
  }
  const nextId = "L" + String(maxIdNum + 1).padStart(3, "0");

  const H = dados[0];
  const novaLinha = new Array(H.length);
  novaLinha[H.indexOf("ID")] = nextId;
  novaLinha[H.indexOf("Nome Linha")] = linha.nome;
  novaLinha[H.indexOf("Órgão/Instituição")] = linha.orgao || "";
  novaLinha[H.indexOf("Finalidade Principal")] = linha.finalidadePrincipal || "Custeio";
  novaLinha[H.indexOf("Finalidades (tags)")] = linha.finalidades || "";
  novaLinha[H.indexOf("Enquadramento (Renda Min/Max)")] = linha.enquadramento || "";
  novaLinha[H.indexOf("Taxa Mín (%)")] = parseFloat(linha.taxaMin) || 0;
  novaLinha[H.indexOf("Taxa Máx (%)")] = parseFloat(linha.taxaMax) || 0;
  novaLinha[H.indexOf("Taxa (descrição)")] = linha.taxaDescricao || "";
  novaLinha[H.indexOf("Prazo (meses)")] = parseInt(linha.prazo) || 0;
  novaLinha[H.indexOf("Carência (meses)")] = parseInt(linha.carencia) || 0;
  novaLinha[H.indexOf("Limite Min (R$)")] = parseFloat(linha.limiteMin) || 0;
  novaLinha[H.indexOf("Limite Máx (R$)")] = parseFloat(linha.limiteMax) || 0;
  novaLinha[H.indexOf("Requisitos")] = linha.requisitos || "";
  novaLinha[H.indexOf("Documentos Necessários")] = linha.documentos || "";
  novaLinha[H.indexOf("Status (Ativa/Inativa)")] = linha.status || "Ativa";
  novaLinha[H.indexOf("Data Atualização")] = new Date();
  novaLinha[H.indexOf("Observações")] = linha.observacoes || "";
  novaLinha[H.indexOf("Itens Financiáveis")] = linha.itensFinanciaveis || "";
  novaLinha[H.indexOf("Culturas Financiadas")] = linha.culturas || "";

  SHEET_LINHAS.appendRow(novaLinha);
  return { success: true, id: nextId };
}

function atualizarLinha(idLinha, novosDados) {
  if (!SHEET_LINHAS) return { success: false, error: "Aba Linhas não encontrada." };

  const dados = SHEET_LINHAS.getDataRange().getValues();
  const headers = dados[0];
  const idIdx = headers.indexOf("ID");

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][idIdx] === idLinha) {
      for (const [chave, valor] of Object.entries(novosDados)) {
        const colIdx = headers.indexOf(chave);
        if (colIdx !== -1) {
          SHEET_LINHAS.getRange(i + 1, colIdx + 1).setValue(valor);
        }
      }
      SHEET_LINHAS.getRange(i + 1, headers.indexOf("Data Atualização") + 1).setValue(new Date());
      return { success: true };
    }
  }
  return { success: false, error: "Linha não encontrada" };
}

function ativarDesativarLinha(id, active) {
  if (!SHEET_LINHAS) return { success: false };
  const dados = SHEET_LINHAS.getDataRange().getValues();
  const colIdx = dados[0].indexOf("Status (Ativa/Inativa)");
  for (let r = 1; r < dados.length; r++) {
    if (String(dados[r][0]) === String(id)) {
      SHEET_LINHAS.getRange(r + 1, colIdx + 1).setValue(active ? "Ativa" : "Inativa");
      return { success: true };
    }
  }
  return { success: false };
}

// ==================== ASSOCIADOS & CRÉDITO TOMADO ====================

/**
 * Busca um associado na aba Base por conta ou CPF/CNPJ (somente dígitos).
 * Observação importante: na base oficial, o campo "vl_anual_fonte_renda_total"
 * contém a renda MENSAL (apesar do nome) — por isso anualizamos (x12).
 */
function buscarAssociado(termo) {
  try {
    if (!SHEET_BASE) return { sucesso: false, erro: "Aba Base não encontrada." };
    if (!termo) return { sucesso: false, erro: "Informe a conta ou o CPF/CNPJ." };
    const alvo = _chaveDoc(termo);
    if (!alvo) return { sucesso: false, erro: "Informe um número válido." };

    const dados = SHEET_BASE.getDataRange().getValues();
    if (!dados || dados.length < 2) {
      return { sucesso: false, erro: "Base de associados vazia. Atualize a base na aba Administrativo." };
    }

    const H = dados[0].map(function (h) { return String(h).trim(); });
    const iCpf = H.indexOf("nr_cpf_cnpj");
    const iConta = H.indexOf("nr_conta_corrente");
    const iNome = H.indexOf("nm_nome");
    const iRenda = H.indexOf("vl_anual_fonte_renda_total");
    const iTipo = H.indexOf("ds_pessoa_tipo");

    for (let r = 1; r < dados.length; r++) {
      const cpf = iCpf !== -1 ? _chaveDoc(dados[r][iCpf]) : "";
      const conta = iConta !== -1 ? _chaveDoc(dados[r][iConta]) : "";
      if ((cpf && cpf === alvo) || (conta && conta === alvo)) {
        const rendaMensal = iRenda !== -1 ? _numBR(dados[r][iRenda]) : 0;
        return {
          sucesso: true,
          nome: iNome !== -1 ? (dados[r][iNome] || "") : "",
          cpfCnpj: iCpf !== -1 ? (dados[r][iCpf] || "") : "",
          conta: iConta !== -1 ? (dados[r][iConta] || "") : "",
          tipo: iTipo !== -1 ? (dados[r][iTipo] || "") : "",
          rendaMensal: rendaMensal,
          rendaAnual: Math.round(rendaMensal * 12)
        };
      }
    }
    return { sucesso: false, erro: "Associado não encontrado na base." };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}

/**
 * Normaliza o valor da coluna "Finalidade Recurso" do SICOR para
 * Custeio / Investimento / Comercialização. Retorna "" quando não reconhece
 * (nesses casos o sistema recorre à classificação por produto/atividade).
 */
function _normalizarFinalidadeRecurso(txt) {
  const s = String(txt || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").trim();
  if (!s) return "";
  if (s.indexOf("custeio") !== -1) return "Custeio";
  if (s.indexOf("investi") !== -1) return "Investimento";
  if (s.indexOf("comerci") !== -1) return "Comercialização";
  // Alguns exports usam códigos: 1=custeio, 2=investimento, 3=comercializacao
  if (s === "1") return "Custeio";
  if (s === "2") return "Investimento";
  if (s === "3") return "Comercialização";
  return "";
}

/**
 * Classifica uma operação do SICOR em Custeio / Investimento / Comercialização.
 * Prioriza a coluna oficial "Finalidade Recurso" (quando informada) e recorre
 * ao produto/atividade apenas como fallback. Apenas o CUSTEIO consome o limite
 * de custeio do enquadramento (PRONAF/PRONAMP); investimento e comercialização
 * possuem limites próprios e NÃO reduzem o custeio.
 */
function _classificarFinalidadeCredito(produto, atividade, finalidadeRecurso) {
  const oficial = _normalizarFinalidadeRecurso(finalidadeRecurso);
  if (oficial) return oficial;

  const s = (String(produto || "") + " " + String(atividade || ""))
    .toLowerCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "");

  const investimento = [
    "trator", "colheitad", "colhedora", "maquin", "implement", "moderfrota",
    "finame", "inovagro", "mais aliment", "investimento", "benfeitor", "galp",
    "armazen", "silo", "resfriad", "ordenha", "caminh", "veicul", "aquisi",
    "constru", "reforma", "irriga", "renovagro", "pca", "bioeconomia",
    "energ", "fotovolt", "solar", "matriz"
  ];
  for (let i = 0; i < investimento.length; i++) {
    if (s.indexOf(investimento[i]) !== -1) return "Investimento";
  }

  const comercial = ["comercial", "estocagem", "desconto", "egf", "amparo"];
  for (let j = 0; j < comercial.length; j++) {
    if (s.indexOf(comercial[j]) !== -1) return "Comercialização";
  }

  // Padrão: operações nomeadas pela cultura/atividade (milho, soja, bovinos...) são custeio.
  return "Custeio";
}

function buscarCreditoTomado(cpf) {
  const vazio = { sucesso: false, items: [], totalFinanciado: 0, totalCusteio: 0, totalInvestimento: 0, totalComercializacao: 0 };
  if (!SHEET_BASE_CREDITO) return vazio;

  const cpfAlvo = String(cpf).replace(/\D/g, "").replace(/^0+/, "");
  if (!cpfAlvo) return vazio;

  const dados = SHEET_BASE_CREDITO.getDataRange().getValues();
  if (dados.length < 2) return { sucesso: true, items: [], totalFinanciado: 0, totalCusteio: 0, totalInvestimento: 0, totalComercializacao: 0 };

  const H = dados[0];
  const items = [];
  let totalFinanciado = 0;
  let totalCusteio = 0;
  let totalInvestimento = 0;
  let totalComercializacao = 0;

  for (let r = 1; r < dados.length; r++) {
    const rowCpf = String(dados[r][H.indexOf("nr_cpf_cnpj")]).replace(/\D/g, "").replace(/^0+/, "");
    if (rowCpf === cpfAlvo) {
      const produto = String(dados[r][H.indexOf("produto")]);
      const atividade = String(dados[r][H.indexOf("atividade")]);
      const iFinRec = H.indexOf("finalidade_recurso");
      const finalidadeRecurso = iFinRec !== -1 ? dados[r][iFinRec] : "";
      const valorFin = _numBR(dados[r][H.indexOf("valor_financiado")]);
      const aliq = _numBR(dados[r][H.indexOf("aliquota_proagro")]);
      // Valor tomado = valor financiado + Alíquota do ProAgro (quando houver).
      const valorTom = valorFin * (aliq > 0 ? (1 + aliq / 100) : 1);
      const finalidade = _classificarFinalidadeCredito(produto, atividade, finalidadeRecurso);

      totalFinanciado += valorTom;
      if (finalidade === "Investimento") totalInvestimento += valorTom;
      else if (finalidade === "Comercialização") totalComercializacao += valorTom;
      else totalCusteio += valorTom;

      items.push({
        anoSafra: String(dados[r][H.indexOf("ano_safra")]),
        produto: produto,
        atividade: atividade,
        ifFin: String(dados[r][H.indexOf("if_fin")]),
        finalidade: finalidade,
        valorFinanciado: valorFin,
        aliquotaProagro: aliq,
        valorTomado: valorTom
      });
    }
  }

  return {
    sucesso: true,
    items: items,
    totalFinanciado: totalFinanciado,
    totalCusteio: totalCusteio,
    totalInvestimento: totalInvestimento,
    totalComercializacao: totalComercializacao
  };
}

// ==================== CHECKLISTS DOCUMENTAÇÃO ====================

function obterChecklistDocs(nomeLinha) {
  if (!SHEET_CHECKLIST) return [];
  const dados = SHEET_CHECKLIST.getDataRange().getValues();
  for (let r = 1; r < dados.length; r++) {
    if (String(dados[r][0]) === String(nomeLinha)) {
      return String(dados[r][1]).split("\n").map(d => d.trim()).filter(Boolean);
    }
  }

  const linhas = listarTodasAsLinhas();
  const linhaMatch = linhas.find(l => l.nome === nomeLinha);
  if (linhaMatch && linhaMatch.documentos) {
    return linhaMatch.documentos.split("\n").map(d => d.trim()).filter(Boolean);
  }
  return [];
}

function salvarChecklist(nomeLinha, documentosTexto) {
  if (!SHEET_CHECKLIST) return { success: false };
  const dados = SHEET_CHECKLIST.getDataRange().getValues();
  for (let r = 1; r < dados.length; r++) {
    if (String(dados[r][0]) === String(nomeLinha)) {
      SHEET_CHECKLIST.getRange(r + 1, 2).setValue(documentosTexto);
      return { success: true };
    }
  }
  SHEET_CHECKLIST.appendRow([nomeLinha, documentosTexto]);
  return { success: true };
}

// Alias mantido para compatibilidade com o frontend (chama _salvarChecklist).
function _salvarChecklist(nomeLinha, documentosTexto) {
  return salvarChecklist(nomeLinha, documentosTexto);
}

// -------- Importação de Checklists (documentos por etapa (P)/(F)) --------

/** Normaliza texto p/ comparação: minúsculas, sem acento, espaços colapsados. */
function _ckNorm(s) {
  return String(s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ").trim();
}

/** Retorna os tokens de faixa presentes no texto, ex.: ["faixa i","faixa iv"]. */
function _ckFaixas(txt) {
  const t = _ckNorm(txt).replace(/\bfx\b/g, "faixa");
  const out = [];
  const re = /faixa\s+(iv|iii|ii|i|v|\d+)/g;
  let m;
  while ((m = re.exec(t))) out.push("faixa " + m[1]);
  return out;
}

/** Parte do cabeçalho antes da primeira "faixa"/"fx" (base do nome da linha). */
function _ckBase(header) {
  const t = _ckNorm(header).replace(/\bfx\b/g, "faixa");
  const i = t.indexOf("faixa");
  return i === -1 ? t : t.slice(0, i).trim();
}

/** Nomes das linhas cadastradas na aba Linhas. */
function _nomesLinhasExistentes() {
  if (!SHEET_LINHAS) return [];
  const dados = SHEET_LINHAS.getDataRange().getValues();
  if (dados.length <= 1) return [];
  const idx = dados[0].indexOf("Nome Linha");
  const out = [];
  for (let r = 1; r < dados.length; r++) {
    const n = String(dados[r][idx] || "").trim();
    if (n) out.push(n);
  }
  return out;
}

/**
 * Resolve o cabeçalho de um bloco de checklist para os nomes de linha reais.
 * Expande enumerações de faixa ("Faixa I, Faixa II, Faixa III e FX IV").
 */
function _resolverLinhasChecklist(header, nomesExistentes) {
  const base = _ckBase(header);
  const faixas = _ckFaixas(header);
  const alvo = [];

  if (faixas.length) {
    nomesExistentes.forEach(function (nome) {
      const nn = _ckNorm(nome).replace(/\bfx\b/g, "faixa");
      if (base && nn.indexOf(base) === -1) return; // a base precisa aparecer no nome
      const f = _ckFaixas(nome);
      if (f.length && faixas.indexOf(f[0]) !== -1) alvo.push(nome);
    });
    return alvo;
  }

  // Sem faixas: tenta igualdade exata; senão, correspondência por conteúdo
  const hn = _ckNorm(header);
  const exatos = nomesExistentes.filter(function (nome) { return _ckNorm(nome) === hn; });
  if (exatos.length) return exatos;
  nomesExistentes.forEach(function (nome) {
    const nn = _ckNorm(nome);
    if (hn && (nn.indexOf(hn) !== -1 || hn.indexOf(nn) !== -1)) alvo.push(nome);
  });
  return alvo;
}

/** Extrai o texto de um arquivo de checklist (.xls/.xlsx, .docx, ou texto puro). */
function _extrairTextoChecklist(conteudo, filename) {
  const nome = String(filename || "").toLowerCase();

  // Planilha (.xls/.xlsx): converte via Drive e usa a 1ª célula não vazia de
  // cada linha como uma linha de texto (cabeçalhos e itens ficam um por linha).
  if (/\.(xlsx|xls)$/.test(nome)) {
    const blob = _base64ParaBlob(conteudo, filename);
    const valores = _xlsxParaValores(blob);
    const linhas = [];
    for (let r = 0; r < valores.length; r++) {
      const row = valores[r] || [];
      let cell = "";
      for (let c = 0; c < row.length; c++) {
        const v = String(row[c] || "").trim();
        if (v) { cell = v; break; }
      }
      linhas.push(cell);
    }
    return linhas.join("\n");
  }

  if (/\.docx$/.test(nome)) {
    const blob = _base64ParaBlob(conteudo, filename);
    blob.setContentType("application/zip");
    const arquivos = Utilities.unzip(blob);
    let docXml = null;
    for (let i = 0; i < arquivos.length; i++) {
      if (arquivos[i].getName() === "word/document.xml") { docXml = arquivos[i].getDataAsString("UTF-8"); break; }
    }
    if (!docXml) return "";
    return docXml
      .replace(/<\/w:p>/g, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&apos;/g, "'");
  }
  // Texto puro (pode vir como base64/data URL ou já como texto)
  let texto = String(conteudo || "");
  if (texto.substring(0, 5) === "data:") {
    const v = texto.indexOf(",");
    if (v !== -1) texto = texto.substring(v + 1);
  }
  const semEsp = texto.replace(/\s/g, "");
  if (semEsp.length > 0 && semEsp.indexOf(",") === -1 && semEsp.indexOf(";") === -1 &&
      /^[A-Za-z0-9+/]+={0,2}$/.test(semEsp)) {
    try { texto = Utilities.newBlob(Utilities.base64Decode(semEsp)).getDataAsString("UTF-8"); } catch (e) {}
  }
  return texto;
}

/**
 * Importa checklists de um documento (.docx ou texto). Cada bloco começa com
 * um cabeçalho (nome/faixas da linha) seguido dos itens marcados com (P) ou (F):
 *   (P) = documentos de PRÉ-CONTRATAÇÃO
 *   (F) = documentos de FORMALIZAÇÃO
 * Os itens são gravados por linha, preservando o prefixo (P)/(F).
 */
function importarChecklists(conteudo, filename) {
  try {
    if (!SHEET_CHECKLIST) return _fail("Aba ChecklistDocs não encontrada.");
    const texto = _extrairTextoChecklist(conteudo, filename);
    if (!texto || !texto.trim()) return _fail("Não foi possível ler o conteúdo do arquivo.");

    const linhasTxt = texto.split(/\r?\n/).map(function (s) { return s.trim(); });
    const reItem = /^\(\s*([pfPF])\s*\)\s*(.*)$/;

    // Monta blocos: cabeçalho + itens (P)/(F)
    const blocos = [];
    let atual = null;
    for (let i = 0; i < linhasTxt.length; i++) {
      const l = linhasTxt[i];
      if (!l) continue;
      const m = l.match(reItem);
      if (m) {
        if (!atual) { atual = { header: "", itens: [] }; blocos.push(atual); }
        const etapa = m[1].toUpperCase();
        const txt = m[2].trim();
        if (txt) atual.itens.push("(" + etapa + ") " + txt);
      } else {
        atual = { header: l, itens: [] };
        blocos.push(atual);
      }
    }

    const nomesExistentes = _nomesLinhasExistentes();
    const relatorio = [];
    let totalLinhas = 0;

    blocos.forEach(function (b) {
      if (!b.itens.length) return;
      const alvos = _resolverLinhasChecklist(b.header, nomesExistentes);
      const textoDocs = b.itens.join("\n");
      alvos.forEach(function (nome) { salvarChecklist(nome, textoDocs); totalLinhas++; });
      relatorio.push({
        header: b.header,
        linhas: alvos,
        naoEncontrado: alvos.length === 0,
        qtdP: b.itens.filter(function (x) { return /^\(P\)/.test(x); }).length,
        qtdF: b.itens.filter(function (x) { return /^\(F\)/.test(x); }).length
      });
    });

    if (totalLinhas === 0) {
      return _fail("Nenhuma linha correspondente encontrada para os títulos do arquivo. Confira se os nomes batem com as linhas cadastradas.");
    }
    return _ok({ linhasAtualizadas: totalLinhas, blocos: relatorio.length, relatorio: relatorio });
  } catch (e) {
    Logger.log("Erro em importarChecklists: " + e);
    return _fail(e.toString());
  }
}

/** Retorna o Nome da Linha a partir do ID. */
function _nomeLinhaPorId(id) {
  if (!SHEET_LINHAS) return "";
  const dados = SHEET_LINHAS.getDataRange().getValues();
  if (dados.length <= 1) return "";
  const iId = dados[0].indexOf("ID");
  const iNome = dados[0].indexOf("Nome Linha");
  for (let r = 1; r < dados.length; r++) {
    if (String(dados[r][iId]) === String(id)) return String(dados[r][iNome] || "").trim();
  }
  return "";
}

/**
 * Importa o checklist de UMA linha específica (por ID). Não faz correspondência
 * por nome: pega TODOS os itens (P)/(F) do arquivo e grava na linha informada.
 * Útil quando linhas diferentes compartilham o mesmo checklist.
 */
function importarChecklistLinha(idLinha, conteudo, filename) {
  try {
    if (!SHEET_CHECKLIST) return _fail("Aba ChecklistDocs não encontrada.");
    const nome = _nomeLinhaPorId(idLinha);
    if (!nome) return _fail("Linha não encontrada (ID " + idLinha + ").");

    const texto = _extrairTextoChecklist(conteudo, filename);
    if (!texto || !texto.trim()) return _fail("Não foi possível ler o conteúdo do arquivo.");

    const reItem = /^\(\s*([pfPF])\s*\)\s*(.*)$/;
    const itens = [];
    texto.split(/\r?\n/).forEach(function (l) {
      const s = l.trim();
      if (!s) return;
      const m = s.match(reItem);
      if (m && m[2].trim()) itens.push("(" + m[1].toUpperCase() + ") " + m[2].trim());
    });

    if (itens.length === 0) {
      return _fail("Nenhum item (P)/(F) encontrado no arquivo. Confira se os documentos começam com (P) ou (F).");
    }

    salvarChecklist(nome, itens.join("\n"));
    const qtdP = itens.filter(function (x) { return /^\(P\)/.test(x); }).length;
    const qtdF = itens.filter(function (x) { return /^\(F\)/.test(x); }).length;
    return _ok({ nome: nome, qtdP: qtdP, qtdF: qtdF, total: itens.length });
  } catch (e) {
    Logger.log("Erro em importarChecklistLinha: " + e);
    return _fail(e.toString());
  }
}

// ==================== IMPORTAÇÃO DE LINHAS VIA .docx DA CRESOL ====================

/**
 * Recebe o conteúdo do .docx da Cresol (em base64), extrai as linhas de
 * crédito e grava numa aba temporária (staging). NÃO altera a base ainda —
 * a confirmação é feita por aplicarAtualizacaoCresol().
 *
 * Aceita: (base64, filename) — enviado pelo frontend React — ou um objeto
 * de formulário { arquivo: Blob } por compatibilidade.
 */
function processarArquivoCresol(base64, filename) {
  try {
    let blob = null;

    // Caso 1: objeto de formulário { arquivo: Blob }
    if (base64 && typeof base64 === "object" && base64.arquivo) {
      blob = base64.arquivo;
    } else if (typeof base64 === "string" && base64.length > 0) {
      // Caso 2: string base64 (pode vir como data URL "data:...;base64,XXXX")
      let b64 = base64;
      const virg = b64.indexOf(",");
      if (b64.substring(0, 5) === "data:" && virg !== -1) b64 = b64.substring(virg + 1);
      const bytes = Utilities.base64Decode(b64);
      blob = Utilities.newBlob(bytes, "application/zip", filename || "cresol.docx");
    }

    if (!blob) return _fail("Nenhum arquivo recebido. Selecione o .docx da Cresol.");

    blob.setContentType("application/zip");
    const arquivos = Utilities.unzip(blob);
    let docXml = null;
    for (let i = 0; i < arquivos.length; i++) {
      if (arquivos[i].getName() === "word/document.xml") {
        docXml = arquivos[i].getDataAsString("UTF-8");
        break;
      }
    }
    if (!docXml) return _fail("Arquivo .docx inválido (document.xml não encontrado).");

    // Extrai texto por parágrafo e decodifica entidades XML
    const texto = docXml
      .replace(/<\/w:p>/g, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/&apos;/g, "'");
    const linhasTxt = texto.split("\n").map(function (s) { return s.trim(); });

    const blocos = _cresolMontarBlocos(linhasTxt);
    if (blocos.length === 0) {
      return _fail("Nenhuma linha encontrada no arquivo. Confira se o documento usa o rótulo \"Nome da Linha de Crédito:\".");
    }

    const rows = [];
    const itens = [];
    let idNum = 0;
    for (let k = 0; k < blocos.length; k++) {
      const rec = _cresolParseBloco(blocos[k]);
      if (!rec.nome) continue;
      idNum++;
      rows.push(_cresolMapear(rec, idNum));
      itens.push({ idx: rows.length - 1, nome: rec.nome, rural: _cresolEhRural(rec.nome) });
    }
    if (rows.length === 0) return _fail("Nenhuma linha válida encontrada no arquivo.");

    // Grava em staging (não altera a base ainda)
    const staging = _cresolStagingSheet(true);
    _cresolEscreverEmSheet(staging, rows);

    return _ok({
      total: rows.length,
      qtdRural: itens.filter(function (i) { return i.rural; }).length,
      qtdNaoRural: itens.filter(function (i) { return !i.rural; }).length,
      itens: itens
    });
  } catch (e) {
    Logger.log("Erro em processarArquivoCresol: " + e.toString());
    return _fail(e.toString());
  }
}

/**
 * Aplica as linhas selecionadas do staging na aba Linhas, criando antes um
 * backup da base atual. Mantém apenas o backup mais recente.
 */
function aplicarAtualizacaoCresol(selecionados) {
  try {
    const staging = SS.getSheetByName("Linhas_Staging");
    if (!staging) return _fail("Nenhuma atualização pendente. Faça o upload do arquivo primeiro.");
    const vals = staging.getDataRange().getValues();
    if (!vals || vals.length <= 1) return _fail("A pré-visualização está vazia. Refaça o upload.");

    if (!selecionados || !selecionados.length) {
      return _fail("Selecione ao menos uma linha para incluir.");
    }

    const headers = vals[0];
    const dados = vals.slice(1);
    // Filtra pelos índices escolhidos e re-sequencia os IDs (L001, L002, ...)
    const escolhidas = [];
    selecionados.forEach(function (i) {
      if (i >= 0 && i < dados.length) escolhidas.push(dados[i].slice());
    });
    if (escolhidas.length === 0) return _fail("Seleção inválida.");
    escolhidas.forEach(function (r, i) { r[0] = "L" + ("000" + (i + 1)).slice(-3); });

    // Backup: remove backups antigos e copia a base atual
    const sheets = SS.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().indexOf("Linhas_Backup") === 0) SS.deleteSheet(sheets[i]);
    }
    const carimbo = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");
    const backup = SHEET_LINHAS.copyTo(SS);
    backup.setName("Linhas_Backup " + carimbo);

    // Aplica na aba Linhas: cabeçalho + linhas escolhidas
    SHEET_LINHAS.clear();
    SHEET_LINHAS.getRange(1, 1, 1, headers.length).setValues([headers]);
    SHEET_LINHAS.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#005c46").setFontColor("white");
    SHEET_LINHAS.getRange(2, 1, escolhidas.length, headers.length).setValues(escolhidas);

    SS.deleteSheet(staging);
    return _ok({ linhas: escolhidas.length, backup: backup.getName() });
  } catch (e) {
    Logger.log("Erro em aplicarAtualizacaoCresol: " + e.toString());
    return _fail(e.toString());
  }
}

/** Descarta a atualização pendente (remove o staging) sem alterar a base. */
function cancelarAtualizacaoCresol() {
  try {
    const staging = SS.getSheetByName("Linhas_Staging");
    if (staging) SS.deleteSheet(staging);
    return _ok();
  } catch (e) {
    return _fail(e.toString());
  }
}

function _cresolStagingSheet(criarSeNao) {
  let sh = SS.getSheetByName("Linhas_Staging");
  if (!sh && criarSeNao) sh = SS.insertSheet("Linhas_Staging");
  return sh;
}

function _cresolMontarBlocos(linhas) {
  const idxs = [];
  for (let i = 0; i < linhas.length; i++) {
    if (linhas[i].indexOf("Nome da Linha de Crédito:") === 0) idxs.push(i);
  }
  const blocos = [];
  for (let k = 0; k < idxs.length; k++) {
    const start = idxs[k];
    const end = (k + 1 < idxs.length) ? idxs[k + 1] : linhas.length;
    blocos.push(linhas.slice(start, end));
  }
  return blocos;
}

function _cresolFieldAfter(block, label) {
  for (let i = 0; i < block.length; i++) {
    if (block[i].indexOf(label) === 0) return block[i].substring(label.length).trim();
  }
  return "";
}

function _cresolSection(block, startLabel, endLabels) {
  const out = [];
  let cap = false;
  for (let i = 0; i < block.length; i++) {
    const l = block[i];
    if (!cap && l.indexOf(startLabel) === 0) {
      cap = true;
      const rest = l.substring(startLabel.length).trim();
      if (rest) out.push(rest);
      continue;
    }
    if (cap) {
      let stop = false;
      for (let j = 0; j < endLabels.length; j++) { if (l.indexOf(endLabels[j]) === 0) { stop = true; break; } }
      if (stop) break;
      if (l.trim()) out.push(l);
    }
  }
  return out;
}

function _cresolParseBloco(b) {
  return {
    nome: _cresolFieldAfter(b, "Nome da Linha de Crédito:"),
    tipo: _cresolFieldAfter(b, "Tipo de Linha:"),
    objetivo: _cresolFieldAfter(b, "Objetivo:"),
    publico: _cresolFieldAfter(b, "Público_resumido:"),
    sistematica: _cresolFieldAfter(b, "Sistemática:"),
    taxa: _cresolFieldAfter(b, "Taxa de Juros Anual:"),
    taxaTexto: _cresolSection(b, "Taxa de Juros Anual:", [
      "IOF", "Limites e Prazos", "Limite de Crédito", "Percentual de Financiamento",
      "Prazo", "Condições", "Modalidades", "Normas", "Circular", "Garantias", "Restrições"
    ]).join(" "),
    iof: _cresolFieldAfter(b, "IOF Complementar:") || _cresolFieldAfter(b, "IOF:"),
    limite: _cresolFieldAfter(b, "Limite de Crédito por Beneficiário:"),
    prazo: _cresolFieldAfter(b, "Prazo Total:"),
    circular: _cresolFieldAfter(b, "Circular BNDES:"),
    tipos: _cresolSection(b, "Tipos:", [
      "Financiamento:", "Requisitos:", "O que financia:", "Produtos Beneficiados:",
      "Sistemática:", "Garantias:", "Taxa", "IOF", "Limite", "Prazo", "Modalidades", "Normas", "Circular"
    ]).join(" "),
    requisitos: _cresolSection(b, "Requisitos:", ["Tipos:", "Financiamento:"]),
    financia: _cresolSection(b, "O que financia:", ["Produtos Beneficiados:", "Sistemática:", "Garantias:"]),
    produtos: _cresolSection(b, "Produtos Beneficiados:", ["Sistemática:", "Taxas e Encargos:", "Garantias:"])
  };
}

/**
 * Classifica a quem a linha se destina a partir do campo "Tipos:" do DOCX.
 * Retorna "PF" (só física), "PJ" (só jurídica), "PF/PJ" (ambas) ou "" (não
 * especificado — não restringe a busca).
 */
function _cresolTipoPessoa(tiposTxt) {
  const s = _ckNorm(tiposTxt);
  if (!s) return "";
  const temPF = s.indexOf("pessoa fisica") !== -1 || /\bfisica\b/.test(s) || /\bpf\b/.test(s);
  const temPJ = s.indexOf("pessoa juridica") !== -1 || /\bjuridica\b/.test(s) || /\bpj\b/.test(s);
  if (temPF && temPJ) return "PF/PJ";
  if (temPF) return "PF";
  if (temPJ) return "PJ";
  return "";
}

function _cresolEhRural(nome) {
  const n = nome.toUpperCase();
  const exclui = ["FINEP", "FUNGETUR", "PROMOVE SUL", "FUNDO CLIMA", "PROCAPCRED",
    "INVESTIMENTO EMPRESARIAL", "CRESOL EMPRESARIAL BNDES", "BNDES FINAME (FINAME BK"];
  for (let i = 0; i < exclui.length; i++) { if (n.indexOf(exclui[i]) !== -1) return false; }
  return true;
}

function _cresolNumTaxa(t) {
  if (!t) return 0;
  const m = String(t).match(/(\d+(?:[.,]\d+)?)/);
  if (!m) return 0;
  const n = parseFloat(m[1].replace(",", "."));
  return isNaN(n) ? 0 : n;
}

function _cresolNumLimite(t) {
  if (!t) return 0;
  const re = /R\$\s*([\d.]+(?:,\d+)?)/g;
  let m, max = 0, achou = false;
  while ((m = re.exec(t))) {
    const v = m[1].replace(/\./g, "").replace(",", ".");
    const n = parseInt(parseFloat(v), 10);
    if (!isNaN(n)) { achou = true; if (n > max) max = n; }
  }
  return achou ? max : 0;
}

function _cresolPrazoMeses(t) {
  if (!t) return 0;
  let meses = 0, m;
  const reAnos = /(\d+)\s*anos?/g;
  while ((m = reAnos.exec(t))) { const v = parseInt(m[1], 10) * 12; if (v > meses) meses = v; }
  const reMes = /(\d+)\s*mes/g;
  while ((m = reMes.exec(t))) { const v = parseInt(m[1], 10); if (v > meses) meses = v; }
  return meses;
}

function _cresolEnquadramento(pub, nome) {
  const p = (pub || "").toUpperCase();
  const n = nome.toUpperCase();
  if (p.indexOf("PRONAF") !== -1 || n.indexOf("PRONAF") !== -1) return "Sem limite/R$ 500 mil";
  if (p.indexOf("PRONAMP") !== -1 || n.indexOf("PRONAMP") !== -1) return "R$ 500 mil/R$ 3.5 mi";
  return "Conforme análise";
}

function _cresolTags(nome, tipo, objetivo) {
  const s = (nome + " " + tipo + " " + objetivo).toLowerCase();
  const t = {};
  if (s.indexOf("custeio") !== -1) t["custeio"] = 1;
  const inv = ["investimento", "tratores", "colheitadeira", "máquina", "maquina", "finame", "moderfrota", "inovagro", "agroind", "habita", "bioeconomia", "pca", "prodecoop", "irriga", "renovagro"];
  for (let i = 0; i < inv.length; i++) { if (s.indexOf(inv[i]) !== -1) { t["investimento"] = 1; break; } }
  if (s.indexOf("pecuár") !== -1 || s.indexOf("pecuar") !== -1) t["pecuaria"] = 1;
  if (s.indexOf("agrícola") !== -1 || s.indexOf("agricola") !== -1) t["agricola"] = 1;
  if (s.indexOf("café") !== -1 || s.indexOf("cafe") !== -1 || s.indexOf("funcaf") !== -1) t["cafe"] = 1;
  if (s.indexOf("irriga") !== -1) t["irrigacao"] = 1;
  const mec = ["tratores", "colheitadeira", "máquina", "maquina", "moderfrota", "finame"];
  for (let j = 0; j < mec.length; j++) { if (s.indexOf(mec[j]) !== -1) { t["mecanizacao"] = 1; break; } }
  if (s.indexOf("pca") !== -1 || s.indexOf("armaz") !== -1) t["armazenagem"] = 1;
  const sus = ["renovagro", "sustentável", "sustentavel", "agroecolog", "bioeconomia", "ambiental", "clima"];
  for (let k = 0; k < sus.length; k++) { if (s.indexOf(sus[k]) !== -1) { t["sustentabilidade"] = 1; break; } }
  if (s.indexOf("agroind") !== -1 || s.indexOf("industrial") !== -1) t["infraestrutura"] = 1;
  let keys = Object.keys(t);
  if (keys.length === 0) keys = ["investimento"];
  keys.sort();
  return keys.join(",");
}

function _cresolOrgao(rec) {
  const src = ((rec.sistematica || "") + " " + (rec.circular || "")).toUpperCase();
  if (src.indexOf("BNDES") !== -1) return "BNDES / Cresol";
  if (src.indexOf("FCO") !== -1 || rec.nome.toUpperCase().indexOf("FCO") !== -1) return "FCO / Cresol";
  if (src.indexOf("POUPAN") !== -1) return "Cresol (Poupança Rural)";
  return "Cresol";
}

function _cresolDocumentos(pub) {
  if ((pub || "").toUpperCase().indexOf("PRONAF") !== -1) return "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda";
  return "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda";
}

function _cresolObs(rec) {
  const p = [];
  if (rec.objetivo) p.push(rec.objetivo);
  if (rec.sistematica) p.push("Sistemática: " + rec.sistematica);
  if (rec.iof) p.push("IOF: " + rec.iof);
  if (rec.prazo) p.push("Prazo: " + rec.prazo);
  if (rec.circular) p.push("Norma: " + rec.circular);
  return p.join(" | ").substring(0, 600);
}

function _cresolMapear(rec, idNum) {
  const rid = "L" + ("000" + idNum).slice(-3);
  const status = (rec.nome.toLowerCase().indexOf("fechado") !== -1) ? "Inativa" : "Ativa";
  const taxa = _cresolNumTaxa(rec.taxa);
  return [
    rid, rec.nome, _cresolOrgao(rec),
    (rec.objetivo || rec.tipo || "Crédito rural"),
    _cresolTags(rec.nome, rec.tipo, rec.objetivo),
    _cresolEnquadramento(rec.publico, rec.nome),
    taxa, taxa,
    _cresolPrazoMeses(rec.prazo), 0,
    0, _cresolNumLimite(rec.limite),
    (rec.requisitos.join("; ")).substring(0, 600) || "Conforme política de crédito da Cresol",
    _cresolDocumentos(rec.publico),
    status, new Date(), _cresolObs(rec),
    (rec.financia.join("; ")).substring(0, 900),
    (rec.produtos.join(", ")).substring(0, 1500),
    String(rec.taxaTexto || rec.taxa || "").substring(0, 800),
    _cresolTipoPessoa(rec.tipos)
  ];
}

function _cresolEscreverEmSheet(sheet, rows) {
  const headers = [
    "ID", "Nome Linha", "Órgão/Instituição", "Finalidade Principal",
    "Finalidades (tags)", "Enquadramento (Renda Min/Max)", "Taxa Mín (%)",
    "Taxa Máx (%)", "Prazo (meses)", "Carência (meses)", "Limite Min (R$)",
    "Limite Máx (R$)", "Requisitos", "Documentos Necessários",
    "Status (Ativa/Inativa)", "Data Atualização", "Observações",
    "Itens Financiáveis", "Culturas Financiadas", "Taxa (descrição)", "Tipo Pessoa"
  ];
  sheet.clear();
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#005c46").setFontColor("white");
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function obterIdDoDrive(link) {
  if (!link) return null;
  var matchFolder = link.match(/\/folders\/([a-zA-Z0-9-_]+)/);
  if (matchFolder) return { type: "folder", id: matchFolder[1] };
  var matchFile = link.match(/\/file\/d\/([a-zA-Z0-9-_]+)/) || link.match(/id=([a-zA-Z0-9-_]+)/);
  if (matchFile) return { type: "file", id: matchFile[1] };
  return null;
}

// ==================== IMPORTAÇÃO DE DADOS ====================

/**
 * Respostas padronizadas com chaves em PT e EN, porque o frontend (React)
 * lê `success`/`error` em alguns handlers e `sucesso`/`erro` em outros.
 * Assim qualquer handler funciona independentemente da chave que consultar.
 */
function _ok(extra) {
  const o = { success: true, sucesso: true };
  if (extra) { for (const k in extra) o[k] = extra[k]; }
  return o;
}

function _fail(msg) {
  const m = String(msg || "Erro desconhecido");
  return { success: false, sucesso: false, error: m, erro: m };
}

/**
 * Converte um valor monetário em formato brasileiro para número.
 * Ex.: "R$ 1.234,56" -> 1234.56 ; "1500" -> 1500.
 */
function _numBR(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  let s = String(v).replace(/[^\d.,\-]/g, "");
  if (s.indexOf(",") !== -1) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/**
 * Chave de comparação de documento/conta: só dígitos, sem zeros à esquerda.
 * Resolve o caso da base que perde zeros iniciais de CPF (ex.: digitar
 * "01234567890" encontra o registro "1234567890").
 */
function _chaveDoc(v) {
  return String(v || "").replace(/\D/g, "").replace(/^0+/, "");
}

/**
 * Lê o conteúdo do CSV a partir de um link de PASTA ou ARQUIVO do Google
 * Drive, de uma URL direta, ou de um ID solto do Drive.
 */
function _lerCsvDoLink(link, filename) {
  // Link de pasta do Drive
  const mFolder = link.match(/folders\/([a-zA-Z0-9_\-]+)/);
  if (mFolder) {
    const folder = DriveApp.getFolderById(mFolder[1]);
    const it = folder.getFilesByName(filename);
    if (it.hasNext()) return it.next().getBlob().getDataAsString("UTF-8");
    const its = folder.getFiles();
    while (its.hasNext()) {
      const f = its.next();
      if (f.getName().toLowerCase().indexOf(".csv") !== -1) return f.getBlob().getDataAsString("UTF-8");
    }
    return null;
  }
  // Link de arquivo do Drive (/d/<id> ou ?id=<id>)
  const idm = link.match(/(?:\/d\/|id=)([a-zA-Z0-9_\-]+)/);
  if (idm) return DriveApp.getFileById(idm[1]).getBlob().getDataAsString("UTF-8");
  // URL direta
  if (/^https?:\/\//.test(link)) return UrlFetchApp.fetch(link).getContentText();
  // Talvez seja só um ID
  const mId = link.match(/^[a-zA-Z0-9_\-]{20,}$/);
  if (mId) return DriveApp.getFileById(link).getBlob().getDataAsString("UTF-8");
  return null;
}

/**
 * Grava o conteúdo de um CSV diretamente na aba Base, preservando os
 * cabeçalhos e colunas originais do arquivo (não força colunas fixas).
 * Detecta o delimitador pela primeira linha e normaliza a largura das linhas.
 */
function _gravarCsvNaBase(conteudo) {
  if (!conteudo || conteudo.trim() === "") {
    return _fail("CSV vazio ou inválido.");
  }

  const primeiraLinha = conteudo.split("\n")[0] || "";
  const delim = (primeiraLinha.split(";").length > primeiraLinha.split(",").length) ? ";" : ",";
  const dados = Utilities.parseCsv(conteudo, delim);
  if (!dados || dados.length < 2) return _fail("CSV vazio ou inválido.");

  // Normaliza a largura das linhas (todas com o mesmo nº de colunas do cabeçalho)
  const largura = dados[0].length;
  const norm = dados.map(function (r) {
    const linha = r.slice(0, largura);
    while (linha.length < largura) linha.push("");
    return linha;
  });

  SHEET_BASE.clear();
  SHEET_BASE.getRange(1, 1, norm.length, largura).setValues(norm);
  SHEET_BASE.getRange(1, 1, 1, largura).setFontWeight("bold").setBackground("#005c46").setFontColor("white");

  return _ok({ registros: norm.length - 1, atualizado: new Date().toLocaleString("pt-BR") });
}

/**
 * Busca o basedepessoas.csv no link configurado e regrava a aba Base.
 * Também é chamada pelo trigger automático de atualização diária.
 */
function atualizarBaseAssociados() {
  try {
    if (!SHEET_BASE) return _fail("Aba Base não encontrada.");

    const link = obterLinkBase();
    if (!link) return _fail("Configure o link da pasta/arquivo da base na aba Administrativo.");

    const conteudo = _lerCsvDoLink(link, "basedepessoas.csv");
    if (!conteudo) return _fail("Arquivo basedepessoas.csv não encontrado no link informado.");

    return _gravarCsvNaBase(conteudo);
  } catch (e) {
    Logger.log("Erro em atualizarBaseAssociados: " + e.toString());
    return _fail(e.toString());
  }
}

/**
 * Faz upload de um blob para o Drive convertendo para um tipo Google
 * (Sheets/Docs) e retorna o ID do arquivo convertido. Necessário para ler
 * .xlsx/.xls (binários) que o Utilities.parseCsv não consegue interpretar.
 */
function _uploadConvert(blob, targetMime) {
  const boundary = "conv" + Date.now();
  const metadata = { name: "tmp_conv_" + Date.now(), mimeType: targetMime };
  const ct = blob.getContentType() || "application/octet-stream";
  const pre = "--" + boundary + "\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n" +
    JSON.stringify(metadata) + "\r\n--" + boundary + "\r\nContent-Type: " + ct + "\r\n\r\n";
  const post = "\r\n--" + boundary + "--";
  const bytes = Utilities.newBlob(pre).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob(post).getBytes());

  const res = UrlFetchApp.fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true", {
    method: "post",
    contentType: "multipart/related; boundary=" + boundary,
    payload: bytes,
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("Falha na conversão do arquivo (HTTP " + code + "): " + res.getContentText().substring(0, 200));
  }
  return JSON.parse(res.getContentText()).id;
}

/** Converte um blob .xlsx/.xls em matriz de valores (lê a 1ª aba). */
function _xlsxParaValores(blob) {
  const fileId = _uploadConvert(blob, "application/vnd.google-apps.spreadsheet");
  try {
    return SpreadsheetApp.openById(fileId).getSheets()[0].getDataRange().getValues();
  } finally {
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
  }
}

/** Remove prefixo de data URL e converte string base64 em Blob. */
function _base64ParaBlob(conteudo, filename) {
  let b64 = String(conteudo || "");
  if (b64.substring(0, 5) === "data:") {
    const v = b64.indexOf(",");
    if (v !== -1) b64 = b64.substring(v + 1);
  }
  b64 = b64.replace(/\s/g, "");
  const bytes = Utilities.base64Decode(b64);
  return Utilities.newBlob(bytes, "application/octet-stream", filename || "arquivo");
}

/**
 * Importa a base de crédito tomado. Aceita .csv (texto ou base64) e .xlsx/.xls
 * (base64). Assinatura compatível com o frontend: (filename, contentText).
 * - .xlsx/.xls: converte via Drive e lê os valores.
 * - .csv: faz parse do texto.
 * Detecta o layout automaticamente: cabeçalhos nomeados OU posições fixas do
 * export SICOR/CACR (C=Ano Safra, D=Valor, G=Produto, K=IF, S=Atividade,
 * W=CPF/CNPJ, AB=Alíquota ProAgro).
 */
function processarArquivoCredito(filename, contentText) {
  if (!SHEET_BASE_CREDITO) return _fail("Aba BaseCredito não encontrada.");

  try {
    // Compatibilidade: aceita objeto { nome, conteudo }
    if (contentText === undefined && filename && typeof filename === "object") {
      contentText = filename.conteudo || filename.content || "";
      filename = filename.nome || "credito.csv";
    }
    if (!contentText) return _fail("Arquivo vazio ou não foi possível ler o conteúdo.");

    const nome = String(filename || "").toLowerCase();
    const ehXlsx = /\.(xlsx|xls)$/.test(nome);

    let valores;
    if (ehXlsx) {
      // .xlsx/.xls -> converte via Drive e lê a matriz de valores
      const blob = _base64ParaBlob(contentText, filename);
      valores = _xlsxParaValores(blob);
    } else {
      // .csv -> pode vir como texto puro ou base64 (data URL)
      let texto = String(contentText);
      if (texto.substring(0, 5) === "data:") {
        const v = texto.indexOf(",");
        if (v !== -1) texto = texto.substring(v + 1);
      }
      const semEsp = texto.replace(/\s/g, "");
      // Se for base64 (sem vírgulas/; e só alfabeto base64), decodifica p/ texto
      if (semEsp.length > 0 && semEsp.indexOf(",") === -1 && semEsp.indexOf(";") === -1 &&
          /^[A-Za-z0-9+/]+={0,2}$/.test(semEsp)) {
        try { texto = Utilities.newBlob(Utilities.base64Decode(semEsp)).getDataAsString("UTF-8"); } catch (e) {}
      }
      const delim = ((texto.split("\n")[0] || "").split(";").length > (texto.split("\n")[0] || "").split(",").length) ? ";" : ",";
      valores = Utilities.parseCsv(texto, delim);
    }

    return _gravarCreditoMatriz(valores);
  } catch (err) {
    Logger.log("Erro ao processar arquivo de crédito: " + err);
    return _fail("Erro ao processar arquivo: " + err.toString());
  }
}

/**
 * Grava uma matriz (de CSV ou XLSX) na aba BaseCredito, detectando o layout:
 * cabeçalhos nomeados OU posições fixas do export SICOR/CACR.
 */
function _gravarCreditoMatriz(valores) {
  const headers = ["nr_cpf_cnpj", "ano_safra", "produto", "atividade", "finalidade_recurso", "if_fin", "valor_financiado", "aliquota_proagro", "valor_tomado"];
  SHEET_BASE_CREDITO.clear();
  SHEET_BASE_CREDITO.getRange(1, 1, 1, headers.length).setValues([headers]);
  SHEET_BASE_CREDITO.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#005c46").setFontColor("white");

  if (!valores || valores.length < 2) {
    return _ok({ registros: 0, atualizado: new Date().toLocaleString('pt-BR') });
  }

  const H = valores[0].map(function (h) { return String(h || "").trim().toLowerCase(); });
  const acha = function (nomes) { for (let i = 0; i < nomes.length; i++) { const p = H.indexOf(nomes[i]); if (p !== -1) return p; } return -1; };
  const cpfIdx = acha(["nr_cpf_cnpj", "cpf_cnpj", "cpf", "cnpj"]);
  const usarNomes = cpfIdx !== -1;

  // Posições fixas (base 0) do layout SICOR/CACR
  const P = { cpf: 22, safra: 2, produto: 6, atividade: 18, ifFin: 10, valFin: 3, aliq: 27 };

  const rows = [];
  for (let r = 1; r < valores.length; r++) {
    const row = valores[r];
    if (!row || row.length === 0) continue;

    let cpf, safra, produto, atividade, finalidade, ifFin, valFin, aliq;
    if (usarNomes) {
      const g = function (nomes, def) { const p = acha(nomes); return p !== -1 ? row[p] : def; };
      cpf = String(g(["nr_cpf_cnpj", "cpf_cnpj", "cpf", "cnpj"], "") || "").trim();
      safra = String(g(["ano_safra", "safra"], "") || "").trim();
      produto = String(g(["produto", "linha"], "") || "").trim();
      atividade = String(g(["atividade", "cultura"], "") || "").trim();
      finalidade = String(g(["finalidade_recurso", "finalidade recurso", "finalidade do recurso", "finalidade_do_recurso", "finalidade"], "") || "").trim();
      ifFin = String(g(["if_fin", "instituicao", "if_financiamento"], "") || "").trim();
      valFin = _numBR(g(["valor_financiado", "valor"], 0));
      aliq = _numBR(g(["aliquota_proagro", "proagro", "aliquota"], 0));
    } else {
      cpf = String(row[P.cpf] || "").trim();
      safra = String(row[P.safra] || "").trim();
      produto = String(row[P.produto] || "").trim();
      atividade = String(row[P.atividade] || "").trim();
      finalidade = P.finalidade !== undefined ? String(row[P.finalidade] || "").trim() : "";
      ifFin = String(row[P.ifFin] || "").trim();
      valFin = _numBR(row[P.valFin]);
      aliq = _numBR(row[P.aliq]);
    }

    if (!_chaveDoc(cpf)) continue;

    // ProAgro (quando há alíquota) é somado ao valor financiado -> valor tomado
    const fator = aliq > 0 ? (1 + aliq / 100) : 1;
    const valTom = valFin * fator;

    // Normaliza a finalidade oficial; se vier vazia, classifica por produto/atividade.
    const finalidadeNorm = _classificarFinalidadeCredito(produto, atividade, finalidade);

    rows.push([
      cpf,
      safra || "2025/2026",
      (produto || "CRÉDITO RURAL").toUpperCase(),
      atividade || "Outros",
      finalidadeNorm,
      ifFin || "Cresol",
      valFin,
      aliq,
      valTom
    ]);
  }

  if (rows.length > 0) {
    SHEET_BASE_CREDITO.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  if (rows.length === 0) {
    return _fail("Nenhuma linha com CPF/CNPJ encontrada. Confira o layout/colunas do arquivo.");
  }
  return _ok({ registros: rows.length, atualizado: new Date().toLocaleString('pt-BR') });
}

// Alias mantido para compatibilidade (usado pelo doPost e versões anteriores).
function processarArquivoCreditoBase(arquivoInfo) {
  const conteudo = (arquivoInfo && (arquivoInfo.conteudo || arquivoInfo.content)) || arquivoInfo;
  return processarArquivoCredito((arquivoInfo && arquivoInfo.nome) || "credito.csv", conteudo);
}

function processarArquivoAssociados(filename, contentText) {
  if (!SHEET_BASE) return _fail("Aba Base não encontrada.");

  try {
    return _gravarCsvNaBase(contentText);
  } catch (err) {
    return _fail("Erro ao processar arquivo de associados: " + err.toString());
  }
}

/* =========================================================================
 * AGENTE DE IA — Sugestão da melhor linha de crédito
 * =======================================================================*/

const IA_PROP_CHAVE = "IA_API_KEY";              // chave atual (provedor selecionado)
const IA_PROP_CHAVE_LEGADO = "ANTHROPIC_API_KEY"; // compatibilidade com versões anteriores
const IA_PROVIDER_PADRAO = "gemini";
const IA_MODELO_GEMINI_PADRAO = "gemini-2.5-flash";
const IA_MODELO_ANTHROPIC_PADRAO = "claude-opus-4-8";
const IA_ENDPOINT_ANTHROPIC = "https://api.anthropic.com/v1/messages";
const IA_ENDPOINT_GEMINI_BASE = "https://generativelanguage.googleapis.com/v1beta/models/";

// Nomes dos parâmetros salvos na aba Configurações (editáveis pelo admin).
const IA_CFG_PROVIDER = "Provedor de IA";
const IA_CFG_MODELO = "Modelo de IA";
const IA_CFG_DIRETRIZES = "Diretrizes do Agente de IA";

// Diretrizes PADRÃO do agente especialista — usadas quando o admin não
// definiu diretrizes próprias na aba Configurações.
const IA_SYSTEM_INSTRUCTION_PADRAO =
"Você é o Especialista em Crédito Rural e Agronegócio da Cresol. Seu papel principal é atuar como um arquiteto de soluções financeiras para gerentes de carteira e associados. Seu objetivo absoluto é: diagnosticar a necessidade do produtor, cruzar com as regras de enquadramento vigentes e recomendar a linha de menor custo financeiro (menor taxa de juros) e prazo mais adequado.\n\n" +
"DIRETRIZES DE COMPORTAMENTO:\n" +
"- Sempre classifique o produtor (PRONAF, PRONAMP, DEMAIS).\n" +
"- Identifique se é custeio ou investimento.\n" +
"- Busque sempre a menor taxa de juros possível para o cenário dele.\n\n" +
"BASE DE DADOS DE LINHAS DE CRÉDITO:\n" +
"[CUSTEIO AGRÍCOLA - PRONAF]\n" +
"- CAF-Pronaf/DAP ativa, área <= 4 módulos fiscais, renda até R$ 500k. Limite: R$ 250k.\n" +
"* Faixa I (3% a.a.): Alface, feijão, arroz, mandioca, milho (até 25 mil), hortaliças, morango, trigo, banana, abacaxi, batata, tomate.\n" +
"* Faixa II (6.5% a.a.): Café, cana, pastagem, milho (acima de 25 mil), uva, soja, eucalipto.\n" +
"* Faixa III (Orgânicos): Taxa 2% a.a.\n" +
"* Faixa IV (8% a.a.): Soja, algodão, búfalos (carne).\n\n" +
"[CUSTEIO PECUÁRIO - PRONAF]\n" +
"* Faixa I (3% a.a.): Apicultura, leite, ovinos.\n" +
"* Faixa II (6.5% a.a.): Suínos, aves corte.\n" +
"* Faixa IV (8% a.a.): Bovinos de corte (confinamento até 6 meses, recria até 12 meses, recria/engorda conjunta até 20 meses).\n\n" +
"[CUSTEIO PRONAMP]\n" +
"- Renda até R$ 3.5M, min 80% agro. Limite: R$ 1.5M.\n" +
"* Comum: Taxa 10% a.a.\n" +
"* Sustentável: Taxa 9.5% a.a.\n\n" +
"[CUSTEIO DEMAIS PRODUTORES]\n" +
"- Limite: R$ 3.0M.\n" +
"* Comum: Taxa 14% a.a.\n" +
"* Sustentável: Taxa 13.5% a.a.\n\n" +
"[FUNCAFÉ]\n" +
"- IOF Isento. Limite: R$ 3M.\n" +
"* Custeio: Taxa 13% a.a. (Até 20 meses)\n" +
"* Comercialização: Taxa 13% a.a. (Até 12 meses)\n\n" +
"[INVESTIMENTOS IMPORTANTES]\n" +
"* Pronaf Grupo B (0.5% a.a.): Limite R$ 12k. Renda <= R$ 50k.\n" +
"* Mais Alimentos Máquinas F3 (2.5% a.a.): Limite R$ 100k. Renda < R$ 150k.\n" +
"* Mais Alimentos F1 (3% a.a.): Limite R$ 250k (ou R$ 450k para suínos/aves/fruticultura).\n" +
"* Mais Alimentos F2 (8% a.a.): Caminhonetes de carga, matrizes, galpões.\n" +
"* Pronamp Investimento (10% a.a.): Limite R$ 600k.\n" +
"* RenovAgro Ambiental / Pastagem (8.5% a.a.): Recuperação ambiental e de pastagem degradada. Limite R$ 5M.\n" +
"* RenovAgro Demais (10% a.a.): Orgânicos, solos, florestas. Limite R$ 5M.\n" +
"* PCA Grãos (8.5% a.a.): Armazenagem de grãos até 12 mil toneladas. Limite R$ 50M.\n" +
"* PCA Demais (10% a.a.): Armazenagem acima de 12 mil toneladas. Limite R$ 50M.\n" +
"* FCO Rural MS (9.05% a.a.): Exclusivo para Mato Grosso do Sul.\n" +
"* Bioeconomia Solar (3% a.a.): Sistemas fotovoltaicos, somente via Finame.\n\n" +
"ORIENTAÇÕES OPERACIONAIS:\n" +
"- Quando o sistema informar as linhas ELEGÍVEIS já filtradas para este produtor, priorize recomendar entre elas; se nenhuma delas atender bem, explique o motivo e indique a linha da base acima que melhor se aplica.\n" +
"- ANÁLISE DE LIMITES: separe SEMPRE as operações de CUSTEIO das de INVESTIMENTO no histórico SICOR. Investimento (tratores, máquinas, benfeitorias) tem limite próprio e NÃO consome o limite de custeio. Calcule a margem remanescente de custeio como: (teto de custeio do grupo) − (total de custeio já tomado, com ProAgro). O valor máximo do novo contrato de custeio é essa margem.\n" +
"- TRAVA DO SICOR: alerte que o SICOR bloqueia contratações de custeio que superem a margem remanescente na safra; o valor proposto deve ser menor ou igual a esse saldo.\n" +
"- ESTRATÉGIA HÍBRIDA: se a necessidade do produtor superar a margem do PRONAF, sugira operação complementar (ex.: contratar a margem no PRONAF à menor taxa e o excedente no PRONAMP, priorizando a modalidade Sustentável quando aplicável), sempre que o produtor também se enquadrar no outro programa.\n" +
"- Sua resposta é um apoio à decisão do analista, não uma aprovação de crédito. Seja objetivo, mas completo.\n" +
"- Responda em português do Brasil, de forma organizada, cobrindo quando fizer sentido: Perfil do associado e enquadramento; Análise de limites (custeio vs investimento e margem remanescente); Linha recomendada e justificativa; Condições financeiras (taxa, limite para o contrato, IOF quando conhecido); Prazo e carência sugeridos; Regras e cuidados (trava do SICOR, ProAgro/seguro, documentação); Alternativa próxima / estratégia híbrida se a demanda superar a margem.";

/** Lê a chave da API das Script Properties (nunca fica na planilha nem no código). */
function _obterChaveIA() {
  try {
    const props = PropertiesService.getScriptProperties();
    let v = props.getProperty(IA_PROP_CHAVE);
    if (!v) v = props.getProperty(IA_PROP_CHAVE_LEGADO); // compatibilidade
    return v ? String(v).trim() : "";
  } catch (e) {
    return "";
  }
}

/** Grava a chave da API nas Script Properties. Chamada apenas por administradores. */
function salvarChaveIA(chave) {
  try {
    if (!usuarioEhAdmin()) return _fail("Apenas administradores podem configurar a chave da IA.");
    const c = String(chave || "").trim();
    const props = PropertiesService.getScriptProperties();
    if (!c) {
      props.deleteProperty(IA_PROP_CHAVE);
      props.deleteProperty(IA_PROP_CHAVE_LEGADO);
      return _ok({ configurada: false, mensagem: "Chave removida." });
    }
    props.setProperty(IA_PROP_CHAVE, c);
    props.deleteProperty(IA_PROP_CHAVE_LEGADO); // consolida na chave nova
    return _ok({ configurada: true, mensagem: "Chave salva com segurança." });
  } catch (err) {
    return _fail("Erro ao salvar chave: " + err.toString());
  }
}

/** Provedor de IA selecionado (gemini | anthropic). */
function _obterProviderIA() {
  const p = (obterValorConfig(IA_CFG_PROVIDER) || "").trim().toLowerCase();
  return p === "anthropic" ? "anthropic" : IA_PROVIDER_PADRAO;
}

/** Modelo configurado; se vazio, usa o padrão do provedor. */
function _obterModeloIA() {
  const m = (obterValorConfig(IA_CFG_MODELO) || "").trim();
  if (m) return m;
  return _obterProviderIA() === "anthropic" ? IA_MODELO_ANTHROPIC_PADRAO : IA_MODELO_GEMINI_PADRAO;
}

/** Diretrizes (system prompt) configuradas; se vazio, usa o padrão embutido. */
function _obterDiretrizesIA() {
  const d = (obterValorConfig(IA_CFG_DIRETRIZES) || "").trim();
  return d || IA_SYSTEM_INSTRUCTION_PADRAO;
}

/** Informa ao front-end se a IA já está configurada (sem revelar a chave). */
function iaConfigurada() {
  return _ok({ configurada: _obterChaveIA() !== "" });
}

/**
 * Retorna a configuração atual da IA para o painel administrativo.
 * NUNCA retorna a chave da API — apenas se está configurada.
 */
function obterConfigIA() {
  return _ok({
    configurada: _obterChaveIA() !== "",
    provider: _obterProviderIA(),
    modelo: _obterModeloIA(),
    diretrizes: _obterDiretrizesIA(),
    diretrizesPadrao: IA_SYSTEM_INSTRUCTION_PADRAO
  });
}

/**
 * Salva provedor, modelo e diretrizes do agente (parametrizável pelo admin).
 * Deixar um campo vazio faz o sistema voltar ao padrão daquele campo.
 */
function salvarConfigIA(cfg) {
  try {
    if (!usuarioEhAdmin()) return _fail("Apenas administradores podem configurar a IA.");
    cfg = cfg || {};
    if (cfg.provider !== undefined) {
      const p = String(cfg.provider).trim().toLowerCase();
      salvarValorConfig(IA_CFG_PROVIDER, p === "anthropic" ? "anthropic" : "gemini");
    }
    if (cfg.modelo !== undefined) {
      salvarValorConfig(IA_CFG_MODELO, String(cfg.modelo).trim());
    }
    if (cfg.diretrizes !== undefined) {
      salvarValorConfig(IA_CFG_DIRETRIZES, String(cfg.diretrizes));
    }
    return _ok({
      provider: _obterProviderIA(),
      modelo: _obterModeloIA(),
      diretrizes: _obterDiretrizesIA(),
      mensagem: "Configurações da IA salvas."
    });
  } catch (err) {
    return _fail("Erro ao salvar configurações da IA: " + err.toString());
  }
}

/** Monta o texto do cenário do produtor sem enviar dados pessoais (sem CPF/nome). */
function _iaMontarCenario(ctx) {
  ctx = ctx || {};
  const linhas = [];
  linhas.push("CENÁRIO DO PRODUTOR:");
  if (ctx.tipoPessoa) linhas.push("- Tipo de pessoa: " + ctx.tipoPessoa);
  if (ctx.enquadramento) linhas.push("- Enquadramento: " + ctx.enquadramento);
  if (ctx.renda !== undefined && ctx.renda !== null && ctx.renda !== "")
    linhas.push("- Renda bruta anual (R$): " + ctx.renda);
  if (ctx.finalidade) linhas.push("- Finalidade/necessidade: " + ctx.finalidade);
  if (ctx.produto) linhas.push("- Produto/atividade: " + ctx.produto);
  if (ctx.cultura) linhas.push("- Cultura: " + ctx.cultura);
  if (ctx.valorPretendido !== undefined && ctx.valorPretendido !== null && ctx.valorPretendido !== "")
    linhas.push("- Valor pretendido (R$): " + ctx.valorPretendido);
  if (ctx.valorTomado !== undefined && ctx.valorTomado !== null && ctx.valorTomado !== "")
    linhas.push("- Custeio já tomado que consome o limite de custeio (R$, inclui ProAgro): " + ctx.valorTomado);

  // Histórico SICOR resumido, separando custeio (consome limite) de investimento.
  const sicor = ctx.sicorResumo || null;
  if (sicor) {
    linhas.push("");
    linhas.push("HISTÓRICO SICOR (safra vigente) — separado por finalidade:");
    linhas.push("- Total CUSTEIO tomado (consome o limite de custeio do grupo): R$ " + (sicor.totalCusteio || 0));
    linhas.push("- Total INVESTIMENTO tomado (limite próprio, NÃO consome custeio): R$ " + (sicor.totalInvestimento || 0));
    if (sicor.totalComercializacao)
      linhas.push("- Total COMERCIALIZAÇÃO tomado: R$ " + sicor.totalComercializacao);
    const itens = Array.isArray(sicor.itens) ? sicor.itens : [];
    if (itens.length) {
      linhas.push("Operações ativas:");
      itens.forEach(function (it) {
        linhas.push("  • " + [it.produto, it.atividade].filter(Boolean).join(" / ") +
          " — finalidade: " + (it.finalidade || "Custeio") +
          " — valor tomado: R$ " + (it.valorTomado || 0));
      });
    }
    linhas.push("IMPORTANTE: ao calcular a margem remanescente de custeio, subtraia do teto de custeio do grupo APENAS o total de custeio — nunca o investimento.");
  }

  const elegiveis = Array.isArray(ctx.linhasElegiveis) ? ctx.linhasElegiveis : [];
  const houveSimulacao = ctx.simulacaoRealizada === true;

  if (elegiveis.length) {
    // O sistema aplicou as regras e encontrou linhas disponíveis.
    linhas.push("");
    linhas.push("VEREDITO DO SISTEMA: " + elegiveis.length + " linha(s) ELEGÍVEL(is) após aplicar as regras (grupo de enquadramento, renda, tipo de pessoa, cultura/tags e limite disponível descontado o valor já tomado):");
    elegiveis.forEach(function (l, i) {
      const partes = [];
      if (l.nome) partes.push(l.nome);
      if (l.orgao) partes.push("órgão: " + l.orgao);
      if (l.taxa) partes.push("taxa: " + l.taxa);
      if (l.prazo) partes.push("prazo: " + l.prazo);
      if (l.carencia) partes.push("carência: " + l.carencia);
      if (l.teto) partes.push("teto: " + l.teto);
      if (l.saldoDisponivel !== undefined && l.saldoDisponivel !== null && l.saldoDisponivel !== "")
        partes.push("saldo disponível: " + l.saldoDisponivel);
      linhas.push((i + 1) + ") " + partes.join(" | "));
    });
    linhas.push("");
    linhas.push("Recomende a melhor opção ENTRE as linhas elegíveis acima (menor custo e prazo adequado).");
    linhas.push("");
    linhas.push("Pergunta: qual a melhor linha de crédito (menor custo e prazo adequado) para este produtor?");
  } else if (houveSimulacao) {
    // O sistema aplicou as regras e NÃO encontrou nenhuma linha disponível.
    linhas.push("");
    linhas.push("VEREDITO DO SISTEMA: NENHUMA linha ficou disponível para este cenário. O sistema aplicou as regras vigentes (grupo de enquadramento, renda, tipo de pessoa, cultura/tags e, principalmente, o limite disponível após descontar o valor já tomado) e nenhuma linha do grupo atendeu.");
    linhas.push("");
    linhas.push("INSTRUÇÕES OBRIGATÓRIAS PARA ESTE CASO:");
    linhas.push("- NÃO apresente nenhuma linha como disponível ou aprovada — isso contrariaria o resultado do sistema.");
    linhas.push("- Explique, de forma objetiva, os prováveis MOTIVOS da ausência de enquadramento (ex.: limite de crédito do grupo já esgotado pelo valor tomado, renda fora da faixa, requisitos/cultura não atendidos).");
    linhas.push("- Se pertinente, indique o que o analista poderia REAVALIAR (ex.: verificar se há saldo em outro grupo/programa, considerar a próxima safra, reduzir o valor pretendido, revisar operações já liquidadas), sempre deixando claro que são hipóteses a confirmar.");
    linhas.push("");
    linhas.push("Pergunta: por que este produtor não tem linha disponível e o que o analista pode reavaliar?");
  } else {
    // Chamada sem simulação prévia (fallback): usar a base das diretrizes.
    linhas.push("");
    linhas.push("O sistema não enviou o resultado de uma simulação; use a base de dados das diretrizes como referência.");
    linhas.push("");
    linhas.push("Pergunta: qual a melhor linha de crédito (menor custo e prazo adequado) para este produtor?");
  }

  return linhas.join("\n");
}

/**
 * Sugere a melhor linha usando o provedor de IA configurado (Gemini ou Anthropic).
 * `contexto` deve conter apenas dados não sensíveis (sem CPF/nome).
 */
function sugerirMelhorLinhaIA(contexto) {
  const chave = _obterChaveIA();
  if (!chave) {
    return _fail("A IA ainda não foi configurada. Um administrador precisa cadastrar a chave da API na aba Configurações do Sistema.");
  }

  try {
    const pergunta = _iaMontarCenario(contexto);
    const system = _obterDiretrizesIA();
    const provider = _obterProviderIA();
    const modelo = _obterModeloIA();

    if (provider === "anthropic") {
      return _iaChamarAnthropic(chave, modelo, system, pergunta);
    }
    return _iaChamarGemini(chave, modelo, system, pergunta);
  } catch (err) {
    return _fail("Erro ao consultar a IA: " + err.toString());
  }
}

/** Chamada à API do Google Gemini (generateContent). */
function _iaChamarGemini(chave, modelo, system, pergunta) {
  const url = IA_ENDPOINT_GEMINI_BASE + encodeURIComponent(modelo) + ":generateContent";

  const generationConfig = { maxOutputTokens: 3072, temperature: 0.2 };
  // Nos modelos "flash" o raciocínio interno consome tokens de saída; desligá-lo
  // evita respostas vazias e reduz custo. Modelos "pro" mantêm o raciocínio.
  if (/flash/i.test(modelo)) generationConfig.thinkingConfig = { thinkingBudget: 0 };

  const payload = {
    system_instruction: { parts: [{ text: system }] },
    contents: [{ role: "user", parts: [{ text: pergunta }] }],
    generationConfig: generationConfig
  };

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { "x-goog-api-key": chave },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const body = resp.getContentText();

  if (code === 400 && /api[_ ]?key|API_KEY_INVALID/i.test(body)) {
    return _fail("Chave da API do Gemini inválida (HTTP 400). Verifique a configuração.");
  }
  if (code === 401 || code === 403) {
    return _fail("Chave da API inválida ou sem permissão (HTTP " + code + "). Verifique a configuração.");
  }
  if (code === 429) {
    return _fail("Limite de uso da IA atingido no momento (HTTP 429). Tente novamente em instantes.");
  }
  if (code === 404) {
    return _fail("Modelo de IA '" + modelo + "' não encontrado (HTTP 404). Ajuste o modelo na aba Configurações do Sistema.");
  }
  if (code < 200 || code >= 300) {
    let detalhe = "";
    try { const e = JSON.parse(body); detalhe = (e.error && e.error.message) ? (" — " + e.error.message) : ""; } catch (x) {}
    return _fail("Falha ao consultar a IA (HTTP " + code + ")" + detalhe);
  }

  const json = JSON.parse(body);
  if (json.promptFeedback && json.promptFeedback.blockReason) {
    return _fail("A solicitação foi bloqueada pela IA (" + json.promptFeedback.blockReason + ").");
  }

  const cand = (json.candidates && json.candidates[0]) || null;
  let texto = "";
  if (cand && cand.content && Array.isArray(cand.content.parts)) {
    for (let i = 0; i < cand.content.parts.length; i++) {
      if (cand.content.parts[i] && cand.content.parts[i].text) texto += cand.content.parts[i].text;
    }
  }
  texto = texto.trim();
  if (!texto) {
    if (cand && cand.finishReason === "MAX_TOKENS") {
      return _fail("A resposta da IA excedeu o limite de tokens. Tente reduzir as diretrizes ou usar outro modelo.");
    }
    if (cand && cand.finishReason === "SAFETY") {
      return _fail("A resposta foi bloqueada por segurança pela IA.");
    }
    return _fail("A IA retornou uma resposta vazia.");
  }

  return _ok({ recomendacao: texto, provedor: "gemini", modelo: modelo });
}

/** Chamada à API da Anthropic (Claude / Messages). */
function _iaChamarAnthropic(chave, modelo, system, pergunta) {
  const payload = {
    model: modelo,
    max_tokens: 2000,
    system: system,
    messages: [{ role: "user", content: pergunta }]
  };

  const resp = UrlFetchApp.fetch(IA_ENDPOINT_ANTHROPIC, {
    method: "post",
    contentType: "application/json",
    headers: { "x-api-key": chave, "anthropic-version": "2023-06-01" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const body = resp.getContentText();

  if (code === 401 || code === 403) {
    return _fail("Chave da API inválida ou sem permissão (HTTP " + code + "). Verifique a configuração.");
  }
  if (code === 429) {
    return _fail("Limite de uso da IA atingido no momento (HTTP 429). Tente novamente em instantes.");
  }
  if (code < 200 || code >= 300) {
    let detalhe = "";
    try { const e = JSON.parse(body); detalhe = (e.error && e.error.message) ? (" — " + e.error.message) : ""; } catch (x) {}
    return _fail("Falha ao consultar a IA (HTTP " + code + ")" + detalhe);
  }

  const json = JSON.parse(body);
  if (json.stop_reason === "refusal") {
    return _fail("A IA não pôde responder a esta solicitação.");
  }

  let texto = "";
  if (Array.isArray(json.content)) {
    for (let i = 0; i < json.content.length; i++) {
      if (json.content[i] && json.content[i].type === "text") texto += json.content[i].text;
    }
  }
  texto = texto.trim();
  if (!texto) return _fail("A IA retornou uma resposta vazia.");

  return _ok({ recomendacao: texto, provedor: "anthropic", modelo: modelo });
}
