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
      "Itens Financiáveis", "Culturas Financiadas", "Taxa (descrição)"
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
      "3% a.a."
    ]);
  }
}

function inicializarSheetConfig() {
  if (SHEET_CONFIG.getLastRow() === 0) {
    SHEET_CONFIG.appendRow(["Parâmetro", "Valor"]);
    SHEET_CONFIG.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#f58220").setFontColor("white");
    SHEET_CONFIG.appendRow(["Link Consulta Crédito (SICOR/CACR)", "https://www.bcb.gov.br/sicor/"]);
    SHEET_CONFIG.appendRow(["Link Pasta Base de Associados", ""]);
    SHEET_CONFIG.appendRow(["E-mails Administradores", "gestor@cresol.com.br, analista@cresol.com.br"]);
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
    const cabecalhos = ["nr_cpf_cnpj", "ano_safra", "produto", "atividade", "if_fin", "valor_financiado", "aliquota_proagro", "valor_tomado"];
    SHEET_BASE_CREDITO.appendRow(cabecalhos);
    SHEET_BASE_CREDITO.getRange(1, 1, 1, cabecalhos.length).setFontWeight("bold").setBackground("#005c46").setFontColor("white");

    SHEET_BASE_CREDITO.appendRow(["12345678909", "2025/2026", "PRONAF CUSTEIO AGRÍCOLA", "Milho", "Cresol", 45000, 3, 46350]);
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
  return emailsText.split(",").map(e => e.trim()).filter(Boolean);
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

          const cap = _capCultura(parametros.produto, culturasTxt);
          if (cap) {
            if (cap.tipo === "max") {
              if (limiteMax === 0 || cap.valor < limiteMax) limiteMax = cap.valor;
            } else if (cap.tipo === "min") {
              if (cap.valor > limiteMin) limiteMin = cap.valor;
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
            valorTomado: parametros.valorTomado || 0
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
  return dados.slice(1).map(function(linha) {
    const nome = String(linha[H.indexOf("Nome Linha")]);
    const documentos = String(linha[H.indexOf("Documentos Necessários")]);
    const temChecklist = checklistLines.indexOf(nome) !== -1 || documentos.trim() !== "";

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

function buscarCreditoTomado(cpf) {
  if (!SHEET_BASE_CREDITO) return { sucesso: false, items: [], totalFinanciado: 0 };

  const cpfAlvo = String(cpf).replace(/\D/g, "").replace(/^0+/, "");
  if (!cpfAlvo) return { sucesso: false, items: [], totalFinanciado: 0 };

  const dados = SHEET_BASE_CREDITO.getDataRange().getValues();
  if (dados.length < 2) return { sucesso: true, items: [], totalFinanciado: 0 };

  const H = dados[0];
  const items = [];
  let totalFinanciado = 0;

  for (let r = 1; r < dados.length; r++) {
    const rowCpf = String(dados[r][H.indexOf("nr_cpf_cnpj")]).replace(/\D/g, "").replace(/^0+/, "");
    if (rowCpf === cpfAlvo) {
      const valorFin = parseFloat(dados[r][H.indexOf("valor_financiado")]) || 0;
      const valorTom = parseFloat(dados[r][H.indexOf("valor_tomado")]) || valorFin;
      totalFinanciado += valorTom;

      items.push({
        anoSafra: String(dados[r][H.indexOf("ano_safra")]),
        produto: String(dados[r][H.indexOf("produto")]),
        atividade: String(dados[r][H.indexOf("atividade")]),
        ifFin: String(dados[r][H.indexOf("if_fin")]),
        valorFinanciado: valorFin,
        aliquotaProagro: parseFloat(dados[r][H.indexOf("aliquota_proagro")]) || 0,
        valorTomado: valorTom
      });
    }
  }

  return { sucesso: true, items: items, totalFinanciado: totalFinanciado };
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

// ==================== PROCESSAMENTO DE ARQUIVOS ====================

function processarArquivoCresol(arquivoInfo) {
  try {
    // Itens padrão de demonstração
    const itens = [
      { idx: 0, nome: "Pronaf Agroindústria (Faixa II)", rural: true },
      { idx: 1, nome: "Pronaf Jovem Empreendedor", rural: true },
      { idx: 2, nome: "RenovAgro Recuperação de Pastagens", rural: true }
    ];

    return {
      sucesso: true,
      total: itens.length,
      itens: itens
    };
  } catch (e) {
    Logger.log("Erro ao processar arquivo Cresol: " + e);
    return {
      sucesso: false,
      erro: "Erro ao processar arquivo: " + e.toString()
    };
  }
}

function aplicarAtualizacaoCresol(selecionados) {
  if (!SHEET_LINHAS) return { success: false };

  const simula = [
    {
      nome: "Pronaf Agroindústria (Faixa II)",
      orgao: "BNDES / Cresol",
      finalidadePrincipal: "Investimento",
      finalidades: "infraestrutura,investimento",
      enquadramento: "Sem limite/R$ 500 mil",
      taxaMin: 8.0,
      taxaMax: 8.0,
      taxaDescricao: "8% a.a.",
      prazo: 120,
      carencia: 36,
      limiteMin: 0,
      limiteMax: 500000,
      requisitos: "Construção ou melhoramento de pequenas agroindústrias Pronaf.",
      documentos: "DAP-Pronaf PJ, RG, CPF dos sócios, projeto de viabilidade.",
      status: "Ativa",
      observacoes: "Extraído via DOCX",
      itensFinanciaveis: "Investimentos em infraestrutura, beneficiamento, processamento de alimentos.",
      culturas: ""
    },
    {
      nome: "Pronaf Jovem Empreendedor",
      orgao: "BNDES / Cresol",
      finalidadePrincipal: "Investimento",
      finalidades: "investimento",
      enquadramento: "Sem limite/R$ 500 mil",
      taxaMin: 3.0,
      taxaMax: 3.0,
      taxaDescricao: "3% a.a.",
      prazo: 120,
      carencia: 24,
      limiteMin: 0,
      limiteMax: 35000,
      requisitos: "Jovens entre 16 e 29 anos com curso técnico ou capacitação na área.",
      documentos: "CAF/DAP-Pronaf Jovem, RG, CPF, Certificado de capacitação.",
      status: "Ativa",
      observacoes: "Extraído via DOCX",
      itensFinanciaveis: "Projetos de investimento de jovens agricultores.",
      culturas: ""
    },
    {
      nome: "RenovAgro Recuperação de Pastagens",
      orgao: "BNDES / Cresol",
      finalidadePrincipal: "Investimento",
      finalidades: "investimento,sustentabilidade",
      enquadramento: "Conforme análise",
      taxaMin: 8.5,
      taxaMax: 8.5,
      taxaDescricao: "8.5% a.a.",
      prazo: 144,
      carencia: 60,
      limiteMin: 0,
      limiteMax: 5000000,
      requisitos: "Análise de solo e recomendação agronômica para recuperação de pastagem degradada.",
      documentos: "RG, CPF, Certidão da propriedade, laudo agronômico.",
      status: "Ativa",
      observacoes: "Extraído via DOCX",
      itensFinanciaveis: "Calcário, sementes de capim, adubo, cercas, curva de nível.",
      culturas: "PASTAGEM, FORRAGEIRAS"
    }
  ];

  let count = 0;
  selecionados.forEach(idx => {
    if (simula[idx]) {
      adicionarLinha(simula[idx]);
      count++;
    }
  });

  return { success: true, linhas: count };
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
    return { sucesso: false, erro: "CSV vazio ou inválido." };
  }

  const primeiraLinha = conteudo.split("\n")[0] || "";
  const delim = (primeiraLinha.split(";").length > primeiraLinha.split(",").length) ? ";" : ",";
  const dados = Utilities.parseCsv(conteudo, delim);
  if (!dados || dados.length < 2) return { sucesso: false, erro: "CSV vazio ou inválido." };

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

  return { sucesso: true, registros: norm.length - 1, atualizado: new Date().toLocaleString("pt-BR") };
}

/**
 * Busca o basedepessoas.csv no link configurado e regrava a aba Base.
 * Também é chamada pelo trigger automático de atualização diária.
 */
function atualizarBaseAssociados() {
  try {
    if (!SHEET_BASE) return { sucesso: false, erro: "Aba Base não encontrada." };

    const link = obterLinkBase();
    if (!link) return { sucesso: false, erro: "Configure o link da pasta/arquivo da base na aba Administrativo." };

    const conteudo = _lerCsvDoLink(link, "basedepessoas.csv");
    if (!conteudo) return { sucesso: false, erro: "Arquivo basedepessoas.csv não encontrado no link informado." };

    return _gravarCsvNaBase(conteudo);
  } catch (e) {
    Logger.log("Erro em atualizarBaseAssociados: " + e.toString());
    return { sucesso: false, erro: e.toString() };
  }
}

function processarArquivoCreditoBase(arquivoInfo) {
  if (!SHEET_BASE_CREDITO) return { sucesso: false, error: "Aba BaseCredito não encontrada." };

  try {
    const contentText = arquivoInfo.conteudo || arquivoInfo;
    if (!contentText) {
      return { sucesso: false, erro: "Arquivo vazio" };
    }

    const headers = ["nr_cpf_cnpj", "ano_safra", "produto", "atividade", "if_fin", "valor_financiado", "aliquota_proagro", "valor_tomado"];
    SHEET_BASE_CREDITO.clearContents();
    SHEET_BASE_CREDITO.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Detectar delimitador
    var delimiter = ",";
    if (contentText.indexOf(";") > -1) {
      delimiter = ";";
    }

    // Fazer parse do CSV com tratamento de erro
    var parsedData;
    try {
      parsedData = Utilities.parseCsv(contentText, delimiter);
    } catch (e) {
      // Se falhar, tentar com outro delimitador
      delimiter = delimiter === "," ? ";" : ",";
      parsedData = Utilities.parseCsv(contentText, delimiter);
    }

    if (!parsedData || parsedData.length <= 1) {
      return { sucesso: true, registros: 0, atualizado: new Date().toLocaleString('pt-BR') };
    }

    var csvHeaders = parsedData[0].map(function(h) { return String(h || "").trim().toLowerCase(); });

    const cpfIdx = csvHeaders.indexOf("nr_cpf_cnpj") !== -1 ? csvHeaders.indexOf("nr_cpf_cnpj") : csvHeaders.indexOf("cpf");
    const safraIdx = csvHeaders.indexOf("ano_safra") !== -1 ? csvHeaders.indexOf("ano_safra") : csvHeaders.indexOf("safra");
    const produtoIdx = csvHeaders.indexOf("produto") !== -1 ? csvHeaders.indexOf("produto") : csvHeaders.indexOf("linha");
    const atividadeIdx = csvHeaders.indexOf("atividade") !== -1 ? csvHeaders.indexOf("atividade") : csvHeaders.indexOf("cultura");
    const ifIdx = csvHeaders.indexOf("if_fin") !== -1 ? csvHeaders.indexOf("if_fin") : csvHeaders.indexOf("instituicao");
    const valorFinIdx = csvHeaders.indexOf("valor_financiado") !== -1 ? csvHeaders.indexOf("valor_financiado") : csvHeaders.indexOf("valor");
    const proagroIdx = csvHeaders.indexOf("aliquota_proagro") !== -1 ? csvHeaders.indexOf("aliquota_proagro") : csvHeaders.indexOf("proagro");
    const valorTomIdx = csvHeaders.indexOf("valor_tomado") !== -1 ? csvHeaders.indexOf("valor_tomado") : csvHeaders.indexOf("tomado");

    var count = 0;
    var rowsToAppend = [];

    for (var i = 1; i < parsedData.length; i++) {
      var row = parsedData[i];
      if (!row || row.length < 2) continue;

      var cpf = cpfIdx !== -1 ? String(row[cpfIdx] || "").trim() : "";
      var safra = safraIdx !== -1 ? String(row[safraIdx] || "").trim() : "2025/2026";
      var produto = produtoIdx !== -1 ? String(row[produtoIdx] || "").trim().toUpperCase() : "CRÉDITO RURAL";
      var atividade = atividadeIdx !== -1 ? String(row[atividadeIdx] || "").trim() : "Outros";
      var ifFin = ifIdx !== -1 ? String(row[ifIdx] || "").trim() : "Cresol";

      var valFinText = valorFinIdx !== -1 ? String(row[valorFinIdx] || "0").trim() : "0";
      valFinText = valFinText.replace(/[R$\s]/g, "").replace(/\./g, "").replace(",", ".");
      var valFin = parseFloat(valFinText) || 0;

      var aliquotaText = proagroIdx !== -1 ? String(row[proagroIdx] || "0").trim() : "0";
      aliquotaText = aliquotaText.replace(/[%\s]/g, "").replace(",", ".");
      var aliquota = parseFloat(aliquotaText) || 0;

      var valTomText = valorTomIdx !== -1 ? String(row[valorTomIdx] || valFin).trim() : String(valFin);
      valTomText = valTomText.replace(/[R$\s]/g, "").replace(/\./g, "").replace(",", ".");
      var valTom = parseFloat(valTomText) || valFin;

      if (cpf) {
        rowsToAppend.push([cpf, safra, produto, atividade, ifFin, valFin, aliquota, valTom]);
        count++;
      }
    }

    if (rowsToAppend.length > 0) {
      SHEET_BASE_CREDITO.getRange(2, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
    }

    return { sucesso: true, registros: count, atualizado: new Date().toLocaleString('pt-BR') };
  } catch (err) {
    Logger.log("Erro ao processar arquivo de crédito: " + err);
    return { sucesso: false, erro: "Erro ao processar arquivo: " + err.toString() };
  }
}

function processarArquivoAssociados(filename, contentText) {
  if (!SHEET_BASE) return { sucesso: false, erro: "Aba Base não encontrada." };

  try {
    return _gravarCsvNaBase(contentText);
  } catch (err) {
    return { sucesso: false, erro: "Erro ao processar arquivo de associados: " + err.toString() };
  }
}
