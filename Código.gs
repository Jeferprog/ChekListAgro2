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

// Manter compatibilidade com código anterior
function inicializarPlanilha() {
  inicializarSistema();

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

    // Linha de exemplo
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

    SHEET_BASE.appendRow(["12345678909", "12345-6", "JEFERSON OLIVEIRA DA SILVA", "Física", 350000]);
    SHEET_BASE.appendRow(["98765432100", "54321-0", "MARIA SOUZA REIS", "Física", 1200000]);
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

// Handler principal para servir a aplicação web
function doGet() {
  inicializarPlanilha();
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Cresol Crédito Rural")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
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

// ==================== LINHAS DE CRÉDITO ====================

function listarTodasAsLinhas() {
  if (!SHEET_LINHAS) return [];
  const dados = SHEET_LINHAS.getDataRange().getValues();
  if (dados.length <= 1) return [];

  // Carrega checklists customizados
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

function buscarAssociado(termo) {
  if (!SHEET_BASE) return { sucesso: false, error: "Aba Base não encontrada." };

  const termoAlvo = String(termo).replace(/\D/g, "").replace(/^0+/, "");
  if (!termoAlvo) return { sucesso: false, error: "Termo de busca inválido." };

  const dados = SHEET_BASE.getDataRange().getValues();
  if (dados.length < 2) return { sucesso: false, error: "Base de associados vazia." };

  const H = dados[0];
  for (let r = 1; r < dados.length; r++) {
    const cpf = String(dados[r][H.indexOf("nr_cpf_cnpj")]).replace(/\D/g, "").replace(/^0+/, "");
    const conta = String(dados[r][H.indexOf("nr_conta_corrente")]).replace(/\D/g, "").replace(/^0+/, "");

    if (cpf === termoAlvo || conta === termoAlvo) {
      const rendaAnual = parseFloat(dados[r][H.indexOf("vl_anual_fonte_renda_total")]) || 0;
      return {
        sucesso: true,
        nome: String(dados[r][H.indexOf("nm_nome")]),
        cpfCnpj: String(dados[r][H.indexOf("nr_cpf_cnpj")]),
        conta: String(dados[r][H.indexOf("nr_conta_corrente")]),
        tipo: String(dados[r][H.indexOf("ds_pessoa_tipo")]),
        rendaAnual: rendaAnual,
        rendaMensal: Math.round(rendaAnual / 12)
      };
    }
  }
  return { sucesso: false, error: "Associado não encontrado." };
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

  // Fallback para documentos padrão cadastrados na linha
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

// EXTRATOR DOCX SIMULADO NO BACKEND (ALINHADO COM SISTEMA ORIGINAL)
function processarArquivoCresol(arquivoInfo) {
  return {
    success: true,
    total: 3,
    itens: [
      { idx: 0, nome: "Pronaf Agroindústria (Faixa II)", rural: true },
      { idx: 1, nome: "Pronaf Jovem Empreendedor", rural: true },
      { idx: 2, nome: "RenovAgro Recuperação de Pastagens", rural: true }
    ]
  };
}

function aplicarAtualizacaoCresol(selecionados) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SS.getSheetByName("Linhas");
  if (!sheet) return { success: false };
  
  // Criar linhas selecionadas exatamente idênticas ao sistema de desenvolvimento
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
  // Match folder ID
  var matchFolder = link.match(/\/folders\/([a-zA-Z0-9-_]+)/);
  if (matchFolder) return { type: "folder", id: matchFolder[1] };
  // Match file ID
  var matchFile = link.match(/\/file\/d\/([a-zA-Z0-9-_]+)/) || link.match(/id=([a-zA-Z0-9-_]+)/);
  if (matchFile) return { type: "file", id: matchFile[1] };
  return null;
}

function processarECarregarCSVBase(sheet, csvContent) {
  const headers = ["nr_cpf_cnpj", "nr_conta_corrente", "nm_nome", "ds_pessoa_tipo", "vl_anual_fonte_renda_total"];
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  let delimiter = ",";
  if (csvContent.indexOf(";") !== -1) {
    delimiter = ";";
  }
  
  const parsedData = Utilities.parseCsv(csvContent, delimiter);
  if (parsedData.length <= 1) return 0;
  
  const csvHeaders = parsedData[0].map(function(h) { return h.trim().toLowerCase(); });
  
  const cpfIdx = csvHeaders.indexOf("nr_cpf_cnpj") !== -1 ? csvHeaders.indexOf("nr_cpf_cnpj") : csvHeaders.indexOf("cpf");
  const contaIdx = csvHeaders.indexOf("nr_conta_corrente") !== -1 ? csvHeaders.indexOf("nr_conta_corrente") : csvHeaders.indexOf("conta");
  const nomeIdx = csvHeaders.indexOf("nm_nome") !== -1 ? csvHeaders.indexOf("nm_nome") : csvHeaders.indexOf("nome");
  const tipoIdx = csvHeaders.indexOf("ds_pessoa_tipo") !== -1 ? csvHeaders.indexOf("ds_pessoa_tipo") : csvHeaders.indexOf("tipo");
  const rendaIdx = csvHeaders.indexOf("vl_anual_fonte_renda_total") !== -1 ? csvHeaders.indexOf("vl_anual_fonte_renda_total") : csvHeaders.indexOf("renda");
  
  var recordsAdded = 0;
  var rowsToAppend = [];
  
  for (var i = 1; i < parsedData.length; i++) {
    var row = parsedData[i];
    if (row.length < 2) continue;
    
    var cpf = cpfIdx !== -1 ? String(row[cpfIdx]).trim() : "";
    var conta = contaIdx !== -1 ? String(row[contaIdx]).trim() : "";
    var nome = nomeIdx !== -1 ? String(row[nomeIdx]).trim().toUpperCase() : "";
    var tipo = tipoIdx !== -1 ? String(row[tipoIdx]).trim() : "Física";
    
    var rendaText = rendaIdx !== -1 ? String(row[rendaIdx]).trim() : "0";
    rendaText = rendaText.replace(/[R$\s]/g, "").replace(/\./g, "").replace(",", ".");
    var renda = parseFloat(rendaText) || 0;
    
    if (cpf || conta) {
      rowsToAppend.push([cpf, conta, nome, tipo, renda]);
      recordsAdded++;
    }
  }
  
  if (rowsToAppend.length > 0) {
    sheet.getRange(2, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
  }
  
  return recordsAdded;
}

// ==================== IMPORTAÇÃO DE DADOS ====================

function atualizarBaseAssociados() {
  if (!SHEET_BASE) return { success: false, error: "Aba Base não encontrada." };

  const linkBase = obterValorConfig("Link Pasta Base de Associados");
  if (!linkBase) {
    return { success: false, error: "Link da Pasta Base de Associados não configurado nas Configurações." };
  }

  const driveInfo = obterIdDoDrive(linkBase);
  if (!driveInfo) {
    return { success: false, error: "Link de pasta inválido. Por favor, insira um link válido do Google Drive." };
  }

  try {
    var file;
    if (driveInfo.type === "file") {
      file = DriveApp.getFileById(driveInfo.id);
    } else {
      var folder = DriveApp.getFolderById(driveInfo.id);
      var files = folder.getFiles();
      while (files.hasNext()) {
        var f = files.next();
        var fName = f.getName().toLowerCase();
        if (fName.includes("basedepessoas") || fName.endsWith(".csv")) {
          file = f;
          break;
        }
      }
    }

    if (!file) {
      return { success: false, error: "Nenhum arquivo 'basedepessoas.csv' ou arquivo .csv correspondente encontrado na pasta do Drive." };
    }

    var contentText = file.getBlob().getDataAsString("UTF-8");
    if (contentText.indexOf("") !== -1 || contentText.indexOf("") !== -1) {
      contentText = file.getBlob().getDataAsString("ISO-8859-1");
    }

    var numLines = processarECarregarCSVBase(SHEET_BASE, contentText);
    return { success: true, registros: numLines };

  } catch (err) {
    return { success: false, error: "Erro ao acessar arquivos do Google Drive: " + err.message };
  }
}

function processarArquivoCredito(filename, contentText) {
  if (!SHEET_BASE_CREDITO) return { success: false, error: "Aba BaseCredito não encontrada." };

  const headers = ["nr_cpf_cnpj", "ano_safra", "produto", "atividade", "if_fin", "valor_financiado", "aliquota_proagro", "valor_tomado"];
  SHEET_BASE_CREDITO.clearContents();
  SHEET_BASE_CREDITO.getRange(1, 1, 1, headers.length).setValues([headers]);

  var delimiter = ",";
  if (contentText.indexOf(";") !== -1) {
    delimiter = ";";
  }

  var parsedData = Utilities.parseCsv(contentText, delimiter);
  if (parsedData.length <= 1) {
    return { success: true, registros: 0 };
  }

  var csvHeaders = parsedData[0].map(function(h) { return h.trim().toLowerCase(); });

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
    if (row.length < 2) continue;

    var cpf = cpfIdx !== -1 ? String(row[cpfIdx]).trim() : "";
    var safra = safraIdx !== -1 ? String(row[safraIdx]).trim() : "2025/2026";
    var produto = produtoIdx !== -1 ? String(row[produtoIdx]).trim().toUpperCase() : "CRÉDITO RURAL";
    var atividade = atividadeIdx !== -1 ? String(row[atividadeIdx]).trim() : "Outros";
    var ifFin = ifIdx !== -1 ? String(row[ifIdx]).trim() : "Cresol";

    var valFinText = valorFinIdx !== -1 ? String(row[valorFinIdx]).trim() : "0";
    valFinText = valFinText.replace(/[R$\s]/g, "").replace(/\./g, "").replace(",", ".");
    var valFin = parseFloat(valFinText) || 0;

    var aliquotaText = proagroIdx !== -1 ? String(row[proagroIdx]).trim() : "0";
    aliquotaText = aliquotaText.replace(/[%\s]/g, "").replace(",", ".");
    var aliquota = parseFloat(aliquotaText) || 0;

    var valTomText = valorTomIdx !== -1 ? String(row[valorTomIdx]).trim() : String(valFin);
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

  return { success: true, registros: count };
}

function processarArquivoAssociados(filename, contentText) {
  if (!SHEET_BASE) return { success: false, error: "Aba Base não encontrada." };

  try {
    var numLines = processarECarregarCSVBase(SHEET_BASE, contentText);
    return { success: true, registros: numLines };
  } catch (err) {
    return { success: false, error: "Erro ao processar arquivo de associados: " + err.message };
  }
}

