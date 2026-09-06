const SHEET_NAME = null; // caso outra aba seja criada.

const COLUMNS = {
  NOME_PESSOA: 'nome da pessoa',
  CARGO_PESSOA: 'cargo da pessoa',
  LINKEDIN_PESSOA: 'linkedin da pessoa',
  EMAIL_PESSOA: 'email da pessoa',
  TELEFONE_PESSOA: 'telefone da pessoa',
  NOME_EMPRESA: 'nome da empresa',
  CNPJ_EMPRESA: 'cnpj da empresa',
  SETOR_EMPRESA: 'setor da empresa',
  SITE_EMPRESA: 'site da empresa',
  EMAIL_EMPRESA: 'email da empresa',
  TELEFONE_EMPRESA: 'telefone da empresa',
  LINKEDIN_EMPRESA: 'Linkedin da Empresa',
  // Colunas de resultado (serão criadas automaticamente se não existirem).
  PERFIL: 'Perfil',
  TIPO_VENDA: 'Tipo de Venda',
  DORES: 'Dores',
  SCRIPT: 'Script'
};

const GEMINI_MODEL = 'gemini-3.5-flash-lite';
const API_KEY_PROP = 'GEMINI_API_KEY';

// MENU
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Analisar empresa')
    .addItem('Analisar linha selecionada', 'analisarEmpresaSelecionada')
    .addSeparator()
    .addItem('Configurar API Key do Gemini', 'configurarApiKey')
    .addToUi();
}

// API KEY
function configurarApiKey() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Configurar API Key do Gemini',
    'Cole aqui sua API Key do Google AI Studio (Gemini):',
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() === ui.Button.OK) {
    const key = resp.getResponseText().trim();
    if (key) {
      PropertiesService.getScriptProperties().setProperty(API_KEY_PROP, key);
      ui.alert('API Key salva com sucesso.');
    } else {
      ui.alert('Nenhuma key foi informada.');
    }
  }
}

// FUNÇÃO PRINCIPAL
function analisarEmpresaSelecionada() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SHEET_NAME
    ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
    : SpreadsheetApp.getActiveSheet();

  const apiKey = PropertiesService.getScriptProperties().getProperty(API_KEY_PROP);
  if (!apiKey) {
    ui.alert('Configure a API Key primeiro em "Analisar empresa" > "Configurar API Key do Gemini".');
    return;
  }

  const activeRange = sheet.getActiveRange();
  if (!activeRange) {
    ui.alert('Selecione uma celula na linha da empresa que deseja analisar.');
    return;
  }
  const row = activeRange.getRow();
  if (row === 1) {
    ui.alert('Selecione uma linha de dados, nao o cabecalho.');
    return;
  }

  const colIndex = obterOuCriarColunas(sheet);

  const obrigatorias = [
    COLUMNS.NOME_PESSOA, COLUMNS.CARGO_PESSOA,
    COLUMNS.NOME_EMPRESA, COLUMNS.SITE_EMPRESA, COLUMNS.LINKEDIN_EMPRESA
  ];
  for (const col of obrigatorias) {
    if (!colIndex[col]) {
      ui.alert('Coluna "' + col + '" nao encontrada no cabecalho (linha 1).');
      return;
    }
  }

  const dados = {
    empresa: sheet.getRange(row, colIndex[COLUMNS.NOME_EMPRESA]).getValue(),
    site: sheet.getRange(row, colIndex[COLUMNS.SITE_EMPRESA]).getValue(),
    linkedinEmpresa: sheet.getRange(row, colIndex[COLUMNS.LINKEDIN_EMPRESA]).getValue(),
    setor: colIndex[COLUMNS.SETOR_EMPRESA] ? sheet.getRange(row, colIndex[COLUMNS.SETOR_EMPRESA]).getValue() : '',
    lead: sheet.getRange(row, colIndex[COLUMNS.NOME_PESSOA]).getValue(),
    cargo: sheet.getRange(row, colIndex[COLUMNS.CARGO_PESSOA]).getValue(),
    linkedinPessoa: colIndex[COLUMNS.LINKEDIN_PESSOA] ? sheet.getRange(row, colIndex[COLUMNS.LINKEDIN_PESSOA]).getValue() : ''
  };

  if (!dados.empresa) {
    ui.alert('A coluna "' + COLUMNS.NOME_EMPRESA + '" esta vazia nessa linha.');
    return;
  }

  const cellPerfil = sheet.getRange(row, colIndex[COLUMNS.PERFIL]);
  cellPerfil.setValue('Analisando...');
  SpreadsheetApp.flush();

  try {
    const resultado = chamarGeminiAPI(dados, apiKey);
    preencherResultado(sheet, row, colIndex, resultado);
    ui.alert('Analise concluida para: ' + dados.empresa);
  } catch (err) {
    cellPerfil.setValue('');
    ui.alert('Erro ao analisar: ' + err.message);
    Logger.log(err);
  }
}

// Garante que as colunas de resultado existam
function obterOuCriarColunas(sheet) {
  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colIndex = {};

  headerRow.forEach((h, i) => {
    if (h) colIndex[h.toString().trim()] = i + 1;
  });

  const saidas = [COLUMNS.PERFIL, COLUMNS.TIPO_VENDA, COLUMNS.DORES, COLUMNS.SCRIPT];
  let proximaCol = lastCol;

  saidas.forEach((nomeColuna) => {
    if (!colIndex[nomeColuna]) {
      proximaCol += 1;
      sheet.getRange(1, proximaCol).setValue(nomeColuna);
      colIndex[nomeColuna] = proximaCol;
    }
  });

  return colIndex;
}

// PROMPT PARA API do GEMINI
function chamarGeminiAPI(dados, apiKey) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;

  const promptBase = `
Aja como um especialista em prospeccao B2B. Analise as informacoes da empresa, seu site e LinkedIn para identificar o perfil de compra e determinar se a venda e consultiva ou transacional.

Identifique de 3 a 5 dores comerciais especificas daquela empresa, evitando dores genericas. Quando possivel, baseie as dores em evidencias encontradas (segmento de atuacao, porte, site, linkedin).

Depois, crie um script personalizado para o lead, adaptado a empresa, cargo, segmento e dores identificadas.

Use os roteiros abaixo como referencia de linguagem e estrutura, mas nao copie-os. O resultado deve parecer uma abordagem personalizada e natural.

Script consultivo (referencia de tom/estrutura):
"Fala, Felipe. Tudo bem? Te acompanho no linkedin e sei que a BMC e uma empresa que trabalha com venda de equipamentos pesados, alem de servicos e suporte tecnico. Hoje, nossos clientes do seu segmento, especialmente os que vendem muito para o agro, trazem desafios vinculados a captacao de novos clientes, exatamente pela dificuldade de acessar os decisores, alem de outras dores no comercial: contratacao de time comercial, gestao de todos os dados (leads, carteira, inativos, ex clientes), entre outros pontos. Queria entender seu momento, considerando inclusive o cenario politico e economico brasileiro, e te apresentar alguns cases de empresas de engenharia e industria que estao conseguindo crescer sem elevar muito investimento em aquisicao. Pode me passar seu contato para conversarmos?"

Script transacional (referencia de tom/estrutura):
"Lead, a gente ja tem atuado com outras industrias, principalmente no ramo equipamentos, e principalmente atuado com a resolucao de alguns desafios de empresas que tem ainda um desafio de gestao de carteira, que nao conseguem ter a taxa de positivacao dentro do que estao buscando, tem esse desafio de garantir a recompra desses clientes ali no dia a dia e tem uma cobertura assertiva ali tambem da carteira de clientes positivados."

Dados da empresa a ser analisada:
- Empresa: ${dados.empresa}
- Setor: ${dados.setor}
- Site: ${dados.site}
- LinkedIn da Empresa: ${dados.linkedinEmpresa}
- Nome do Lead: ${dados.lead}
- Cargo do Lead: ${dados.cargo}
- LinkedIn do Lead: ${dados.linkedinPessoa}

Responda EXCLUSIVAMENTE em formato JSON valido, sem markdown, sem texto adicional, seguindo exatamente esta estrutura:
{
  "perfil": "Consultivo ou Transacional",
  "tipo_venda": "Consultiva ou Transacional",
  "dores": "lista das dores em texto corrido ou com marcadores separados por ; ",
  "script": "mensagem personalizada completa para o lead"
}
`.trim();

  const payload = {
    contents: [
      {
        role: 'user',
        parts: [{ text: promptBase }]
      }
    ],
    generationConfig: {
      temperature: 0.7,
      responseMimeType: 'application/json'
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code !== 200) {
    throw new Error('Falha na API Gemini (HTTP ' + code + '): ' + text);
  }

  const json = JSON.parse(text);
  const rawText = json && json.candidates && json.candidates[0] &&
    json.candidates[0].content && json.candidates[0].content.parts &&
    json.candidates[0].content.parts[0] && json.candidates[0].content.parts[0].text;

  if (!rawText) {
    throw new Error('Resposta da API vazia ou em formato inesperado.');
  }

  let parsed;
  try {
    parsed = JSON.parse(rawText);
  } catch (e) {
    const match = rawText.match(/\{[\s\S]*\}/);
    if (match) {
      parsed = JSON.parse(match[0]);
    } else {
      throw new Error('Nao foi possivel interpretar o JSON retornado pela IA.');
    }
  }

  return parsed;
}

// Preenche a planilha.
function preencherResultado(sheet, row, colIndex, resultado) {
  const setIfExists = (colName, valor) => {
    if (colIndex[colName]) {
      sheet.getRange(row, colIndex[colName]).setValue(valor || '');
    }
  };

  setIfExists(COLUMNS.PERFIL, resultado.perfil);
  setIfExists(COLUMNS.TIPO_VENDA, resultado.tipo_venda);
  setIfExists(COLUMNS.DORES, resultado.dores);
  setIfExists(COLUMNS.SCRIPT, resultado.script);
}
