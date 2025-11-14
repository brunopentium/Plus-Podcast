/**
 * Plus Podcast - Finanças Apps Script backend.
 */

// Nome das abas utilizadas na planilha.
const NOME_ABA_RESULTADO = 'Resultado';
const NOME_ABA_SAIDAS = 'Saidas';
const NOME_ABA_ENTRADAS = 'Entradas';
const NOME_ABA_CLIENTES = 'Clientes Ativos';

/**
 * Função executada ao acessar o web app.
 * Carrega o template HTML principal e configura metadados da página.
 *
 * @return {HtmlOutput} Página renderizada para o web app.
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  const htmlOutput = template.evaluate();
  htmlOutput.setTitle('Plus Podcast - Finanças');
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return htmlOutput;
}

/**
 * Permite incluir arquivos HTML parciais dentro do template principal.
 *
 * @param {string} filename Nome do arquivo HTML parcial.
 * @return {string} Conteúdo renderizado do arquivo.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Recupera os dados do resumo financeiro para o dashboard.
 *
 * @return {{saldoAtual: number, totalEntradas: number, totalSaidas: number}}
 */
function getDashboardData() {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheetResultado = spreadsheet.getSheetByName(NOME_ABA_RESULTADO);
  const saldoAtual = Number(sheetResultado.getRange(2, 3).getValue()) || 0;

  const totalEntradas = somarValoresColuna(spreadsheet.getSheetByName(NOME_ABA_ENTRADAS));
  const totalSaidas = somarValoresColuna(spreadsheet.getSheetByName(NOME_ABA_SAIDAS));

  return {
    saldoAtual: saldoAtual,
    totalEntradas: totalEntradas,
    totalSaidas: totalSaidas,
  };
}

/**
 * Soma todos os valores numéricos da coluna C a partir da linha 5 de uma aba.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Aba alvo.
 * @return {number} Soma dos valores encontrados.
 */
function somarValoresColuna(sheet) {
  if (!sheet) {
    return 0;
  }

  const primeiraLinhaDados = 5;
  const ultimaLinha = sheet.getLastRow();

  if (ultimaLinha < primeiraLinhaDados) {
    return 0;
  }

  const numeroLinhas = ultimaLinha - primeiraLinhaDados + 1;
  const valores = sheet.getRange(primeiraLinhaDados, 3, numeroLinhas, 1).getValues();

  return valores.reduce(function (acumulador, linha) {
    const valor = Number(linha[0]);
    return isNaN(valor) ? acumulador : acumulador + valor;
  }, 0);
}

/**
 * Retorna os lançamentos mais recentes combinando entradas e saídas.
 *
 * @param {number} [limit=20] Quantidade máxima de lançamentos.
 * @return {Array<{tipo: string, data: Date, descricao: string, valor: number}>}
 */
function getLancamentosRecentes(limit) {
  const maxRegistros = limit || 20;
  const spreadsheet = SpreadsheetApp.getActive();
  const sheetEntradas = spreadsheet.getSheetByName(NOME_ABA_ENTRADAS);
  const sheetSaidas = spreadsheet.getSheetByName(NOME_ABA_SAIDAS);

  const lancamentos = [];
  lancamentos.push.apply(lancamentos, lerLancamentosDaAba(sheetEntradas, 'Entrada'));
  lancamentos.push.apply(lancamentos, lerLancamentosDaAba(sheetSaidas, 'Saída'));

  lancamentos.sort(function (a, b) {
    return b.data - a.data;
  });

  return lancamentos.slice(0, maxRegistros);
}

/**
 * Lê os lançamentos de uma aba específica a partir da linha 5.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Aba alvo.
 * @param {'Entrada'|'Saída'} tipo Tipo do lançamento.
 * @return {Array<{tipo: string, data: Date, descricao: string, valor: number}>}
 */
function lerLancamentosDaAba(sheet, tipo) {
  if (!sheet) {
    return [];
  }

  const primeiraLinhaDados = 5;
  const ultimaLinha = sheet.getLastRow();

  if (ultimaLinha < primeiraLinhaDados) {
    return [];
  }

  const numeroLinhas = ultimaLinha - primeiraLinhaDados + 1;
  const valores = sheet.getRange(primeiraLinhaDados, 1, numeroLinhas, 3).getValues();

  return valores.reduce(function (acumulador, linha) {
    const dataCelula = linha[0];
    const descricao = linha[1];
    const valor = Number(linha[2]);

    if (!dataCelula && !descricao && isNaN(valor)) {
      return acumulador;
    }

    const dataObj = dataCelula instanceof Date ? dataCelula : new Date(dataCelula);

    acumulador.push({
      tipo: tipo,
      data: dataObj,
      descricao: descricao || '',
      valor: isNaN(valor) ? 0 : valor,
    });

    return acumulador;
  }, []);
}

/**
 * Adiciona um novo lançamento de entrada ou saída na planilha correspondente.
 *
 * @param {{tipo: string, data: string, descricao: string, valor: number}} lancamento Lançamento informado pelo front-end.
 * @return {{success: boolean, message: string}}
 */
function addLancamento(lancamento) {
  try {
    const spreadsheet = SpreadsheetApp.getActive();
    const sheet = obterAbaPorTipo(lancamento.tipo, spreadsheet);

    if (!sheet) {
      return {
        success: false,
        message: 'Tipo de lançamento inválido.',
      };
    }

    const dataObjeto = converterStringParaData(lancamento.data);
    const proximaLinha = Math.max(sheet.getLastRow() + 1, 5);

    sheet.getRange(proximaLinha, 1).setValue(dataObjeto);
    sheet.getRange(proximaLinha, 2).setValue(lancamento.descricao);
    sheet.getRange(proximaLinha, 3).setValue(Number(lancamento.valor));

    return {
      success: true,
      message: 'Lançamento adicionado com sucesso.',
    };
  } catch (erro) {
    Logger.log('Erro ao adicionar lançamento: ' + erro);
    return {
      success: false,
      message: 'Não foi possível adicionar o lançamento. Tente novamente.',
    };
  }
}

/**
 * Retorna a aba correspondente ao tipo de lançamento.
 *
 * @param {'Entrada'|'Saída'} tipo Tipo de lançamento informado.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Planilha ativa.
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function obterAbaPorTipo(tipo, spreadsheet) {
  if (tipo === 'Entrada') {
    return spreadsheet.getSheetByName(NOME_ABA_ENTRADAS);
  }
  if (tipo === 'Saída') {
    return spreadsheet.getSheetByName(NOME_ABA_SAIDAS);
  }
  return null;
}

/**
 * Converte uma string no formato yyyy-MM-dd para um objeto Date na timezone padrão.
 *
 * @param {string} dataISO String de data em formato ISO (yyyy-MM-dd).
 * @return {Date}
 */
function converterStringParaData(dataISO) {
  if (!dataISO) {
    return new Date();
  }
  const partes = dataISO.split('-').map(Number);
  const ano = partes[0];
  const mes = partes[1] - 1;
  const dia = partes[2];
  return new Date(ano, mes, dia);
}
