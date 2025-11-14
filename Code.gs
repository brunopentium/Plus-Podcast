/**
 * Plus Podcast - Finanças Apps Script backend.
 */

// Nome das abas utilizadas na planilha.
const NOMES_POSSIVEIS_SAIDAS = ['Saídas', 'Saidas'];
const NOMES_POSSIVEIS_ENTRADAS = ['Entradas'];
const NOMES_MESES = [
  'Janeiro',
  'Fevereiro',
  'Março',
  'Abril',
  'Maio',
  'Junho',
  'Julho',
  'Agosto',
  'Setembro',
  'Outubro',
  'Novembro',
  'Dezembro',
];
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
 * Retorna a primeira aba disponível dentre os nomes informados.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Planilha ativa.
 * @param {string[]} nomes Lista de nomes possíveis da aba.
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function obterAbaPorNomes(spreadsheet, nomes) {
  if (!spreadsheet || !Array.isArray(nomes)) {
    return null;
  }

  for (var i = 0; i < nomes.length; i++) {
    var sheet = spreadsheet.getSheetByName(nomes[i]);
    if (sheet) {
      return sheet;
    }
  }

  return null;
}

/**
 * Recupera os dados do resumo financeiro para o dashboard.
 *
 * @param {string} [mesReferencia] Mês no formato yyyy-MM para filtrar os dados.
 * @return {{saldoAtual: number, totalEntradas: number, totalSaidas: number}}
 */
function getDashboardData(mesReferencia) {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheetEntradas = obterAbaPorNomes(spreadsheet, NOMES_POSSIVEIS_ENTRADAS);
  const sheetSaidas = obterAbaPorNomes(spreadsheet, NOMES_POSSIVEIS_SAIDAS);

  const lancamentosEntradas = lerLancamentosDaAba(sheetEntradas, 'Entrada');
  const lancamentosSaidas = lerLancamentosDaAba(sheetSaidas, 'Saída');

  const referencia = obterReferenciaMes(mesReferencia);
  const inicioMes = new Date(referencia.ano, referencia.mes, 1);
  const fimMes = obterFimDoMes(referencia.ano, referencia.mes);

  const totalEntradas = somarLancamentosNoPeriodo(lancamentosEntradas, inicioMes, fimMes);
  const totalSaidas = somarLancamentosNoPeriodo(lancamentosSaidas, inicioMes, fimMes);
  const saldoAtual =
    somarLancamentosAteData(lancamentosEntradas, fimMes) -
    somarLancamentosAteData(lancamentosSaidas, fimMes);

  return {
    saldoAtual: saldoAtual,
    totalEntradas: totalEntradas,
    totalSaidas: totalSaidas,
  };
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
  const sheetEntradas = obterAbaPorNomes(spreadsheet, NOMES_POSSIVEIS_ENTRADAS);
  const sheetSaidas = obterAbaPorNomes(spreadsheet, NOMES_POSSIVEIS_SAIDAS);

  const lancamentos = [];
  lancamentos.push.apply(lancamentos, lerLancamentosDaAba(sheetEntradas, 'Entrada'));
  lancamentos.push.apply(lancamentos, lerLancamentosDaAba(sheetSaidas, 'Saída'));

  lancamentos.sort(function (a, b) {
    return b.data - a.data;
  });

  return lancamentos.slice(0, maxRegistros);
}

/**
 * Retorna a lista de meses disponíveis entre entradas e saídas.
 *
 * @return {Array<{valor: string, rotulo: string}>}
 */
function getMesesDisponiveis() {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheetEntradas = obterAbaPorNomes(spreadsheet, NOMES_POSSIVEIS_ENTRADAS);
  const sheetSaidas = obterAbaPorNomes(spreadsheet, NOMES_POSSIVEIS_SAIDAS);

  const meses = {};

  lerLancamentosDaAba(sheetEntradas, 'Entrada').forEach(function (item) {
    registrarMes(item, meses);
  });
  lerLancamentosDaAba(sheetSaidas, 'Saída').forEach(function (item) {
    registrarMes(item, meses);
  });

  if (Object.keys(meses).length === 0) {
    const hoje = new Date();
    const chave = criarChaveMes(hoje.getFullYear(), hoje.getMonth());
    meses[chave] = { ano: hoje.getFullYear(), mes: hoje.getMonth() };
  }

  return Object.keys(meses)
    .map(function (chave) {
      const info = meses[chave];
      return {
        valor: chave,
        rotulo: formatarRotuloMes(info.ano, info.mes),
      };
    })
    .sort(function (a, b) {
      return a.valor < b.valor ? 1 : -1;
    });
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
    return obterAbaPorNomes(spreadsheet, NOMES_POSSIVEIS_ENTRADAS);
  }
  if (tipo === 'Saída') {
    return obterAbaPorNomes(spreadsheet, NOMES_POSSIVEIS_SAIDAS);
  }
  return null;
}

/**
 * Retorna o mês de referência informado ou o mês atual caso inválido.
 *
 * @param {string} mesReferencia Mês no formato yyyy-MM.
 * @return {{ano: number, mes: number}}
 */
function obterReferenciaMes(mesReferencia) {
  if (typeof mesReferencia === 'string') {
    const partes = mesReferencia.split('-');
    if (partes.length === 2) {
      const ano = Number(partes[0]);
      const mes = Number(partes[1]) - 1;
      if (!isNaN(ano) && !isNaN(mes) && mes >= 0 && mes < 12) {
        return { ano: ano, mes: mes };
      }
    }
  }

  const hoje = new Date();
  return {
    ano: hoje.getFullYear(),
    mes: hoje.getMonth(),
  };
}

/**
 * Obtém o último instante do mês informado.
 *
 * @param {number} ano Ano de referência.
 * @param {number} mes Índice do mês (0-11).
 * @return {Date}
 */
function obterFimDoMes(ano, mes) {
  return new Date(ano, mes + 1, 0, 23, 59, 59, 999);
}

/**
 * Soma os valores de lançamentos dentro de um intervalo.
 *
 * @param {Array<{data: Date, valor: number}>} lancamentos Lista de lançamentos.
 * @param {Date} inicio Data inicial inclusiva.
 * @param {Date} fim Data final inclusiva.
 * @return {number}
 */
function somarLancamentosNoPeriodo(lancamentos, inicio, fim) {
  if (!Array.isArray(lancamentos) || !inicio || !fim) {
    return 0;
  }

  return lancamentos.reduce(function (total, item) {
    const data = item && item.data instanceof Date ? item.data : null;
    if (!data || isNaN(data)) {
      return total;
    }
    if (data >= inicio && data <= fim) {
      const valor = Number(item.valor);
      return total + (isNaN(valor) ? 0 : valor);
    }
    return total;
  }, 0);
}

/**
 * Soma os valores de lançamentos até uma data limite.
 *
 * @param {Array<{data: Date, valor: number}>} lancamentos Lista de lançamentos.
 * @param {Date} limite Data limite inclusiva.
 * @return {number}
 */
function somarLancamentosAteData(lancamentos, limite) {
  if (!Array.isArray(lancamentos) || !limite) {
    return 0;
  }

  return lancamentos.reduce(function (total, item) {
    const data = item && item.data instanceof Date ? item.data : null;
    if (!data || isNaN(data)) {
      return total;
    }
    if (data <= limite) {
      const valor = Number(item.valor);
      return total + (isNaN(valor) ? 0 : valor);
    }
    return total;
  }, 0);
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

/**
 * Registra um mês no conjunto informado.
 *
 * @param {{data: Date}} item Lançamento contendo a data.
 * @param {Object<string, {ano: number, mes: number}>} meses Mapa de meses já registrados.
 */
function registrarMes(item, meses) {
  const data = item && item.data instanceof Date ? item.data : null;
  if (!data || isNaN(data)) {
    return;
  }

  const chave = criarChaveMes(data.getFullYear(), data.getMonth());
  if (!meses[chave]) {
    meses[chave] = {
      ano: data.getFullYear(),
      mes: data.getMonth(),
    };
  }
}

/**
 * Cria uma chave no formato yyyy-MM para o mês informado.
 *
 * @param {number} ano Ano de referência.
 * @param {number} mes Índice do mês (0-11).
 * @return {string}
 */
function criarChaveMes(ano, mes) {
  const mesFormatado = ('0' + (mes + 1)).slice(-2);
  return ano + '-' + mesFormatado;
}

/**
 * Formata um rótulo legível para o mês informado.
 *
 * @param {number} ano Ano de referência.
 * @param {number} mes Índice do mês (0-11).
 * @return {string}
 */
function formatarRotuloMes(ano, mes) {
  const nomeMes = NOMES_MESES[mes] || '';
  return nomeMes ? nomeMes + ' de ' + ano : ano + '-' + ('0' + (mes + 1)).slice(-2);
}
