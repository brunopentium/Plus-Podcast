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
const USUARIOS_SISTEMA = {
  bruno: { senha: 'Cesar177*', papel: 'admin', nome: 'Bruno' },
  alexandre: { senha: 'plus123', papel: 'viewer', nome: 'Alexandre' },
};
const CHAVE_SESSOES = 'PLUS_PODCAST_SESSOES_ATIVAS';
const MENSAGEM_SESSAO_INVALIDA = 'Sessão inválida. Faça login novamente.';

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
 * Realiza a autenticação do usuário a partir das credenciais informadas.
 *
 * @param {{usuario: string, senha: string}} credenciais
 * @return {{success: boolean, message?: string, token?: string, usuario?: {identificador: string, nome: string, papel: string}}}
 */
function autenticarUsuario(credenciais) {
  if (!credenciais || !credenciais.usuario || !credenciais.senha) {
    return {
      success: false,
      message: 'Informe usuário e senha.',
    };
  }

  const identificador = String(credenciais.usuario).trim().toLowerCase();
  const registro = USUARIOS_SISTEMA[identificador];

  if (!registro || registro.senha !== credenciais.senha) {
    return {
      success: false,
      message: 'Usuário ou senha inválidos.',
    };
  }

  const sessoes = obterSessoesAtivas();
  let token = obterTokenPorUsuario(sessoes, identificador);

  if (!token) {
    token = gerarTokenSessao();
  }

  sessoes[token] = {
    usuario: identificador,
    papel: registro.papel,
    nome: registro.nome,
  };
  salvarSessoesAtivas(sessoes);

  return {
    success: true,
    token: token,
    usuario: {
      identificador: identificador,
      nome: registro.nome,
      papel: registro.papel,
    },
  };
}

/**
 * Valida se uma sessão permanece ativa e retorna os dados do usuário.
 *
 * @param {string} token
 * @return {{valid: boolean, token?: string, usuario?: {identificador: string, nome: string, papel: string}}}
 */
function validarSessao(token) {
  if (!token) {
    return { valid: false };
  }

  const sessoes = obterSessoesAtivas();
  const sessao = sessoes[token];

  if (!sessao) {
    return { valid: false };
  }

  const configuracaoUsuario = USUARIOS_SISTEMA[sessao.usuario];
  if (!configuracaoUsuario) {
    delete sessoes[token];
    salvarSessoesAtivas(sessoes);
    return { valid: false };
  }

  let atualizado = false;

  if (sessao.papel !== configuracaoUsuario.papel) {
    sessao.papel = configuracaoUsuario.papel;
    atualizado = true;
  }

  if (sessao.nome !== configuracaoUsuario.nome) {
    sessao.nome = configuracaoUsuario.nome;
    atualizado = true;
  }

  if (atualizado) {
    salvarSessoesAtivas(sessoes);
  }

  return {
    valid: true,
    token: token,
    usuario: {
      identificador: sessao.usuario,
      nome: sessao.nome,
      papel: sessao.papel,
    },
  };
}

/**
 * Remove uma sessão ativa, caso exista.
 *
 * @param {string} token
 * @return {{success: boolean}}
 */
function encerrarSessao(token) {
  if (!token) {
    return { success: true };
  }

  const sessoes = obterSessoesAtivas();
  if (sessoes[token]) {
    delete sessoes[token];
    salvarSessoesAtivas(sessoes);
  }

  return { success: true };
}

/**
 * Garante que o token informado pertença a uma sessão ativa.
 *
 * @param {string} token
 * @return {{usuario: string, papel: string, nome: string}}
 */
function exigirSessaoAtiva(token) {
  if (!token) {
    throw new Error(MENSAGEM_SESSAO_INVALIDA);
  }

  const sessoes = obterSessoesAtivas();
  const sessao = sessoes[token];

  if (!sessao) {
    throw new Error(MENSAGEM_SESSAO_INVALIDA);
  }

  const configuracaoUsuario = USUARIOS_SISTEMA[sessao.usuario];
  if (!configuracaoUsuario) {
    delete sessoes[token];
    salvarSessoesAtivas(sessoes);
    throw new Error(MENSAGEM_SESSAO_INVALIDA);
  }

  let atualizado = false;

  if (sessao.papel !== configuracaoUsuario.papel) {
    sessao.papel = configuracaoUsuario.papel;
    atualizado = true;
  }

  if (sessao.nome !== configuracaoUsuario.nome) {
    sessao.nome = configuracaoUsuario.nome;
    atualizado = true;
  }

  if (atualizado) {
    salvarSessoesAtivas(sessoes);
  }

  return sessao;
}

/**
 * Recupera o objeto de sessões armazenado nas propriedades do script.
 *
 * @return {Object<string, {usuario: string, papel: string, nome: string}>}
 */
function obterSessoesAtivas() {
  const propriedades = PropertiesService.getScriptProperties();
  const dado = propriedades.getProperty(CHAVE_SESSOES);

  if (!dado) {
    return {};
  }

  try {
    const recuperado = JSON.parse(dado);
    return recuperado && typeof recuperado === 'object' ? recuperado : {};
  } catch (erro) {
    Logger.log('Falha ao interpretar sessões ativas: ' + erro);
    return {};
  }
}

/**
 * Persiste as sessões ativas nas propriedades do script.
 *
 * @param {Object<string, {usuario: string, papel: string, nome: string}>} sessoes
 */
function salvarSessoesAtivas(sessoes) {
  const propriedades = PropertiesService.getScriptProperties();
  propriedades.setProperty(CHAVE_SESSOES, JSON.stringify(sessoes || {}));
}

/**
 * Retorna o token associado ao usuário, se existir.
 *
 * @param {Object<string, {usuario: string}>} sessoes
 * @param {string} usuario
 * @return {string|null}
 */
function obterTokenPorUsuario(sessoes, usuario) {
  for (var token in sessoes) {
    if (Object.prototype.hasOwnProperty.call(sessoes, token)) {
      var sessao = sessoes[token];
      if (sessao && sessao.usuario === usuario) {
        return token;
      }
    }
  }
  return null;
}

/**
 * Gera um token único para identificação da sessão.
 *
 * @return {string}
 */
function gerarTokenSessao() {
  return Utilities.getUuid();
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
 * @param {string} token Token de sessão válido.
 * @return {{saldoAtual: number, totalEntradas: number, totalSaidas: number}}
 */
function getDashboardData(mesReferencia, token) {
  exigirSessaoAtiva(token);

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
 * @param {string} token Token de sessão válido.
 * @return {Array<{tipo: string, data: Date, descricao: string, valor: number}>}
 */
function getLancamentosRecentes(limit, token) {
  exigirSessaoAtiva(token);

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
 * @param {string} token Token de sessão válido.
 * @return {Array<{valor: string, rotulo: string}>}
 */
function getMesesDisponiveis(token) {
  exigirSessaoAtiva(token);

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
 * @param {string} token Token de sessão válido.
 * @return {{success: boolean, message: string}}
 */
function addLancamento(lancamento, token) {
  const sessao = exigirSessaoAtiva(token);
  if (!sessao || sessao.papel !== 'admin') {
    return {
      success: false,
      message: 'Usuário não possui permissão para adicionar lançamentos.',
    };
  }

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
