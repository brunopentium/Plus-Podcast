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

const PAPEL_COMPLETO = 'completo';
const PAPEL_SOMENTE_LEITURA = 'leitura';

const USUARIOS_CONFIG = {
  bruno: {
    senha: 'Cesar177*',
    nome: 'Bruno',
    papel: PAPEL_COMPLETO,
  },
  alexandre: {
    senha: 'plus123',
    nome: 'Alexandre',
    papel: PAPEL_SOMENTE_LEITURA,
  },
};

const PREFIXO_SESSAO = 'sessao.';

let CACHE_FUSO_HORARIO_PLANILHA = null;

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
 * Realiza a autenticação de um usuário e gera um token de sessão persistente.
 *
 * @param {{usuario: string, senha: string}} credenciais
 * @return {{success: boolean, message?: string, token?: string, usuario?: {id: string, nome: string, papel: string}}}
 */
function login(credenciais) {
  if (!credenciais || typeof credenciais.usuario !== 'string' || typeof credenciais.senha !== 'string') {
    return { success: false, message: 'Informe usuário e senha.' };
  }

  const usuarioId = credenciais.usuario.trim().toLowerCase();
  const senhaInformada = credenciais.senha;
  const configuracao = USUARIOS_CONFIG[usuarioId];

  if (!configuracao || configuracao.senha !== senhaInformada) {
    return { success: false, message: 'Usuário ou senha inválidos.' };
  }

  const token = Utilities.getUuid();
  registrarSessaoAtiva(usuarioId, token);

  return {
    success: true,
    token: token,
    usuario: {
      id: usuarioId,
      nome: configuracao.nome,
      papel: configuracao.papel,
    },
  };
}

/**
 * Valida uma sessão existente.
 *
 * @param {{usuario: string, token: string}} sessao
 * @return {{success: boolean, message?: string, token?: string, usuario?: {id: string, nome: string, papel: string}}}
 */
function validarSessao(sessao) {
  const validacao = validarSessaoInterno(sessao);
  if (!validacao.valida) {
    return { success: false, message: 'Sessão inválida. Faça login novamente.' };
  }

  return {
    success: true,
    token: validacao.token,
    usuario: {
      id: validacao.id,
      nome: validacao.nome,
      papel: validacao.papel,
    },
  };
}

/**
 * Encerra uma sessão ativa.
 *
 * @param {{usuario: string, token: string}} sessao
 * @return {{success: boolean}}
 */
function logout(sessao) {
  const validacao = validarSessaoInterno(sessao);
  if (validacao.valida) {
    removerSessaoAtiva(validacao.id);
  }
  return { success: true };
}

/**
 * Confere se a sessão informada é válida.
 *
 * @param {{usuario: string, token: string}} sessao
 * @return {{valida: boolean, id?: string, nome?: string, papel?: string, token?: string}}
 */
function validarSessaoInterno(sessao) {
  if (!sessao || typeof sessao.usuario !== 'string' || typeof sessao.token !== 'string') {
    return { valida: false };
  }

  const usuarioId = sessao.usuario.trim().toLowerCase();
  const configuracao = USUARIOS_CONFIG[usuarioId];

  if (!configuracao) {
    return { valida: false };
  }

  const registro = obterRegistroSessao(usuarioId);
  if (!registro || registro.token !== sessao.token) {
    return { valida: false };
  }

  return {
    valida: true,
    id: usuarioId,
    nome: configuracao.nome,
    papel: configuracao.papel,
    token: registro.token,
  };
}

/**
 * Prepara o contexto de execução validando a sessão informada.
 *
 * @param {Object} argumento Parâmetros recebidos do front-end.
 * @param {boolean} exigeEscrita Indica se a operação exige permissão de escrita.
 * @return {{parametros: Object, sessao: {id: string, nome: string, papel: string, token: string}}}
 */
function prepararContexto(argumento, exigeEscrita) {
  const parametros = argumento && typeof argumento === 'object' ? Object.assign({}, argumento) : {};
  const sessaoInformada = parametros.sessao || null;

  if (sessaoInformada) {
    delete parametros.sessao;
  }

  const validacao = validarSessaoInterno(sessaoInformada);

  if (!validacao.valida) {
    throw new Error('Sessão inválida.');
  }

  if (exigeEscrita && validacao.papel === PAPEL_SOMENTE_LEITURA) {
    throw new Error('Permissão insuficiente para realizar esta operação.');
  }

  return {
    parametros: parametros,
    sessao: validacao,
  };
}

/**
 * Persiste um token de sessão para o usuário informado.
 *
 * @param {string} usuarioId
 * @param {string} token
 */
function registrarSessaoAtiva(usuarioId, token) {
  const propriedade = PREFIXO_SESSAO + usuarioId;
  PropertiesService.getScriptProperties().setProperty(
    propriedade,
    JSON.stringify({ token: token, atualizadoEm: new Date().toISOString() })
  );
}

/**
 * Recupera os dados persistidos da sessão do usuário.
 *
 * @param {string} usuarioId
 * @return {{token: string, atualizadoEm: string}|null}
 */
function obterRegistroSessao(usuarioId) {
  if (!usuarioId) {
    return null;
  }

  const propriedade = PropertiesService.getScriptProperties().getProperty(PREFIXO_SESSAO + usuarioId);
  if (!propriedade) {
    return null;
  }

  try {
    return JSON.parse(propriedade);
  } catch (erro) {
    Logger.log('Erro ao ler sessão: ' + erro);
    return null;
  }
}

/**
 * Remove a sessão armazenada do usuário.
 *
 * @param {string} usuarioId
 */
function removerSessaoAtiva(usuarioId) {
  if (!usuarioId) {
    return;
  }
  PropertiesService.getScriptProperties().deleteProperty(PREFIXO_SESSAO + usuarioId);
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

  var nomesNormalizados = nomes
    .map(function (nome) {
      return normalizarNomeAbaComparacao(nome);
    })
    .filter(function (texto) {
      return !!texto;
    });

  if (nomesNormalizados.length === 0) {
    return null;
  }

  var abas = spreadsheet.getSheets();
  for (var j = 0; j < abas.length; j++) {
    var abaAtual = abas[j];
    var nomeNormalizado = normalizarNomeAbaComparacao(abaAtual.getName());
    if (nomesNormalizados.indexOf(nomeNormalizado) !== -1) {
      return abaAtual;
    }
  }

  return null;
}

/**
 * Retorna a aba correspondente aos nomes informados ou cria uma nova com o nome padrão.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string[]} nomes
 * @param {string} nomePadrao
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function obterAbaPorNomesOuCriar(spreadsheet, nomes, nomePadrao) {
  var sheet = obterAbaPorNomes(spreadsheet, nomes);
  if (sheet) {
    configurarEstruturaAba(sheet, true);
    return sheet;
  }
  return criarAbaPadrao(spreadsheet, nomePadrao);
}

/**
 * Cria uma aba padrão com cabeçalho e formatação básica.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string} nomePadrao
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function criarAbaPadrao(spreadsheet, nomePadrao) {
  if (!spreadsheet || !nomePadrao) {
    return null;
  }

  var sheetExistente = spreadsheet.getSheetByName(nomePadrao);
  if (sheetExistente) {
    configurarEstruturaAba(sheetExistente, true);
    return sheetExistente;
  }

  var sheet = spreadsheet.insertSheet(nomePadrao);
  configurarEstruturaAba(sheet, false);
  return sheet;
}

/**
 * Garante que a aba possua o cabeçalho e formatação esperados.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {boolean} preservarCabecalho Existindo cabeçalho, evita sobrescrever valores.
 */
function configurarEstruturaAba(sheet, preservarCabecalho) {
  if (!sheet) {
    return;
  }

  var intervaloCabecalho = sheet.getRange(4, 1, 1, 3);
  var valoresAtuais = intervaloCabecalho.getValues();

  if (!preservarCabecalho || !valoresAtuais || valoresAtuais.length === 0) {
    intervaloCabecalho.setValues([['Data', 'Descrição', 'Valor']]);
  } else {
    var linha = valoresAtuais[0] || [];
    if (!linha[0] || !linha[1] || !linha[2]) {
      intervaloCabecalho.setValues([['Data', 'Descrição', 'Valor']]);
    }
  }

  sheet.getRange('A:A').setNumberFormat('dd/MM/yyyy');
  sheet.getRange('C:C').setNumberFormat('R$ #,##0.00');
  sheet.setFrozenRows(4);
}

/**
 * Recupera os dados do resumo financeiro para o dashboard.
 *
 * @param {string} [mesReferencia] Mês no formato yyyy-MM para filtrar os dados.
 * @return {{saldoAtual: number, totalEntradas: number, totalSaidas: number}}
 */
function getDashboardData(requisicao) {
  const contexto = prepararContexto(requisicao, false);
  const mesReferencia = contexto.parametros.mesReferencia;

  const spreadsheet = SpreadsheetApp.getActive();
  const sheetEntradas = obterAbaPorNomesOuCriar(
    spreadsheet,
    NOMES_POSSIVEIS_ENTRADAS,
    'Entradas'
  );
  const sheetSaidas = obterAbaPorNomesOuCriar(spreadsheet, NOMES_POSSIVEIS_SAIDAS, 'Saídas');

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
function getLancamentosRecentes(requisicao) {
  const contexto = prepararContexto(requisicao, false);
  const parametros = contexto.parametros || {};

  const maxRegistrosInformado = Number(parametros.limite || parametros.limit);
  const maxRegistros = isFinite(maxRegistrosInformado) && maxRegistrosInformado > 0
    ? Math.floor(maxRegistrosInformado)
    : 200;

  const intervalo = determinarIntervaloLancamentos(parametros);
  const filtroPersonalizado = Boolean(
    (typeof parametros.dataInicio === 'string' && parametros.dataInicio) ||
      (typeof parametros.dataFim === 'string' && parametros.dataFim) ||
      (typeof parametros.mesReferencia === 'string' && parametros.mesReferencia)
  );

  const spreadsheet = SpreadsheetApp.getActive();
  const sheetEntradas = obterAbaPorNomesOuCriar(
    spreadsheet,
    NOMES_POSSIVEIS_ENTRADAS,
    'Entradas'
  );
  const sheetSaidas = obterAbaPorNomesOuCriar(spreadsheet, NOMES_POSSIVEIS_SAIDAS, 'Saídas');

  const lancamentos = [];
  lancamentos.push.apply(lancamentos, lerLancamentosDaAba(sheetEntradas, 'Entrada'));
  lancamentos.push.apply(lancamentos, lerLancamentosDaAba(sheetSaidas, 'Saída'));

  lancamentos.sort(function (a, b) {
    return b.data - a.data;
  });

  const filtrados = lancamentos.filter(function (item) {
    if (!(item.data instanceof Date) || isNaN(item.data)) {
      return false;
    }
    if (!intervalo.inicio && !intervalo.fim) {
      return true;
    }
    if (intervalo.inicio && item.data < intervalo.inicio) {
      return false;
    }
    if (intervalo.fim && item.data > intervalo.fim) {
      return false;
    }
    return true;
  });

  if (filtrados.length === 0 && !filtroPersonalizado) {
    return lancamentos.slice(0, maxRegistros);
  }

  return filtrados.slice(0, maxRegistros);
}

function listarLancamentos(requisicao) {
  const contexto = prepararContexto(requisicao, false);
  const parametros = contexto.parametros || {};

  const maxRegistrosInformado = Number(parametros.limite || parametros.limit);
  const limite = isFinite(maxRegistrosInformado) && maxRegistrosInformado > 0
    ? Math.floor(maxRegistrosInformado)
    : 200;

  const intervalo = determinarIntervaloLancamentos(parametros);
  const filtroPersonalizado = Boolean(
    (typeof parametros.dataInicio === 'string' && parametros.dataInicio) ||
      (typeof parametros.dataFim === 'string' && parametros.dataFim) ||
      (typeof parametros.mesReferencia === 'string' && parametros.mesReferencia)
  );

  const spreadsheet = SpreadsheetApp.getActive();
  const sheetEntradas = obterAbaPorNomesOuCriar(
    spreadsheet,
    NOMES_POSSIVEIS_ENTRADAS,
    'Entradas'
  );
  const sheetSaidas = obterAbaPorNomesOuCriar(spreadsheet, NOMES_POSSIVEIS_SAIDAS, 'Saídas');

  const todosLancamentos = [];
  todosLancamentos.push.apply(
    todosLancamentos,
    lerLancamentosDaAba(sheetEntradas, 'Entrada')
  );
  todosLancamentos.push.apply(
    todosLancamentos,
    lerLancamentosDaAba(sheetSaidas, 'Saída')
  );

  todosLancamentos.sort(function (a, b) {
    return b.data - a.data;
  });

  const filtrados = todosLancamentos.filter(function (item) {
    if (!(item.data instanceof Date) || isNaN(item.data)) {
      return false;
    }
    if (!intervalo.inicio && !intervalo.fim) {
      return true;
    }
    if (intervalo.inicio && item.data < intervalo.inicio) {
      return false;
    }
    if (intervalo.fim && item.data > intervalo.fim) {
      return false;
    }
    return true;
  });

  const base = filtrados.length === 0 && !filtroPersonalizado ? todosLancamentos : filtrados;

  return base.slice(0, limite).map(function (item) {
    return {
      tipo: item.tipo,
      data: item.data,
      descricao: item.descricao,
      valor: item.valor,
    };
  });
}

/**
 * Retorna a lista de meses disponíveis entre entradas e saídas.
 *
 * @return {Array<{valor: string, rotulo: string}>}
 */
function getMesesDisponiveis(requisicao) {
  prepararContexto(requisicao, false);

  const spreadsheet = SpreadsheetApp.getActive();
  const sheetEntradas = obterAbaPorNomesOuCriar(
    spreadsheet,
    NOMES_POSSIVEIS_ENTRADAS,
    'Entradas'
  );
  const sheetSaidas = obterAbaPorNomesOuCriar(spreadsheet, NOMES_POSSIVEIS_SAIDAS, 'Saídas');

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
 * Lê os lançamentos de uma aba específica respeitando o cabeçalho padrão.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Aba alvo.
 * @param {'Entrada'|'Saída'} tipo Tipo do lançamento.
 * @return {Array<{tipo: string, data: Date, descricao: string, valor: number, linha: number}>}
 */
function lerLancamentosDaAba(sheet, tipo) {
  if (!sheet) {
    return [];
  }

  const ultimaLinha = sheet.getLastRow();
  if (ultimaLinha === 0) {
    return [];
  }

  const ultimaColuna = sheet.getLastColumn();
  const numeroColunas = Math.max(3, ultimaColuna);
  const dados = sheet.getRange(1, 1, ultimaLinha, numeroColunas).getValues();
  if (!dados || dados.length === 0) {
    return [];
  }

  const registros = [];

  let indiceInicial = 0;
  let mapeamentoColunas = { tipo: null, data: 0, descricao: 1, valor: 2 };
  for (let i = 0; i < dados.length; i++) {
    const infoCabecalho = obterMapeamentoColunasLancamentos(dados[i]);
    if (infoCabecalho) {
      indiceInicial = i + 1;
      mapeamentoColunas = infoCabecalho;
      break;
    }
  }

  for (let i = indiceInicial; i < dados.length; i++) {
    const linhaAtual = dados[i];
    const dataCelula = obterValorDaLinha(linhaAtual, mapeamentoColunas.data, linhaAtual[0]);
    const descricaoCelula = obterValorDaLinha(
      linhaAtual,
      mapeamentoColunas.descricao,
      linhaAtual[1]
    );
    const valorCelula = obterValorDaLinha(linhaAtual, mapeamentoColunas.valor, linhaAtual[2]);

    if (!possuiConteudoLancamento(dataCelula, descricaoCelula, valorCelula)) {
      continue;
    }

    const dataObj = converterValorParaDataPlanilha(dataCelula);
    if (!(dataObj instanceof Date) || isNaN(dataObj)) {
      continue;
    }

    const tipoCelula = obterValorDaLinha(linhaAtual, mapeamentoColunas.tipo, null);
    const tipoNormalizado = limparTextoCelula(tipoCelula);

    registros.push({
      tipo: tipoNormalizado || tipo,
      data: dataObj,
      descricao: limparTextoCelula(descricaoCelula),
      valor: normalizarValorMonetario(valorCelula),
      linha: i + 1,
    });
  }

  return registros;
}

function possuiConteudoLancamento(dataCelula, descricaoCelula, valorCelula) {
  if (dataCelula instanceof Date && !isNaN(dataCelula)) {
    return true;
  }
  if (typeof dataCelula === 'string' && dataCelula.trim()) {
    return true;
  }
  if (limparTextoCelula(descricaoCelula)) {
    return true;
  }
  if (typeof valorCelula === 'number' && !isNaN(valorCelula)) {
    return true;
  }
  if (typeof valorCelula === 'string' && valorCelula.trim()) {
    return true;
  }
  return false;
}

function ehCabecalhoLancamentos(linha) {
  return Boolean(obterMapeamentoColunasLancamentos(linha));
}

function obterMapeamentoColunasLancamentos(linha) {
  if (!Array.isArray(linha) || linha.length < 3) {
    return null;
  }

  const estrategias = [
    mapearCabecalhoVersao1,
    mapearCabecalhoVersao2,
    mapearCabecalhoVersao3,
  ];

  for (let i = 0; i < estrategias.length; i++) {
    const resultado = estrategias[i](linha);
    if (
      resultado &&
      indiceCabecalhoValido(resultado.data) &&
      indiceCabecalhoValido(resultado.descricao) &&
      indiceCabecalhoValido(resultado.valor)
    ) {
      return {
        tipo: indiceCabecalhoValido(resultado.tipo) ? resultado.tipo : null,
        data: resultado.data,
        descricao: resultado.descricao,
        valor: resultado.valor,
      };
    }
  }

  return null;
}

function mapearCabecalhoVersao1(linha) {
  const sinonimos = {
    tipo: criarSinonimosCabecalho(['tipo', 'categoria', 'classificacao', 'classe', 'natureza']),
    data: criarSinonimosCabecalho([
      'data',
      'data do lancamento',
      'data lancamento',
      'data pagamento',
      'data do pagamento',
      'data emissao',
      'data de emissao',
      'dia',
    ]),
    descricao: criarSinonimosCabecalho([
      'descricao',
      'descricao do lancamento',
      'historico',
      'detalhes',
      'detalhe',
      'referencia',
      'observacao',
      'observacoes',
    ]),
    valor: criarSinonimosCabecalho([
      'valor',
      'valor total',
      'valor bruto',
      'valor liquido',
      'valor r$',
      'valor rs',
      'montante',
      'quantia',
      'total',
      'valor pago',
      'amount',
      'preco',
      'preco final',
    ]),
  };

  const indices = { tipo: null, data: null, descricao: null, valor: null };

  for (let i = 0; i < linha.length; i++) {
    const textoNormalizado = normalizarTextoParaComparacao(linha[i]);
    if (!textoNormalizado) {
      continue;
    }

    const chave = padronizarChaveComparacao(textoNormalizado);

    if (indices.tipo === null && sinonimos.tipo.indexOf(chave) !== -1) {
      indices.tipo = i;
      continue;
    }

    if (indices.data === null && sinonimos.data.indexOf(chave) !== -1) {
      indices.data = i;
      continue;
    }

    if (indices.descricao === null && sinonimos.descricao.indexOf(chave) !== -1) {
      indices.descricao = i;
      continue;
    }

    if (indices.valor === null && sinonimos.valor.indexOf(chave) !== -1) {
      indices.valor = i;
    }
  }

  if (indices.data === null || indices.descricao === null || indices.valor === null) {
    return null;
  }

  return {
    tipo: indices.tipo,
    data: indices.data,
    descricao: indices.descricao,
    valor: indices.valor,
  };
}

function mapearCabecalhoVersao2(linha) {
  const tokens = {
    tipo: ['tipo', 'categoria', 'classific', 'naturez', 'grupo'],
    data: ['data', 'dt', 'competencia', 'pagamento', 'emissao', 'vencimento', 'vencto'],
    descricao: ['descricao', 'descri', 'histor', 'detal', 'referenc', 'observac', 'memo'],
    valor: ['valor', 'quantia', 'montante', 'total', 'bruto', 'liquido', 'amount', 'preco', 'custo'],
  };

  const indices = { tipo: null, data: null, descricao: null, valor: null };

  for (let i = 0; i < linha.length; i++) {
    const textoNormalizado = normalizarTextoParaComparacao(linha[i]);
    if (!textoNormalizado) {
      continue;
    }

    const chave = padronizarChaveComparacao(textoNormalizado);

    if (indices.tipo === null && contemTokenCabecalho(chave, tokens.tipo)) {
      indices.tipo = i;
    }

    if (indices.data === null && contemTokenCabecalho(chave, tokens.data)) {
      indices.data = i;
      continue;
    }

    if (indices.descricao === null && contemTokenCabecalho(chave, tokens.descricao)) {
      indices.descricao = i;
      continue;
    }

    if (indices.valor === null && contemTokenCabecalho(chave, tokens.valor)) {
      indices.valor = i;
    }
  }

  if (indices.data === null || indices.descricao === null || indices.valor === null) {
    return null;
  }

  return {
    tipo: indices.tipo,
    data: indices.data,
    descricao: indices.descricao,
    valor: indices.valor,
  };
}

function mapearCabecalhoVersao3(linha) {
  const candidatos = [];

  for (let i = 0; i < linha.length; i++) {
    const textoNormalizado = normalizarTextoParaComparacao(linha[i]);
    if (!textoNormalizado) {
      continue;
    }

    candidatos.push({
      indice: i,
      chave: padronizarChaveComparacao(textoNormalizado),
    });
  }

  if (candidatos.length === 0) {
    return null;
  }

  const configuracao = {
    data: {
      fortes: criarSinonimosCabecalho([
        'data',
        'datadolancamento',
        'datapagamento',
        'dataemissao',
        'datadeemissao',
      ]),
      medios: criarSinonimosCabecalho([
        'competencia',
        'vencimento',
        'vencto',
        'mesref',
        'periodo',
      ]),
      fracos: criarSinonimosCabecalho(['dia', 'mes']),
    },
    descricao: {
      fortes: criarSinonimosCabecalho([
        'descricao',
        'descricaodolancamento',
        'historico',
        'detalhes',
      ]),
      medios: criarSinonimosCabecalho([
        'observacao',
        'memo',
        'comentario',
        'referencia',
      ]),
      fracos: criarSinonimosCabecalho(['resumo', 'texto', 'informacao']),
    },
    valor: {
      fortes: criarSinonimosCabecalho([
        'valor',
        'valortotal',
        'valorbruto',
        'valorliquido',
        'valorpago',
      ]),
      medios: criarSinonimosCabecalho([
        'quantia',
        'montante',
        'total',
        'saldo',
      ]),
      fracos: criarSinonimosCabecalho(['amount', 'preco', 'custo', 'despesa']),
    },
    tipo: {
      fortes: criarSinonimosCabecalho(['tipo', 'categoria', 'natureza']),
      medios: criarSinonimosCabecalho(['classificacao', 'grupo', 'entrada', 'saida']),
      fracos: criarSinonimosCabecalho(['credito', 'debito']),
    },
  };

  const disponiveis = candidatos.slice();

  const indiceData = selecionarIndiceCabecalho(disponiveis, configuracao.data);
  if (indiceData === -1) {
    return null;
  }
  removerCandidatoCabecalho(disponiveis, indiceData);

  const indiceDescricao = selecionarIndiceCabecalho(disponiveis, configuracao.descricao);
  if (indiceDescricao === -1) {
    return null;
  }
  removerCandidatoCabecalho(disponiveis, indiceDescricao);

  const indiceValor = selecionarIndiceCabecalho(disponiveis, configuracao.valor);
  if (indiceValor === -1) {
    return null;
  }

  const indiceTipo = selecionarIndiceCabecalho(candidatos, configuracao.tipo);

  return {
    tipo: indiceTipo === -1 ? null : indiceTipo,
    data: indiceData,
    descricao: indiceDescricao,
    valor: indiceValor,
  };
}

function indiceCabecalhoValido(valor) {
  return typeof valor === 'number' && valor >= 0;
}

function padronizarChaveComparacao(valor) {
  if (!valor) {
    return '';
  }
  return valor.replace(/[^a-z0-9]/g, '');
}

function normalizarSinonimoCabecalho(valor) {
  return padronizarChaveComparacao(normalizarTextoParaComparacao(valor));
}

function criarSinonimosCabecalho(lista) {
  return lista.map(function (item) {
    return normalizarSinonimoCabecalho(item);
  });
}

function contemTokenCabecalho(chave, tokens) {
  if (!chave) {
    return false;
  }

  for (let i = 0; i < tokens.length; i++) {
    if (chave.indexOf(tokens[i]) !== -1) {
      return true;
    }
  }

  return false;
}

function selecionarIndiceCabecalho(candidatos, configuracao) {
  let melhorIndice = -1;
  let melhorPontuacao = 0;

  candidatos.forEach(function (item) {
    const pontuacao = pontuarChaveCabecalho(item.chave, configuracao);
    if (pontuacao > melhorPontuacao) {
      melhorPontuacao = pontuacao;
      melhorIndice = item.indice;
    }
  });

  return melhorPontuacao > 0 ? melhorIndice : -1;
}

function pontuarChaveCabecalho(chave, configuracao) {
  if (!chave) {
    return 0;
  }

  const grupos = [
    { lista: configuracao.fortes || [], valorIgual: 6, valorContem: 4 },
    { lista: configuracao.medios || [], valorIgual: 3, valorContem: 2 },
    { lista: configuracao.fracos || [], valorIgual: 2, valorContem: 1 },
  ];

  let pontuacao = 0;

  grupos.forEach(function (grupo) {
    grupo.lista.forEach(function (token) {
      if (!token) {
        return;
      }
      if (chave === token) {
        pontuacao += grupo.valorIgual;
      } else if (chave.indexOf(token) !== -1) {
        pontuacao += grupo.valorContem;
      }
    });
  });

  return pontuacao;
}

function removerCandidatoCabecalho(candidatos, indice) {
  for (let i = candidatos.length - 1; i >= 0; i--) {
    if (candidatos[i].indice === indice) {
      candidatos.splice(i, 1);
      break;
    }
  }
}

function obterValorDaLinha(linha, indice, padrao) {
  if (!linha) {
    return padrao;
  }
  if (typeof indice === 'number' && indice >= 0 && indice < linha.length) {
    return linha[indice];
  }
  return padrao;
}

function normalizarTextoParaComparacao(valor) {
  if (typeof valor !== 'string') {
    return '';
  }
  return valor
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function normalizarNomeAbaComparacao(nome) {
  if (!nome) {
    return '';
  }

  return normalizarTextoParaComparacao(nome).replace(/[\s_-]+/g, '');
}

function limparTextoCelula(valor) {
  if (valor === null || typeof valor === 'undefined') {
    return '';
  }
  if (typeof valor === 'string') {
    return valor.trim();
  }
  return String(valor).trim();
}

function normalizarValorMonetario(valor) {
  if (typeof valor === 'number' && !isNaN(valor)) {
    return valor;
  }

  if (typeof valor === 'string') {
    const textoNormalizado = valor
      .replace(/\s/g, '')
      .replace(/R\$/gi, '')
      .replace(/\./g, '')
      .replace(',', '.');

    const numero = Number(textoNormalizado);
    if (!isNaN(numero)) {
      return numero;
    }
  }

  return 0;
}

/**
 * Adiciona um novo lançamento de entrada ou saída na planilha correspondente.
 *
 * @param {{tipo: string, data: string, descricao: string, valor: number}} lancamento Lançamento informado pelo front-end.
 * @return {{success: boolean, message: string}}
 */
function addLancamento(requisicao) {
  try {
    const contexto = prepararContexto(requisicao, true);
    const lancamento = contexto.parametros.lancamento;

    if (!lancamento) {
      return {
        success: false,
        message: 'Dados do lançamento não informados.',
      };
    }

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
    return obterAbaPorNomesOuCriar(spreadsheet, NOMES_POSSIVEIS_ENTRADAS, 'Entradas');
  }
  if (tipo === 'Saída') {
    return obterAbaPorNomesOuCriar(spreadsheet, NOMES_POSSIVEIS_SAIDAS, 'Saídas');
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
 * Determina o intervalo de datas utilizado para filtrar lançamentos recentes.
 *
 * @param {{dataInicio?: string, dataFim?: string, mesReferencia?: string}} parametros
 * @return {{inicio: Date|null, fim: Date|null}}
 */
function determinarIntervaloLancamentos(parametros) {
  const dados = parametros || {};
  let inicio = null;
  let fim = null;

  if (typeof dados.dataInicio === 'string' && dados.dataInicio) {
    inicio = converterStringParaData(dados.dataInicio);
  }

  if (typeof dados.dataFim === 'string' && dados.dataFim) {
    fim = converterStringParaData(dados.dataFim);
  }

  if (!inicio && !fim) {
    if (typeof dados.mesReferencia === 'string' && dados.mesReferencia) {
      const referencia = obterReferenciaMes(dados.mesReferencia);
      inicio = normalizarInicioDoDia(new Date(referencia.ano, referencia.mes, 1));
      fim = obterFimDoMes(referencia.ano, referencia.mes);
    } else {
      const fimPadrao = normalizarFimDoDia(new Date());
      const inicioPadrao = new Date(fimPadrao.getTime());
      inicioPadrao.setDate(inicioPadrao.getDate() - 29);
      inicio = normalizarInicioDoDia(inicioPadrao);
      fim = fimPadrao;
    }
  } else {
    if (inicio) {
      inicio = normalizarInicioDoDia(inicio);
    }
    if (fim) {
      fim = normalizarFimDoDia(fim);
    } else if (inicio) {
      fim = normalizarFimDoDia(new Date());
    }
  }

  return { inicio: inicio, fim: fim };
}

/**
 * Ajusta uma data para o início do dia (00:00:00.000).
 *
 * @param {Date} data Data alvo.
 * @return {Date}
 */
function normalizarInicioDoDia(data) {
  const clone = new Date(data.getTime());
  clone.setHours(0, 0, 0, 0);
  return clone;
}

/**
 * Ajusta uma data para o fim do dia (23:59:59.999).
 *
 * @param {Date} data Data alvo.
 * @return {Date}
 */
function normalizarFimDoDia(data) {
  const clone = new Date(data.getTime());
  clone.setHours(23, 59, 59, 999);
  return clone;
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
 * Converte um valor retornado pela planilha para um objeto Date válido.
 * Aceita objetos Date, strings no formato brasileiro (dd/MM/yyyy) ou ISO, e números serializados.
 *
 * @param {Date|string|number} valor Valor da célula que representa uma data.
 * @return {Date|null}
 */
function converterValorParaDataPlanilha(valor) {
  if (valor instanceof Date && !isNaN(valor)) {
    return valor;
  }

  if (typeof valor === 'string') {
    const texto = valor.trim();
    if (!texto) {
      return null;
    }

    const padraoBR = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/;
    const correspondenciaBR = texto.match(padraoBR);

    if (correspondenciaBR) {
      let ano = Number(correspondenciaBR[3]);
      const mes = Number(correspondenciaBR[2]) - 1;
      const dia = Number(correspondenciaBR[1]);

      if (ano < 100) {
        ano += ano >= 70 ? 1900 : 2000;
      }

      const dataBR = new Date(ano, mes, dia);
      if (!isNaN(dataBR)) {
        return dataBR;
      }
    }

    const dataGenerica = new Date(texto);
    if (!isNaN(dataGenerica)) {
      return dataGenerica;
    }
  }

  if (typeof valor === 'number' && !isNaN(valor)) {
    const epocaExcel = new Date(Math.round((valor - 25569) * 86400 * 1000));
    if (!isNaN(epocaExcel)) {
      const fusoHorario = obterFusoHorarioPlanilha();
      if (fusoHorario) {
        const dataNormalizada = Utilities.formatDate(epocaExcel, fusoHorario, 'yyyy-MM-dd');
        return converterStringParaData(dataNormalizada);
      }
      return new Date(epocaExcel.getUTCFullYear(), epocaExcel.getUTCMonth(), epocaExcel.getUTCDate());
    }
  }

  return null;
}

/**
 * Recupera e memoriza o fuso horário configurado na planilha.
 *
 * @return {string}
 */
function obterFusoHorarioPlanilha() {
  if (CACHE_FUSO_HORARIO_PLANILHA === null) {
    try {
      const planilhaAtiva = SpreadsheetApp.getActive();
      CACHE_FUSO_HORARIO_PLANILHA =
        (planilhaAtiva && planilhaAtiva.getSpreadsheetTimeZone()) || Session.getScriptTimeZone() || 'GMT';
    } catch (erro) {
      CACHE_FUSO_HORARIO_PLANILHA = Session.getScriptTimeZone() || 'GMT';
    }
  }
  return CACHE_FUSO_HORARIO_PLANILHA;
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
