/**
 * CONTROLE DE SEPARAÇÃO — Web App (Google Apps Script + Google Sheets)
 *
 * - Banco de dados: planilha do Google Sheets (abas Registros, Itens_Lote, Produtos, Colaboradores).
 * - Integração com o Amber (Aton ERP): SOMENTE CONSULTA (GET). Nada é gravado/alterado no Amber.
 * - Token do Amber: Propriedades do script (chave AMBER_TOKEN). Nunca coloque o token no código.
 *
 * Primeiros passos: veja o README.md.
 */

// =====================================================================
// CONFIGURAÇÃO
// =====================================================================
const CONFIG = {
  // true  = gera itens fictícios (não chama o Amber). Mude para false quando o endpoint estiver pronto.
  MODO_TESTE: true,

  // URL do endpoint do Amber que retorna os itens de um lote de separação.
  // Use {lote} no lugar do número do lote. >>> AJUSTE AQUI <<<
  AMBER_URL: 'https://COLOQUE_AQUI_A_URL_DO_ENDPOINT/{lote}',

  // Somente consulta. Qualquer valor diferente de GET é recusado por segurança.
  AMBER_METODO: 'GET',

  // Cabeçalhos da API (conforme a documentação: token em Authorization, Integrador = ATONAPI).
  AMBER_HEADER_AUTH: 'Authorization',
  AMBER_PREFIXO_TOKEN: '',        // ex.: 'Bearer ' se a API exigir. Vazio = token puro.
  AMBER_INTEGRADOR: 'ATONAPI',

  // Lote usado pela função testarAmber().
  TESTE_LOTE: '12345',

  // Nome da propriedade do script que guarda o token.
  PROP_TOKEN: 'AMBER_TOKEN',
  // (Opcional) Propriedade com o ID da planilha, se o script NÃO estiver vinculado a ela.
  PROP_PLANILHA: 'PLANILHA_ID'
};

const ABAS = {
  REGISTROS: {
    nome: 'Registros',
    cab: ['ID', 'Data', 'Lote', 'Separador', 'Inicio', 'Termino', 'Duracao_min', 'Total_Itens', 'Criado_em'],
    texto: [1, 2, 3, 4, 5, 6, 9]            // colunas (1-based) gravadas como texto
  },
  ITENS: {
    nome: 'Itens_Lote',
    cab: ['ID_Registro', 'Data', 'Lote', 'Separador', 'SKU', 'Descricao', 'Quantidade'],
    texto: [1, 2, 3, 4, 5]
  },
  PRODUTOS: {
    nome: 'Produtos',
    cab: ['SKU', 'Descricao', 'Dificuldade', 'Cadastrado_em'],
    texto: [1, 4]
  },
  COLABORADORES: {
    nome: 'Colaboradores',
    cab: ['Nome', 'Ativo'],
    texto: [1]
  }
};

const NOMES_DIF = { 1: 'Simples', 2: 'Médio', 3: 'Difícil' };

// =====================================================================
// WEB APP
// =====================================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Controle de Separação')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// =====================================================================
// SETUP (rode uma vez)
// =====================================================================
function setup() {
  const ss = planilha_();

  Object.keys(ABAS).forEach(function (k) {
    const def = ABAS[k];
    const sh = ss.getSheetByName(def.nome) || ss.insertSheet(def.nome);
    def.texto.forEach(function (c) {
      sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@');   // evita conversão automática
    });
    sh.getRange(1, 1, 1, def.cab.length)
      .setValues([def.cab])
      .setFontWeight('bold')
      .setBackground('#e8eaf6');
    sh.setFrozenRows(1);
  });

  // Produtos: Dificuldade só aceita 1, 2 ou 3 (número, não texto).
  const prod = ss.getSheetByName(ABAS.PRODUTOS.nome);
  prod.getRange('C2:C').setNumberFormat('0');
  prod.getRange('C2:C').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 3)
      .setAllowInvalid(false)
      .setHelpText('1 = Simples, 2 = Médio, 3 = Difícil')
      .build()
  );
  prod.setColumnWidth(2, 320);

  // Colaboradores: Ativo = SIM / NÃO.
  const col = ss.getSheetByName(ABAS.COLABORADORES.nome);
  col.getRange('B2:B').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['SIM', 'NÃO'], true)
      .setAllowInvalid(false)
      .build()
  );
  if (col.getLastRow() < 2) {
    col.getRange('A2:B2').setValues([['(escreva aqui o nome do 1º separador)', 'SIM']]);
  }

  // Remove a aba padrão vazia, se existir.
  const padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 1 && padrao.getLastRow() === 0) ss.deleteSheet(padrao);

  Logger.log('Setup concluído. Cadastre os separadores na aba "Colaboradores".');
}

// =====================================================================
// API CHAMADA PELO FRONT-END (google.script.run)
// =====================================================================
function getInitData() {
  return {
    hoje: hoje_(),
    modoTeste: CONFIG.MODO_TESTE,
    separadores: listarSeparadores_(),
    pendentes: contarPendentes_(),
    ultimos: ultimosRegistros_(12)
  };
}

function salvarLancamento(dados) {
  const d = validarLancamento_(dados);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('O sistema está ocupado com outro lançamento. Aguarde alguns segundos e tente novamente.');
  }
  try {
    if (separadorDoLote_(d.lote)) {
      throw new Error('O lote ' + d.lote + ' já foi lançado (separador: ' + separadorDoLote_(d.lote) +
        '). Um lote não pode ser lançado duas vezes.');
    }

    const itens = agruparPorSku_(buscarItensLote_(d.lote));
    const novos = cadastrarProdutosNovos_(itens);

    const id = 'R' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss') +
      '-' + Utilities.getUuid().slice(0, 4);
    const totalItens = itens.reduce(function (s, i) { return s + i.qtd; }, 0);
    const agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    const shItens = aba_(ABAS.ITENS);
    const shReg = aba_(ABAS.REGISTROS);
    const linhaIni = shItens.getLastRow() + 1;

    const linhas = itens.map(function (i) {
      return [id, d.data, d.lote, d.separador, i.sku, i.descricao, i.qtd];
    });
    shItens.getRange(linhaIni, 1, linhas.length, 7).setValues(linhas);
    try {
      shReg.getRange(shReg.getLastRow() + 1, 1, 1, 9)
        .setValues([[id, d.data, d.lote, d.separador, d.inicio, d.termino, d.duracao, totalItens, agora]]);
    } catch (e) {
      shItens.deleteRows(linhaIni, linhas.length);   // desfaz para não deixar itens órfãos
      throw e;
    }
    SpreadsheetApp.flush();

    return {
      id: id,
      lote: d.lote,
      separador: d.separador,
      duracao: d.duracao,
      totalItens: totalItens,
      skus: itens.length,
      novosProdutos: novos,
      pendentes: contarPendentes_(),
      ultimos: ultimosRegistros_(12)
    };
  } finally {
    lock.releaseLock();
  }
}

function listarProdutos() {
  const sh = aba_(ABAS.PRODUTOS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 3).getValues()
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) {
      return { sku: String(r[0]).trim(), descricao: String(r[1]), dificuldade: difValida_(r[2]) };
    });
}

function classificarProduto(sku, dificuldade) {
  const dif = Number(dificuldade);
  if ([1, 2, 3].indexOf(dif) < 0) throw new Error('Dificuldade inválida. Use 1 (Simples), 2 (Médio) ou 3 (Difícil).');
  sku = String(sku || '').trim();
  if (!sku) throw new Error('SKU não informado.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('O sistema está ocupado. Tente novamente em alguns segundos.');
  try {
    const sh = aba_(ABAS.PRODUTOS);
    const last = sh.getLastRow();
    const skus = last < 2 ? [] : sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < skus.length; i++) {
      if (String(skus[i][0]).trim() === sku) {
        sh.getRange(i + 2, 3).setValue(dif);
        SpreadsheetApp.flush();
        return { sku: sku, dificuldade: dif, pendentes: contarPendentes_() };
      }
    }
    throw new Error('Produto "' + sku + '" não encontrado na aba Produtos.');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Painel de análise.
 * @param {string} periodo 'dia' | 'semana' | 'mes'
 * @param {string} ref     data de referência yyyy-MM-dd
 */
function getPainel(periodo, ref) {
  const faixa = faixa_(periodo, ref);
  const prod = mapaProdutos_();
  const regs = lerPeriodo_(ABAS.REGISTROS, faixa);
  const itens = lerPeriodo_(ABAS.ITENS, faixa);

  const colab = {};
  function obter(nome) {
    if (!colab[nome]) colab[nome] = { nome: nome, lotes: 0, min: 0, itens: 0, s: 0, m: 0, d: 0, pontos: 0, diario: {} };
    return colab[nome];
  }

  regs.forEach(function (r) {
    const c = obter(String(r[3]));
    c.lotes += 1;
    c.min += Number(r[6]) || 0;
  });

  const pendentes = {};
  itens.forEach(function (r) {
    const c = obter(String(r[3]));
    const sku = String(r[4]).trim();
    const qtd = Number(r[6]) || 0;
    let peso = prod[sku] ? prod[sku].dif : null;
    if (!peso) {                                  // sem classificação: peso 1 + aviso
      peso = 1;
      if (!pendentes[sku]) pendentes[sku] = { sku: sku, descricao: String(r[5]), qtd: 0 };
      pendentes[sku].qtd += qtd;
    }
    c.itens += qtd;
    if (peso === 1) c.s += qtd; else if (peso === 2) c.m += qtd; else c.d += qtd;
    c.pontos += qtd * peso;
    const dia = txtData_(r[1]);
    c.diario[dia] = (c.diario[dia] || 0) + qtd * peso;
  });

  const r2 = function (n) { return Math.round(n * 100) / 100; };
  const ranking = Object.keys(colab).map(function (k) {
    const c = colab[k];
    const horas = c.min / 60;
    return {
      nome: c.nome, lotes: c.lotes, horas: r2(horas), itens: r2(c.itens),
      simples: r2(c.s), medios: r2(c.m), dificeis: r2(c.d),
      difMedia: c.itens ? r2(c.pontos / c.itens) : 0,
      pontos: r2(c.pontos),
      itensHora: horas ? r2(c.itens / horas) : 0,
      pontosHora: horas ? r2(c.pontos / horas) : 0,
      diario: c.diario
    };
  }).sort(function (a, b) { return b.pontosHora - a.pontosHora || b.pontos - a.pontos; });

  const tot = ranking.reduce(function (t, c) {
    t.lotes += c.lotes; t.horas += c.horas; t.itens += c.itens; t.pontos += c.pontos;
    return t;
  }, { lotes: 0, horas: 0, itens: 0, pontos: 0 });

  const listaPend = Object.keys(pendentes).map(function (k) { return pendentes[k]; })
    .sort(function (a, b) { return b.qtd - a.qtd; });

  return {
    ini: faixa.ini,
    fim: faixa.fim,
    dias: faixa.dias,
    equipe: {
      lotes: tot.lotes,
      horas: r2(tot.horas),
      itens: r2(tot.itens),
      pontos: r2(tot.pontos),
      pontosHora: tot.horas ? r2(tot.pontos / tot.horas) : 0,   // pontos totais / horas totais trabalhadas
      itensHora: tot.horas ? r2(tot.itens / tot.horas) : 0
    },
    ranking: ranking.map(function (c) {
      return {
        nome: c.nome, lotes: c.lotes, horas: c.horas, itens: c.itens, simples: c.simples,
        medios: c.medios, dificeis: c.dificeis, difMedia: c.difMedia, pontos: c.pontos,
        itensHora: c.itensHora, pontosHora: c.pontosHora,
        serie: faixa.dias.map(function (d) { return r2(c.diario[d] || 0); })
      };
    }),
    pendentesTotal: listaPend.length,
    pendentes: listaPend.slice(0, 30)
  };
}

// =====================================================================
// INTEGRAÇÃO COM O AMBER  (SOMENTE CONSULTA — nunca grava nada no Amber)
// =====================================================================
function buscarItensLote_(lote) {
  if (CONFIG.MODO_TESTE) return gerarItensTeste_(lote);
  const resposta = consultarAmber_(lote);
  const itens = mapearItensAmber_(resposta.json);
  if (!itens.length) {
    throw new Error('O Amber não retornou nenhum item para o lote ' + lote +
      '. Confira o número do lote (ou ajuste mapearItensAmber_ se o formato da resposta for diferente).');
  }
  return itens;
}

/** Única função que fala com o Amber. Faz apenas GET, sem corpo (payload). */
function consultarAmber_(lote) {
  if (String(CONFIG.AMBER_METODO).toUpperCase() !== 'GET') {
    throw new Error('Por segurança este sistema só faz consultas (GET) no Amber.');
  }
  const token = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_TOKEN);
  if (!token) {
    throw new Error('Token do Amber não configurado. Em Configurações do projeto > Propriedades do script, ' +
      'crie a propriedade ' + CONFIG.PROP_TOKEN + '.');
  }
  if (CONFIG.AMBER_URL.indexOf('COLOQUE_AQUI') >= 0 || CONFIG.AMBER_URL.indexOf('{lote}') < 0) {
    throw new Error('Endpoint do Amber não configurado. Ajuste CONFIG.AMBER_URL no Code.gs (use {lote}).');
  }
  const url = CONFIG.AMBER_URL.replace('{lote}', encodeURIComponent(lote));
  if (url.indexOf('https://') !== 0) throw new Error('O endpoint do Amber deve começar com https://');

  const headers = { Accept: 'application/json', Integrador: CONFIG.AMBER_INTEGRADOR };
  headers[CONFIG.AMBER_HEADER_AUTH] = CONFIG.AMBER_PREFIXO_TOKEN + token;
  const opcoes = { method: 'get', headers: headers, muteHttpExceptions: true, followRedirects: true };

  let resp;
  for (let tentativa = 1; tentativa <= 3; tentativa++) {
    try {
      resp = UrlFetchApp.fetch(url, opcoes);
    } catch (e) {
      if (tentativa === 3) throw new Error('Não foi possível conectar ao Amber: ' + e.message);
      Utilities.sleep(1500 * tentativa);
      continue;
    }
    const code = resp.getResponseCode();
    if ((code === 429 || code >= 500) && tentativa < 3) { Utilities.sleep(1500 * tentativa); continue; }
    break;
  }

  const code = resp.getResponseCode();
  const texto = resp.getContentText();
  if (code === 401 || code === 403) {
    throw new Error('O Amber recusou o acesso (HTTP ' + code + '). Verifique o token AMBER_TOKEN e o cabeçalho Integrador.');
  }
  if (code === 404) throw new Error('Lote ' + lote + ' não encontrado no Amber (HTTP 404).');
  if (code < 200 || code >= 300) {
    throw new Error('O Amber respondeu com erro HTTP ' + code + '. Tente novamente; se persistir, avise o suporte.');
  }
  let json;
  try {
    json = JSON.parse(texto);
  } catch (e) {
    throw new Error('A resposta do Amber não é um JSON válido.');
  }
  return { code: code, texto: texto, json: json };
}

/**
 * MAPEAMENTO ISOLADO: JSON do Amber -> [{sku, descricao, qtd}].
 * Se os nomes dos campos do seu Amber forem outros, ajuste as listas abaixo.
 * Rode testarAmber() para ver a resposta crua e conferir os campos.
 */
function mapearItensAmber_(json) {
  const CAMPOS_SKU = ['sku', 'codigo', 'codigoProduto', 'codigo_produto', 'cod_produto', 'codProduto', 'referencia'];
  const CAMPOS_DESC = ['descricao', 'descricaoProduto', 'descricao_produto', 'nome', 'nomeProduto', 'produto'];
  const CAMPOS_QTD = ['quantidade', 'qtd', 'qtde', 'quantidadeSeparar', 'quantidade_separar', 'qtdSeparar'];

  const lista = encontrarLista_(json, 0);
  const itens = [];
  lista.forEach(function (el) {
    if (!el || typeof el !== 'object') return;
    // Se o produto vier aninhado (ex.: { produto: { sku, descricao }, quantidade }), achata um nível.
    const plano = Object.assign({}, el.produto && typeof el.produto === 'object' ? el.produto : {}, el);
    const sku = pegar_(plano, CAMPOS_SKU);
    const qtd = paraNumero_(pegar_(plano, CAMPOS_QTD));
    if (sku === '' || sku == null || !(qtd > 0)) return;
    let desc = pegar_(plano, CAMPOS_DESC);
    if (desc && typeof desc === 'object') desc = '';
    itens.push({ sku: String(sku).trim(), descricao: String(desc || '').trim(), qtd: qtd });
  });
  return itens;
}

function encontrarLista_(obj, nivel) {
  if (Array.isArray(obj)) return obj;
  if (!obj || typeof obj !== 'object' || nivel > 3) return [];
  const preferidas = ['itens', 'items', 'produtos', 'data', 'dados', 'result', 'resultado', 'lista'];
  for (let i = 0; i < preferidas.length; i++) {
    const v = obj[preferidas[i]];
    if (v != null && typeof v === 'object') {
      const r = encontrarLista_(v, nivel + 1);
      if (r.length) return r;
    }
  }
  const chaves = Object.keys(obj);
  for (let j = 0; j < chaves.length; j++) {
    const v = obj[chaves[j]];
    if (v != null && typeof v === 'object') {
      const r = encontrarLista_(v, nivel + 1);
      if (r.length) return r;
    }
  }
  return [];
}

function pegar_(obj, chaves) {
  const mapa = {};
  Object.keys(obj).forEach(function (k) { mapa[k.toLowerCase()] = k; });
  for (let i = 0; i < chaves.length; i++) {
    const real = mapa[chaves[i].toLowerCase()];
    if (real !== undefined && obj[real] !== null && obj[real] !== '') return obj[real];
  }
  return '';
}

/** Rode no editor (Executar > testarAmber) e veja o resultado em Execuções/Registro. */
function testarAmber() {
  const lote = String(PropertiesService.getScriptProperties().getProperty('TESTE_LOTE') || CONFIG.TESTE_LOTE);
  Logger.log('MODO_TESTE = ' + CONFIG.MODO_TESTE + ' | lote = ' + lote);
  if (CONFIG.MODO_TESTE) {
    Logger.log('MODO_TESTE está ligado: a API NÃO será chamada. Itens fictícios: ' + JSON.stringify(gerarItensTeste_(lote)));
    return;
  }
  const r = consultarAmber_(lote);
  Logger.log('HTTP ' + r.code + ' — resposta crua (até 8000 caracteres):\n' + r.texto.slice(0, 8000));
  Logger.log('Itens mapeados: ' + JSON.stringify(mapearItensAmber_(r.json)));
}

function gerarItensTeste_(lote) {
  const catalogo = [
    ['SKU-1001', 'Camiseta Básica Branca M'], ['SKU-1002', 'Camiseta Básica Preta G'],
    ['SKU-1003', 'Calça Jeans Slim 42'], ['SKU-1004', 'Tênis Esportivo 40'],
    ['SKU-1005', 'Boné Aba Reta'], ['SKU-1006', 'Mochila Executiva 20L'],
    ['SKU-1007', 'Garrafa Térmica 500ml'], ['SKU-1008', 'Fone Bluetooth'],
    ['SKU-1009', 'Carregador USB-C 20W'], ['SKU-1010', 'Cabo HDMI 2m'],
    ['SKU-1011', 'Mouse Sem Fio'], ['SKU-1012', 'Teclado Compacto'],
    ['SKU-1013', 'Luminária de Mesa LED'], ['SKU-1014', 'Caixa Organizadora 30L'],
    ['SKU-1015', 'Jogo de Panelas 5 peças'], ['SKU-1016', 'Liquidificador 700W'],
    ['SKU-1017', 'Ventilador de Mesa 30cm'], ['SKU-1018', 'Cadeira Gamer'],
    ['SKU-1019', 'Colchão Casal Espuma'], ['SKU-1020', 'Mesa Dobrável 1,2m'],
    ['SKU-1021', 'Pilha AA (cartela 4)'], ['SKU-1022', 'Fita Adesiva 48mm'],
    ['SKU-1023', 'Caderno 10 matérias'], ['SKU-1024', 'Kit Canetas 12 cores']
  ];
  let seed = 0;
  String(lote).split('').forEach(function (ch) { seed = (seed * 31 + ch.charCodeAt(0)) % 2147483647; });
  const rnd = function () { seed = (seed * 48271) % 2147483647; return seed / 2147483647; };
  const n = 8 + Math.floor(rnd() * 9);
  const itens = [];
  for (let i = 0; i < n; i++) {
    const p = catalogo[Math.floor(rnd() * catalogo.length)];
    itens.push({ sku: p[0], descricao: p[1], qtd: 1 + Math.floor(rnd() * 12) });
  }
  return itens;       // pode repetir SKU — agruparPorSku_ soma
}

// =====================================================================
// REGRAS DE NEGÓCIO / AUXILIARES
// =====================================================================
function validarLancamento_(dados) {
  dados = dados || {};
  const data = String(dados.data || '').trim();
  const lote = String(dados.lote || '').trim();
  const separador = String(dados.separador || '').trim();
  const inicio = String(dados.inicio || '').trim();
  const termino = String(dados.termino || '').trim();

  if (!dataValida_(data)) throw new Error('Informe uma data válida.');
  if (!lote) throw new Error('Informe o número do lote.');
  if (!/^[A-Za-z0-9._\-]+$/.test(lote)) throw new Error('O número do lote só pode ter letras, números, ponto, hífen ou sublinhado.');
  if (!separador) throw new Error('Selecione o separador.');
  if (listarSeparadores_().indexOf(separador) < 0) {
    throw new Error('O separador "' + separador + '" não está ativo na aba Colaboradores.');
  }
  const mi = minutos_(inicio);
  const mt = minutos_(termino);
  if (isNaN(mi)) throw new Error('Horário de início inválido.');
  if (isNaN(mt)) throw new Error('Horário de término inválido.');
  if (mi === mt) throw new Error('O horário de término é igual ao de início. Confira os horários.');

  let duracao = mt - mi;
  if (duracao < 0) duracao += 1440;           // virou a meia-noite

  return {
    data: data, lote: lote, separador: separador,
    inicio: pad2_(Math.floor(mi / 60)) + ':' + pad2_(mi % 60),
    termino: pad2_(Math.floor(mt / 60)) + ':' + pad2_(mt % 60),
    duracao: duracao
  };
}

function agruparPorSku_(itens) {
  const mapa = {};
  const ordem = [];
  itens.forEach(function (i) {
    const sku = String(i.sku).trim();
    if (!mapa[sku]) { mapa[sku] = { sku: sku, descricao: i.descricao || '', qtd: 0 }; ordem.push(sku); }
    mapa[sku].qtd += Number(i.qtd) || 0;
    if (!mapa[sku].descricao && i.descricao) mapa[sku].descricao = i.descricao;
  });
  return ordem.map(function (s) { return mapa[s]; });
}

/** Cadastra na aba Produtos os SKUs inéditos, com dificuldade em branco (pendente). */
function cadastrarProdutosNovos_(itens) {
  const sh = aba_(ABAS.PRODUTOS);
  const last = sh.getLastRow();
  const existentes = {};
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { existentes[String(r[0]).trim()] = true; });
  }
  const agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const novos = itens.filter(function (i) { return !existentes[i.sku]; })
    .map(function (i) { return [i.sku, i.descricao, '', agora]; });
  if (novos.length) sh.getRange(Math.max(last, 1) + 1, 1, novos.length, 4).setValues(novos);
  return novos.length;
}

function mapaProdutos_() {
  const sh = aba_(ABAS.PRODUTOS);
  const last = sh.getLastRow();
  const mapa = {};
  if (last < 2) return mapa;
  sh.getRange(2, 1, last - 1, 3).getValues().forEach(function (r) {
    const sku = String(r[0]).trim();
    if (sku) mapa[sku] = { descricao: String(r[1]), dif: difValida_(r[2]) };
  });
  return mapa;
}

function contarPendentes_() {
  const sh = aba_(ABAS.PRODUTOS);
  const last = sh.getLastRow();
  if (last < 2) return 0;
  return sh.getRange(2, 1, last - 1, 3).getValues().filter(function (r) {
    return String(r[0]).trim() !== '' && !difValida_(r[2]);
  }).length;
}

function difValida_(v) {
  const n = Number(v);
  return (v !== '' && v !== null && n >= 1 && n <= 3 && Math.floor(n) === n) ? n : null;
}

function listarSeparadores_() {
  const sh = aba_(ABAS.COLABORADORES);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 2).getValues()
    .filter(function (r) {
      const ativo = String(r[1]).trim().toUpperCase().replace('Ã', 'A');
      return String(r[0]).trim() !== '' && ativo !== 'NAO';
    })
    .map(function (r) { return String(r[0]).trim(); });
}

/** Retorna o separador do lote se ele já foi lançado, senão ''. */
function separadorDoLote_(lote) {
  const sh = aba_(ABAS.REGISTROS);
  const last = sh.getLastRow();
  if (last < 2) return '';
  const v = sh.getRange(2, 3, last - 1, 2).getValues();      // colunas Lote e Separador
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === lote) return String(v[i][1]);
  }
  return '';
}

function ultimosRegistros_(n) {
  const sh = aba_(ABAS.REGISTROS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const ini = Math.max(2, last - n + 1);
  return sh.getRange(ini, 1, last - ini + 1, 9).getValues().reverse().map(function (r) {
    return {
      data: txtData_(r[1]), lote: String(r[2]), separador: String(r[3]),
      inicio: txtHora_(r[4]), termino: txtHora_(r[5]),
      duracao: Number(r[6]) || 0, itens: Number(r[7]) || 0
    };
  });
}

/** Lê só o bloco de linhas cujo campo Data (coluna B) cai no período. */
function lerPeriodo_(def, faixa) {
  const sh = aba_(def);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const datas = sh.getRange(2, 2, last - 1, 1).getValues();
  let a = -1, b = -1;
  for (let i = 0; i < datas.length; i++) {
    const t = txtData_(datas[i][0]);
    if (t >= faixa.ini && t <= faixa.fim) { if (a < 0) a = i; b = i; }
  }
  if (a < 0) return [];
  return sh.getRange(2 + a, 1, b - a + 1, def.cab.length).getValues().filter(function (r) {
    const t = txtData_(r[1]);
    return t >= faixa.ini && t <= faixa.fim;
  });
}

function faixa_(periodo, ref) {
  if (!dataValida_(ref)) throw new Error('Data de referência inválida.');
  const p = parseData_(ref);
  let ini, fim;
  if (periodo === 'dia') {
    ini = p; fim = p;
  } else if (periodo === 'semana') {                   // segunda a domingo
    const dow = (p.getUTCDay() + 6) % 7;
    ini = somarDias_(p, -dow); fim = somarDias_(ini, 6);
  } else if (periodo === 'mes') {
    ini = new Date(Date.UTC(p.getUTCFullYear(), p.getUTCMonth(), 1));
    fim = new Date(Date.UTC(p.getUTCFullYear(), p.getUTCMonth() + 1, 0));
  } else {
    throw new Error('Período inválido.');
  }
  const dias = [];
  for (let d = ini; d <= fim; d = somarDias_(d, 1)) dias.push(fmtData_(d));
  return { ini: fmtData_(ini), fim: fmtData_(fim), dias: dias };
}

// ---- datas e horas (tudo como texto yyyy-MM-dd / HH:mm) ----
function hoje_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function parseData_(s) { const p = s.split('-'); return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])); }
function fmtData_(d) { return d.getUTCFullYear() + '-' + pad2_(d.getUTCMonth() + 1) + '-' + pad2_(d.getUTCDate()); }
function somarDias_(d, n) { return new Date(d.getTime() + n * 86400000); }
function dataValida_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s || ''))) return false;
  return fmtData_(parseData_(s)) === s;
}
function minutos_(hhmm) {
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(hhmm || ''));
  if (!m) return NaN;
  const h = +m[1], mi = +m[2];
  return (h > 23 || mi > 59) ? NaN : h * 60 + mi;
}
function txtData_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}
function txtHora_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  return String(v == null ? '' : v).trim();
}
function paraNumero_(v) {
  if (typeof v === 'number') return v;
  let s = String(v == null ? '' : v).trim();
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');   // "1.234,5" -> "1234.5"
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

// ---- planilha ----
function planilha_() {
  const id = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_PLANILHA);
  if (id) return SpreadsheetApp.openById(id);
  const ativa = SpreadsheetApp.getActiveSpreadsheet();
  if (!ativa) {
    throw new Error('Planilha não encontrada. Use um script vinculado à planilha ou crie a propriedade ' +
      CONFIG.PROP_PLANILHA + ' com o ID dela.');
  }
  return ativa;
}

function aba_(def) {
  const sh = planilha_().getSheetByName(def.nome);
  if (!sh) throw new Error('A aba "' + def.nome + '" não existe. Execute a função setup() no editor do Apps Script.');
  return sh;
}
