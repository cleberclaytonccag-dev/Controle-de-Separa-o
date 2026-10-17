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
  // true  = gera itens fictícios (não chama o Amber). Mude para false quando for usar o Amber de verdade.
  MODO_TESTE: true,

  // Endereço base da API (documentação: "API Aton ERP (Pública)"). Não precisa mudar.
  AMBER_BASE_URL: 'https://api.ambarxcall.com.br/AtonSNIsapi.dll/atonerp',

  // Como achar os pedidos de um lote: busca pedidos por período e posição e filtra por "lotesep".
  AMBER_TIPO_DATA: 'data_pedido',                 // data_pedido | data_aprovacao | data_limite_postagem
  AMBER_DIAS_ANTES: 7,                            // busca pedidos da data do lançamento até N dias antes (pedidos antigos também entram em lote)
  AMBER_POSICOES: ['PENDENTE', 'ABERTA', 'FECHADO', 'EMITIDO'],   // posições dos pedidos a varrer
  AMBER_PAGINA: 50,                               // máximo permitido pela API
  AMBER_CACHE_SEGUNDOS: 1800,                     // quanto tempo reaproveitar a varredura (30 min)
  AMBER_PARALELO: 2,                              // consultas simultâneas ao Amber (o servidor parece processar uma por vez; menos = mais estável)
  AMBER_OFFSET_MODO: 'auto',                      // auto | pagina | registro (como a API interpreta "offset")

  // Cabeçalhos da API (token em Authorization; Integrador = ATONAPI para integração direta).
  AMBER_HEADER_AUTH: 'Authorization',
  AMBER_PREFIXO_TOKEN: '',        // ex.: 'Bearer ' se a API exigir. Vazio = token puro.
  AMBER_INTEGRADOR: 'ATONAPI',

  // Lote usado pela função testarAmber().
  TESTE_LOTE: '12345',

  // Pedido usado pela função testarPedido() (diagnóstico de um pedido que não apareceu no lote).
  TESTE_PEDIDO: '',

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

  // Consulta ao Amber (só leitura) ANTES de travar a planilha: pode demorar e não deve bloquear os outros.
  verificarLoteNaoLancado_(d.lote);
  const itens = completarDescricoes_(agruparPorSku_(buscarItensLote_(d.lote, d.data)));

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('O sistema está ocupado com outro lançamento. Aguarde alguns segundos e tente novamente.');
  }
  try {
    verificarLoteNaoLancado_(d.lote);       // confere de novo, já dentro do bloqueio (evita duplicidade)
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
// INTEGRAÇÃO COM O AMBER / ATON ERP  (SOMENTE CONSULTA — nunca altera nada no Amber)
//
// A API do Aton usa POST também para consultas. Por isso a segurança aqui é uma LISTA FIXA
// de endpoints de consulta (ENDPOINTS_PERMITIDOS). Endpoints que gravam (PUT: atualizar pedidos,
// estoques, baixas etc.) não existem neste código e são recusados por requisicaoAmber_().
//
// Como achamos os itens de um lote: a API não tem "itens do lote". Cada pedido de venda traz o
// campo "lotesep" (lote de separação) e a lista "itens". Então buscamos os pedidos do período
// (por posição, paginado) e filtramos os que têm lotesep == número do lote.
// =====================================================================
const ENDPOINTS_PERMITIDOS = [
  '/pedidosvenda/consulta',     // consulta de pedidos de venda (itens + lotesep)
  '/produtos/listagemid'        // consulta de produtos por ID (descrição)
];

function buscarItensLote_(lote, data) {
  if (CONFIG.MODO_TESTE) return gerarItensTeste_(lote);
  if (!/^\d+$/.test(lote) || Number(lote) < 1) {
    throw new Error('No Amber o lote de separação é um número maior que zero. Confira o número digitado.');
  }
  // Uma varredura traz TODOS os lotes do período; guardamos num cache curto para os próximos lançamentos.
  const chave = 'amber_idx_' + data + '_' + CONFIG.AMBER_DIAS_ANTES + '_' + CONFIG.AMBER_POSICOES.join('');
  let indice = lerIndiceCache_(chave);
  let busca = null;
  if (!indice || !indice[lote]) {
    busca = buscarPedidosAmber_(data);
    indice = montarIndice_(busca.pedidos);
    gravarIndiceCache_(chave, indice);
  }
  if (!indice[lote]) {
    throw new Error('Nenhum pedido do lote ' + lote + ' foi encontrado no Amber (pedidos de ' + busca.dataIni +
      ' a ' + busca.dataFim + ', posições: ' + CONFIG.AMBER_POSICOES.join(', ') + '; ' + busca.pedidos.length +
      ' pedidos analisados). Confira o número do lote. Se ele for mais antigo, aumente CONFIG.AMBER_DIAS_ANTES ' +
      'ou inclua a posição correta em CONFIG.AMBER_POSICOES.');
  }
  return indice[lote].map(function (x) { return { sku: x[0], descricao: '', qtd: x[1], codigoAton: x[2] }; });
}

/** { lote: [[sku, qtd, codigoAton], ...] } para todos os lotes (lotesep > 0) dos pedidos. */
function montarIndice_(pedidos) {
  const grupos = {};
  pedidos.forEach(function (p) {
    const l = Number(p.lotesep);
    if (l > 0) (grupos[l] = grupos[l] || []).push(p);
  });
  const indice = {};
  Object.keys(grupos).forEach(function (l) {
    indice[l] = agruparPorSku_(mapearItensAmber_(grupos[l], l)).map(function (i) { return [i.sku, i.qtd, i.codigoAton]; });
  });
  return indice;
}

function lerIndiceCache_(chave) {
  try {
    const cache = CacheService.getScriptCache();
    const n = Number(cache.get(chave + ':n'));
    if (!n) return null;
    const partes = [];
    for (let i = 0; i < n; i++) {
      const parte = cache.get(chave + ':' + i);
      if (parte == null) return null;
      partes.push(parte);
    }
    return JSON.parse(partes.join(''));
  } catch (e) {
    return null;
  }
}

function gravarIndiceCache_(chave, indice) {
  try {
    const txt = JSON.stringify(indice);
    const TAM = 90000, n = Math.ceil(txt.length / TAM);
    if (n > 9) return;                                   // grande demais para o cache: segue sem ele
    const itens = {};
    for (let i = 0; i < n; i++) itens[chave + ':' + i] = txt.slice(i * TAM, (i + 1) * TAM);
    itens[chave + ':n'] = String(n);
    CacheService.getScriptCache().putAll(itens, CONFIG.AMBER_CACHE_SEGUNDOS);
  } catch (e) {
    Logger.log('Cache não gravado: ' + e.message);
  }
}

/**
 * MAPEAMENTO ISOLADO: pedidos do Amber -> [{sku, descricao, qtd, codigoAton}] só do lote pedido.
 * sku = codigo_interno do item. A descrição não vem no pedido; é completada por completarDescricoes_().
 */
function mapearItensAmber_(pedidos, lote) {
  const alvo = Number(lote);
  const itens = [];
  pedidos.forEach(function (p) {
    if (Number(p.lotesep) !== alvo) return;
    (p.itens || []).forEach(function (i) {
      const sku = String(i.codigo_interno == null ? '' : i.codigo_interno).trim();
      const qtd = Number(i.quantidade);
      if (!sku || !(qtd > 0)) return;
      itens.push({ sku: sku, descricao: '', qtd: qtd, codigoAton: i.codigo_aton });
    });
  });
  return itens;
}

/** Busca os pedidos do período em todas as posições configuradas (paginado, em paralelo). */
function buscarPedidosAmber_(data) {
  const inicio = Date.now();
  const fim = parseData_(data);
  const ini = somarDias_(fim, -CONFIG.AMBER_DIAS_ANTES);
  const brIni = brData_(fmtData_(ini)), brFim = brData_(fmtData_(fim));
  const limite = CONFIG.AMBER_PAGINA;
  const mapa = {};
  const stats = { porPosicao: {}, modoOffset: CONFIG.AMBER_OFFSET_MODO, chamadas: 0, amostra: null };

  CONFIG.AMBER_POSICOES.forEach(function (pos) {
    const corpo = function (offset) {
      return { tipo_data: CONFIG.AMBER_TIPO_DATA, data_inicial: brIni, data_final: brFim, posicao: pos, offset: offset, limit: limite };
    };
    const pedir = function (offsets) {
      stats.chamadas += offsets.length;
      return chamarAmber_(offsets.map(function (o) { return requisicaoAmber_('/pedidosvenda/consulta', corpo(o)); }))
        .map(resultadoPedidos_);
    };
    const juntar = function (r) {
      r.pedidos.forEach(function (p) { mapa[p.pedido] = p; });
      if (!stats.amostra && r.pedidos.length) stats.amostra = r.pedidos[0];
    };

    const r1 = pedir([1])[0];
    juntar(r1);
    stats.porPosicao[pos] = r1.total;
    if (r1.total <= limite) return;

    // offset = nº da página OU nº do registro? ("auto" testa a página 2 e vê se repete pedidos da 1)
    let modo = CONFIG.AMBER_OFFSET_MODO;
    let r2 = null;
    if (modo === 'auto') {
      r2 = pedir([2])[0];
      juntar(r2);
      const ids1 = {}; r1.pedidos.forEach(function (p) { ids1[p.pedido] = true; });
      modo = r2.pedidos.some(function (p) { return ids1[p.pedido]; }) ? 'registro' : 'pagina';
      stats.modoOffset = 'auto→' + modo;
    }
    const paginas = Math.ceil(r1.total / limite);
    const offsets = [];
    for (let k = 1; k < paginas; k++) {
      const off = modo === 'registro' ? 1 + k * limite : 1 + k;
      if (r2 && modo === 'pagina' && off === 2) continue;      // a página 2 já veio na sondagem
      offsets.push(off);
    }
    for (let i = 0; i < offsets.length; i += CONFIG.AMBER_PARALELO) {
      if (Date.now() - inicio > 270000) {
        throw new Error('A consulta ao Amber demorou demais. Reduza CONFIG.AMBER_DIAS_ANTES ou as posições em CONFIG.AMBER_POSICOES.');
      }
      pedir(offsets.slice(i, i + CONFIG.AMBER_PARALELO)).forEach(juntar);
    }
  });

  stats.segundos = Math.round((Date.now() - inicio) / 1000);
  return { pedidos: Object.keys(mapa).map(function (k) { return mapa[k]; }), dataIni: brIni, dataFim: brFim, stats: stats };
}

function resultadoPedidos_(json) {
  const status = String(json && json.status || '').toLowerCase();
  const res = (json && json.resultado) || {};
  if (status !== 'sucesso') {
    if (/n[ãa]o.*encontrad|nenhum/i.test(String(json && json.mensagem || ''))) return { total: 0, pedidos: [] };
    throw new Error('O Amber respondeu: ' + String(json && json.mensagem || 'erro desconhecido'));
  }
  const pedidos = res.pedidos || [];
  return { total: Number(res.total) || pedidos.length, pedidos: pedidos };
}

/** Preenche a descrição dos itens: primeiro pela aba Produtos; o que faltar, consulta no Amber (50 por vez). */
function completarDescricoes_(itens) {
  if (CONFIG.MODO_TESTE) return itens;
  const prod = mapaProdutos_();
  itens.forEach(function (i) {
    if (!i.descricao && prod[i.sku] && prod[i.sku].descricao) i.descricao = prod[i.sku].descricao;
  });
  const faltam = itens.filter(function (i) { return !i.descricao && i.codigoAton; });
  const ids = [];
  faltam.forEach(function (i) { if (ids.indexOf(i.codigoAton) < 0) ids.push(i.codigoAton); });
  if (!ids.length) return itens;
  try {
    const lotes = [];
    for (let k = 0; k < ids.length; k += 50) lotes.push(ids.slice(k, k + 50));
    const resp = chamarAmber_(lotes.map(function (l) { return requisicaoAmber_('/produtos/listagemid', { produtos: l }); }));
    const desc = {};
    resp.forEach(function (r) {
      ((r.resultado && r.resultado.produtos_encontrados) || []).forEach(function (p) { desc[p.codigo_aton] = p.descricao; });
    });
    faltam.forEach(function (i) { if (desc[i.codigoAton]) i.descricao = String(desc[i.codigoAton]); });
  } catch (e) {
    Logger.log('Não foi possível obter descrições no Amber (o lançamento segue sem elas): ' + e.message);
  }
  return itens;
}

/** Monta uma requisição de CONSULTA. Recusa qualquer endpoint fora da lista permitida. */
function requisicaoAmber_(caminho, corpo) {
  if (ENDPOINTS_PERMITIDOS.indexOf(caminho) < 0) {
    throw new Error('Por segurança este sistema só chama endpoints de consulta do Amber. Endpoint recusado: ' + caminho);
  }
  const token = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_TOKEN);
  if (!token) {
    throw new Error('Token do Amber não configurado. Em Configurações do projeto > Propriedades do script, ' +
      'crie a propriedade ' + CONFIG.PROP_TOKEN + '.');
  }
  if (CONFIG.AMBER_BASE_URL.indexOf('https://') !== 0) throw new Error('CONFIG.AMBER_BASE_URL deve começar com https://');
  const headers = { Integrador: CONFIG.AMBER_INTEGRADOR };
  headers[CONFIG.AMBER_HEADER_AUTH] = CONFIG.AMBER_PREFIXO_TOKEN + token;
  return {
    url: CONFIG.AMBER_BASE_URL + caminho,
    method: 'post',                       // a API do Aton consulta via POST (ver nota acima)
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify(corpo),
    muteHttpExceptions: true,
    followRedirects: true
  };
}

/** Executa as requisições (em paralelo) com até 3 tentativas para erros temporários. Retorna os JSONs. */
function chamarAmber_(requisicoes) {
  const saida = new Array(requisicoes.length);
  let pendentes = requisicoes.map(function (_, i) { return i; });
  for (let tentativa = 1; tentativa <= 3 && pendentes.length; tentativa++) {
    let respostas;
    try {
      respostas = UrlFetchApp.fetchAll(pendentes.map(function (i) { return requisicoes[i]; }));
    } catch (e) {
      if (tentativa === 3) throw new Error('Não foi possível conectar ao Amber: ' + e.message);
      Utilities.sleep(1500 * tentativa);
      continue;
    }
    const outra = [];
    respostas.forEach(function (resp, k) {
      const i = pendentes[k];
      const code = resp.getResponseCode();
      // 429/5xx e também 400: o Aton às vezes devolve "Solicitação Inválida" de forma intermitente; tenta de novo.
      if ((code === 400 || code === 429 || code >= 500) && tentativa < 3) { outra.push(i); return; }
      if (code === 401 || code === 403) {
        throw new Error('O Amber recusou o acesso (HTTP ' + code + '). Verifique o token AMBER_TOKEN e o cabeçalho Integrador.');
      }
      if (code < 200 || code >= 300) {
        throw new Error('O Amber recusou a consulta (HTTP ' + code + '): ' + motivoAmber_(resp) + detalheConsulta_(requisicoes[i]) +
          (code >= 500 ? ' Tente novamente; se persistir, avise o suporte.' : ''));
      }
      try {
        saida[i] = JSON.parse(resp.getContentText());
      } catch (e) {
        throw new Error('A resposta do Amber não é um JSON válido.');
      }
    });
    pendentes = outra;
    if (pendentes.length) Utilities.sleep(2000 * tentativa);
  }
  return saida;
}

/** Texto de erro devolvido pelo Amber (campo "mensagem" do JSON, ou o início da resposta). */
function motivoAmber_(resp) {
  const t = String(resp.getContentText() || '');
  try {
    const j = JSON.parse(t);
    if (j && j.mensagem) return String(j.mensagem).slice(0, 300);
  } catch (e) { /* não é JSON */ }
  return t.replace(/\s+/g, ' ').slice(0, 300) || '(sem mensagem)';
}

/** Parâmetros da consulta (sem token) para ajudar a diagnosticar erros. */
function detalheConsulta_(req) {
  try {
    const b = JSON.parse(req.payload);
    delete b.produtos;
    return ' [consulta: ' + JSON.stringify(b) + ']';
  } catch (e) { return ''; }
}

/** Rode no editor (Executar > testarAmber) e veja o Registro de execução. Só faz consultas. */
function testarAmber() {
  const lote = String(PropertiesService.getScriptProperties().getProperty('TESTE_LOTE') || CONFIG.TESTE_LOTE);
  Logger.log('MODO_TESTE = ' + CONFIG.MODO_TESTE + ' | lote = ' + lote);
  if (CONFIG.MODO_TESTE) {
    Logger.log('MODO_TESTE está ligado: a API NÃO será chamada. Itens fictícios: ' + JSON.stringify(gerarItensTeste_(lote)));
    return;
  }
  const busca = buscarPedidosAmber_(hoje_());
  const s = busca.stats;
  Logger.log('Período ' + busca.dataIni + ' a ' + busca.dataFim + ' | ' + busca.pedidos.length + ' pedidos | chamadas: ' + s.chamadas +
    ' em ' + s.segundos + 's | total por posição: ' + JSON.stringify(s.porPosicao) + ' | modo do offset: ' + s.modoOffset);
  const porLote = {};
  busca.pedidos.forEach(function (p) { const k = String(p.lotesep); porLote[k] = (porLote[k] || 0) + 1; });
  Logger.log('Pedidos por lotesep (0 = sem lote): ' + JSON.stringify(porLote));
  const histograma = function (campo) {
    const h = {};
    busca.pedidos.forEach(function (p) {
      if (Number(p.lotesep) > 0) { const d = String(p[campo]).slice(0, 10); h[d] = (h[d] || 0) + 1; }
    });
    const chave = function (d) { return d.split('/').reverse().join(''); };
    return Object.keys(h).sort(function (x, y) { return chave(x).localeCompare(chave(y)); })
      .map(function (d) { return d + '=' + h[d]; }).join(' | ');
  };
  Logger.log('Pedidos EM LOTE por data_pedido (para escolher AMBER_DIAS_ANTES): ' + histograma('data_pedido'));
  Logger.log('Pedidos EM LOTE por data_limite_postagem (30/12/1899 = sem data): ' + histograma('data_limite_postagem'));
  if (s.amostra) {
    const a = JSON.parse(JSON.stringify(s.amostra)); delete a.etiquetas;
    Logger.log('Exemplo de pedido (sem etiquetas):\n' + JSON.stringify(a, null, 1).slice(0, 4000));
  }
  const doLote = busca.pedidos.filter(function (p) { return Number(p.lotesep) === Number(lote); });
  Logger.log('Pedidos do lote ' + lote + ': ' + doLote.length);
  doLote.forEach(function (p) {
    Logger.log('  pedido ' + p.pedido + ' | posição ' + p.posicao + ' | status ' + p.status + ' | data_pedido ' + p.data_pedido +
      ' | itens: ' + (p.itens || []).map(function (i) { return '[' + i.codigo_interno + ' x' + i.quantidade + ']'; }).join(' '));
    (p.itens || []).forEach(function (i) {
      if (!String(i.codigo_interno == null ? '' : i.codigo_interno).trim() || !(Number(i.quantidade) > 0)) {
        Logger.log('  ⚠️ item DESCARTADO no pedido ' + p.pedido + ' (sem codigo_interno ou quantidade <= 0): ' + JSON.stringify(i));
      }
    });
  });
  const itens = agruparPorSku_(mapearItensAmber_(busca.pedidos, lote));
  Logger.log('Itens do lote ' + lote + ' (' + itens.length + ' SKUs): ' + JSON.stringify(completarDescricoes_(itens).slice(0, 30)));
}

/**
 * Diagnóstico: consulta UM pedido pelo número (só leitura) e mostra posição, datas, lote e itens.
 * Use quando um pedido que deveria estar no lote não apareceu. Informe o número em CONFIG.TESTE_PEDIDO
 * (ou na propriedade do script TESTE_PEDIDO). Não mostra dados do cliente.
 */
function testarPedido() {
  const num = Number(PropertiesService.getScriptProperties().getProperty('TESTE_PEDIDO') || CONFIG.TESTE_PEDIDO);
  if (!num) throw new Error('Informe o número do pedido em CONFIG.TESTE_PEDIDO.');
  const r = chamarAmber_([requisicaoAmber_('/pedidosvenda/consulta', { pedido: num })])[0];
  const lista = (r && r.resultado && r.resultado.pedidos) || [];
  if (!lista.length) { Logger.log('Amber: ' + JSON.stringify(r && r.mensagem) + ' (pedido ' + num + ' não encontrado)'); return; }
  lista.forEach(function (p) {
    Logger.log('Pedido ' + p.pedido + ' | posição ' + p.posicao + ' | status ' + p.status + ' | lotesep ' + p.lotesep +
      ' | data_pedido ' + p.data_pedido + ' | data_aprovacao ' + p.data_aprovacao + ' | limite_postagem ' + p.data_limite_postagem +
      ' | posicao_etiqueta ' + p.posicao_etiqueta + '\nitens: ' + JSON.stringify((p.itens || []).map(function (i) {
        return { codigo_interno: i.codigo_interno, codigo_aton: i.codigo_aton, quantidade: i.quantidade };
      })));
  });
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
    if (!mapa[sku]) { mapa[sku] = { sku: sku, descricao: i.descricao || '', qtd: 0, codigoAton: i.codigoAton }; ordem.push(sku); }
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

function verificarLoteNaoLancado_(lote) {
  const sep = separadorDoLote_(lote);
  if (sep) {
    throw new Error('O lote ' + lote + ' já foi lançado (separador: ' + sep + '). Um lote não pode ser lançado duas vezes.');
  }
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
function brData_(iso) { const p = iso.split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }
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
