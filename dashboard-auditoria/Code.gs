/**
 * Dashboard de Auditoria de Atendimento (Filiais + Central de Relacionamento)
 *
 * Fluxo: Google Forms -> Google Sheets -> este Web App (Apps Script).
 * O gestor abre o link, filtra e baixa o PDF da auditoria (ou de todas as
 * auditorias de um AR) com um clique.
 *
 * Segurança:
 *  - Implantar como "Executar como: Eu" + "Quem pode acessar: usuários do domínio".
 *  - Só entra quem está na aba "Acessos" da planilha de controle, e cada pessoa
 *    só recebe as linhas das filiais/equipes liberadas para ela.
 *  - Toda checagem é feita aqui no servidor; o navegador nunca recebe o que não pode ver.
 *  - Aberturas e downloads ficam registrados na aba "Log".
 *
 * Tudo que depende da sua planilha está no CONFIG abaixo.
 */

const CONFIG = {
  // Uma entrada por formulário/planilha. A chave é usada na URL (?origem=central).
  origens: {
    filiais: {
      titulo: 'Filiais',
      planilhaId: '',                       // vazio = planilha onde o script está vinculado
      aba: 'Respostas ao formulário 1',     // nome da aba com as respostas
      colunas: {},                          // sobrescreve CONFIG.colunas só para esta origem
      rotulos: {}                           // nomes exibidos na tela (ex.: { filial: 'Equipe' })
    },
    central: {
      titulo: 'Central de Relacionamento',
      planilhaId: '',                       // ID da planilha da Central (trecho da URL entre /d/ e /edit)
      aba: 'Respostas ao formulário 1',
      // Na Central não há filial: agrupe por equipe/supervisor, se a planilha tiver essas colunas.
      colunas: { filial: ['Equipe', 'Célula', 'Supervisor'], regional: ['Coordenação', 'Turno'] },
      rotulos: { filial: 'Equipe', regional: 'Coordenação' }
    }
  },

  // Cabeçalho de cada campo na planilha. Pode ser texto ou lista de alternativas.
  // A busca ignora maiúsculas, acentos e espaços extras.
  colunas: {
    carimbo:  ['Carimbo de data/hora', 'Timestamp'],
    data:     ['Data da Auditoria', 'Data do Atendimento'],
    ar:       ['Nome do AR', 'AR', 'Nome do Atendente', 'Nome do Consultor'],
    auditor:  ['Nome do Auditor', 'Auditor'],
    filial:   ['Filial', 'Unidade'],
    regional: ['Regional'],
    nota:     ['Nota', 'Nota Atendimento', 'Nota Final', 'Pontuação'],
    pontos:   ['Pontos de Auditoria', 'Pontos da Auditoria', 'Feedback', 'Observações']
  },

  // Toda coluna cujo cabeçalho tenha "(Vale X)" é tratada como pergunta valendo X pontos.
  padraoPergunta: /\(\s*vale\s*(\d+(?:[.,]\d+)?)\s*\)/i,

  // Quanto cada resposta vale (fração do "Vale"). null = não entra na conta.
  // Se a resposta já for um número, ele é usado como pontos diretamente.
  pesos: {
    'Sim': 1,
    'Parcialmente': 0.5,
    'Não': 0,
    'N/C': 1          // troque para null se "N/C" não deve contar a favor nem contra
  },

  // Só considera auditorias a partir desta data (aaaa-mm-dd). Vazio = todas.
  dataInicio: '2026-05-05',

  // Meta de nota (linha de referência e destaque nas tabelas).
  meta: 70,

  // Só e-mails deste domínio entram (checagem extra além da implantação restrita ao domínio).
  dominio: 'ibeu.com.br',

  // Planilha de controle (compartilhada só com quem administra os acessos).
  controle: {
    planilhaId: '',           // ID da planilha de controle
    abaAcessos: 'Acessos',    // E-mail | Painel | Filiais/Equipes | Ativo
    abaLog: 'Log'             // preenchida automaticamente
  },

  // Quem procurar para pedir acesso (aparece na tela de acesso negado).
  contatoAcesso: 'a equipe de Qualidade',

  // Máximo de auditorias por PDF.
  limitePdf: 60,

  // Minutos que a lista de acessos fica em cache (alterações na aba levam até isso para valer).
  cacheMinutos: 5
};

/* ------------------------------------------------------------------ */
/*
 * IMPORTANTE (segurança): no Apps Script, toda função global cujo nome NÃO
 * termina com "_" pode ser chamada pelo navegador (google.script.run).
 * Só as 5 funções abaixo são públicas, e todas conferem o usuário logado.
 * Qualquer função auxiliar nova deve terminar com "_".
 */

function doGet(e) {
  let email = '', perm = {};
  try {
    email = usuarioAtual_();
    perm = permissoes_(email);
  } catch (err) {
    perm = {};
  }
  const permitidas = Object.keys(perm);
  if (!permitidas.length) {
    const negado = HtmlService.createTemplateFromFile('Negado');
    negado.email = email;
    negado.contato = CONFIG.contatoAcesso;
    negado.logo = LOGO_PNG;
    registrar_(email || '(desconhecido)', 'acesso negado', '', '', '');
    return negado.evaluate().setTitle('Acesso não autorizado')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const pedida = e && e.parameter && e.parameter.origem;
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.origemInicial = permitidas.indexOf(pedida) >= 0 ? pedida : permitidas[0];
  tpl.usuario = email;
  tpl.logo = LOGO_PNG;
  return tpl.evaluate()
    .setTitle('Auditoria de Atendimento · IBEU')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Painéis que o usuário logado pode ver (para o seletor no topo). */
function listarOrigens() {
  const perm = permissoes_(usuarioAtual_());
  return Object.keys(CONFIG.origens)
    .filter(function (id) { return perm[id]; })
    .map(function (id) { return { id: id, titulo: CONFIG.origens[id].titulo }; });
}

/** Dados do painel, já filtrados pelas filiais/equipes liberadas para o usuário. */
function getDados(origemId) {
  const email = usuarioAtual_();
  const grupos = exigirPainel_(email, origemId);
  const base = lerAuditorias_(origemId);
  const auditorias = base.auditorias.filter(function (a) { return podeVer_(grupos, a); });
  registrar_(email, 'abriu painel', origemId, '', auditorias.length + ' auditorias');
  return montarRetorno_(origemId, base.perguntas, auditorias, email);
}

/** HTML do relatório para a pré-visualização na tela. */
function previaRelatorio(origemId, ids) {
  const email = usuarioAtual_();
  const r = auditoriasAutorizadas_(email, origemId, ids);
  registrar_(email, 'visualizou', origemId, nomesAr_(r.lista), idsTexto_(r.lista));
  return montarRelatorio_(r.meta, r.lista, { tela: true, logo: LOGO_PNG });
}

/**
 * Gera o PDF das auditorias pedidas. Recebe só IDs: o conteúdo é relido da
 * planilha e cada ID é conferido contra a permissão de quem pediu.
 */
function gerarPdf(origemId, ids) {
  const email = usuarioAtual_();
  const r = auditoriasAutorizadas_(email, origemId, ids);
  const agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const html = montarRelatorio_(r.meta, r.lista, { tela: false, logo: LOGO_PNG, geradoPor: email, geradoEm: agora });
  const nome = nomeArquivo_(r.meta.titulo, r.lista);
  const pdf = Utilities.newBlob(html, MimeType.HTML, 'relatorio.html').getAs(MimeType.PDF).setName(nome);
  registrar_(email, 'baixou PDF', origemId, nomesAr_(r.lista), idsTexto_(r.lista));
  return { nome: nome, base64: Utilities.base64Encode(pdf.getBytes()) };
}

/* ---------------------------- permissões ---------------------------- */

/** E-mail de quem está usando o app. Recusa anônimo e outros domínios. */
function usuarioAtual_() {
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  if (!email) throw new Error('Acesso não autorizado.');
  if (CONFIG.dominio && !email.endsWith('@' + CONFIG.dominio.toLowerCase())) throw new Error('Acesso não autorizado.');
  return email;
}

/**
 * Lê a aba Acessos e devolve { origemId: ['*'] | ['botafogo', ...] } para o e-mail.
 * Colunas: E-mail | Painel (Filiais / Central / Ambos) | Filiais ou Equipes ("Todas" ou lista) | Ativo (Sim/Não)
 */
function permissoes_(email) {
  const cache = CacheService.getScriptCache();
  const chave = 'perm:' + email;
  const salvo = cache.get(chave);
  if (salvo) return JSON.parse(salvo);

  const perm = {};
  const c = CONFIG.controle;
  if (!c.planilhaId) throw new Error('Configure CONFIG.controle.planilhaId.');
  const aba = SpreadsheetApp.openById(c.planilhaId).getSheetByName(c.abaAcessos);
  if (!aba) throw new Error('Aba de acessos não encontrada: ' + c.abaAcessos);
  const linhas = aba.getDataRange().getValues().slice(1);

  linhas.forEach(function (l) {
    if (String(l[0] || '').toLowerCase().trim() !== email) return;
    const ativo = normalizar_(l[3]);
    if (ativo && ativo !== 'sim') return;
    const paineis = idsDoPainel_(l[1]);
    const lista = String(l[2] || '').split(/[,;\n]/).map(normalizar_).filter(Boolean);
    const todas = !lista.length || lista.indexOf('todas') >= 0 || lista.indexOf('todos') >= 0;
    paineis.forEach(function (id) {
      if (todas) { perm[id] = ['*']; return; }
      if (perm[id] && perm[id][0] === '*') return;
      perm[id] = (perm[id] || []).concat(lista);
    });
  });

  cache.put(chave, JSON.stringify(perm), CONFIG.cacheMinutos * 60);
  return perm;
}

/** "Filiais", "Central", "Central de Relacionamento", "Ambos" -> ids do CONFIG. */
function idsDoPainel_(valor) {
  const n = normalizar_(valor);
  const ids = Object.keys(CONFIG.origens);
  if (!n) return [];
  if (n === 'ambos' || n === 'todos' || n === 'todas') return ids;
  return ids.filter(function (id) {
    const titulo = normalizar_(CONFIG.origens[id].titulo);
    return n === id || n === titulo || titulo.indexOf(n) === 0;
  });
}

function exigirPainel_(email, origemId) {
  if (!Object.prototype.hasOwnProperty.call(CONFIG.origens, origemId)) throw new Error('Acesso não autorizado.');
  const grupos = permissoes_(email)[origemId];
  if (!grupos) throw new Error('Acesso não autorizado.');
  return grupos;
}

function podeVer_(grupos, auditoria) {
  return grupos[0] === '*' || grupos.indexOf(normalizar_(auditoria.filial)) >= 0;
}

/** Relê a planilha e devolve as auditorias pedidas, recusando qualquer uma fora da permissão. */
function auditoriasAutorizadas_(email, origemId, ids) {
  const grupos = exigirPainel_(email, origemId);
  if (!Array.isArray(ids) || !ids.length) throw new Error('Nenhuma auditoria selecionada.');
  if (ids.length > CONFIG.limitePdf) throw new Error('Máximo de ' + CONFIG.limitePdf + ' auditorias por arquivo.');
  const base = lerAuditorias_(origemId);
  const porId = {};
  base.auditorias.forEach(function (a) { porId[a.id] = a; });
  const lista = ids.map(function (id) {
    const a = porId[Number(id)];
    if (!a || !podeVer_(grupos, a)) throw new Error('Acesso não autorizado.');
    return a;
  });
  return { meta: montarRetorno_(origemId, base.perguntas, [], email), lista: lista };
}

/* ------------------------------- log -------------------------------- */

function registrar_(email, acao, origemId, ar, detalhe) {
  try {
    const c = CONFIG.controle;
    if (!c.planilhaId) return;
    const ss = SpreadsheetApp.openById(c.planilhaId);
    const aba = ss.getSheetByName(c.abaLog) || ss.insertSheet(c.abaLog);
    if (aba.getLastRow() === 0) aba.appendRow(['Data/hora', 'E-mail', 'Ação', 'Painel', 'AR', 'Detalhe']);
    aba.appendRow([new Date(), email, acao, origemId, ar, detalhe].map(seguroParaCelula_));
  } catch (err) {
    console.error('Falha ao registrar_ log: ' + err);
  }
}

/** Evita que um texto vindo do formulário (ex.: "=IMPORTXML(...)") vire fórmula no log. */
function seguroParaCelula_(v) {
  if (v instanceof Date) return v;
  const s = String(v === null || v === undefined ? '' : v);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function nomesAr_(lista) {
  const vistos = {};
  return lista.map(function (a) { return a.ar; })
    .filter(function (n) { return n && !vistos[n] && (vistos[n] = true); }).join(', ');
}

function idsTexto_(lista) {
  return 'linhas ' + lista.map(function (a) { return a.id; }).join(', ');
}

function nomeArquivo_(titulo, lista) {
  const ars = nomesAr_(lista).split(', ');
  let nome;
  if (lista.length === 1) nome = 'Auditoria ' + titulo + ' - ' + lista[0].ar + ' - ' + relDataBR_(lista[0].data).replace(/\//g, '-');
  else if (ars.length === 1) nome = 'Auditorias ' + titulo + ' - ' + ars[0] + ' (' + lista.length + ')';
  else nome = 'Auditorias ' + titulo + ' - ' + lista.length + ' registros';
  return nome.replace(/[\\/:*?"<>|]+/g, '-').slice(0, 150) + '.pdf';
}

/* ------------------------- leitura da planilha ----------------------- */

/** Lê a planilha da origem e devolve perguntas + todas as auditorias pontuadas (sem filtro de acesso). */
function lerAuditorias_(origemId) {
  const origem = CONFIG.origens[origemId];
  const ss = origem.planilhaId
    ? SpreadsheetApp.openById(origem.planilhaId)
    : SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(origem.aba) || ss.getSheets()[0];
  const valores = aba.getDataRange().getValues();
  if (valores.length < 2) return { perguntas: [], auditorias: [] };

  const cabecalho = valores.shift().map(function (h) { return String(h).trim(); });
  const nomes = Object.assign({}, CONFIG.colunas, origem.colunas || {});
  const idx = {};
  Object.keys(nomes).forEach(function (campo) { idx[campo] = acharColuna_(cabecalho, nomes[campo]); });

  const perguntas = [];
  cabecalho.forEach(function (h, i) {
    const m = h.match(CONFIG.padraoPergunta);
    if (m) perguntas.push({ col: i, titulo: h, vale: Number(m[1].replace(',', '.')) });
  });

  const inicio = CONFIG.dataInicio ? new Date(CONFIG.dataInicio + 'T00:00:00') : null;
  const tz = Session.getScriptTimeZone();
  const auditorias = [];

  valores.forEach(function (linha, n) {
    if (linha.every(function (v) { return v === '' || v === null; })) return;

    const dataBruta = pegar_(linha, idx.data) || pegar_(linha, idx.carimbo);
    const data = paraData_(dataBruta);
    if (inicio && data && data < inicio) return;

    let obtido = 0, possivel = 0;
    const respostas = perguntas.map(function (p) {
      const bruto = linha[p.col];
      const texto = bruto === '' || bruto === null ? '' : String(bruto).trim();
      const pts = pontuar_(bruto, p.vale);
      if (pts !== null) { obtido += pts; possivel += p.vale; }
      return { r: texto, p: pts };
    });

    let nota = paraNumero_(pegar_(linha, idx.nota));
    if (nota !== null && nota > 0 && nota <= 1) nota = nota * 100;   // célula formatada como %
    if (nota === null && possivel > 0) nota = obtido / possivel * 100;

    auditorias.push({
      id: n + 2,                                  // número da linha na planilha
      data: data ? Utilities.formatDate(data, tz, 'yyyy-MM-dd') : '',
      ar: texto_(pegar_(linha, idx.ar)),
      auditor: texto_(pegar_(linha, idx.auditor)),
      filial: texto_(pegar_(linha, idx.filial)),
      regional: texto_(pegar_(linha, idx.regional)),
      nota: nota === null ? null : Math.round(nota * 10) / 10,
      respostas: respostas,
      pontos: texto_(pegar_(linha, idx.pontos))
    });
  });

  return { perguntas: perguntas, auditorias: auditorias };
}

/* ------------------------------ helpers ----------------------------- */

function montarRetorno_(origemId, perguntas, auditorias, usuario) {
  return {
    origem: origemId,
    usuario: usuario || '',
    titulo: CONFIG.origens[origemId].titulo,
    rotulos: Object.assign({ ar: 'AR', filial: 'Filial', regional: 'Regional', auditor: 'Auditor' },
                           CONFIG.origens[origemId].rotulos || {}),
    meta: CONFIG.meta,
    dataInicio: CONFIG.dataInicio,
    atualizadoEm: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    perguntas: perguntas.map(function (p) { return { titulo: p.titulo, vale: p.vale }; }),
    auditorias: auditorias
  };
}

function normalizar_(s) {
  return String(s === null || s === undefined ? '' : s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

function acharColuna_(cabecalho, nomes) {
  const alvos = [].concat(nomes).map(normalizar_);
  const cab = cabecalho.map(normalizar_);
  for (let a = 0; a < alvos.length; a++) {
    const i = cab.indexOf(alvos[a]);
    if (i >= 0) return i;
  }
  for (let a = 0; a < alvos.length; a++) {
    for (let i = 0; i < cab.length; i++) if (cab[i].indexOf(alvos[a]) === 0) return i;
  }
  return -1;
}

function pegar_(linha, i) { return i >= 0 ? linha[i] : ''; }

function texto_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return v === null || v === undefined ? '' : String(v).trim();
}

function paraNumero_(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').replace('%', '').replace(',', '.').trim();
  if (!s) return null;
  const n = Number(s);
  return isNaN(n) ? null : n;
}

function paraData_(v) {
  if (v instanceof Date) return v;
  const m = String(v || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function pontuar_(valor, vale) {
  if (typeof valor === 'number') return valor;
  const n = normalizar_(valor);
  if (!n) return null;
  const num = paraNumero_(n);
  if (num !== null) return num;
  const chaves = Object.keys(CONFIG.pesos);
  for (let k = 0; k < chaves.length; k++) {
    if (normalizar_(chaves[k]) === n) {
      const peso = CONFIG.pesos[chaves[k]];
      return peso === null ? null : peso * vale;
    }
  }
  return null;
}
