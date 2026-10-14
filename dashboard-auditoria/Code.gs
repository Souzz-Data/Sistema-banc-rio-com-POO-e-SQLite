/**
 * Dashboard de Auditoria de Atendimento (Filiais + Central de Relacionamento)
 *
 * Fluxo: Google Forms -> Google Sheets -> este Web App (Apps Script).
 * O gestor abre o link, filtra e baixa o PDF da auditoria (ou de todas as
 * auditorias de um AR) com um clique.
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
  meta: 70
};

/* ------------------------------------------------------------------ */

function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.origemInicial = (e && e.parameter && e.parameter.origem) || Object.keys(CONFIG.origens)[0];
  return tpl.evaluate()
    .setTitle('Auditoria de Atendimento')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Lista as origens configuradas (para o seletor no topo). */
function listarOrigens() {
  return Object.keys(CONFIG.origens).map(function (id) {
    return { id: id, titulo: CONFIG.origens[id].titulo };
  });
}

/** Lê a planilha e devolve perguntas + auditorias já pontuadas. */
function getDados(origemId) {
  const origem = CONFIG.origens[origemId];
  if (!origem) throw new Error('Origem desconhecida: ' + origemId);

  const ss = origem.planilhaId
    ? SpreadsheetApp.openById(origem.planilhaId)
    : SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(origem.aba) || ss.getSheets()[0];
  const valores = aba.getDataRange().getValues();
  if (valores.length < 2) return montarRetorno(origemId, [], []);

  const cabecalho = valores.shift().map(function (h) { return String(h).trim(); });
  const nomes = Object.assign({}, CONFIG.colunas, origem.colunas || {});
  const idx = {};
  Object.keys(nomes).forEach(function (campo) { idx[campo] = acharColuna(cabecalho, nomes[campo]); });

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

    const dataBruta = pegar(linha, idx.data) || pegar(linha, idx.carimbo);
    const data = paraData(dataBruta);
    if (inicio && data && data < inicio) return;

    let obtido = 0, possivel = 0;
    const respostas = perguntas.map(function (p) {
      const bruto = linha[p.col];
      const texto = bruto === '' || bruto === null ? '' : String(bruto).trim();
      const pts = pontuar(bruto, p.vale);
      if (pts !== null) { obtido += pts; possivel += p.vale; }
      return { r: texto, p: pts };
    });

    let nota = paraNumero(pegar(linha, idx.nota));
    if (nota !== null && nota > 0 && nota <= 1) nota = nota * 100;   // célula formatada como %
    if (nota === null && possivel > 0) nota = obtido / possivel * 100;

    auditorias.push({
      id: n + 2,                                  // número da linha na planilha
      data: data ? Utilities.formatDate(data, tz, 'yyyy-MM-dd') : '',
      ar: texto(pegar(linha, idx.ar)),
      auditor: texto(pegar(linha, idx.auditor)),
      filial: texto(pegar(linha, idx.filial)),
      regional: texto(pegar(linha, idx.regional)),
      nota: nota === null ? null : Math.round(nota * 10) / 10,
      respostas: respostas,
      pontos: texto(pegar(linha, idx.pontos))
    });
  });

  return montarRetorno(origemId, perguntas, auditorias);
}

/**
 * Converte o HTML do relatório (montado no navegador) em PDF.
 * Devolve base64 para o navegador baixar.
 */
function gerarPdf(html, nomeArquivo) {
  const nome = String(nomeArquivo || 'auditoria').replace(/[\\/:*?"<>|]+/g, '-') + '.pdf';
  const pdf = Utilities.newBlob(html, MimeType.HTML, 'relatorio.html').getAs(MimeType.PDF).setName(nome);
  return { nome: nome, base64: Utilities.base64Encode(pdf.getBytes()) };
}

/* ------------------------------ helpers ----------------------------- */

function montarRetorno(origemId, perguntas, auditorias) {
  return {
    origem: origemId,
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

function normalizar(s) {
  return String(s === null || s === undefined ? '' : s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

function acharColuna(cabecalho, nomes) {
  const alvos = [].concat(nomes).map(normalizar);
  const cab = cabecalho.map(normalizar);
  for (let a = 0; a < alvos.length; a++) {
    const i = cab.indexOf(alvos[a]);
    if (i >= 0) return i;
  }
  for (let a = 0; a < alvos.length; a++) {
    for (let i = 0; i < cab.length; i++) if (cab[i].indexOf(alvos[a]) === 0) return i;
  }
  return -1;
}

function pegar(linha, i) { return i >= 0 ? linha[i] : ''; }

function texto(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return v === null || v === undefined ? '' : String(v).trim();
}

function paraNumero(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').replace('%', '').replace(',', '.').trim();
  if (!s) return null;
  const n = Number(s);
  return isNaN(n) ? null : n;
}

function paraData(v) {
  if (v instanceof Date) return v;
  const m = String(v || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function pontuar(valor, vale) {
  if (typeof valor === 'number') return valor;
  const n = normalizar(valor);
  if (!n) return null;
  const num = paraNumero(n);
  if (num !== null) return num;
  const chaves = Object.keys(CONFIG.pesos);
  for (let k = 0; k < chaves.length; k++) {
    if (normalizar(chaves[k]) === n) {
      const peso = CONFIG.pesos[chaves[k]];
      return peso === null ? null : peso * vale;
    }
  }
  return null;
}
