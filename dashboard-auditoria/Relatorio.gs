/**
 * Monta o HTML do "Relatório de Feedback" (frente e verso) que vira PDF.
 *
 * JS puro, sem APIs do Apps Script: o mesmo arquivo é usado no servidor
 * (PDF e prévia) e na página de demonstração.
 *
 * Layout pensado para impressão frente e verso:
 *   frente -> logo, nota, dados do atendimento e perguntas
 *   verso  -> Pontos de Auditoria (continua na folha seguinte se for longo)
 * Com várias auditorias no mesmo PDF, cada uma começa sempre numa frente:
 * quando o verso estimado tem número par de páginas, entra uma página em branco.
 */

const REL_COR = { azul: '#025B94', vermelho: '#F04131', texto: '#1c2733', suave: '#5b6b7b', linha: '#d5e0ea', fundo: '#eef4f9',
                  bom: '#1e7b4f', medio: '#b7791f', ruim: '#c62828' };
const REL_LINHAS_POR_PAGINA = 44;   // linhas de texto do verso que cabem numa página A4
const REL_CARACTERES_POR_LINHA = 95;

/**
 * @param dados  { titulo, meta, rotulos, perguntas: [{titulo, vale}] }
 * @param lista  auditorias [{ data, ar, auditor, filial, regional, nota, respostas: [{r, p}], pontos }]
 * @param op     { tela: bool (prévia na tela), logo: dataURI, geradoPor: email, geradoEm: texto }
 */
function montarRelatorio_(dados, lista, op) {
  op = op || {};
  const C = REL_COR, r = dados.rotulos;
  const td = 'padding:5px 7px;border-bottom:1px solid ' + C.linha + ';font-size:10pt;vertical-align:top;';
  const th = 'padding:6px 7px;background:' + C.azul + ';color:#ffffff;font-size:9.5pt;text-align:left;';
  const quebra = 'page-break-before:always;';
  const folha = op.tela
    ? 'background:#ffffff;max-width:760px;margin:0 auto 18px;padding:30px 34px;box-shadow:0 2px 10px rgba(0,0,0,.14);min-height:900px;'
    : '';

  const rodape = '<div style="margin-top:18px;padding-top:6px;border-top:1px solid ' + C.linha + ';font-size:7.5pt;color:' + C.suave + '">' +
    'Confidencial · uso interno IBEU' +
    (op.geradoEm ? ' · gerado em ' + relEsc_(op.geradoEm) : '') +
    (op.geradoPor ? ' por ' + relEsc_(op.geradoPor) : '') + '</div>';

  const logo = op.logo ? '<img src="' + relEsc_(op.logo) + '" width="54" height="54" style="width:54px;height:54px" alt="IBEU">' : '';

  const cabecalho = function (a, verso) {
    return '<table style="width:100%;border-collapse:collapse;border-bottom:3px solid ' + C.vermelho + ';margin-bottom:12px"><tr>' +
      '<td style="width:62px;padding:0 8px 8px 0;vertical-align:middle">' + logo + '</td>' +
      '<td style="padding:0 0 8px;vertical-align:middle">' +
        '<div style="font-size:' + (verso ? '12' : '15') + 'pt;font-weight:bold;color:' + C.azul + '">' +
          (verso ? 'Pontos de Auditoria' : 'Relatório de Feedback') + '</div>' +
        '<div style="font-size:9.5pt;color:' + C.suave + '">' +
          (verso ? relEsc_(a.ar) + ' · ' + relDataBR_(a.data) : 'Auditoria de Atendimento · ' + relEsc_(dados.titulo)) + '</div>' +
      '</td>' +
      '<td style="width:150px;padding:0 0 8px;text-align:right;vertical-align:middle">' +
        '<div style="font-size:8.5pt;color:' + C.suave + '">Nota do atendimento</div>' +
        '<div style="font-size:' + (verso ? '16' : '26') + 'pt;font-weight:bold;color:' + relCorNota_(a.nota, dados.meta) + '">' +
          (a.nota === null || a.nota === undefined ? '–' : relFmt_(a.nota, 1)) + '</div>' +
      '</td></tr></table>';
  };

  const frente = function (a) {
    const info = [[r.ar, a.ar], [r.filial, a.filial], [r.regional, a.regional], [r.auditor, a.auditor], ['Data da auditoria', relDataBR_(a.data)]]
      .filter(function (x) { return x[1]; });
    const perguntas = dados.perguntas.map(function (p, k) {
      const resp = a.respostas[k] || { r: '', p: null };
      return '<tr>' +
        '<td style="' + td + '">' + relEsc_(p.titulo) + '</td>' +
        '<td style="' + td + 'color:' + relCorResposta_(resp.r) + ';font-weight:bold;white-space:nowrap">' + relEsc_(resp.r || '—') + '</td>' +
        '<td style="' + td + 'text-align:right;white-space:nowrap">' + (resp.p === null ? '–' : relFmt_(resp.p, 1)) + ' / ' + relFmt_(p.vale, 0) + '</td>' +
        '</tr>';
    }).join('');
    return cabecalho(a, false) +
      '<table style="width:100%;border-collapse:collapse;margin-bottom:14px">' +
        info.map(function (x) {
          return '<tr><td style="' + td + 'width:150px;color:' + C.suave + '">' + relEsc_(x[0]) + '</td>' +
                 '<td style="' + td + 'font-weight:bold">' + relEsc_(x[1]) + '</td></tr>';
        }).join('') +
      '</table>' +
      '<table style="width:100%;border-collapse:collapse">' +
        '<tr><th style="' + th + '">Pergunta</th><th style="' + th + 'width:95px">Resposta</th><th style="' + th + 'width:70px;text-align:right">Pontos</th></tr>' +
        perguntas +
      '</table>' +
      (a.pontos ? '<div style="margin-top:10px;font-size:9pt;color:' + C.suave + '">Pontos de Auditoria no verso →</div>' : '') +
      rodape;
  };

  const verso = function (a) {
    return cabecalho(a, true) +
      '<div style="font-size:10.5pt;line-height:1.5;color:' + C.texto + '">' + relFormatarPontos_(a.pontos) + '</div>' +
      rodape;
  };

  const emBranco = '<div style="' + quebra + 'padding-top:300px;text-align:center;font-size:9pt;color:#aab6c2">' +
    'Página em branco para impressão frente e verso</div>';

  let corpo = '';
  lista.forEach(function (a, i) {
    const ultima = i === lista.length - 1;
    corpo += '<div style="' + (i > 0 ? quebra : '') + folha + '">' + frente(a) + '</div>';
    let paginasVerso = 0;
    if (a.pontos) {
      corpo += '<div style="' + quebra + folha + '">' + verso(a) + '</div>';
      paginasVerso = relEstimarPaginas_(a.pontos);
    }
    // frente (1) + verso precisa ser par para a próxima auditoria começar numa frente
    if (!ultima && !op.tela && (1 + paginasVerso) % 2 === 1) corpo += emBranco;
  });

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Relatório de Feedback</title>' +
    '<style>@page { size: A4; margin: 14mm; } body { margin:0; -webkit-print-color-adjust:exact; print-color-adjust:exact; }</style></head>' +
    '<body style="font-family:Arial,Helvetica,sans-serif;color:' + C.texto + ';' + (op.tela ? 'background:#e9eef3;padding:18px 12px;' : '') + '">' +
    corpo + '</body></html>';
}

/** Estimativa de páginas do verso (o conversor de PDF não informa a contagem real). */
function relEstimarPaginas_(txt) {
  const linhas = String(txt || '').split(/\r?\n/).reduce(function (soma, l) {
    return soma + Math.max(1, Math.ceil(l.trim().length / REL_CARACTERES_POR_LINHA));
  }, 0) + 6;   // + cabeçalho e rodapé do verso
  return Math.max(1, Math.ceil(linhas / REL_LINHAS_POR_PAGINA));
}

/** Quebras de linha preservadas; títulos ("Pontos fortes:", "1. Diagnóstico…") em negrito; "-" vira marcador. */
function relFormatarPontos_(txt) {
  return String(txt || '').split(/\r?\n/).map(function (l) {
    const s = l.trim();
    if (!s) return '<div style="height:8px"></div>';
    if (/^[-•*]\s+/.test(s)) return '<div style="padding-left:14px">• ' + relEsc_(s.replace(/^[-•*]\s+/, '')) + '</div>';
    const titulo = s.length <= 90 && (/:$/.test(s) || /^\d+[.)]\s+\S/.test(s) ||
      /^(pontos fortes|oportunidades|plano de|pontos de melhoria)/i.test(s));
    return titulo
      ? '<div style="font-weight:bold;color:' + REL_COR.azul + ';margin-top:6px">' + relEsc_(s) + '</div>'
      : '<div>' + relEsc_(s) + '</div>';
  }).join('');
}

function relEsc_(s) {
  return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function relFmt_(v, casas) {
  return Number(v).toFixed(casas).replace('.', ',');
}

function relDataBR_(iso) {
  if (!iso) return '–';
  const p = String(iso).split('-');
  return p[2] + '/' + p[1] + '/' + p[0];
}

function relCorNota_(v, meta) {
  if (v === null || v === undefined) return REL_COR.suave;
  return v >= meta ? REL_COR.bom : v >= meta - 15 ? REL_COR.medio : REL_COR.ruim;
}

function relCorResposta_(s) {
  const n = String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
  return { sim: REL_COR.bom, nao: REL_COR.ruim, parcialmente: REL_COR.medio }[n] || '#555555';
}
