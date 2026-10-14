// Testes de controle de acesso (rodar com: node dashboard-auditoria/testes/teste-permissoes.js)
// Testa Code.gs + Relatorio.gs + Logo.gs com stubs das APIs do Apps Script
const fs = require('fs'), vm = require('vm');
const D = process.argv[2] || require('path').join(__dirname, '..');
const PERG = ['Saudar o Lead (Vale 5)', 'Identificar a dor ou desejo do lead (Vale 10)'];
const planilhas = {
  filiais: [['Carimbo de data/hora','Nome do Auditor','Nome do AR','Filial','Regional','Data da Auditoria',...PERG,'Pontos de Auditoria'],
            [new Date(2026,5,1),'Aud','=IMPORTXML("x")','Botafogo','Zona Sul',new Date(2026,5,1),'Sim','Não','Pontos fortes:\nok'],
            [new Date(2026,5,2),'Aud','Bruno','Méier','Zona Norte',new Date(2026,5,2),'Sim','Sim',''],
            [new Date(2026,3,2),'Aud','Antigo','Botafogo','Zona Sul',new Date(2026,3,2),'Sim','Sim','']],
  central: [['Carimbo de data/hora','Nome do AR','Equipe','Data da Auditoria',...PERG],
            [new Date(2026,6,1),'Carla','Equipe WhatsApp',new Date(2026,6,1),'Parcialmente','Sim']],
  acessos: [['E-mail','Painel','Filiais/Equipes','Ativo'],
            ['gestor.bota@ibeu.com.br','Filiais','Botafogo','Sim'],
            ['diretor@ibeu.com.br','Ambos','Todas','Sim'],
            ['ex@ibeu.com.br','Filiais','Todas','Não'],
            ['central@ibeu.com.br','Central de Relacionamento','','Sim'],
            ['fora@gmail.com','Filiais','Todas','Sim']]
};
const log = [];
const aba = rows => ({ getDataRange: () => ({ getValues: () => rows.map(r => r.slice()) }), getLastRow: () => log.length, appendRow: r => log.push(r) });
let email = '';
const ctx = {
  console, JSON, Object, Array, String, Number, Math, Date, RegExp, Error, isNaN,
  Session: { getActiveUser: () => ({ getEmail: () => email }), getScriptTimeZone: () => 'America/Sao_Paulo' },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ({ getSheetByName: () => aba(planilhas.filiais), getSheets: () => [] }),
    openById: id => id === 'CTRL'
      ? { getSheetByName: n => n === 'Acessos' ? aba(planilhas.acessos) : aba(log), insertSheet: () => aba(log) }
      : { getSheetByName: () => aba(planilhas.central), getSheets: () => [] }
  },
  Utilities: { formatDate: d => d.toISOString().slice(0, 10), newBlob: h => ({ getAs: () => ({ setName: () => ({ getBytes: () => [1] }) }) }), base64Encode: () => 'AA==' },
  MimeType: { HTML: 'h', PDF: 'p' },
  HtmlService: { createTemplateFromFile: n => ({ evaluate() { const t = { nome: n, dados: this }; t.setTitle = () => t; t.addMetaTag = () => t; return t; } }) }
};
vm.createContext(ctx);
for (const f of ['Logo.gs', 'Relatorio.gs', 'Code.gs']) vm.runInContext(fs.readFileSync(D + '/' + f, 'utf8'), ctx, { filename: f });
vm.runInContext("CONFIG.controle.planilhaId='CTRL'; CONFIG.origens.central.planilhaId='CENT';", ctx);
const run = (quem, code) => { email = quem; try { return vm.runInContext(code, ctx); } catch (e) { return 'ERRO: ' + e.message; } };
let ok = 0, falhas = 0;
const confere = (nome, cond, info) => { cond ? ok++ : falhas++; console.log((cond ? 'OK   ' : 'FALHA') + ' ' + nome + (cond ? '' : ' -> ' + JSON.stringify(info))); };

let r;
r = run('gestor.bota@ibeu.com.br', 'getDados("filiais")');
confere('gestor de Botafogo só vê Botafogo', r.auditorias.length === 1 && r.auditorias[0].filial === 'Botafogo', r);
confere('dataInicio exclui auditoria antiga', !r.auditorias.some(a => a.ar === 'Antigo'), r.auditorias);
r = run('gestor.bota@ibeu.com.br', 'listarOrigens()');
confere('gestor de Botafogo não vê a Central', JSON.stringify(r.map(o => o.id)) === '["filiais"]', r);
r = run('gestor.bota@ibeu.com.br', 'getDados("central")');
confere('gestor de Botafogo não abre a Central', String(r).startsWith('ERRO'), r);
r = run('gestor.bota@ibeu.com.br', 'gerarPdf("filiais",[3])');
confere('PDF de outra filial (Méier) é recusado', String(r).startsWith('ERRO: Acesso'), r);
r = run('gestor.bota@ibeu.com.br', 'gerarPdf("filiais",[2])');
confere('PDF da própria filial funciona', r && r.nome && r.nome.endsWith('.pdf'), r);
r = run('gestor.bota@ibeu.com.br', 'gerarPdf("filiais",[2,999])');
confere('ID inexistente é recusado', String(r).startsWith('ERRO'), r);
r = run('gestor.bota@ibeu.com.br', 'gerarPdf("__proto__",[2])');
confere('painel inválido é recusado', String(r).startsWith('ERRO'), r);
r = run('gestor.bota@ibeu.com.br', 'gerarPdf("filiais",Array.from({length:61},(_, i)=>2))');
confere('limite de 60 por PDF', String(r).includes('Máximo'), r);
r = run('diretor@ibeu.com.br', 'getDados("filiais").auditorias.length + "/" + getDados("central").auditorias.length');
confere('diretor (Ambos/Todas) vê tudo', r === '2/1', r);
r = run('ex@ibeu.com.br', 'listarOrigens()');
confere('acesso inativo não entra', Array.isArray(r) && r.length === 0, r);
r = run('ex@ibeu.com.br', 'doGet({parameter:{}}).nome');
confere('inativo cai na página Negado', r === 'Negado', r);
r = run('naolistado@ibeu.com.br', 'getDados("filiais")');
confere('e-mail fora da lista é recusado', String(r).startsWith('ERRO'), r);
r = run('fora@gmail.com', 'getDados("filiais")');
confere('outro domínio é recusado mesmo na lista', String(r).startsWith('ERRO'), r);
r = run('', 'getDados("filiais")');
confere('usuário anônimo é recusado', String(r).startsWith('ERRO'), r);
r = run('central@ibeu.com.br', 'listarOrigens().map(o=>o.id).join()');
confere('"Central de Relacionamento" com equipes vazias = todas da Central', r === 'central', r);
r = run('gestor.bota@ibeu.com.br', 'doGet({parameter:{origem:"central"}}).dados.origemInicial');
confere('?origem=central sem permissão volta para filiais', r === 'filiais', r);
const linhaPdf = log.find(l => l[2] === 'baixou PDF');
confere('log registra download e neutraliza fórmula do AR', linhaPdf && linhaPdf[4] === "'=IMPORTXML(\"x\")", linhaPdf);
r = run('gestor.bota@ibeu.com.br', 'previaRelatorio("filiais",[2])');
confere('prévia escapa HTML do AR', typeof r === 'string' && r.includes('Relatório de Feedback') && !r.includes('<script'), String(r).slice(0, 80));
const publicas = vm.runInContext("Object.keys(this).filter(k => typeof this[k] === 'function' && !k.endsWith('_'))", ctx)
  .filter(k => !['console','JSON','Object','Array','String','Number','Math','Date','RegExp','Error','isNaN'].includes(k)).sort().join();
confere('só 5 funções públicas (google.script.run)', publicas === 'doGet,gerarPdf,getDados,listarOrigens,previaRelatorio', publicas);
r = run('', 'seguroParaCelula_(["=IMPORTXML(1)"])');
confere('log neutraliza fórmula mesmo vinda como lista', r === "'=IMPORTXML(1)", r);
console.log(`\n${ok} ok, ${falhas} falha(s)`);
process.exit(falhas ? 1 : 0);
