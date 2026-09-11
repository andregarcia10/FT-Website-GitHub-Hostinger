const CONTATOS_COLA = Object.freeze({
  spreadsheetId: '1HHQSeR3SbTAZ7ol3Mc81zxchlgn3ZAP1c6aNVCUDDGI',
  abaNome: 'Cadastros',
  timezone: 'America/Sao_Paulo',
  versaoPadraoConsentimento: 'cola-lgpd-v2-2026-09-11'
});

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.nome && !p.email && !p.telefone) {
    return ContentService.createTextOutput('Endpoint de cadastro da Cola Eleitoral ativo.');
  }
  return processarCadastro_(p);
}

function doPost(e) {
  return processarCadastro_((e && e.parameter) || {});
}

function processarCadastro_(p) {
  const token = texto_(p.token).replace(/[^A-Za-z0-9_-]/g, '').slice(0, 100);

  try {
    // Honeypot simples contra envios automatizados triviais.
    if (texto_(p.website)) return resposta_(false, 'requisicao_invalida', token);

    const nome = texto_(p.nome);
    const email = texto_(p.email).toLowerCase();
    const telefone = texto_(p.telefone);
    const consentimento = String(p.consentimento).toLowerCase() === 'true';
    const versao = texto_(p.versaoConsentimento) || CONTATOS_COLA.versaoPadraoConsentimento;
    const origem = texto_(p.origem) || 'site-cola-eleitoral';
    const acao = texto_(p.acao) || 'acao-nao-informada';

    if (!nomeValido_(nome) || !emailValido_(email) || !telefoneValido_(telefone) || !consentimento) {
      return resposta_(false, 'dados_invalidos', token);
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      // O frontend envia por dois caminhos para aumentar a confiabilidade.
      // Este cache impede duas linhas para o mesmo clique.
      if (token) {
        const cache = CacheService.getScriptCache();
        if (cache.get('cola_token_' + token)) {
          return resposta_(true, 'duplicado_ignorado', token);
        }
        cache.put('cola_token_' + token, '1', 600);
      }

      const ss = SpreadsheetApp.openById(CONTATOS_COLA.spreadsheetId);
      const sh = obterAba_(ss);
      sh.appendRow([
        seguroPlanilha_(nome),
        seguroPlanilha_(email),
        seguroPlanilha_(telefone),
        Utilities.formatDate(new Date(), CONTATOS_COLA.timezone, 'dd/MM/yyyy HH:mm:ss'),
        'SIM',
        seguroPlanilha_(versao),
        seguroPlanilha_(origem),
        seguroPlanilha_(acao)
      ]);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    return resposta_(true, '', token);
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return resposta_(false, 'erro_interno', token);
  }
}

function obterAba_(ss) {
  let sh = ss.getSheetByName(CONTATOS_COLA.abaNome);
  if (!sh) sh = ss.insertSheet(CONTATOS_COLA.abaNome);

  const cabecalho = [[
    'Nome completo', 'E-mail', 'Telefone / WhatsApp', 'Data e hora',
    'Consentimento LGPD', 'Versão do consentimento', 'Origem', 'Ação liberada'
  ]];
  const atual = sh.getRange(1, 1, 1, 8).getDisplayValues()[0];
  if (atual.join('|') !== cabecalho[0].join('|')) {
    sh.getRange(1, 1, 1, 8).setValues(cabecalho);
    sh.setFrozenRows(1);
  }
  return sh;
}

function nomeValido_(v) {
  return texto_(v).split(/\s+/).filter(Boolean).length >= 2;
}
function emailValido_(v) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(texto_(v));
}
function telefoneValido_(v) {
  const d = texto_(v).replace(/\D/g, '');
  return d.length >= 10 && d.length <= 13;
}
function texto_(v) {
  return String(v == null ? '' : v).trim();
}
function seguroPlanilha_(v) {
  const s = texto_(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function resposta_(ok, error, token) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: !!ok, error: error || '', token: token || '' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Execute manualmente UMA VEZ no editor para confirmar que o script tem
// permissão de escrita na planilha. Depois apague a linha de teste se quiser.
function testarGravacao() {
  return processarCadastro_({
    nome: 'Teste de integração',
    email: 'teste@exemplo.com',
    telefone: '21999999999',
    consentimento: 'true',
    versaoConsentimento: CONTATOS_COLA.versaoPadraoConsentimento,
    origem: 'teste-manual-apps-script',
    acao: 'teste',
    token: 'teste_' + Date.now(),
    website: ''
  });
}
