const CONTATOS_COLA = Object.freeze({
  spreadsheetId: '1HHQSeR3SbTAZ7ol3Mc81zxchlgn3ZAP1c6aNVCUDDGI',
  abaNome: 'Cadastros',
  timezone: 'America/Sao_Paulo',
  versaoPadraoConsentimento: 'cola-lgpd-v2-2026-09-11'
});

function doGet() {
  return HtmlService.createHtmlOutput('Endpoint de cadastro da Cola Eleitoral ativo.');
}

function doPost(e) {
  const p = (e && e.parameter) || {};
  const token = texto_(p.token).replace(/[^A-Za-z0-9_-]/g, '').slice(0, 100);

  try {
    // Honeypot simples contra envios automatizados triviais.
    if (texto_(p.website)) return respostaIframe_(token, false, 'requisicao_invalida');

    const nome = texto_(p.nome);
    const email = texto_(p.email).toLowerCase();
    const telefone = texto_(p.telefone);
    const consentimento = String(p.consentimento).toLowerCase() === 'true';
    const versao = texto_(p.versaoConsentimento) || CONTATOS_COLA.versaoPadraoConsentimento;
    const origem = texto_(p.origem) || 'site-cola-eleitoral';
    const acao = texto_(p.acao) || 'acao-nao-informada';

    if (!nomeValido_(nome) || !emailValido_(email) || !telefoneValido_(telefone) || !consentimento) {
      return respostaIframe_(token, false, 'dados_invalidos');
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const ss = SpreadsheetApp.openById(CONTATOS_COLA.spreadsheetId);
      const sh = obterAba_(ss);
      sh.appendRow([
        seguroPlanilha_(nome),
        seguroPlanilha_(email),
        seguroPlanilha_(telefone),
        Utilities.formatDate(new Date(), CONTATOS_COLA.timezone, "dd/MM/yyyy HH:mm:ss"),
        'SIM',
        seguroPlanilha_(versao),
        seguroPlanilha_(origem),
        seguroPlanilha_(acao)
      ]);
    } finally {
      lock.releaseLock();
    }

    return respostaIframe_(token, true, '');
  } catch (err) {
    console.error(err);
    return respostaIframe_(token, false, 'erro_interno');
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
function respostaIframe_(token, ok, error) {
  const payload = JSON.stringify({
    type: 'cola-eleitoral-registro',
    token: token,
    ok: !!ok,
    error: error || ''
  }).replace(/</g, '\\u003c');

  return HtmlService.createHtmlOutput(
    '<!doctype html><meta charset="utf-8"><script>' +
    'try{parent.postMessage(' + payload + ',"*");}catch(e){}' +
    '<\\/script>'
  );
}
