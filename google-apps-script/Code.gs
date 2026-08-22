/**
 * AGENDA PÚBLICA — FABIANO TROMPETISTA 13007
 *
 * 1) Execute criarEstruturaAgendaFabiano() uma única vez.
 * 2) O script criará uma pasta e uma planilha no Google Drive.
 * 3) Depois, implante este projeto como Web App:
 *    - Executar como: você
 *    - Quem tem acesso: qualquer pessoa
 * 4) Copie a URL terminada em /exec e cole no atributo
 *    data-agenda-api-url da seção #agenda no index.html.
 */

const AGENDA_CONFIG = Object.freeze({
  pastaNome: 'Agenda – Fabiano 13007',
  planilhaNome: 'Agenda Pública – Fabiano 13007',
  abaNome: 'Agenda',
  propriedadePlanilhaId: 'AGENDA_SPREADSHEET_ID',
  cacheSegundos: 300,
  timezone: 'America/Sao_Paulo'
});

function criarEstruturaAgendaFabiano() {
  const props = PropertiesService.getScriptProperties();
  const existente = props.getProperty(AGENDA_CONFIG.propriedadePlanilhaId);

  if (existente) {
    try {
      const ssExistente = SpreadsheetApp.openById(existente);
      Logger.log('A estrutura já existe: ' + ssExistente.getUrl());
      return {
        planilhaId: existente,
        planilhaUrl: ssExistente.getUrl()
      };
    } catch (erro) {
      props.deleteProperty(AGENDA_CONFIG.propriedadePlanilhaId);
    }
  }

  const pasta = DriveApp.createFolder(AGENDA_CONFIG.pastaNome);
  const ss = SpreadsheetApp.create(AGENDA_CONFIG.planilhaNome);
  const arquivo = DriveApp.getFileById(ss.getId());

  pasta.addFile(arquivo);
  DriveApp.getRootFolder().removeFile(arquivo);

  const sheet = ss.getSheets()[0];
  sheet.setName(AGENDA_CONFIG.abaNome);

  const cabecalhos = [
    'Exibir',
    'Data',
    'Dia da semana',
    'Horário',
    'Evento',
    'Local',
    'Tipo',
    'Recorrência',
    'Observações'
  ];

  const linhas = [
    ['Sim', new Date(2026, 7, 18), 'Terça-feira', '18:30', 'Lançamento da Campanha para Deputada Federal de Rosilene Corrêa', 'Sinpro – SIG Q.6, 2260', 'Evento', '', ''],
    ['Sim', new Date(2026, 7, 19), 'Quarta-feira', '19:00', 'Lançamento Oficial da Campanha Agnelo Queiroz – Deputado Federal', 'Teatro dos Bancários – EQS 314/315 Sul', 'Evento', '', ''],
    ['Sim', new Date(2026, 8, 6), 'Quarta-feira', 'Manhã', 'Feirinha de adoção com a Candidata a Deputada Federal Vanessa Bicho', 'Eixão do Lazer – Asa Norte', 'Evento', '', ''],
    ['Sim', '', '', '10:30', 'Banca na Feira do MDA', 'Esplanada dos Ministérios', 'Agenda fixa', 'Todas as terças', ''],
    ['Sim', '', '', '8:30', 'Banca na Feira do Pontão Norte', '', 'Agenda fixa', 'Todos os sábados', '']
  ];

  sheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
  sheet.getRange(2, 1, linhas.length, cabecalhos.length).setValues(linhas);

  sheet.setFrozenRows(1);
  sheet.getRange('A1:I1')
    .setBackground('#531414')
    .setFontColor('#F4E3C8')
    .setFontWeight('bold');

  sheet.getRange('B2:B').setNumberFormat('dd/MM/yyyy');
  sheet.getRange('A2:A').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Sim', 'Não'], true)
      .setAllowInvalid(false)
      .build()
  );
  sheet.getRange('G2:G').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Evento', 'Agenda fixa'], true)
      .setAllowInvalid(false)
      .build()
  );

  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 95);
  sheet.setColumnWidth(5, 430);
  sheet.setColumnWidth(6, 320);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 180);
  sheet.setColumnWidth(9, 320);

  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 30), cabecalhos.length)
    .setVerticalAlignment('middle');
  sheet.getRange(2, 5, sheet.getMaxRows() - 1, 5).setWrap(true);

  const instrucoes = ss.insertSheet('Instruções');
  instrucoes.getRange('A1:A8').setValues([
    ['COMO ATUALIZAR A AGENDA DO SITE'],
    ['1. Edite somente a aba "Agenda".'],
    ['2. Use "Sim" na coluna Exibir para publicar o compromisso; use "Não" para ocultá-lo.'],
    ['3. Para compromisso com data, preencha Data, Dia da semana, Horário, Evento e Local.'],
    ['4. Para agenda recorrente, deixe Data em branco, escolha "Agenda fixa" e preencha Recorrência.'],
    ['5. Compromissos com data anterior ao dia atual deixam de ser publicados automaticamente.'],
    ['6. O site busca a agenda automaticamente quando é acessado. O serviço usa cache curto de até 5 minutos.'],
    ['7. Não altere os nomes das colunas da aba "Agenda".']
  ]);
  instrucoes.getRange('A1').setFontWeight('bold').setFontSize(14).setBackground('#DC1413').setFontColor('#F4E3C8');
  instrucoes.setColumnWidth(1, 760);
  instrucoes.getRange('A1:A8').setWrap(true);

  props.setProperty(AGENDA_CONFIG.propriedadePlanilhaId, ss.getId());
  CacheService.getScriptCache().remove('agenda_publica_json');

  Logger.log('Pasta criada: ' + AGENDA_CONFIG.pastaNome);
  Logger.log('Planilha: ' + ss.getUrl());

  return {
    pastaId: pasta.getId(),
    planilhaId: ss.getId(),
    planilhaUrl: ss.getUrl()
  };
}

function doGet(e) {
  try {
    const callback = sanitizarCallback_(e && e.parameter ? e.parameter.callback : '');
    const payload = obterAgendaPublica_();
    return responder_(payload, callback);
  } catch (erro) {
    const callback = sanitizarCallback_(e && e.parameter ? e.parameter.callback : '');
    return responder_({
      ok: false,
      erro: 'Não foi possível carregar a agenda.',
      detalhe: String(erro && erro.message ? erro.message : erro)
    }, callback);
  }
}

function obterAgendaPublica_() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'agenda_publica_json';
  const emCache = cache.get(cacheKey);

  if (emCache) {
    return JSON.parse(emCache);
  }

  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty(AGENDA_CONFIG.propriedadePlanilhaId);

  if (!spreadsheetId) {
    throw new Error('A planilha ainda não foi criada. Execute criarEstruturaAgendaFabiano().');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(AGENDA_CONFIG.abaNome);

  if (!sheet) {
    throw new Error('A aba "' + AGENDA_CONFIG.abaNome + '" não foi encontrada.');
  }

  const valores = sheet.getDataRange().getValues();
  if (valores.length < 2) {
    return {
      ok: true,
      ultimaAtualizacao: formatarAtualizacao_(spreadsheetId),
      itens: []
    };
  }

  const cabecalhos = valores[0].map(String);
  const idx = {};
  cabecalhos.forEach((nome, i) => idx[nome] = i);

  const obrigatorios = ['Exibir', 'Data', 'Dia da semana', 'Horário', 'Evento', 'Local', 'Tipo', 'Recorrência', 'Observações'];
  obrigatorios.forEach(nome => {
    if (idx[nome] === undefined) {
      throw new Error('Coluna obrigatória ausente: ' + nome);
    }
  });

  const hoje = inicioDoDia_(new Date());

  const itens = valores.slice(1)
    .map((linha, posicao) => mapearLinha_(linha, idx, posicao))
    .filter(item => item)
    .filter(item => item.fixa || !item.dataObj || item.dataObj >= hoje)
    .sort(ordenarItens_)
    .map(({ dataObj, ordemOriginal, ...publico }) => publico);

  const payload = {
    ok: true,
    ultimaAtualizacao: formatarAtualizacao_(spreadsheetId),
    itens
  };

  cache.put(cacheKey, JSON.stringify(payload), AGENDA_CONFIG.cacheSegundos);
  return payload;
}

function mapearLinha_(linha, idx, posicao) {
  const exibir = normalizarTexto_(linha[idx['Exibir']]).toLowerCase();
  if (!['sim', 's', 'yes', '1', 'true'].includes(exibir)) return null;

  const evento = normalizarTexto_(linha[idx['Evento']]);
  if (!evento) return null;

  const tipo = normalizarTexto_(linha[idx['Tipo']]);
  const fixa = tipo.toLowerCase() === 'agenda fixa';
  const dataObj = normalizarData_(linha[idx['Data']]);

  let data = '';
  let dia = '';
  let mes = '';

  if (dataObj) {
    data = Utilities.formatDate(dataObj, AGENDA_CONFIG.timezone, 'yyyy-MM-dd');
    dia = Utilities.formatDate(dataObj, AGENDA_CONFIG.timezone, 'dd');
    mes = mesCurto_(Number(Utilities.formatDate(dataObj, AGENDA_CONFIG.timezone, 'M')));
  }

  return {
    data,
    dia,
    mes,
    diaSemana: normalizarTexto_(linha[idx['Dia da semana']]),
    horario: normalizarTexto_(linha[idx['Horário']]),
    evento,
    local: normalizarTexto_(linha[idx['Local']]),
    fixa,
    recorrencia: normalizarTexto_(linha[idx['Recorrência']]),
    observacoes: normalizarTexto_(linha[idx['Observações']]),
    dataObj,
    ordemOriginal: posicao
  };
}

function ordenarItens_(a, b) {
  if (a.fixa !== b.fixa) return a.fixa ? 1 : -1;

  if (!a.fixa && !b.fixa) {
    const ta = a.dataObj ? a.dataObj.getTime() : Number.MAX_SAFE_INTEGER;
    const tb = b.dataObj ? b.dataObj.getTime() : Number.MAX_SAFE_INTEGER;
    if (ta !== tb) return ta - tb;
  }

  return a.ordemOriginal - b.ordemOriginal;
}

function normalizarData_(valor) {
  if (!valor) return null;

  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return inicioDoDia_(valor);
  }

  const texto = String(valor).trim();
  const br = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (br) {
    return inicioDoDia_(new Date(Number(br[3]), Number(br[2]) - 1, Number(br[1])));
  }

  const iso = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    return inicioDoDia_(new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])));
  }

  return null;
}

function inicioDoDia_(data) {
  return new Date(data.getFullYear(), data.getMonth(), data.getDate());
}

function normalizarTexto_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim();
}

function mesCurto_(mes) {
  return ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ'][mes - 1] || '';
}

function formatarAtualizacao_(spreadsheetId) {
  const data = DriveApp.getFileById(spreadsheetId).getLastUpdated();
  return Utilities.formatDate(data, AGENDA_CONFIG.timezone, "dd/MM/yyyy 'às' HH:mm");
}

function sanitizarCallback_(callback) {
  if (!callback) return '';
  const valor = String(callback).trim();
  return /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(valor) ? valor : '';
}

function responder_(payload, callback) {
  const json = JSON.stringify(payload);

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Use esta função somente se quiser forçar a atualização imediata
 * após editar a planilha, sem aguardar os até 5 minutos de cache.
 */
function limparCacheAgenda() {
  CacheService.getScriptCache().remove('agenda_publica_json');
  Logger.log('Cache da agenda limpo.');
}
