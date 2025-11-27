const SHEET_TAREFAS = 'Tarefas';
const SHEET_INFRACOES = 'Infracoes';
const SHEET_CONFIG = 'Config';

function doGet(e) {
  ensureSetup_();
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Controle de Infrações');
}

function ensureSetup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let tarefasSheet = ss.getSheetByName(SHEET_TAREFAS);
  if (!tarefasSheet) {
    tarefasSheet = ss.insertSheet(SHEET_TAREFAS);
    tarefasSheet.appendRow(['ID', 'Nome', 'Responsavel', 'Ativa']);
  } else {
    const headers = tarefasSheet.getRange(1, 1, 1, 4).getValues()[0];
    if (headers.join('') === '') {
      tarefasSheet.getRange(1, 1, 1, 4).setValues([
        ['ID', 'Nome', 'Responsavel', 'Ativa'],
      ]);
    }
  }

  let infracoesSheet = ss.getSheetByName(SHEET_INFRACOES);
  if (!infracoesSheet) {
    infracoesSheet = ss.insertSheet(SHEET_INFRACOES);
    infracoesSheet.appendRow([
      'ID',
      'TarefaID',
      'NomeTarefa',
      'Responsavel',
      'RegistradoPor',
      'DataInfracao',
      'DataRegistro',
      'DentroPrazo',
      'ContaPonto',
      'Observacao',
    ]);
  } else {
    const headers = infracoesSheet.getRange(1, 1, 1, 10).getValues()[0];
    if (headers.join('') === '') {
      infracoesSheet.getRange(1, 1, 1, 10).setValues([
        [
          'ID',
          'TarefaID',
          'NomeTarefa',
          'Responsavel',
          'RegistradoPor',
          'DataInfracao',
          'DataRegistro',
          'DentroPrazo',
          'ContaPonto',
          'Observacao',
        ],
      ]);
    }
  }

  let configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEET_CONFIG);
    configSheet.appendRow(['chave', 'valor']);
    configSheet.appendRow(['dias_limite_registro', 1]);
  } else {
    const headers = configSheet.getRange(1, 1, 1, 2).getValues()[0];
    if (headers.join('') === '') {
      configSheet.getRange(1, 1, 1, 2).setValues([['chave', 'valor']]);
    }
    const lastRow = configSheet.getLastRow();
    const keys = configSheet.getRange(2, 1, Math.max(lastRow - 1, 1), 1).getValues().flat();
    if (!keys.includes('dias_limite_registro')) {
      configSheet.appendRow(['dias_limite_registro', 1]);
    }
  }
}

function getConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) {
    ensureSetup_();
  }
  const lastRow = sheet.getLastRow();
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];
  const config = {};
  values.forEach((row) => {
    const [key, value] = row;
    if (key) {
      config[key] = value;
    }
  });
  if (config.dias_limite_registro === undefined || config.dias_limite_registro === '') {
    config.dias_limite_registro = 1;
  }
  return config;
}

function getTarefas() {
  ensureSetup_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TAREFAS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return data.map((row) => ({
    id: row[0] != null ? row[0].toString() : '',
    nome: row[1],
    responsavel: row[2],
    ativa: row[3] === true || row[3] === 'TRUE' || row[3] === 'true',
  }));
}

function saveTarefa(tarefa) {
  ensureSetup_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TAREFAS);
  const lastRow = sheet.getLastRow();
  const ativa = tarefa.ativa === true || tarefa.ativa === 'true' || tarefa.ativa === 'Sim';

  if (tarefa.id) {
    const range = sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 1).getValues();
    const rowIndex = range.findIndex((r) => (r[0] != null ? r[0].toString() : '') === tarefa.id.toString());
    if (rowIndex !== -1) {
      const rowNumber = rowIndex + 2;
      sheet.getRange(rowNumber, 1, 1, 4).setValues([
        [tarefa.id, tarefa.nome, tarefa.responsavel, ativa],
      ]);
      return { success: true, message: 'Tarefa atualizada com sucesso.' };
    }
  }

  const newId = new Date().getTime().toString();
  sheet.appendRow([newId, tarefa.nome, tarefa.responsavel, ativa]);
  return { success: true, message: 'Tarefa criada com sucesso.' };
}

function deleteTarefa(id) {
  ensureSetup_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TAREFAS);
  const lastRow = sheet.getLastRow();
  const ids = sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 1).getValues();
  const index = ids.findIndex((r) => (r[0] != null ? r[0].toString() : '') === id.toString());
  if (index !== -1) {
    sheet.deleteRow(index + 2);
    return { success: true, message: 'Tarefa excluída.' };
  }
  return { success: false, message: 'Tarefa não encontrada.' };
}

function registerInfracao(data) {
  ensureSetup_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tarefasSheet = ss.getSheetByName(SHEET_TAREFAS);
  const infracoesSheet = ss.getSheetByName(SHEET_INFRACOES);

  const tarefas = getTarefas();
  const tarefa = tarefas.find((t) => t.id === data.tarefaId);
  if (!tarefa) {
    return { success: false, message: 'Tarefa não encontrada.' };
  }

  const config = getConfig_();
  const diasLimite = Number(config.dias_limite_registro) || 1;
  const tz = Session.getScriptTimeZone();
  const dataRegistro = new Date();
  const dataRegistroDia = new Date(Utilities.formatDate(dataRegistro, tz, 'yyyy-MM-dd'));
  const dataInfracaoDate = new Date(`${data.dataInfracao}T00:00:00`);
  dataInfracaoDate.setHours(0, 0, 0, 0);

  const dataLimite = new Date(dataInfracaoDate);
  dataLimite.setDate(dataLimite.getDate() + diasLimite);
  dataLimite.setHours(0, 0, 0, 0);

  const dentroPrazo = dataRegistroDia.getTime() <= dataLimite.getTime();
  const contaPonto = dentroPrazo;

  const id = new Date().getTime().toString();
  const dataRegistroStr = Utilities.formatDate(dataRegistro, tz, 'yyyy-MM-dd');

  infracoesSheet.appendRow([
    id,
    tarefa.id,
    tarefa.nome,
    tarefa.responsavel,
    data.registradoPor,
    data.dataInfracao,
    dataRegistroStr,
    dentroPrazo,
    contaPonto,
    data.observacao || '',
  ]);

  return {
    success: true,
    message: dentroPrazo
      ? 'Infração registrada e contou ponto negativo.'
      : 'Infração registrada, mas fora do prazo (não contou ponto).',
  };
}

function getInfracoes() {
  ensureSetup_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INFRACOES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const list = data.map((row) => ({
    id: row[0],
    tarefaId: row[1],
    nomeTarefa: row[2],
    responsavel: row[3],
    registradoPor: row[4],
    dataInfracao: row[5],
    dataRegistro: row[6],
    dentroPrazo: row[7] === true || row[7] === 'TRUE' || row[7] === 'true',
    contaPonto: row[8] === true || row[8] === 'TRUE' || row[8] === 'true',
    observacao: row[9] || '',
  }));
  return list.sort((a, b) => Number(b.id) - Number(a.id));
}
