/**
 * ============================================================================
 *  BANCO DE HORAS — Backend (Google Apps Script)
 * ============================================================================
 *
 * O QUE ESTE FICHEIRO FAZ:
 *  - Na primeira vez que a aplicação é aberta, cria automaticamente uma
 *    Google Sheet chamada "Banco de Horas - Base de Dados" no Drive da conta
 *    que publicou a app, com todas as folhas (Users, Invited, Entries,
 *    Comments, Closed) já preparadas.
 *  - Serve a página (index.html).
 *  - Disponibiliza todas as funções que o frontend chama via
 *    google.script.run (login, criar conta, guardar horas, folgas,
 *    comentários, definições, admin, exportações, etc.)
 *  - Tem a função dailyFillTrigger(), que deve ser agendada para correr
 *    1x por dia (ver INSTRUCOES.md) e que preenche automaticamente o
 *    horário padrão de cada funcionário nos dias úteis sem registo.
 *
 * NÃO É PRECISO CRIAR NADA À MÃO NO GOOGLE SHEETS — tudo é automático.
 * ============================================================================
 */

const DAILY_MINS = 240;           // 4 horas = carga diária
const START_DATE = '2026-06-01';  // data a partir da qual tudo é contabilizado
const SPREADSHEET_NAME = 'Banco de Horas - Base de Dados';
const MAX_BLOCKS = 3;             // nº máximo de períodos (entrada/saída) por dia
const RECOVERY_TTL_MS = 10 * 60 * 1000; // 10 minutos


// ============================================================================
// SERVIR A PÁGINA WEB
// ============================================================================

function doGet(e) {
  getSpreadsheet(); // garante que a base de dados existe antes de servir a página
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Banco de Horas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================================
// CRIAÇÃO / ACESSO À SPREADSHEET
// ============================================================================

function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  let ss = null;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }
  initializeSheets(ss);
  return ss;
}

function initializeSheets(ss) {
  ensureSheet(ss, 'Users',
    ['Email', 'Nome', 'Password', 'Role', 'Active', 'DefaultEntry', 'DefaultExit'],
    [['admin@example.com', 'Administrador', 'changeme123', 'admin', true, '', '']]
  );
  ensureSheet(ss, 'Invited',
    ['Email', 'Nome'],
    [
      ['employee1@example.com', 'Funcionário 1'],
      ['employee2@example.com', 'Funcionário 2']
    ]
  );
  ensureSheet(ss, 'Entries',
    ['Email', 'Data', 'Entrada1', 'Saida1', 'Entrada2', 'Saida2', 'Entrada3', 'Saida3', 'Folga', 'MinEntrada', 'MaxSaida', 'FolgaStatus']
  );
  // Migração: projetos antigos podem não ter as colunas "MinEntrada"/"MaxSaida"/"FolgaStatus" — adiciona-as se faltarem.
  ensureColumn(ss, 'Entries', 'MinEntrada');
  ensureColumn(ss, 'Entries', 'MaxSaida');
  ensureColumn(ss, 'Entries', 'FolgaStatus');
  ensureSheet(ss, 'Comments',
    ['Email', 'Data', 'Comentario']
  );
  ensureSheet(ss, 'Closed',
    ['Nome', 'Inicio', 'Fim'],
    [
      ['Pausa letiva', '2025-08-01', '2025-08-31'],
      ['Pausa letiva', '2026-08-01', '2026-08-31']
    ]
  );

  // Remove a folha em branco criada por defeito ("Sheet1"/"Folha1"), se existir
  ['Sheet1', 'Folha1'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > 1) {
      const dr = s.getDataRange();
      if (dr.getNumRows() <= 1 && dr.getNumColumns() <= 1) {
        try { ss.deleteSheet(s); } catch (e) {}
      }
    }
  });
}

function ensureSheet(ss, name, headers, defaultRows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    if (defaultRows) defaultRows.forEach(row => sheet.appendRow(row));
    // Coluna "Data" e datas de Closed como texto, para evitar conversões automáticas
    if (name === 'Entries') { sheet.getRange('B:B').setNumberFormat('@'); sheet.getRange('C:L').setNumberFormat('@'); }
    if (name === 'Users') sheet.getRange('F:G').setNumberFormat('@');
    if (name === 'Closed') sheet.getRange('B:C').setNumberFormat('@');
  }
  return sheet;
}

// Garante que a folha `name` tem uma coluna com o cabeçalho `header`.
// Se não existir, adiciona-a no fim (como texto simples).
function ensureColumn(ss, name, header) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf(header) !== -1) return; // já existe
  const newCol = lastCol + 1;
  sheet.getRange(1, newCol).setValue(header).setFontWeight('bold');
  sheet.getRange(1, newCol, sheet.getMaxRows(), 1).setNumberFormat('@');
}


// ============================================================================
// UTILITÁRIOS DE DATA / FERIADOS / DIAS ÚTEIS
// ============================================================================

function pad2(n) { return String(n).padStart(2, '0'); }
function fmtKey(y, m, d) { return `${y}-${pad2(m)}-${pad2(d)}`; }

function easterDate(year) {
  const a = year % 19, b = Math.floor(year / 100), c = year % 100, d = Math.floor(b / 4),
        e = b % 4, f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3),
        h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4,
        l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451),
        n = h + l - 7 * m + 114;
  const month = Math.floor(n / 31), day = (n % 31) + 1;
  return new Date(year, month - 1, day);
}

function ptHolidays(year) {
  const e = easterDate(year);
  const addD = (d, n) => { const r = new Date(d); r.setDate(r.getDate() + n); return r; };
  const key = (d) => fmtKey(d.getFullYear(), d.getMonth() + 1, d.getDate());
  const map = {};
  map[fmtKey(year, 1, 1)] = 'Ano Novo';
  map[key(addD(e, -2))] = 'Sexta-Feira Santa';
  map[key(e)] = 'Páscoa';
  map[fmtKey(year, 4, 25)] = '25 de Abril';
  map[fmtKey(year, 5, 1)] = 'Dia do Trabalhador';
  map[fmtKey(year, 6, 10)] = 'Dia de Portugal';
  map[key(addD(e, 60))] = 'Corpo de Deus';
  map[fmtKey(year, 8, 15)] = 'Assunção de N. Sra.';
  map[fmtKey(year, 10, 5)] = 'Implantação da República';
  map[fmtKey(year, 11, 1)] = 'Dia de Todos os Santos';
  map[fmtKey(year, 12, 1)] = 'Restauração da Independência';
  map[fmtKey(year, 12, 8)] = 'Imaculada Conceição';
  map[fmtKey(year, 12, 25)] = 'Natal';
  map[fmtKey(year, 6, 24)] = 'Feriado Municipal'; // substitui pelo feriado local da tua cidade
  return map;
}

function getHolidaysAround(dateKey) {
  const year = parseInt(dateKey.split('-')[0], 10);
  return Object.assign({}, ptHolidays(year - 1), ptHolidays(year), ptHolidays(year + 1));
}

function isClosedDay(key, closedList) {
  return closedList.some(c => key >= c.start && key <= c.end);
}

function isWorkDay(key, closedList) {
  const d = new Date(key + 'T00:00:00');
  const dw = d.getDay();
  if (dw === 0 || dw === 6) return false;
  const hols = getHolidaysAround(key);
  if (hols[key]) return false;
  if (isClosedDay(key, closedList)) return false;
  return true;
}

function todayKey() {
  const d = new Date();
  return fmtKey(d.getFullYear(), d.getMonth() + 1, d.getDate());
}

function formatDateValue(v) {
  if (!v) return '';
  if (v instanceof Date) return fmtKey(v.getFullYear(), v.getMonth() + 1, v.getDate());
  return String(v).trim();
}

// Converte um valor de hora vindo do Sheets (pode ser Date "1899-12-30 HH:mm" se a célula
// foi auto-convertida para tipo "hora") para uma string "HH:mm". Strings já no formato
// correto passam normalmente.
function formatTimeValue(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) {
    const hh = String(v.getHours()).padStart(2, '0');
    const mm = String(v.getMinutes()).padStart(2, '0');
    return hh + ':' + mm;
  }
  return String(v).trim();
}


// ============================================================================
// LEITURA GENÉRICA DE FOLHAS
// ============================================================================

function readSheet(name) {
  const sheet = getSpreadsheet().getSheetByName(name);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { sheet, headers: data[0] || [], rows: [] };
  return { sheet, headers: data[0], rows: data.slice(1) };
}

function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

// Cache curto (30s) para folhas que mudam raramente ('Closed' e 'Users'),
// partilhado por todas as execuções do script. É invalidado de imediato em
// qualquer escrita a essas folhas, por isso nunca devolve dados desatualizados
// após uma alteração — só evita re-leituras repetidas em chamadas seguidas.
const CACHE_TTL_SECONDS = 30;
function invalidateClosedCache() { CacheService.getScriptCache().remove('closed_list'); }
function invalidateUsersCache() { CacheService.getScriptCache().remove('users_map'); }

function getClosedList() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('closed_list');
  if (cached) return JSON.parse(cached);
  const { headers, rows } = readSheet('Closed');
  const list = rows
    .map(r => rowToObj(headers, r))
    .filter(o => o.Nome)
    .map(o => ({ name: o.Nome, start: formatDateValue(o.Inicio), end: formatDateValue(o.Fim) }))
    .filter(c => c.start && c.end);
  cache.put('closed_list', JSON.stringify(list), CACHE_TTL_SECONDS);
  return list;
}

function getUsersMap() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('users_map');
  if (cached) return JSON.parse(cached);
  const { headers, rows } = readSheet('Users');
  const map = {};
  rows.forEach(r => {
    const o = rowToObj(headers, r);
    if (!o.Email) return;
    map[String(o.Email).toLowerCase().trim()] = {
      email: String(o.Email).toLowerCase().trim(),
      name: o.Nome,
      password: String(o.Password),
      role: o.Role,
      active: o.Active === true || o.Active === 'TRUE' || o.Active === 'true',
      defaultEntry: formatTimeValue(o.DefaultEntry),
      defaultExit: formatTimeValue(o.DefaultExit)
    };
  });
  cache.put('users_map', JSON.stringify(map), CACHE_TTL_SECONDS);
  return map;
}

function getInvitedMap() {
  const { headers, rows } = readSheet('Invited');
  const map = {};
  rows.forEach(r => {
    const o = rowToObj(headers, r);
    if (!o.Email) return;
    map[String(o.Email).toLowerCase().trim()] = { email: String(o.Email).toLowerCase().trim(), name: o.Nome };
  });
  return map;
}

// Lê a folha 'Entries' UMA SÓ VEZ e devolve tudo o que é derivado dela para
// este utilizador: entradas/saídas, hora mínima de entrada e hora máxima de
// saída. Usar esta função em vez de chamar getEntriesForEmail/getMinEntriesForEmail/
// getMaxExitsForEmail em separado evita ler a mesma folha 3x por utilizador
// (cada leitura da Sheet é uma chamada lenta ao Google Sheets).
function getEntryDataForEmail(email) {
  const { headers, rows } = readSheet('Entries');
  const entries = {}, minEntries = {}, maxExits = {}, folgaStatus = {};
  rows.forEach(r => {
    const o = rowToObj(headers, r);
    if (!o.Email || String(o.Email).toLowerCase().trim() !== email) return;
    const dateKey = formatDateValue(o.Data);
    if (!dateKey || dateKey < START_DATE) return; // ignora linhas residuais anteriores ao início do sistema

    const minEntry = formatTimeValue(o.MinEntrada);
    if (minEntry) minEntries[dateKey] = minEntry;
    const maxExit = formatTimeValue(o.MaxSaida);
    if (maxExit) maxExits[dateKey] = maxExit;
    const fStatus = String(o.FolgaStatus || '').trim();
    if (fStatus) folgaStatus[dateKey] = fStatus;

    if (o.Folga === true || o.Folga === 'TRUE' || o.Folga === 'true') {
      entries[dateKey] = { folga: true };
      return;
    }
    const blocks = [];
    for (let i = 1; i <= MAX_BLOCKS; i++) {
      const en = formatTimeValue(o['Entrada' + i]), ex = formatTimeValue(o['Saida' + i]);
      if (en && ex) blocks.push({ entry: en, exit: ex });
    }
    if (blocks.length === 0) return;
    entries[dateKey] = blocks.length === 1 ? blocks[0] : blocks;
  });
  return { entries, minEntries, maxExits, folgaStatus };
}

// Devolve { 'YYYY-MM-DD': 'HH:mm' } — dias em que o admin definiu uma hora mínima
// de entrada para este utilizador (o utilizador não pode escolher entrada antes disso).
function getMinEntriesForEmail(email) {
  return getEntryDataForEmail(email).minEntries;
}

// Devolve { 'YYYY-MM-DD': 'HH:mm' } — dias em que o admin definiu uma hora máxima
// de saída para este utilizador (o utilizador não pode escolher saída depois disso).
function getMaxExitsForEmail(email) {
  return getEntryDataForEmail(email).maxExits;
}

// Devolve { 'YYYY-MM-DD': {entry,exit} | [{entry,exit},...] | {folga:true} }
function getEntriesForEmail(email) {
  return getEntryDataForEmail(email).entries;
}

// Devolve { 'YYYY-MM-DD': 'texto' } — só para o próprio utilizador
function getCommentsForEmail(email) {
  const { headers, rows } = readSheet('Comments');
  const out = {};
  rows.forEach(r => {
    const o = rowToObj(headers, r);
    if (!o.Email || String(o.Email).toLowerCase().trim() !== email) return;
    const dateKey = formatDateValue(o.Data);
    if (dateKey && dateKey >= START_DATE && o.Comentario) out[dateKey] = String(o.Comentario);
  });
  return out;
}


// ============================================================================
// AUTENTICAÇÃO / CONTA
// ============================================================================

function gsPing(payload) {
  Logger.log('gsPing CHAMADO. payload=' + JSON.stringify(payload));
  return { ok: true, recebido: payload, agora: new Date().toString() };
}

function gsAuthenticate(payload) {
  const emailRaw = payload && payload.email;
  const password = payload && payload.pass;
  Logger.log('gsAuthenticate CHAMADO. payload=' + JSON.stringify(payload) + ' typeof payload=' + typeof payload + ' emailRaw=' + emailRaw + ' typeof password=' + typeof password);
  try {
    const email = String(emailRaw).toLowerCase().trim();
    const users = getUsersMap();
    Logger.log('users keys=' + JSON.stringify(Object.keys(users)));
    const user = users[email];
    if (!user || String(user.password) !== String(password)) {
      const r1 = { ok: false, error: 'Email ou palavra-passe incorretos.' };
      Logger.log('RETORNO r1=' + JSON.stringify(r1));
      return r1;
    }
    if (user.role === 'admin') {
      let adminData = null;
      try {
        adminData = getAdminBootstrap();
      } catch (e) {
        // Se o bootstrap falhar por algum motivo, o login continua válido;
        // o cliente faz fallback para gsGetAdminBootstrap (com retries).
        Logger.log('getAdminBootstrap falhou no login combinado: ' + e.message);
      }
      const r2 = { ok: true, user: publicUser(user), admin: adminData };
      Logger.log('RETORNO r2 (admin, com bootstrap=' + (adminData ? 'sim' : 'não') + ') success=' + r2.ok);
      return r2;
    }
    if (!user.active) {
      const r3 = { ok: false, error: 'Conta inativa. Contacta o administrador.' };
      Logger.log('RETORNO r3=' + JSON.stringify(r3));
      return r3;
    }
    const closed = runDailyFillForUser(user); // garante que os dias até hoje estão preenchidos
    const entryData = getEntryDataForEmail(email);
    const r4 = {
      ok: true,
      user: publicUser(user),
      entries: entryData.entries,
      comments: getCommentsForEmail(email),
      closed: closed,
      minEntries: entryData.minEntries,
      maxExits: entryData.maxExits,
      folgaStatus: entryData.folgaStatus
    };
    Logger.log('RETORNO r4 success=' + r4.ok);
    return r4;
  } catch (err) {
    const r5 = { ok: false, error: 'Erro interno no gsAuthenticate: ' + err.message };
    Logger.log('CATCH r5=' + JSON.stringify(r5) + ' stack=' + err.stack);
    return r5;
  }
}

function publicUser(u) {
  return {
    email: u.email, name: u.name, role: u.role,
    defaultEntry: u.defaultEntry, defaultExit: u.defaultExit
  };
}

function gsCheckSignupEmail(emailRaw) {
  const email = String(emailRaw).toLowerCase().trim();
  const users = getUsersMap();
  if (users[email]) return { authorized: false, error: 'Esta conta já existe. Faz login.' };
  const invited = getInvitedMap();
  if (!invited[email]) return { authorized: false, error: 'Email não autorizado. Contacta o administrador.' };
  return { authorized: true, name: invited[email].name || '' };
}

function gsCreateAccount(emailRaw, name, password, defaultEntry, defaultExit) {
  const email = String(emailRaw).toLowerCase().trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let closed;
  try {
    const invited = getInvitedMap();
    if (!invited[email]) return { success: false, error: 'Email não autorizado.' };
    const users = getUsersMap();
    if (users[email]) return { success: false, error: 'Esta conta já existe.' };

    // Adiciona à folha Users
    const usersSheet = getSpreadsheet().getSheetByName('Users');
    usersSheet.appendRow([email, name, password, 'user', true, defaultEntry, defaultExit]);
    invalidateUsersCache();

    // Remove da folha Invited
    removeRowsWhere('Invited', o => String(o.Email).toLowerCase().trim() === email);

    // Preenche automaticamente desde START_DATE até hoje com o horário padrão
    closed = fillDefaultEntriesForUser(email, defaultEntry, defaultExit);

  } finally {
    lock.releaseLock();
  }

  return {
    success: true,
    user: { email, name, role: 'user', defaultEntry, defaultExit },
    entries: getEntriesForEmail(email),
    comments: {},
    closed: closed
  };
}

// Devolve a lista de "Closed" (pausas), já lida aqui, para quem chamar
// poder reutilizá-la sem precisar de outra leitura à folha.
function fillDefaultEntriesForUser(email, defaultEntry, defaultExit) {
  const closed = getClosedList();
  if (!defaultEntry || !defaultExit) return closed;
  const existing = getEntriesForEmail(email);
  const sheet = getSpreadsheet().getSheetByName('Entries');
  const today = todayKey();
  const newRows = [];
  let d = new Date(START_DATE + 'T00:00:00');
  const end = new Date(today + 'T00:00:00');
  while (d <= end) {
    const k = fmtKey(d.getFullYear(), d.getMonth() + 1, d.getDate());
    if (isWorkDay(k, closed) && !existing[k]) {
      newRows.push([email, k, defaultEntry, defaultExit, '', '', '', '', false]);
    }
    d.setDate(d.getDate() + 1);
  }
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
  return closed;
}


// ============================================================================
// RECUPERAÇÃO DE PALAVRA-PASSE (envia email real via MailApp)
// ============================================================================

function gsSendRecoveryCode(emailRaw) {
  const email = String(emailRaw).toLowerCase().trim();
  const users = getUsersMap();
  if (!users[email]) return { success: false, error: 'Email não encontrado ou conta ainda não criada.' };

  const code = String(Math.floor(100000 + Math.random() * 900000));
  const props = PropertiesService.getScriptProperties();
  props.setProperty('recovery_' + email, JSON.stringify({ code, expires: Date.now() + RECOVERY_TTL_MS }));

  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Código de recuperação — Banco de Horas',
      body: `Olá ${users[email].name},\n\n` +
            `Recebemos um pedido para recuperar o acesso à tua conta do Banco de Horas.\n\n` +
            `O teu código de verificação é: ${code}\n\n` +
            `Este código é válido durante 10 minutos. Se não fizeste este pedido, ignora este email.`
    });
  } catch (e) {
    return { success: false, error: 'Não foi possível enviar o email: ' + e.message };
  }

  return { success: true };
}

function gsResetPassword(emailRaw, code, newPassword) {
  const email = String(emailRaw).toLowerCase().trim();
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('recovery_' + email);
  if (!raw) return { success: false, error: 'Código inválido ou expirado.' };
  const rec = JSON.parse(raw);
  if (rec.code !== String(code) || Date.now() > rec.expires) {
    return { success: false, error: 'Código inválido ou expirado.' };
  }
  if (!newPassword || newPassword.length < 4) {
    return { success: false, error: 'A palavra-passe deve ter pelo menos 4 caracteres.' };
  }
  updateRowWhere('Users', o => String(o.Email).toLowerCase().trim() === email, { Password: newPassword });
  invalidateUsersCache();
  props.deleteProperty('recovery_' + email);
  return { success: true };
}


// ============================================================================
// REGISTO DE HORAS / FOLGAS
// ============================================================================

// payload: {email, date:'YYYY-MM-DD', minEntry:'HH:mm'|'', maxExit:'HH:mm'|''}
// Define (ou remove, se string vazia) a hora mínima de entrada e/ou a hora máxima
// de saída que o admin impõe a este utilizador, neste dia. Requer que já exista
// uma linha de Entries para esse email+data (normalmente já existe, criada pelo
// preenchimento automático diário).
function gsSetEntryLock(payload) {
  try {
    const email = String(payload && payload.email).toLowerCase().trim();
    const dateKey = String(payload && payload.date).trim();
    const minEntry = formatTimeValue(payload && payload.minEntry);
    const maxExit = formatTimeValue(payload && payload.maxExit);

    const sheet = getSpreadsheet().getSheetByName('Entries');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const colIdx = {};
    headers.forEach((h, i) => colIdx[h] = i);
    if (colIdx['MinEntrada'] === undefined || colIdx['MaxSaida'] === undefined) {
      return { ok: false, error: 'Colunas MinEntrada/MaxSaida não encontradas (faz setup outra vez).' };
    }

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colIdx['Email']]).toLowerCase().trim() === email &&
          formatDateValue(data[i][colIdx['Data']]) === dateKey) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1) {
      return { ok: false, error: 'Ainda não existe um registo para este dia (só é possível trancar dias de hoje ou já passados).' };
    }

    sheet.getRange(rowIndex + 1, colIdx['MinEntrada'] + 1).setNumberFormat('@').setValue(minEntry);
    sheet.getRange(rowIndex + 1, colIdx['MaxSaida'] + 1).setNumberFormat('@').setValue(maxExit);
    return { ok: true, email: email, date: dateKey, minEntry: minEntry, maxExit: maxExit };
  } catch (err) {
    return { ok: false, error: 'Erro ao gravar: ' + err.message };
  }
}

// ============================================================================
// PEDIDOS DE FOLGA (o funcionário pede, o admin aprova/rejeita)
// ============================================================================

// Define FolgaStatus numa linha existente. Devolve {ok:false} se a linha não existir.
function setFolgaStatusCell(emailRaw, dateKey, value) {
  const email = String(emailRaw).toLowerCase().trim();
  dateKey = String(dateKey).trim();
  const sheet = getSpreadsheet().getSheetByName('Entries');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIdx = {};
  headers.forEach((h, i) => colIdx[h] = i);
  if (colIdx['FolgaStatus'] === undefined) return { ok: false, error: 'Coluna FolgaStatus não encontrada (faz setup outra vez).' };
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIdx['Email']]).toLowerCase().trim() === email &&
        formatDateValue(data[i][colIdx['Data']]) === dateKey) {
      sheet.getRange(i + 1, colIdx['FolgaStatus'] + 1).setNumberFormat('@').setValue(value);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Registo não encontrado.' };
}

// O funcionário pede para marcar um dia como folga. Fica "Pendente" até o
// admin aprovar (Folga=true) ou rejeitar (FolgaStatus='Rejeitado'). Se ainda
// não existir registo para esse dia (ex: dia futuro), cria-se uma linha nova
// só com o pedido.
function gsRequestFolga(emailRaw, dateKey) {
  const email = String(emailRaw).toLowerCase().trim();
  dateKey = String(dateKey).trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet().getSheetByName('Entries');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const colIdx = {};
    headers.forEach((h, i) => colIdx[h] = i);
    if (colIdx['FolgaStatus'] === undefined) return { ok: false, error: 'Coluna FolgaStatus não encontrada (faz setup outra vez).' };

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colIdx['Email']]).toLowerCase().trim() === email &&
          formatDateValue(data[i][colIdx['Data']]) === dateKey) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1) {
      const rowValues = new Array(headers.length).fill('');
      rowValues[colIdx['Email']] = email;
      rowValues[colIdx['Data']] = dateKey;
      rowValues[colIdx['Folga']] = false;
      rowValues[colIdx['FolgaStatus']] = 'Pendente';
      sheet.appendRow(rowValues);
      sheet.getRange(sheet.getLastRow(), colIdx['Data'] + 1).setNumberFormat('@').setValue(dateKey);
    } else {
      sheet.getRange(rowIndex + 1, colIdx['FolgaStatus'] + 1).setNumberFormat('@').setValue('Pendente');
    }
    return { ok: true, date: dateKey, status: 'Pendente' };
  } finally {
    lock.releaseLock();
  }
}

// O admin marca/desmarca um dia como folga diretamente (sem precisar de aprovação).
function gsAdminSetFolga(emailRaw, dateKey, folga) {
  const email = String(emailRaw).toLowerCase().trim();
  dateKey = String(dateKey).trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Usa upsertEntryRow para preservar colunas existentes (MinEntrada, MaxSaida, etc.)
    upsertEntryRow(email, dateKey, [], !!folga);
    // Se era folga obrigatória, limpa também qualquer pedido pendente/rejeitado
    if (folga) setFolgaStatusCell(email, dateKey, '');
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// O funcionário retira um pedido pendente, ou dispensa uma rejeição (volta ao estado normal).
function gsClearFolgaRequest(emailRaw, dateKey) {
  return setFolgaStatusCell(emailRaw, dateKey, '');
}

// O admin aprova (Folga=true, FolgaStatus='') ou rejeita (FolgaStatus='Rejeitado') um pedido.
function gsRespondFolgaRequest(emailRaw, dateKey, approve) {
  const email = String(emailRaw).toLowerCase().trim();
  dateKey = String(dateKey).trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet().getSheetByName('Entries');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const colIdx = {};
    headers.forEach((h, i) => colIdx[h] = i);
    if (colIdx['FolgaStatus'] === undefined || colIdx['Folga'] === undefined) {
      return { ok: false, error: 'Colunas necessárias não encontradas (faz setup outra vez).' };
    }
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colIdx['Email']]).toLowerCase().trim() === email &&
          formatDateValue(data[i][colIdx['Data']]) === dateKey) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1) return { ok: false, error: 'Pedido não encontrado.' };

    if (approve) {
      sheet.getRange(rowIndex + 1, colIdx['Folga'] + 1).setValue(true);
      sheet.getRange(rowIndex + 1, colIdx['FolgaStatus'] + 1).setNumberFormat('@').setValue('');
    } else {
      sheet.getRange(rowIndex + 1, colIdx['FolgaStatus'] + 1).setNumberFormat('@').setValue('Rejeitado');
    }
    return { ok: true, date: dateKey, approved: !!approve };
  } finally {
    lock.releaseLock();
  }
}

// blocks: [{entry,exit}, ...] (até MAX_BLOCKS) ; folga: boolean
function gsSaveEntry(emailRaw, dateKey, blocks, folga) {
  const email = String(emailRaw).toLowerCase().trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    upsertEntryRow(email, dateKey, blocks, folga);
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}

function gsSaveBulkEntries(emailRaw, dateKeys, blocks, folga) {
  const email = String(emailRaw).toLowerCase().trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    dateKeys.forEach(k => upsertEntryRow(email, k, blocks, folga));
  } finally {
    lock.releaseLock();
  }
  return { success: true, count: dateKeys.length };
}

function gsDeleteEntry(emailRaw, dateKey) {
  const email = String(emailRaw).toLowerCase().trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    removeRowsWhere('Entries', o =>
      String(o.Email).toLowerCase().trim() === email && formatDateValue(o.Data) === dateKey
    );
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}

// Cria ou atualiza a linha de Entries para email+data
function upsertEntryRow(email, dateKey, blocks, folga) {
  const sheet = getSpreadsheet().getSheetByName('Entries');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIdx = {};
  headers.forEach((h, i) => colIdx[h] = i);

  let rowIndex = -1; // índice 0-based dentro de `data` (inclui header em 0)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIdx['Email']]).toLowerCase().trim() === email &&
        formatDateValue(data[i][colIdx['Data']]) === dateKey) {
      rowIndex = i;
      break;
    }
  }

  const rowValues = new Array(headers.length).fill('');
  rowValues[colIdx['Email']] = email;
  rowValues[colIdx['Data']] = dateKey;
  rowValues[colIdx['Folga']] = !!folga;
  if (!folga && blocks && blocks.length) {
    blocks.slice(0, MAX_BLOCKS).forEach((b, i) => {
      rowValues[colIdx['Entrada' + (i + 1)]] = b.entry;
      rowValues[colIdx['Saida' + (i + 1)]] = b.exit;
    });
  }
  // Preserva a "hora mínima de entrada" / "hora máxima de saída" (definidas pelo admin)
  // e o estado de pedido de folga, ao reescrever a linha.
  if (rowIndex !== -1) {
    if (colIdx['MinEntrada'] !== undefined) rowValues[colIdx['MinEntrada']] = formatTimeValue(data[rowIndex][colIdx['MinEntrada']]);
    if (colIdx['MaxSaida'] !== undefined) rowValues[colIdx['MaxSaida']] = formatTimeValue(data[rowIndex][colIdx['MaxSaida']]);
    if (colIdx['FolgaStatus'] !== undefined) rowValues[colIdx['FolgaStatus']] = String(data[rowIndex][colIdx['FolgaStatus']] || '').trim();
  }

  if (rowIndex === -1) {
    sheet.appendRow(rowValues);
    // Garante que a coluna Data fica como texto (evita conversão automática)
    sheet.getRange(sheet.getLastRow(), colIdx['Data'] + 1).setNumberFormat('@').setValue(dateKey);
  } else {
    sheet.getRange(rowIndex + 1, 1, 1, rowValues.length).setValues([rowValues]);
    sheet.getRange(rowIndex + 1, colIdx['Data'] + 1).setNumberFormat('@').setValue(dateKey);
  }
}


// ============================================================================
// COMENTÁRIOS (privados — só o próprio utilizador)
// ============================================================================

function gsSaveComment(emailRaw, dateKey, text) {
  const email = String(emailRaw).toLowerCase().trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!text || !text.trim()) {
      removeRowsWhere('Comments', o =>
        String(o.Email).toLowerCase().trim() === email && formatDateValue(o.Data) === dateKey
      );
    } else {
      upsertCommentRow(email, dateKey, text.trim());
    }
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}

function gsDeleteComment(emailRaw, dateKey) {
  return gsSaveComment(emailRaw, dateKey, '');
}

function upsertCommentRow(email, dateKey, text) {
  const sheet = getSpreadsheet().getSheetByName('Comments');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === email && formatDateValue(data[i][1]) === dateKey) {
      sheet.getRange(i + 1, 3).setValue(text);
      return;
    }
  }
  sheet.appendRow([email, dateKey, text]);
  sheet.getRange(sheet.getLastRow(), 2).setNumberFormat('@').setValue(dateKey);
}


// ============================================================================
// DEFINIÇÕES DO UTILIZADOR
// ============================================================================

// Atualiza o horário padrão (aplica-se só a dias futuros — dias passados não são tocados)
// e, opcionalmente, a palavra-passe.
function gsSaveSettings(emailRaw, defaultEntry, defaultExit, currentPass, newPass, newName) {
  const email = String(emailRaw).toLowerCase().trim();
  const users = getUsersMap();
  const user = users[email];
  if (!user) return { success: false, error: 'Utilizador não encontrado.' };

  if (newPass) {
    if (String(currentPass) !== String(user.password)) {
      return { success: false, error: 'Palavra-passe atual incorreta.' };
    }
    if (newPass.length < 4) {
      return { success: false, error: 'A nova palavra-passe deve ter pelo menos 4 caracteres.' };
    }
  }

  const finalName = (newName && String(newName).trim()) ? String(newName).trim() : user.name;
  const updates = { Nome: finalName, DefaultEntry: defaultEntry, DefaultExit: defaultExit };
  if (newPass) updates.Password = newPass;
  updateRowWhere('Users', o => String(o.Email).toLowerCase().trim() === email, updates);
  invalidateUsersCache();

  return { success: true, name: finalName };
}


// ============================================================================
// ADMIN — alterar o próprio email de acesso (login + recuperação de password)
// O email "admin@empresa.pt" usado por defeito é fictício; o admin deve
// trocá-lo por um email real para conseguir receber o código de recuperação.
// ============================================================================
function gsUpdateAdminEmail(currentEmailRaw, newEmailRaw, currentPass, newPass, newName) {
  const currentEmail = String(currentEmailRaw).toLowerCase().trim();
  const newEmail = String(newEmailRaw).toLowerCase().trim();

  const users = getUsersMap();
  const user = users[currentEmail];
  if (!user) return { success: false, error: 'Utilizador não encontrado.' };
  if (user.role !== 'admin') return { success: false, error: 'Apenas o administrador pode alterar o email de acesso.' };

  if (!newEmail || newEmail.indexOf('@') < 0) {
    return { success: false, error: 'Introduz um email válido.' };
  }
  const finalName = (newName && String(newName).trim()) ? String(newName).trim() : user.name;
  const emailChanging = newEmail !== currentEmail;
  if (emailChanging || newPass) {
    if (String(currentPass) !== String(user.password)) {
      return { success: false, error: 'Palavra-passe atual incorreta.' };
    }
  }
  if (newPass && newPass.length < 4) {
    return { success: false, error: 'A nova palavra-passe deve ter pelo menos 4 caracteres.' };
  }

  if (newEmail !== currentEmail) {
    const invited = getInvitedMap();
    if (users[newEmail] || invited[newEmail]) {
      return { success: false, error: 'Esse email já está a ser usado por outra conta.' };
    }
  }

  const updates = { Email: newEmail, Nome: finalName };
  if (newPass) updates.Password = newPass;
  updateRowWhere('Users', o => String(o.Email).toLowerCase().trim() === currentEmail, updates);
  invalidateUsersCache();

  return { success: true, email: newEmail, name: finalName };
}


// ============================================================================
// TRIGGER DIÁRIO — preenche automaticamente o horário padrão
// ============================================================================

// Corre 1x por dia (agendar em Acionadores / Triggers — ver INSTRUCOES.md).
// Também é chamada no login de cada utilizador, para garantir que está
// sempre atualizado mesmo que o trigger ainda não tenha corrido.
function dailyFillTrigger() {
  const users = getUsersMap();
  Object.values(users).forEach(u => {
    if (u.role !== 'admin' && u.active) runDailyFillForUser(u);
  });
}

function runDailyFillForUser(user) {
  return fillDefaultEntriesForUser(user.email, user.defaultEntry, user.defaultExit);
}


// ============================================================================
// ADMINISTRAÇÃO
// ============================================================================

function getAdminBootstrap() {
  const users = getUsersMap();
  const employees = Object.values(users).filter(u => u.role !== 'admin' && u.active);
  const entriesByEmail = {};
  const minEntriesByEmail = {};
  const maxExitsByEmail = {};
  const commentsByEmail = {};
  const folgaStatusByEmail = {};
  const pendingRequests = [];
  employees.forEach(u => {
    const entryData = getEntryDataForEmail(u.email);
    entriesByEmail[u.email] = entryData.entries;
    minEntriesByEmail[u.email] = entryData.minEntries;
    maxExitsByEmail[u.email] = entryData.maxExits;
    commentsByEmail[u.email] = getCommentsForEmail(u.email);
    folgaStatusByEmail[u.email] = entryData.folgaStatus;
    Object.keys(entryData.folgaStatus).forEach(dateKey => {
      if (entryData.folgaStatus[dateKey] === 'Pendente') {
        pendingRequests.push({ email: u.email, name: u.name, date: dateKey });
      }
    });
  });
  pendingRequests.sort((a, b) => a.date < b.date ? -1 : (a.date > b.date ? 1 : 0));

  return {
    employees: employees.map(u => ({
      email: u.email, name: u.name, defaultEntry: u.defaultEntry, defaultExit: u.defaultExit
    })),
    entriesByEmail,
    minEntriesByEmail,
    maxExitsByEmail,
    commentsByEmail,
    folgaStatusByEmail,
    pendingRequests,
    closed: getClosedList(),
    invited: Object.values(getInvitedMap()),
    users: Object.values(users).filter(u => u.role !== 'admin').map(u => ({
      email: u.email, name: u.name, active: u.active
    }))
  };
}

function gsGetAdminBootstrap() {
  Logger.log('gsGetAdminBootstrap CHAMADO');
  try {
    const admin = getAdminBootstrap();
    Logger.log('gsGetAdminBootstrap OK, employees=' + (admin.employees ? admin.employees.length : 'n/a'));
    return { ok: true, admin: admin };
  } catch (err) {
    Logger.log('gsGetAdminBootstrap ERRO: ' + err.message + ' stack=' + err.stack);
    return { ok: false, error: 'Erro ao carregar dados de admin: ' + err.message };
  }
}

function gsAddInvite(name, emailRaw) {
  const email = String(emailRaw).toLowerCase().trim();
  const users = getUsersMap(), invited = getInvitedMap();
  if (users[email] || invited[email]) return { success: false, error: 'Email já existe.' };
  getSpreadsheet().getSheetByName('Invited').appendRow([email, name]);
  return { success: true };
}

function gsRemoveUser(emailRaw) {
  const email = String(emailRaw).toLowerCase().trim();
  removeRowsWhere('Users', o => String(o.Email).toLowerCase().trim() === email);
  invalidateUsersCache();
  return { success: true };
}

// Verifica a password de um utilizador (ex: confirmar identidade do admin antes de ações sensíveis).
function gsVerifyPassword(emailRaw, password) {
  const email = String(emailRaw).toLowerCase().trim();
  const users = getUsersMap();
  const user = users[email];
  if (!user) return { ok: false, error: 'Utilizador não encontrado.' };
  if (String(password) !== String(user.password)) return { ok: false, error: 'Password incorreta.' };
  return { ok: true };
}

function gsRemoveInvite(emailRaw) {
  const email = String(emailRaw).toLowerCase().trim();
  removeRowsWhere('Invited', o => String(o.Email).toLowerCase().trim() === email);
  return { success: true };
}

function gsAddClosedPeriod(name, start, end) {
  if (!name || !start || !end || end < start) return { success: false, error: 'Dados inválidos.' };
  const sheet = getSpreadsheet().getSheetByName('Closed');
  sheet.appendRow([name, start, end]);
  const row = sheet.getLastRow();
  sheet.getRange(row, 2, 1, 2).setNumberFormat('@').setValues([[start, end]]);
  invalidateClosedCache();
  return { success: true, closed: getClosedList() };
}

function gsRemoveClosedPeriod(index) {
  const sheet = getSpreadsheet().getSheetByName('Closed');
  const rowNum = index + 2; // +1 header, +1 para 1-based
  if (rowNum >= 2 && rowNum <= sheet.getLastRow()) sheet.deleteRow(rowNum);
  invalidateClosedCache();
  return { success: true, closed: getClosedList() };
}


// ============================================================================
// EXPORTAÇÃO (CSV / TXT) — sem comentários, por respeito à privacidade
// ============================================================================

const MONTH_NAMES_PT = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

function gsExportData(userSel, period) {
  const users = getUsersMap();
  const targets = userSel === 'all'
    ? Object.values(users).filter(u => u.role !== 'admin' && u.active)
    : [users[String(userSel).toLowerCase().trim()]].filter(Boolean);

  const closed = getClosedList();
  const now = new Date();
  const cm = now.getMonth(), cy = now.getFullYear();

  // Parsear período — pode ser múltiplos separados por vírgula
  const periods = String(period).split(',').map(p => p.trim()).filter(Boolean);

  // Função que decide se uma dateKey (YYYY-MM-DD) está incluída num período
  function inPeriod(k, p) {
    const [y, m] = k.split('-').map(Number);
    if (p === 'all') return true;
    if (p === 'current') return y === cy && m - 1 === cm;
    if (p === 'prev') {
      const pm = cm === 0 ? 11 : cm - 1, py = cm === 0 ? cy - 1 : cy;
      return y === py && m - 1 === pm;
    }
    if (p.startsWith('month:')) {
      const parts = p.replace('month:', '').split('-');
      return y === parseInt(parts[0]) && m - 1 === parseInt(parts[1]) - 1;
    }
    if (p.startsWith('week:')) {
      const wStart = p.replace('week:', '');
      const wEnd = (function() {
        const d = new Date(wStart + 'T00:00:00'); d.setDate(d.getDate() + 4);
        return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
      })();
      return k >= wStart && k <= wEnd;
    }
    return false;
  }

  function includeRow(k) { return periods.some(p => inPeriod(k, p)); }

  // Rótulo e tag do período para o nome do ficheiro
  let periodLabel, periodFileTag;
  if (periods.length === 1) {
    const p = periods[0];
    if (p === 'all') { periodLabel = 'Tudo'; periodFileTag = 'tudo'; }
    else if (p === 'current') { periodLabel = MONTH_NAMES_PT[cm]+' '+cy; periodFileTag = cy+'-'+String(cm+1).padStart(2,'0'); }
    else if (p === 'prev') {
      const pm = cm===0?11:cm-1, py = cm===0?cy-1:cy;
      periodLabel = MONTH_NAMES_PT[pm]+' '+py; periodFileTag = py+'-'+String(pm+1).padStart(2,'0');
    } else if (p.startsWith('month:')) {
      const parts = p.replace('month:','').split('-');
      const my = parseInt(parts[0]), mm2 = parseInt(parts[1])-1;
      periodLabel = MONTH_NAMES_PT[mm2]+' '+my; periodFileTag = parts[0]+'-'+parts[1];
    } else if (p.startsWith('week:')) {
      const ws = p.replace('week:','');
      const d = new Date(ws+'T00:00:00'), de = new Date(d); de.setDate(d.getDate()+4);
      const fmt = x => String(x.getDate()).padStart(2,'0')+'/'+String(x.getMonth()+1).padStart(2,'0');
      periodLabel = 'Semana '+fmt(d)+' – '+fmt(de)+'/'+de.getFullYear(); periodFileTag = 'semana_'+ws;
    } else { periodLabel = p; periodFileTag = p; }
  } else {
    periodLabel = periods.length + ' períodos selecionados';
    periodFileTag = 'multiplo_' + new Date().toISOString().slice(0,10);
  }

  let txt = `RELATÓRIO BANCO DE HORAS\nGerado em: ${now.toLocaleDateString('pt-PT')}\nPeríodo: ${periodLabel}\n${'='.repeat(40)}\n\n`;
  const usersData = [];

  targets.forEach(u => {
    const entries = getEntriesForEmail(u.email);
    let totalMins = 0, periodMins = 0;
    let cOver = 0, cExact = 0, cUnder = 0, cFolga = 0, cTotalDays = 0;
    const dayRows = [];

    Object.keys(entries).sort().forEach(k => {
      if (!isWorkDay(k, closed)) return;
      const { mins, isFolga } = entryMinsAndFolga(entries[k]);
      const delta = isFolga ? -DAILY_MINS : mins - DAILY_MINS;
      totalMins += delta;
      if (includeRow(k)) {
        const [yy, mm, dd] = k.split('-');
        dayRows.push({ date: `${dd}/${mm}/${yy}`, dateKey: k, mins, isFolga, delta });
        periodMins += delta;
        cTotalDays++;
        if (isFolga) cFolga++;
        else if (delta > 0) cOver++;
        else if (delta < 0) cUnder++;
        else cExact++;
      }
    });

    const pct = n => cTotalDays > 0 ? Math.round((n / cTotalDays) * 100) : 0;
    txt += `${u.name.toUpperCase()}\n${'-'.repeat(30)}\n` +
      `Email:                 ${u.email}\n` +
      `Dias úteis no período: ${cTotalDays}\n` +
      `  - A mais:            ${cOver} (${pct(cOver)}%)\n` +
      `  - Exato:             ${cExact} (${pct(cExact)}%)\n` +
      `  - A menos:           ${cUnder} (${pct(cUnder)}%)\n` +
      `  - Folga:             ${cFolga} (${pct(cFolga)}%)\n` +
      `Saldo do período:      ${fmtDelta(periodMins)}\n` +
      `Saldo total acumulado: ${fmtDelta(totalMins)} = ${(totalMins/DAILY_MINS).toFixed(2)} dias\n\n`;

    usersData.push({ user: u, dayRows, totalMins, periodMins, cOver, cExact, cUnder, cFolga, cTotalDays });
  });

  return { txt, periodLabel, periodFileTag, usersData };
}

// Gera um PDF estilizado (via HTML→Google Doc) e devolve base64.
function gsExportPDF(userSel, period) {
  const data = gsExportData(userSel, period);

  // Construir HTML do relatório
  const now = new Date();
  const WEEK_PT = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];

  let usersHtml = data.usersData.map(ud => {
    const u = ud.user;
    const pct = n => ud.cTotalDays > 0 ? Math.round((n / ud.cTotalDays) * 100) : 0;
    const saldoColor = ud.periodMins >= 0 ? '#16A34A' : '#DC2626';
    const totalColor = ud.totalMins >= 0 ? '#16A34A' : '#DC2626';

    // Tabela de dias
    const dayRowsHtml = ud.dayRows.map(r => {
      const d = new Date(r.dateKey + 'T00:00:00');
      const dow = WEEK_PT[d.getDay()];
      const rowColor = r.isFolga ? '#EEF4FF' : (r.delta > 0 ? '#F0FDF4' : (r.delta < 0 ? '#FEF2F2' : '#fff'));
      const saldoStr = `<span style="color:${r.delta >= 0 ? '#16A34A' : '#DC2626'};font-weight:500">${fmtDelta(r.delta)}</span>`;
      return `<tr style="background:${rowColor}">
        <td style="padding:5px 8px;font-size:11px;color:#555">${dow}</td>
        <td style="padding:5px 8px;font-size:11px">${r.date}</td>
        <td style="padding:5px 8px;font-size:11px;text-align:center">${r.isFolga ? '🏖️ Folga' : fmtMins(r.mins)}</td>
        <td style="padding:5px 8px;font-size:11px;text-align:right">${saldoStr}</td>
      </tr>`;
    }).join('');

    return `
<div style="margin-bottom:28px;page-break-inside:avoid">
  <div style="background:#185FA5;color:#fff;padding:12px 16px;border-radius:8px 8px 0 0">
    <div style="font-size:15px;font-weight:700">${u.name}</div>
    <div style="font-size:11px;opacity:0.8;margin-top:2px">${u.email}</div>
  </div>
  <table style="width:100%;border-collapse:collapse;border:1px solid #e5e5e5;border-top:none">
    <thead><tr style="background:#f0f4f8">
      <th style="padding:6px 8px;font-size:11px;text-align:left;color:#185FA5">Dia</th>
      <th style="padding:6px 8px;font-size:11px;text-align:left;color:#185FA5">Data</th>
      <th style="padding:6px 8px;font-size:11px;text-align:center;color:#185FA5">Horas</th>
      <th style="padding:6px 8px;font-size:11px;text-align:right;color:#185FA5">Saldo</th>
    </tr></thead>
    <tbody>${dayRowsHtml || '<tr><td colspan="4" style="padding:10px;text-align:center;color:#999;font-size:11px">Sem registos no período</td></tr>'}</tbody>
  </table>
  <div style="background:#f9f9f7;border:1px solid #e5e5e5;border-top:none;padding:10px 16px;border-radius:0 0 8px 8px;display:flex;gap:24px;flex-wrap:wrap">
    <div style="font-size:11px;color:#555">✅ Exato: <strong>${ud.cExact}</strong> (${pct(ud.cExact)}%)</div>
    <div style="font-size:11px;color:#555">🟢 A mais: <strong>${ud.cOver}</strong> (${pct(ud.cOver)}%)</div>
    <div style="font-size:11px;color:#555">🔴 A menos: <strong>${ud.cUnder}</strong> (${pct(ud.cUnder)}%)</div>
    <div style="font-size:11px;color:#555">🏖️ Folga: <strong>${ud.cFolga}</strong> (${pct(ud.cFolga)}%)</div>
    <div style="font-size:11px;color:#555">Saldo período: <strong style="color:${saldoColor}">${fmtDelta(ud.periodMins)}</strong></div>
    <div style="font-size:11px;color:#555">Saldo total: <strong style="color:${totalColor}">${fmtDelta(ud.totalMins)}</strong> (${(ud.totalMins/DAILY_MINS).toFixed(2)} dias)</div>
  </div>
</div>`;
  }).join('');

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  body{font-family:Arial,sans-serif;margin:32px;color:#1a1a1a;}
  h1{font-size:20px;color:#185FA5;margin:0 0 4px}
  .subtitle{font-size:12px;color:#76746C;margin:0 0 20px}
  .divider{border:none;border-top:1px solid #e5e5e5;margin:16px 0}
</style></head><body>
<h1>Relatório Banco de Horas</h1>
<p class="subtitle">Período: ${data.periodLabel} &nbsp;·&nbsp; Gerado em: ${now.toLocaleDateString('pt-PT')}</p>
<hr class="divider">
${usersHtml}
</body></html>`;

  // Criar ficheiro HTML temporário e exportar como PDF
  const blob = Utilities.newBlob(html, 'text/html', '_bh_rpt.html');
  const file = DriveApp.createFile(blob);
  const pdfBlob = file.getAs('application/pdf');
  const base64 = Utilities.base64Encode(pdfBlob.getBytes());
  file.setTrashed(true);
  return { ok: true, base64: base64, filename: 'relatorio_' + data.periodFileTag + '.pdf' };
}

// ============================================================================
// RESUMO MENSAL POR EMAIL
// ============================================================================

const MONTHLY_SUMMARY_HANDLER = 'sendMonthlySummaries';

// Estatísticas de um utilizador para um mês específico (0-indexed) + saldo total acumulado.
function computeMonthStats(email, year, month) {
  const entries = getEntriesForEmail(email);
  const closed = getClosedList();
  let totalMins = 0, periodMins = 0;
  let cOver = 0, cExact = 0, cUnder = 0, cFolga = 0, cTotalDays = 0;
  Object.keys(entries).sort().forEach(k => {
    if (!isWorkDay(k, closed)) return;
    const { mins, isFolga } = entryMinsAndFolga(entries[k]);
    const delta = isFolga ? -DAILY_MINS : mins - DAILY_MINS;
    totalMins += delta;
    const [y, m] = k.split('-').map(Number);
    if (y === year && (m - 1) === month) {
      periodMins += delta;
      cTotalDays++;
      if (isFolga) cFolga++;
      else if (delta > 0) cOver++;
      else if (delta < 0) cUnder++;
      else cExact++;
    }
  });
  return { totalMins, periodMins, cOver, cExact, cUnder, cFolga, cTotalDays };
}

function buildSummaryBlock(user, s, monthLabel) {
  const pct = n => s.cTotalDays > 0 ? Math.round((n / s.cTotalDays) * 100) : 0;
  // Versão texto simples (usada no PDF e no corpo plain-text)
  return `${user.name.toUpperCase()}\n${'-'.repeat(30)}\n` +
    `Email:                  ${user.email}\n` +
    `Dias úteis em ${monthLabel}: ${s.cTotalDays}\n` +
    `  - A mais:             ${s.cOver} (${pct(s.cOver)}%)\n` +
    `  - Exato:              ${s.cExact} (${pct(s.cExact)}%)\n` +
    `  - A menos:            ${s.cUnder} (${pct(s.cUnder)}%)\n` +
    `  - Folga:              ${s.cFolga} (${pct(s.cFolga)}%)\n` +
    `Saldo de ${monthLabel}:  ${fmtDelta(s.periodMins)}\n` +
    `Saldo total acumulado:  ${fmtDelta(s.totalMins)} (${(s.totalMins / DAILY_MINS).toFixed(2)} dias)\n`;
}

// Versão HTML do resumo de um utilizador — usada no corpo do email.
function buildSummaryHtml(user, s, monthLabel) {
  const pct = n => s.cTotalDays > 0 ? Math.round((n / s.cTotalDays) * 100) : 0;
  const saldoColor = s.periodMins >= 0 ? '#16A34A' : '#DC2626';
  const totalColor = s.totalMins >= 0 ? '#16A34A' : '#DC2626';
  const row = (label, val, color) =>
    `<tr><td style="padding:6px 12px;color:#555;font-size:13px">${label}</td><td style="padding:6px 12px;font-weight:600;font-size:13px;color:${color||'#1a1a1a'}">${val}</td></tr>`;
  return `
<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;background:#f9f9f7;padding:24px;border-radius:10px">
  <div style="background:#185FA5;border-radius:8px;padding:20px 24px;margin-bottom:20px">
    <h1 style="color:#fff;margin:0;font-size:18px;font-weight:600">Banco de Horas</h1>
    <p style="color:#c8ddf5;margin:4px 0 0;font-size:13px">Resumo de ${monthLabel}</p>
  </div>
  <p style="font-size:14px;color:#333;margin:0 0 16px">Olá <strong>${user.name}</strong>, aqui está o teu resumo do mês anterior.</p>
  <div style="background:#fff;border-radius:8px;border:1px solid #e5e5e5;overflow:hidden;margin-bottom:16px">
    <div style="background:#f0f4f8;padding:10px 12px;font-size:12px;font-weight:600;color:#185FA5;text-transform:uppercase;letter-spacing:0.5px">
      Dias úteis em ${monthLabel} — ${s.cTotalDays} dias
    </div>
    <table style="width:100%;border-collapse:collapse">
      ${row('🟢 A mais (saldo positivo)', `${s.cOver} dias (${pct(s.cOver)}%)`)}
      ${row('✅ Exato', `${s.cExact} dias (${pct(s.cExact)}%)`)}
      ${row('🔴 A menos (saldo negativo)', `${s.cUnder} dias (${pct(s.cUnder)}%)`)}
      ${row('🏖️ Folga', `${s.cFolga} dias (${pct(s.cFolga)}%)`)}
    </table>
  </div>
  <div style="background:#fff;border-radius:8px;border:1px solid #e5e5e5;overflow:hidden;margin-bottom:20px">
    <div style="background:#f0f4f8;padding:10px 12px;font-size:12px;font-weight:600;color:#185FA5;text-transform:uppercase;letter-spacing:0.5px">Saldo</div>
    <table style="width:100%;border-collapse:collapse">
      ${row(`Saldo de ${monthLabel}`, fmtDelta(s.periodMins), saldoColor)}
      ${row('Saldo total acumulado', `${fmtDelta(s.totalMins)} (${(s.totalMins/DAILY_MINS).toFixed(2)} dias)`, totalColor)}
    </table>
  </div>
  <p style="font-size:11px;color:#999;margin:0;text-align:center">Email automático do Banco de Horas. Não respondas a este email.</p>
</div>`;
}

// Versão HTML do resumo de todos (para o admin)
function buildAdminSummaryHtml(blocks, monthLabel) {
  return `
<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;background:#f9f9f7;padding:24px;border-radius:10px">
  <div style="background:#185FA5;border-radius:8px;padding:20px 24px;margin-bottom:20px">
    <h1 style="color:#fff;margin:0;font-size:18px;font-weight:600">Banco de Horas — Resumo de todos</h1>
    <p style="color:#c8ddf5;margin:4px 0 0;font-size:13px">${monthLabel}</p>
  </div>
  ${blocks}
  <p style="font-size:11px;color:#999;margin:16px 0 0;text-align:center">Email automático do Banco de Horas. Não respondas a este email.</p>
</div>`;
}

// Gera um PDF do resumo mensal de UM utilizador (ou todos, se adminBody passado).
function buildSummaryPDF(content, title) {
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>body{font-family:Arial,sans-serif;margin:32px;color:#1a1a1a;}
h1{font-size:18px;color:#185FA5;text-align:center;margin-bottom:4px;}
pre{font-family:'Courier New',monospace;font-size:10px;white-space:pre-wrap;line-height:1.5;}
hr{border:none;border-top:1px solid #e5e5e5;margin:12px 0;}</style></head>
<body><h1>${title}</h1><hr><pre>${content.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</pre></body></html>`;
  const blob = Utilities.newBlob(html, 'text/html', '_bh_sum.html');
  const file = DriveApp.createFile(blob);
  const pdfBlob = file.getAs('application/pdf').setName(title.replace(/[^a-zA-Z0-9_\- ]/g, '') + '.pdf');
  file.setTrashed(true);
  return pdfBlob;
}

// Calcula e envia, por email, o resumo do MÊS ANTERIOR a cada funcionário
// (o seu próprio resumo) e a cada admin (resumo de todos). Pensada para
// correr via trigger no dia 1 de cada mês.
function sendMonthlySummaries(overrideYear, overrideMonth) {
  const users = getUsersMap();
  const employees = Object.values(users).filter(u => u.role !== 'admin' && u.active);
  const admins = Object.values(users).filter(u => u.role === 'admin');

  const now = new Date();
  let month, year;
  if (overrideYear !== undefined && overrideMonth !== undefined) {
    year = overrideYear;
    month = overrideMonth; // 0-indexed
  } else {
    // Automático: mês anterior
    month = now.getMonth() - 1;
    year = now.getFullYear();
    if (month < 0) { month = 11; year--; }
  }
  const monthLabel = MONTH_NAMES_PT[month] + ' ' + year;

  const allBlocks = [];
  const allHtmlBlocks = [];
  employees.forEach(u => {
    const s = computeMonthStats(u.email, year, month);
    const block = buildSummaryBlock(u, s, monthLabel);
    const htmlBlock = buildSummaryHtml(u, s, monthLabel);
    allBlocks.push(block);
    allHtmlBlocks.push(buildSummaryHtml(u, s, monthLabel));
    try {
      const pdfTitle = `Banco de Horas — ${u.name} — ${monthLabel}`;
      const pdfBlob = buildSummaryPDF(block, pdfTitle);
      MailApp.sendEmail({
        to: u.email,
        subject: `Resumo de ${monthLabel} — Banco de Horas`,
        htmlBody: htmlBlock,
        attachments: [pdfBlob]
      });
    } catch (e) {
      Logger.log('Erro ao enviar resumo a ' + u.email + ': ' + e.message);
    }
  });

  if (admins.length && allBlocks.length) {
    const adminContent = `RESUMO MENSAL DE TODOS — ${monthLabel}\n${'='.repeat(40)}\n\n` + allBlocks.join('\n');
    const adminHtml = buildAdminSummaryHtml(allHtmlBlocks.join('<hr style="border:none;border-top:1px solid #e5e5e5;margin:16px 0">'), monthLabel);
    const pdfTitle = `Banco de Horas — Resumo de Todos — ${monthLabel}`;
    let pdfBlob = null;
    try { pdfBlob = buildSummaryPDF(adminContent, pdfTitle); } catch(e) { Logger.log('Erro PDF admin: ' + e.message); }
    admins.forEach(a => {
      try {
        const mailOpts = {
          to: a.email,
          subject: `Resumo de todos — ${monthLabel} — Banco de Horas`,
          htmlBody: adminHtml
        };
        if (pdfBlob) mailOpts.attachments = [pdfBlob];
        MailApp.sendEmail(mailOpts);
      } catch (e) {
        Logger.log('Erro ao enviar resumo ao admin ' + a.email + ': ' + e.message);
      }
    });
  }

  return { success: true, monthLabel: monthLabel, employees: employees.length };
}

// Ativa/desativa/altera o trigger de envio automático de resumos por email.
// frequency: 'none' | 'weekly' | 'monthly' | 'yearly'
function gsSetSummaryFrequency(frequency) {
  // Remove triggers anteriores
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === MONTHLY_SUMMARY_HANDLER) ScriptApp.deleteTrigger(t);
  });
  if (frequency === 'weekly') {
    ScriptApp.newTrigger(MONTHLY_SUMMARY_HANDLER).timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  } else if (frequency === 'monthly') {
    ScriptApp.newTrigger(MONTHLY_SUMMARY_HANDLER).timeBased()
      .onMonthDay(1).atHour(7).create();
  } else if (frequency === 'yearly') {
    ScriptApp.newTrigger(MONTHLY_SUMMARY_HANDLER).timeBased()
      .onMonthDay(1).atHour(7).inMonth(ScriptApp.Month.JANUARY).create();
  }
  // Guarda a preferência nas PropertiesService para poder relê-la
  PropertiesService.getScriptProperties().setProperty('summary_frequency', frequency || 'none');
  return { success: true, frequency: frequency || 'none' };
}

function gsGetSummaryFrequency() {
  const stored = PropertiesService.getScriptProperties().getProperty('summary_frequency') || 'none';
  // Verifica também se o trigger ainda existe (pode ter sido apagado manualmente)
  const hastrigger = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === MONTHLY_SUMMARY_HANDLER);
  return { frequency: hastrigger ? stored : 'none' };
}

// Mantidos por compatibilidade (podem ainda ser chamados se houver sessões abertas com código antigo)
function gsSetupMonthlyEmailSummary() { return gsSetSummaryFrequency('monthly'); }
function gsDisableMonthlyEmailSummary() { return gsSetSummaryFrequency('none'); }
function gsIsMonthlyEmailSummaryActive() { const r = gsGetSummaryFrequency(); return { active: r.frequency !== 'none' }; }

// Envia já (para testar), com o mês que o admin escolheu.
// yearMonth: string 'YYYY-M' (ex: '2026-5') ou null para usar mês anterior.
function gsSendMonthlySummariesNow(yearMonth) {
  try {
    if (yearMonth) {
      const parts = String(yearMonth).split('-');
      const year = parseInt(parts[0]);
      const month = parseInt(parts[1]); // 0-indexed
      return sendMonthlySummaries(year, month);
    }
    return sendMonthlySummaries();
  } catch (e) {
    return { success: false, error: 'Erro ao enviar: ' + e.message };
  }
}

function entryMinsAndFolga(entry) {
  if (!entry) return { mins: 0, isFolga: false };
  if (entry.folga) return { mins: 0, isFolga: true };
  const blocks = Array.isArray(entry) ? entry : [entry];
  const mins = blocks.reduce((s, b) => s + (toMins(b.exit) - toMins(b.entry)), 0);
  return { mins, isFolga: false };
}
function toMins(t) { const [h, m] = t.split(':').map(Number); return h * 60 + m; }
function fmtMins(m) { const h = Math.floor(m / 60), mn = m % 60; return mn > 0 ? `${h}h ${mn}m` : `${h}h`; }
function fmtDelta(m) { const a = Math.abs(m), h = Math.floor(a / 60), mn = a % 60, s = m >= 0 ? '+' : '-'; return mn > 0 ? `${s}${h}h ${mn}m` : `${s}${h}h`; }


// ============================================================================
// HELPERS GENÉRICOS DE FOLHA (procurar / atualizar / remover linhas)
// ============================================================================

// Remove todas as linhas em que predicate(objLinha) === true
function removeRowsWhere(sheetName, predicate) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = data.length - 1; i >= 1; i--) {
    if (predicate(rowToObj(headers, data[i]))) sheet.deleteRow(i + 1);
  }
}

// Atualiza a 1ª linha em que predicate(objLinha) === true, com os campos de `updates`
function updateRowWhere(sheetName, predicate, updates) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIdx = {};
  headers.forEach((h, i) => colIdx[h] = i);
  for (let i = 1; i < data.length; i++) {
    if (predicate(rowToObj(headers, data[i]))) {
      Object.keys(updates).forEach(key => {
        if (colIdx[key] !== undefined) {
          sheet.getRange(i + 1, colIdx[key] + 1).setValue(updates[key]);
        }
      });
      return true;
    }
  }
  return false;
}


// ============================================================================
// FUNÇÃO DE INSTALAÇÃO DO TRIGGER DIÁRIO
// Executa esta função UMA VEZ a partir do editor do Apps Script
// (menu "Executar" → escolher "setupDailyTrigger") para agendar o
// preenchimento automático todos os dias à 1h da manhã.
// ============================================================================

function setupDailyTrigger() {
  // Remove triggers antigos desta função, para não duplicar
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyFillTrigger') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyFillTrigger')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
  return 'Trigger diário instalado com sucesso (corre todos os dias por volta da 1h).';
}


// ============================================================================
// FUNÇÃO DE CONVENIÊNCIA — corre as duas coisas de uma vez:
//  1) cria a Base de Dados (Google Sheet) se ainda não existir
//  2) instala o trigger diário
// Seleciona "setup" no menu de funções (topo do editor) e clica em ▶ Executar.
// ============================================================================

function setup() {
  getSpreadsheet();           // cria a Sheet "Banco de Horas - Base de Dados" se necessário
  const msg = setupDailyTrigger();
  return 'Base de dados pronta! ' + msg;
}
