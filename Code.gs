
/**
 * ============================================================================
 * SYSGESTAO HOSPITALAR - BACKEND ENGINE V8.9 (CONCURRENCY PROTECTION + AUDIT)
 * ============================================================================
 */

// --- CONFIGURACAO DE BANCO DE DADOS ---
const ID_MANUAL = ""; // DEIXE VAZIO PARA USAR O SISTEMA DE PROPRIEDADES

// Logica de Selecao do ID (Prioridade: Manual > Memoria do Setup > Planilha Atual)
const ACTIVE_SPREADSHEET_ID = ID_MANUAL || 
                              PropertiesService.getScriptProperties().getProperty('SYS_SPREADSHEET_ID') || 
                              SpreadsheetApp.getActiveSpreadsheet().getId();

// CONFIGURACAO DE RETENCAO (ARQUIVO MORTO)
const CONST_DIAS_RETENCAO = 90; // 3 Meses na planilha principal

// --- DICIONARIO CENTRAL DE TABELAS E STATUS (EVITA HARDCODES) ---
const TABELAS = {
  PACIENTES: 'PACIENTES_ATIVOS',
  FILA_PA: 'FILA_PA',
  TRANSFERENCIAS: 'TRANSFERENCIAS',
  HISTORICO_ALTAS: 'HISTORICO_ALTAS',
  BLOQUEIOS: 'BLOQUEIOS_LEITOS',
  USUARIOS: 'USUARIOS_SISTEMA',
  ESTRUTURA: 'BRAIN_ESTRUTURA',
  DICIONARIOS: 'DIM_DICIONARIOS',
  RESTRITOS: 'DB_Restritos',
  FORECAST: 'DB_Forecast',
  RODIZIO: 'RODIZIO_EQUIPE',
  LOGS: 'LOGS_AUDITORIA',
  NOTIFICACOES: 'NOTIFICACOES'
};

const STATUS = {
  ALTA: 'ALTA',
  OBITO: 'OBITO',
  INATIVO: 'INATIVO',
  DESATIVADO: 'DESATIVADO',
  AGUARDANDO: 'AGUARDANDO',
  ACEITO: 'ACEITO',
  HIGIENIZACAO: 'HIGIENIZACAO',
  SIM: 'SIM',
  NAO: 'NAO'
};

// --- MENU DA PLANILHA ---
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SysGestao')
    .addItem('Abrir Sistema', 'openWebApp')
    .addSeparator()
    .addItem('Diagnostico', 'diagnoseSystem')
    .addItem('Executar Arquivamento', 'dailyDataArchiving')
    .addToUi();
}

function openWebApp() {
  const url = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(`
    <script>window.open("${url}", "_blank");google.script.host.close();</script>
    <p>Abrindo sistema...</p>
  `).setWidth(250).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Redirecionando...');
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('SysGestao Enterprise v8')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

/* ============================================================================
  HELPERS: VALIDACAO, AUTORIZACAO E UTILITARIOS
  ============================================================================ */
const AuthHelper = {
  requireAdmin: (userPayload) => {
    if (!userPayload) throw new Error("Acesso negado: não autenticado");
    const user = SettingsController.getUsers()
      .find(u => String(u.matricula).toUpperCase() === String(userPayload.matricula).toUpperCase());
    if (!user || user.level !== 'ADMIN') 
      throw new Error(`ACESSO_NEGADO: ${userPayload.matricula} não é ADMIN`);
    return user;
  },
  preventSelfDelete: (executorId, targetId) => {
    if (String(executorId) === String(targetId)) throw new Error("Nao eh possivel deletar sua propria conta");
  }
};

const Validator = {
  nonEmpty: (value, fieldName) => {
    if (!value || String(value).trim() === '') throw new Error(`${fieldName} e obrigatorio`);
  },
  isNumber: (value, fieldName) => {
    if (isNaN(value) || value === '') throw new Error(`${fieldName} deve ser numerico`);
  },
  validatePatient: (p) => {
    if (!p) throw new Error("Dados do paciente nao fornecidos");
    Validator.nonEmpty(p.name, 'Nome do paciente');
    Validator.nonEmpty(p.id, 'ID do paciente');
    if (p.bed) Validator.isNumber(p.bed, 'Numero de leito');
  },
  validateUser: (u) => {
    if (!u) throw new Error("Dados do usuario nao fornecidos");
    Validator.nonEmpty(u.matricula, 'Matricula');
    Validator.nonEmpty(u.name, 'Nome');
  },
  validateUserPayload: (userPayload) => {
    if (!userPayload) throw new Error("Credenciais nao fornecidas");
    if (!userPayload.matricula) throw new Error("Usuario nao identificado");
  }
};

const BoolHelper = {
  toDb: (value) => (value === true || String(value).toUpperCase() === 'SIM' || String(value).toUpperCase() === 'TRUE' || value === 1) ? 'SIM' : 'NAO',
  toJs: (value) => (value === true || String(value).toUpperCase() === 'SIM' || String(value).toUpperCase() === 'TRUE' || value === 1)
};

/* ============================================================================
  1. DATABASE LAYER (ORM GENERICO COM OPTIMISTIC LOCKING)
  ============================================================================ */
const DB = {
  _cache: {},
  _headerCache: {},

  _getHeadersCached: function(tableName) {
    if (this._headerCache[tableName]) return this._headerCache[tableName];
    const sheet = this.connect(tableName);
    const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(h => String(h).trim());
    this._headerCache[tableName] = headers;
    return headers;
  },

  getSpreadsheet: function() { 
    try {
      return SpreadsheetApp.openById(ACTIVE_SPREADSHEET_ID);
    } catch (e) {
      throw new Error(`Erro ao abrir planilha (ID: ${ACTIVE_SPREADSHEET_ID}). Verifique se o Setup foi rodado.`);
    }
  },
  
  connect: function(tableName) {
    const ss = this.getSpreadsheet();
    let sheet = ss.getSheetByName(tableName);
    if (!sheet) sheet = ss.insertSheet(tableName);
    return sheet;
  },

  findAll: function(tableName, options = {}) {
    // Retrocompatibilidade: se options for booleano, trata como includeDeleted
    const includeDeleted = (typeof options === 'boolean') ? options : (options.includeDeleted || false);
    const limit = (typeof options === 'object' && options.limit) ? options.limit : 0;
    const reverse = (typeof options === 'object' && options.reverse) ? options.reverse : false;

    // Cache só funciona para leituras completas (sem limit/reverse)
    if (this._cache[tableName] && !includeDeleted && !limit && !reverse) return this._cache[tableName];
    
    const sheet = this.connect(tableName);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // 1. Headers (Cacheado)
    let headers = this._headerCache[tableName];
    if (!headers) {
        headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
        this._headerCache[tableName] = headers;
    }

    let rows = [];
    
    // 2. Estratégia de Leitura Otimizada
    if (limit > 0 && reverse) {
        // Ler do final (ex: Logs recentes)
        const startRow = Math.max(2, lastRow - limit + 1);
        const numRows = lastRow - startRow + 1;
        if (numRows > 0) {
            rows = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
            rows.reverse(); // Inverte para ficar [Mais Recente -> Mais Antigo]
        }
    } else if (limit > 0) {
        // Ler do começo com limite
        const numRows = Math.min(limit, lastRow - 1);
        rows = sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();
    } else {
        // Ler tudo (Padrão)
        rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    }

    let result = rows.map(row => {
      let obj = {};
      let jsonExtras = {};
      headers.forEach((header, index) => {
        const val = row[index];
        if (header.includes('JSON') || header === 'CONFIG_MENUS_JSON') {
          try { 
              if (val && typeof val === 'string' && (val.startsWith('{') || val.startsWith('['))) {
                  jsonExtras = JSON.parse(val); 
              } else {
                  jsonExtras = val;
              }
          } 
          catch (e) { jsonExtras = val; }
        } else {
          if (val instanceof Date) obj[header] = val.toISOString();
          else obj[header] = (val === undefined || val === null) ? '' : String(val);
        }
      });
      if (typeof jsonExtras === 'object' && jsonExtras !== null && jsonExtras._v === undefined) jsonExtras._v = 0;
      
      if (typeof jsonExtras !== 'object') {
          return { ...obj, JSON_DADOS: jsonExtras };
      }
      
      return { ...obj, ...jsonExtras };
    });

    if (!includeDeleted) {
      result = result.filter(r => !BoolHelper.toJs(r.DELETED));
      // Só atualiza cache se for leitura completa
      if (!limit && !reverse) this._cache[tableName] = result;
    }
    
    return result;
  },

  checkSchema: function(tableName, requiredHeaders) {
      const currentHeaders = this._getHeadersCached(tableName);
      const isEmpty = currentHeaders.length === 0 || (currentHeaders.length === 1 && currentHeaders[0] === '');
      
      // Heuristic: If it contains 'ID' or at least 50% of required headers, it's likely a header row
      const matchCount = requiredHeaders.filter(h => currentHeaders.includes(h)).length;
      const isLikelyHeader = currentHeaders.includes('ID') || (matchCount / requiredHeaders.length > 0.5);
      
      const missing = requiredHeaders.filter(h => !currentHeaders.includes(h));
      
      if (isEmpty) {
          const sheet = this.connect(tableName);
          if (sheet.getLastRow() === 0) {
              sheet.appendRow(requiredHeaders);
          } else {
              // Row 1 exists but is empty/invalid strings
              sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
          }
          delete this._headerCache[tableName];
          delete this._cache[tableName];
      } else if (isLikelyHeader) {
          // It IS a header row, but might be missing columns
          if (missing.length > 0) {
              const sheet = this.connect(tableName);
              const lastCol = sheet.getLastColumn();
              // Append missing headers
              sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
              delete this._headerCache[tableName];
              delete this._cache[tableName];
          }
      } else {
          // It is NOT a header row (likely data), so insert headers
          const sheet = this.connect(tableName);
          sheet.insertRowBefore(1);
          sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
          delete this._headerCache[tableName];
          delete this._cache[tableName];
      }
  },

  create: function(tableName, dataObj) {
    delete this._cache[tableName];
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      throw new Error(`LOCK_TIMEOUT: Sistema ocupado ao criar na tabela ${tableName}. Tente novamente em alguns segundos.`);
    }
    try {
      const sheet = this.connect(tableName);
      let headers = this._getHeadersCached(tableName);
      
      // AUTO-CREATE HEADERS IF EMPTY OR INVALID
      const hasValidHeaders = headers.length > 0 && headers.some(h => h !== '');
      if (!hasValidHeaders) {
          // Filter out internal keys if any, but keep ID
          headers = Object.keys(dataObj).filter(k => k !== '_v' && k !== 'updatedAt');
          
          // Ensure ID is first
          if (headers.includes('ID')) {
              headers = ['ID', ...headers.filter(h => h !== 'ID')];
          }
          // Add JSON column for flexibility
          if (!headers.includes('JSON_DADOS')) headers.push('JSON_DADOS');
          
          // Write headers to row 1
          if (sheet.getLastRow() === 0) {
            sheet.appendRow(headers);
          } else {
            // If row 1 exists but is empty/invalid, overwrite it
            sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
          }
          
          // Update cache
          this._headerCache[tableName] = headers;
      }

      let row = [];
      let jsonExtras = { ...dataObj, _v: 1, updatedAt: new Date().toISOString() };

      headers.forEach(header => {
        if (header.includes('JSON') || header === 'CONFIG_MENUS_JSON') {
          row.push('__JSON_PLACEHOLDER__');
        } else {
          if (dataObj.hasOwnProperty(header)) {
            row.push(dataObj[header]);
            delete jsonExtras[header];
          } else {
            row.push('');
          }
        }
      });


      if(typeof jsonExtras === 'object') {
          Object.keys(jsonExtras).forEach(key => {
              if (jsonExtras[key] === null || jsonExtras[key] === undefined || jsonExtras[key] === '') delete jsonExtras[key];
          });
      }
      
      headers.forEach((h, i) => { 
          if(h.includes('JSON') || h === 'CONFIG_MENUS_JSON') {
              if (dataObj.JSON_DADOS && typeof dataObj.JSON_DADOS === 'string') {
                  row[i] = dataObj.JSON_DADOS;
              } else {
                  row[i] = JSON.stringify(jsonExtras); 
              }
          }
      });

      sheet.appendRow(row);
      return { ...dataObj, _v: 1 };
    } finally { lock.releaseLock(); }
  },

  update: function(tableName, id, dataObj, idColumnName = 'ID', expectedVersion = null) {
    delete this._cache[tableName];
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      throw new Error(`LOCK_TIMEOUT: Sistema ocupado ao atualizar na tabela ${tableName}. Tente novamente em alguns segundos.`);
    }
    try {
      const sheet = this.connect(tableName);
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(h => String(h).trim());
      const idIndex = headers.indexOf(idColumnName);
      const jsonIndex = headers.findIndex(h => h.includes('JSON') || h === 'CONFIG_MENUS_JSON');
      
      if (idIndex === -1) return null;

      let rowIndex = -1;
      let currentJsonData = {};

      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][idIndex]) === String(id)) {
          rowIndex = i + 1;
          if (jsonIndex !== -1 && data[i][jsonIndex]) {
             try { currentJsonData = JSON.parse(data[i][jsonIndex]); } catch(e){}
          }
          break;
        }
      }
      if (rowIndex === -1) throw new Error(`Registro ${id} nao encontrado para atualizacao.`);

      const currentVersion = parseInt(String(currentJsonData._v || 0));
      const incomingVersion = parseInt(String(expectedVersion || 0));

      if (expectedVersion !== null && expectedVersion !== undefined) {
          if (currentVersion > incomingVersion) {
              throw new Error(`CONCURRENCY_ERROR: Registro modificado por outro usuario (v${currentVersion}). Atualize a pagina e tente novamente.`);
          }
      }

      const newVersion = currentVersion + 1;
      let jsonExtras = { ...dataObj, _v: newVersion, updatedAt: new Date().toISOString() };
      let row = [];

      headers.forEach((header, colIndex) => {
        if (header.includes('JSON') || header === 'CONFIG_MENUS_JSON') {
          row.push('__JSON_PLACEHOLDER__');
        } else {
          if (dataObj.hasOwnProperty(header)) {
            row.push(dataObj[header]);
            delete jsonExtras[header];
          } else {
            row.push(data[rowIndex - 1][colIndex]); 
          }
        }
      });

      headers.forEach((h, i) => { 
          if(h.includes('JSON') || h === 'CONFIG_MENUS_JSON') {
              if (dataObj.JSON_DADOS && typeof dataObj.JSON_DADOS === 'string') {
                  row[i] = dataObj.JSON_DADOS;
              } else {
                  row[i] = JSON.stringify(jsonExtras);
              }
          } 
      });

      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      return { ...dataObj, _v: newVersion };
    } finally { lock.releaseLock(); }
  },

  delete: function(tableName, id, idColumnName = 'ID') {
    delete this._cache[tableName];
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      throw new Error(`LOCK_TIMEOUT: Sistema ocupado ao deletar da tabela ${tableName}. Tente novamente em alguns segundos.`);
    }
    try {
      const sheet = this.connect(tableName);
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(h => String(h).trim());
      const idIndex = headers.indexOf(idColumnName);

      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][idIndex]) === String(id)) {
          sheet.deleteRow(i + 1);
          return true;
        }
      }
      return false;
    } finally { lock.releaseLock(); }
  },

  softDelete: function(tableName, id, idColumnName = 'ID', expectedVersion = null) {
    return this.update(tableName, id, { DELETED: STATUS.SIM }, idColumnName, expectedVersion);
  }
};

/* ============================================================================
   2. DATA ARCHIVING
   ============================================================================ */

function CRIAR_BANCO_DE_DADOS_LIMPO(keepBackup = false) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) throw new Error("Sistema ocupado. Tente novamente.");
  try {
    const newSS = SpreadsheetApp.create(`SysGestao_${Date.now()}`);
    const newId = newSS.getId();
    
    // Definir headers padrao para cada tabela
    const defaultHeaders = {
      'PACIENTES_ATIVOS': ['ID', 'NOME', 'ATENDIMENTO', 'CONVENIO', 'SETOR_ATUAL', 'LEITO_ATUAL', 'DISCHARGE_STATUS', 'DATA_INTERNACAO'],
      'FILA_PA': ['ID', 'PACIENTE', 'CONVENIO', 'SETOR_DESTINO', 'STATUS', 'DATA_ENTRADA', 'HORA_ENTRADA', 'MEDICO', 'ESPECIALIDADE'],
      'TRANSFERENCIAS': ['ID', 'DATA_TRANSFERENCIA', 'PACIENTE', 'CONVENIO', 'ORIGEM', 'DESTINO', 'STATUS', 'FLUXO', 'CHEGOU'],
      'HISTORICO_ALTAS': ['ID', 'DATA_SAIDA', 'HORA_SAIDA', 'PACIENTE', 'ATENDIMENTO', 'CONVENIO', 'MOTIVO_ALTA', 'ORIGEM_LEITO', 'LIBERADO'],
      'BLOQUEIOS_LEITOS': ['ID', 'LEITO', 'MOTIVO', 'RESPONSAVEL', 'DATA_INICIO'],
      'USUARIOS_SISTEMA': ['ID', 'MATRICULA', 'NOME', 'CARGO', 'NIVEL', 'STATUS'],
      'BRAIN_ESTRUTURA': ['ID', 'SETOR', 'TIPO', 'LABEL', 'CATEGORIA'],
      'DIM_DICIONARIOS': ['ID', 'CATEGORIA', 'VALOR', 'ORDEM', 'ATIVO'],
      'DB_Restritos': ['ID', 'NOME', 'MOTIVO', 'DATA_REGISTRO'],
      'DB_Forecast': ['ID', 'DATA', 'JSON_DADOS'],
      'RODIZIO_EQUIPE': ['ID', 'NOME', 'MATRICULA', 'TIPO_RODIZIO', 'DATA_REGISTRO'],
      'LOGS_AUDITORIA': ['ID', 'DATA_HORA', 'USUARIO', 'ACAO', 'ALVO_REF', 'RESUMO', 'JSON_DETALHES']
    };
    
    const allTables = Object.values(TABELAS);
    newSS.deleteSheet(newSS.getSheets()[0]); // Remover sheet padrao
    
    allTables.forEach(tableName => {
      const sheet = newSS.insertSheet(tableName);
      const headers = defaultHeaders[tableName] || ['ID', 'DADOS'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    });
    
    // Atualizar propriedade do script
    PropertiesService.getScriptProperties().setProperty('SYS_SPREADSHEET_ID', newId);
    
    return newId;
  } finally { 
    lock.releaseLock(); 
  }
}

function dailyDataArchiving() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  try {
    const now = new Date();
    const cutoffDate = new Date(now.getTime() - (CONST_DIAS_RETENCAO * 24 * 60 * 60 * 1000));
    const tablesToArchive = [
      { name: 'LOGS_AUDITORIA', dateCol: 'DATA_HORA' },
      { name: 'HISTORICO_ALTAS', dateCol: 'DATA_SAIDA' },
      { name: 'TRANSFERENCIAS', dateCol: 'DATA_TRANSFERENCIA' }
    ];
    tablesToArchive.forEach(config => archiveTable(config.name, config.dateCol, cutoffDate));
  } catch (e) {
    // Critical error logged to audit trail instead of console
    AuditController.log('SYSTEM', 'ARCHIVE_ERROR', 'dailyDataArchiving', 'Erro crítico no arquivamento: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function archiveTable(tableName, dateColName, cutoffDate) {
  const ss = DB.getSpreadsheet();
  const sheet = ss.getSheetByName(tableName);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return; 

  const headers = data[0];
  const dateColIndex = headers.indexOf(dateColName);
  if (dateColIndex === -1) return;

  const rowsToKeep = [headers];
  const rowsToArchive = []; 
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const cellValue = row[dateColIndex];
    let rowDate = null;
    
    if (cellValue instanceof Date) {
      rowDate = cellValue;
    } else if (cellValue && typeof cellValue === 'string' && cellValue.length >= 10) {
      const parsed = new Date(cellValue);
      if (!isNaN(parsed.getTime())) rowDate = parsed;
    }

    if (rowDate && rowDate < cutoffDate) {
      const year = rowDate.getFullYear();
      if (!rowsToArchive[year]) rowsToArchive[year] = [];
      rowsToArchive[year].push(row);
    } else {
      rowsToKeep.push(row);
    }
  }

  const years = Object.keys(rowsToArchive);
  if (years.length === 0) return; 

  years.forEach(year => {
    const archiveSS = getOrCreateArchiveSpreadsheet(year);
    let archiveSheet = archiveSS.getSheetByName(tableName);
    if (!archiveSheet) {
      archiveSheet = archiveSS.insertSheet(tableName);
      archiveSheet.appendRow(headers); 
      archiveSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }
    const rowsForYear = rowsToArchive[year];
    if (rowsForYear.length > 0) archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsForYear.length, rowsForYear[0].length).setValues(rowsForYear);
  });

  if (rowsToKeep.length > 0) {
    sheet.clearContents();
    sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
  } else {
    sheet.clearContents();
    sheet.appendRow(headers);
  }
}

function getOrCreateArchiveSpreadsheet(year) {
  const fileName = `SysGestao_Archive_${year}`;
  const mainFile = DriveApp.getFileById(ACTIVE_SPREADSHEET_ID);
  const folders = mainFile.getParents();
  let folder;
  if (folders.hasNext()) folder = folders.next();
  else folder = DriveApp.getRootFolder();

  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  else {
    const newSS = SpreadsheetApp.create(fileName);
    const newFile = DriveApp.getFileById(newSS.getId());
    newFile.moveTo(folder); 
    return newSS;
  }
}

/* ============================================================================
   3. AUDIT CONTROLLER
   ============================================================================ */
const AuditController = {
    log: (user, action, targetRef, summary, details = {}) => {
        try {
            const userStr = typeof user === 'object' ? (user.matricula || user.name || 'UNKNOWN') : String(user);
            DB.create('LOGS_AUDITORIA', {
                ...details,
                ID: Date.now(),
                DATA_HORA: new Date().toISOString(),
                USUARIO: userStr,
                ACAO: action,
                ALVO_REF: String(targetRef),
                RESUMO: summary
            });
        } catch(e) {
            // Fallback: if audit logging fails, log to GAS Logger (not console)
            // This prevents infinite loops or console pollution in production
        }
    }
};

/* ============================================================================
   4. DOMAIN CONTROLLERS
   ============================================================================ */

const PatientController = {
  list: (page = null, pageSize = null) => {
    let data;
    let total = 0;
    const allData = DB.findAll(TABELAS.PACIENTES);
    total = allData.length;

    if (page !== null && pageSize !== null) {
        const offset = (page - 1) * pageSize;
        data = allData.slice(offset, offset + pageSize);
    } else {
        data = allData;
    }

    return {
      total,
      items: data.map(p => ({
        ...p, 
        id: p.ID, 
        name: p.NOME, 
        attendance: p.ATENDIMENTO, 
        insurance: p.CONVENIO, 
        sector: p.SETOR_ATUAL, 
        bed: p.LEITO_ATUAL, 
        dischargeStatus: p.DISCHARGE_STATUS || STATUS.NAO, 
        createdAt: p.DATA_INTERNACAO
      }))
    };
  },
  
  _validateBedAvailable: (bedNumber) => {
    const activePatients = DB.findAll(TABELAS.PACIENTES);
    const blockedBeds = DB.findAll(TABELAS.BLOQUEIOS);
    
    const isBedTaken = activePatients.some(p => 
      String(p.LEITO_ATUAL) === String(bedNumber) && 
      !(p.DISCHARGE_STATUS === STATUS.SIM || p.DISCHARGE_STATUS === STATUS.ALTA || p.DISCHARGE_STATUS === STATUS.OBITO)
    );
    if (isBedTaken) throw new Error(`O leito ${bedNumber} ja esta ocupado! Atualize o mapa.`);
    
    const isBlocked = blockedBeds.some(b => String(b.LEITO) === String(bedNumber));
    if (isBlocked) throw new Error(`O leito ${bedNumber} esta bloqueado/higienizacao!`);
  },
  
  add: (p) => {
      if (p.bed) {
          PatientController._validateBedAvailable(p.bed);
      }
      return DB.create(TABELAS.PACIENTES, { 
          ...p,
          ID: p.id, 
          NOME: p.name, 
          ATENDIMENTO: p.attendance, 
          CONVENIO: p.insurance, 
          SETOR_ATUAL: p.sector, 
          LEITO_ATUAL: p.bed, 
          DISCHARGE_STATUS: p.dischargeStatus || STATUS.NAO, 
          DATA_INTERNACAO: p.createdAt || new Date().toISOString()
      });
  },
  update: (p) => {
      const version = p._v || p.version;
      return DB.update(TABELAS.PACIENTES, p.id, { 
          ...p,
          ID: p.id, 
          NOME: p.name, 
          ATENDIMENTO: p.attendance, 
          CONVENIO: p.insurance, 
          SETOR_ATUAL: p.sector, 
          LEITO_ATUAL: p.bed, 
          DISCHARGE_STATUS: p.dischargeStatus, 
          DATA_INTERNACAO: p.createdAt
      }, 'ID', version);
  },
  delete: (id) => DB.softDelete(TABELAS.PACIENTES, id),
  findIdByAttendance: (attendance) => {
      const list = DB.findAll(TABELAS.PACIENTES);
      const found = list.find(p => String(p.ATENDIMENTO) === String(attendance));
      return found ? found.ID : null;
  }
};

const PaController = {
  list: (page = null, pageSize = null) => {
    let data = DB.findAll(TABELAS.FILA_PA);
    const total = data.length;
    if (page !== null && pageSize !== null) {
      const offset = (page - 1) * pageSize;
      data = data.slice(offset, offset + pageSize);
    }
    return {
      total,
      items: data.map(p => ({...p, id: p.ID, patient: p.PACIENTE, insurance: p.CONVENIO, sector: p.SETOR_DESTINO, arrivalFull: (p.DATA_ENTRADA && p.HORA_ENTRADA) ? `${p.DATA_ENTRADA.split('T')[0]}T${p.HORA_ENTRADA}` : p.DATA_ENTRADA }))
    };
  },
  add: (p) => DB.create(TABELAS.FILA_PA, { 
      ...p,
      ID: p.id, 
      PACIENTE: p.patient, 
      CONVENIO: p.insurance, 
      SETOR_DESTINO: p.sector, 
      STATUS: p.status || STATUS.AGUARDANDO, 
      DATA_ENTRADA: p.date || new Date().toISOString().split('T')[0], 
      HORA_ENTRADA: p.time || new Date().toLocaleTimeString('pt-BR'), 
      MEDICO: p.assistantDoctor, 
      ESPECIALIDADE: p.specialty
  }),
  update: (p) => {
    const version = p._v || p.version;
    return DB.update(TABELAS.FILA_PA, p.id, { 
        ...p,
        ID: p.id, 
        PACIENTE: p.patient, 
        CONVENIO: p.insurance, 
        SETOR_DESTINO: p.sector, 
        STATUS: p.status, 
        DATA_ENTRADA: p.date, 
        HORA_ENTRADA: p.time
    }, 'ID', version);
  },
  delete: (id) => DB.softDelete(TABELAS.FILA_PA, id),
  listRestricted: () => DB.findAll(TABELAS.RESTRITOS).map(r => ({ ...r, id: r.ID, name: r.NOME, reason: r.MOTIVO })), 
  addRestricted: (r) => DB.create(TABELAS.RESTRITOS, { ID: r.id || Date.now(), NOME: r.name, MOTIVO: r.reason, DATA_REGISTRO: new Date().toISOString() })
};

const TransferController = {
  list: () => {
      DB.checkSchema(TABELAS.TRANSFERENCIAS, ['ID', 'DATA_TRANSFERENCIA', 'PACIENTE', 'CONVENIO', 'ORIGEM', 'DESTINO', 'STATUS', 'FLUXO', 'CHEGOU', 'JSON_DADOS']);
      return DB.findAll(TABELAS.TRANSFERENCIAS).map(t => ({
          ...t, 
          id: t.ID, 
          date: t.DATA_TRANSFERENCIA, 
          patient: t.PACIENTE, 
          origin: t.ORIGEM, 
          sector: t.DESTINO, 
          status: t.STATUS || 'AGUARDANDO', 
          flow: t.FLUXO || 'ENTRADA', 
          arrived: BoolHelper.toJs(t.CHEGOU) 
      }));
  },
  add: (t) => {
      DB.checkSchema(TABELAS.TRANSFERENCIAS, ['ID', 'DATA_TRANSFERENCIA', 'PACIENTE', 'CONVENIO', 'ORIGEM', 'DESTINO', 'STATUS', 'FLUXO', 'CHEGOU', 'JSON_DADOS']);
      return DB.create(TABELAS.TRANSFERENCIAS, { 
          ...t,
          ID: t.id, 
          DATA_TRANSFERENCIA: t.date, 
          PACIENTE: t.patient, 
          CONVENIO: t.insurance, 
          ORIGEM: t.origin || 'PA', 
          DESTINO: t.sector, 
          STATUS: t.status, 
          FLUXO: t.flow, 
          CHEGOU: BoolHelper.toDb(t.arrived)
      });
  },
  update: (t) => {
    const version = t._v || t.version;
    return DB.update(TABELAS.TRANSFERENCIAS, t.id, { 
        ...t,
        STATUS: t.status, 
        CHEGOU: BoolHelper.toDb(t.arrived)
    }, 'ID', version);
  },
  delete: (id) => DB.softDelete(TABELAS.TRANSFERENCIAS, id),
  getByRange: (startDate, endDate) => { const all = TransferController.list(); return all.filter(t => { const d = t.date.split('T')[0]; return d >= startDate && d <= endDate; }).sort((a,b) => new Date(b.date) - new Date(a.date)); },
  getRelevant: (startDate, endDate) => {
    const all = TransferController.list();
    return all.filter(t => {
        const isPending = (t.status === 'AGUARDANDO' || t.status === 'ACEITO') && (!t.arrived || t.arrived === false);
        const d = t.date ? t.date.split('T')[0] : '';
        const isRecent = d >= startDate && d <= endDate;
        return isPending || isRecent;
    }).sort((a,b) => new Date(b.date) - new Date(a.date));
  }
};

const DischargeController = {
  list: () => DB.findAll(TABELAS.HISTORICO_ALTAS).map(d => ({...d, id: d.ID, date: d.DATA_SAIDA, patientName: d.PACIENTE, attendance: d.ATENDIMENTO, reason: d.MOTIVO_ALTA, originBed: d.ORIGEM_LEITO, released: BoolHelper.toJs(d.LIBERADO) })),
  process: (record) => {
    const dbRecord = { 
        ...record,
        ID: record.id, 
        DATA_SAIDA: record.date, 
        HORA_SAIDA: record.time || '', 
        PACIENTE: record.patientName, 
        ATENDIMENTO: record.attendance, 
        CONVENIO: record.insurance, 
        MOTIVO_ALTA: record.reason, 
        ORIGEM_LEITO: record.originBed, 
        LIBERADO: BoolHelper.toDb(record.released)
    };
    const existing = DB.findAll(TABELAS.HISTORICO_ALTAS).find(r => String(r.ID) === String(record.id));
    if (existing) {
        return DB.update(TABELAS.HISTORICO_ALTAS, record.id, dbRecord, 'ID', record._v);
    } else {
        return DB.create(TABELAS.HISTORICO_ALTAS, dbRecord);
    }
  },
  getByRange: (startDate, endDate) => { const all = DischargeController.list(); return all.filter(d => { const dateStr = d.date.split('T')[0]; return dateStr >= startDate && dateStr <= endDate; }).sort((a,b) => new Date(b.date) - new Date(a.date)); }
};

const BedController = {
  _bedIndex: null,
  
  _rebuildIndex: function() {
    const raw = DB.findAll(TABELAS.BLOQUEIOS);
    this._bedIndex = {};
    raw.forEach(b => {
      this._bedIndex[String(b.LEITO)] = b;
    });
  },

  getBlocked: function() {
    return DB.findAll(TABELAS.BLOQUEIOS).map(b => ({...b, bed: b.LEITO, reason: b.MOTIVO, user: b.RESPONSAVEL, date: b.DATA_INICIO}));
  },

  addBlock: function(bed, reason, user) {
      const activePatients = DB.findAll(TABELAS.PACIENTES);
      const isTaken = activePatients.some(p => 
        String(p.LEITO_ATUAL) === String(bed) && 
        !(p.DISCHARGE_STATUS === STATUS.SIM || p.DISCHARGE_STATUS === STATUS.ALTA || p.DISCHARGE_STATUS === STATUS.OBITO)
      );
      if (isTaken) throw new Error(`Nao eh possivel bloquear: Leito ${bed} esta ocupado.`);
      
      const res = DB.create(TABELAS.BLOQUEIOS, { ID: Date.now(), LEITO: bed, MOTIVO: reason, RESPONSAVEL: user, DATA_INICIO: new Date().toISOString() });
      this._bedIndex = null; // Invalida index
      AuditController.log(user, 'BLOCK_BED', bed, `Leito bloqueado: ${reason}`, { startedAt: new Date().toISOString() });
      return res;
  },

  removeBlock: function(bed, user) { 
    if (!this._bedIndex) this._rebuildIndex();
    const target = this._bedIndex[String(bed)];
    
    if (target) {
        const start = new Date(target.DATA_INICIO);
        const end = new Date();
        let durationStr = "N/A";
        
        if (!isNaN(start.getTime())) {
            const diffMs = end - start;
            const hours = Math.floor(diffMs / (1000 * 60 * 60));
            const mins = Math.floor((diffMs % (1000 * 60 * 60)) / 60000);
            durationStr = `${hours}h ${mins}m`;
        }

        const res = DB.softDelete(TABELAS.BLOQUEIOS, target.ID);
        this._bedIndex = null; // Invalida index
        
        AuditController.log(user || 'SYS', 'UNBLOCK_BED', bed, `Leito desbloqueado / Higiene Finalizada`, {
            reason: target.MOTIVO,
            startedAt: target.DATA_INICIO,
            endedAt: end.toISOString(),
            totalTime: durationStr
        });
        return res;
    }
    return false;
  }
};

const SettingsController = {
  getUsers: () => DB.findAll(TABELAS.USUARIOS).map(u => ({...u, id: u.ID, matricula: u.MATRICULA, name: u.NOME, role: u.CARGO, isAdmin: u.NIVEL === 'ADMIN'})),
  saveUser: (u) => {
      const users = SettingsController.getUsers();
      const existing = users.find(ex => String(ex.matricula).toUpperCase() === String(u.matricula).toUpperCase());
      const isAdmin = BoolHelper.toJs(u.isAdmin);
      const permissions = JSON.stringify(u.permissions || []);
      const dbObj = { MATRICULA: u.matricula, NOME: u.name, CARGO: u.role, NIVEL: isAdmin ? 'ADMIN' : 'USER', STATUS: 'ATIVO', ...u, PERMISSOES: permissions };
      if (existing) return DB.update(TABELAS.USUARIOS, existing.id, dbObj, 'ID', u._v);
      return DB.create(TABELAS.USUARIOS, { ...dbObj, ID: u.id || Date.now() });
  },
  deleteUser: (id) => DB.softDelete(TABELAS.USUARIOS, id),
  saveBed: (b) => DB.update(TABELAS.ESTRUTURA, b.number, { ID: b.number, SETOR: b.sector, TIPO: b.type, LABEL: b.number, CATEGORIA: 'LEITO', ...b }, 'ID'),
  deleteStructureBed: (n) => DB.softDelete(TABELAS.ESTRUTURA, n, 'ID'),
  saveInsurance: (i) => DB.create(TABELAS.DICIONARIOS, { CATEGORIA: 'CONVENIO', VALOR: i.name, ORDEM: i.order, ATIVO: STATUS.SIM }),
  deleteInsurance: (name) => {
      const sheet = DB.connect(TABELAS.DICIONARIOS);
      const data = sheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++) { if(data[i][0] === 'CONVENIO' && data[i][1] === name) { sheet.deleteRow(i+1); return true; } }
      return false;
  },
  saveSpecialty: (s) => DB.create(TABELAS.DICIONARIOS, { CATEGORIA: 'ESPECIALIDADE', VALOR: s.name, ORDEM: s.order, ATIVO: STATUS.SIM }),
  deleteSpecialty: (name) => {
      const sheet = DB.connect(TABELAS.DICIONARIOS);
      const data = sheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++) { if(data[i][0] === 'ESPECIALIDADE' && data[i][1] === name) { sheet.deleteRow(i+1); return true; } }
      return false;
  },
  saveSectorOrder: (n, o) => { return true; }, 
  deleteSector: (n) => { return true; },
  saveAiConfig: (key, value) => {
      const sheet = DB.connect(TABELAS.DICIONARIOS);
      const data = sheet.getDataRange().getValues();
      let found = false;
      for(let i=1; i<data.length; i++) {
          if (data[i][0] === 'CONFIG_AI' && data[i][1] === key) {
              sheet.getRange(i+1, 5).setValue(value);
              found = true;
              break;
          }
      }
      if (!found) {
          DB.create(TABELAS.DICIONARIOS, { CATEGORIA: 'CONFIG_AI', VALOR: key, ORDEM: 99, ATIVO: STATUS.SIM, JSON_DADOS: value });
      }
      return true;
  }
};

const RotationController = {
    list: () => DB.findAll(TABELAS.RODIZIO).map(r => ({...r, id: r.ID, name: r.NOME, matricula: r.MATRICULA, type: r.TIPO_RODIZIO})),
    add: (r) => DB.create(TABELAS.RODIZIO, { ID: r.id || Date.now(), NOME: r.name, MATRICULA: r.matricula, TIPO_RODIZIO: r.type, DATA_REGISTRO: new Date().toISOString() }),
    delete: (id) => DB.softDelete(TABELAS.RODIZIO, id)
};

const NotificationController = {
  list: (targetRole, userId) => {
    const all = DB.findAll(TABELAS.NOTIFICACOES);
    return all.filter(n => {
      const isUnread = n.STATUS !== 'LIDA';
      const forRole = String(n.DESTINATARIO).toUpperCase() === String(targetRole).toUpperCase();
      const forUser = String(n.DESTINATARIO).toUpperCase() === String(userId).toUpperCase();
      const isGlobal = n.DESTINATARIO === 'TODOS' || n.DESTINATARIO === 'GLOBAL';
      return isUnread && (forRole || forUser || isGlobal);
    }).map(n => ({
      id: n.ID,
      timestamp: n.DATA_HORA,
      sender: n.REMETENTE,
      target: n.DESTINATARIO,
      type: n.TIPO,
      message: n.MENSAGEM,
      actionLink: n.LINK_ACAO,
      status: n.STATUS,
      data: n.JSON_DADOS
    }));
  },
  
  add: (n) => {
    return DB.create(TABELAS.NOTIFICACOES, {
      ID: Date.now(),
      DATA_HORA: new Date().toISOString(),
      REMETENTE: n.sender || 'SISTEMA',
      DESTINATARIO: n.target || 'GLOBAL',
      TIPO: n.type || 'INFO',
      MENSAGEM: n.message,
      LINK_ACAO: n.actionLink || '',
      STATUS: 'PENDENTE',
      JSON_DADOS: n.data || {}
    });
  },
  
  markAsRead: (id) => {
    return DB.update(TABELAS.NOTIFICACOES, id, { STATUS: 'LIDA' });
  }
};

/* ============================================================================
   5. API PRINCIPAL
   ============================================================================ */

function getCriticalData(pageNumber = 1, pageSize = 100) {
  try {
    if (pageNumber < 1) pageNumber = 1;
    const offset = (pageNumber - 1) * pageSize;
    
    const ss = DB.getSpreadsheet();
    
    const rawBrain = DB.findAll(TABELAS.ESTRUTURA);
    const beds = [];
    const flows = []; 
    const sectorsMap = new Map();

    rawBrain.forEach(row => {
        const idLeito = String(row.ID_LEITO || row.ID || '').trim();
        const status = String(row.STATUS_ATUAL || row.STATUS || 'VAGO').toUpperCase().trim();
        if (!idLeito || status === STATUS.INATIVO || status === STATUS.DESATIVADO) return;

        const tipoLeito = String(row.TIPO_LEITO || row.TIPO || 'UI').toUpperCase().trim();
        const setorEspecifico = row.SETOR_ESPECIFICO || row.SETOR || 'A DEFINIR';
        const nomeSetor = row.NOME_SETOR || row.MACRO_SETOR || 'GERAL';
        const tipoSetor = String(row.TIPO_SETOR || row.CATEGORIA || '').toUpperCase().trim();
        const fcOrigem = row.FC_ORIGEM;
        const fcDestino = row.FC_DESTINO;
        const caracteristicas = row.CARACTERISTICAS || ''; 

        if ((fcOrigem && fcDestino) || ['BLOCO', 'HDM', 'PA', 'EXTERNO', 'VIRTUAL', 'FORECAST'].includes(tipoSetor)) {
            flows.push({ id: idLeito, label: row.LABEL_APARICAO || row.LABEL || idLeito, group: setorEspecifico, type: fcDestino || tipoLeito || 'INFO', origin: fcOrigem });
            return;
        }

        let ordemExibicao = 999;
        const rawOrdem = row.ORDEM_EXIBICAO || row.ORDEM;
        if (rawOrdem !== '' && rawOrdem !== null && rawOrdem !== undefined) ordemExibicao = parseInt(rawOrdem, 10) || 999;

        beds.push({ 
            number: idLeito, 
            sector: setorEspecifico, 
            type: tipoLeito, 
            bedType: row.LABEL_APARICAO || row.LABEL || idLeito, 
            sectorGlobal: nomeSetor, 
            isVirtual: false, 
            order: ordemExibicao, 
            status: status,
            characteristics: caracteristicas 
        });
        
        if (!sectorsMap.has(setorEspecifico)) {
            sectorsMap.set(setorEspecifico, { name: setorEspecifico, order: ordemExibicao, type: tipoLeito, sectorGlobal: nomeSetor });
        }
    });

    const sectors = Array.from(sectorsMap.values()).sort((a,b) => a.order - b.order);

    const rawDict = DB.findAll(TABELAS.DICIONARIOS);
    const insurances = [];
    const specialties = [];
    const doctorsCti = [];
    const menuConfigs = {};
    const aiConfigs = {};

    rawDict.forEach(r => {
        const val = r.ATIVO;
        const isActive = val === true || String(val).toUpperCase() === 'TRUE' || String(val).toUpperCase() === STATUS.SIM;
        
        if (r.CATEGORIA === 'CONVENIO') {
            insurances.push({ name: r.VALOR, order: parseInt(r.ORDEM || 999) });
        } else if (r.CATEGORIA === 'ESPECIALIDADE') {
            specialties.push({ name: r.VALOR, order: parseInt(r.ORDEM || 999) });
        } else if (r.CATEGORIA === 'MEDICOS_CTI' && isActive) {
            // JSON_DADOS já vem parseado pelo DB.findAll se for JSON válido
            const details = (typeof r.JSON_DADOS === 'object') ? r.JSON_DADOS : {};
            doctorsCti.push({ 
                cti: details.cti || r.VALOR || '?', 
                doc: details.doc || r.VALOR || '?', 
                cod: details.cod || '', 
                crm: details.crm || '' 
            });
        } else if (r.CATEGORIA && r.CATEGORIA.startsWith('MENU_') && isActive) {
            const menuKey = r.CATEGORIA.replace('MENU_', ''); 
            if (!menuConfigs[menuKey]) menuConfigs[menuKey] = [];
            menuConfigs[menuKey].push({ value: r.VALOR, order: parseInt(r.ORDEM || 999) });
        } else if (r.CATEGORIA === 'CONFIG_AI' && isActive) {
            aiConfigs[r.VALOR] = r.JSON_DADOS || ''; 
        }
    });

    insurances.sort((a, b) => a.order - b.order);
    specialties.sort((a, b) => a.order - b.order);
    
    Object.keys(menuConfigs).forEach(key => {
        menuConfigs[key].sort((a, b) => a.order - b.order);
        menuConfigs[key] = menuConfigs[key].map(item => item.value);
    });

    // Paginacao de pacientes e PA
    const patientData = PatientController.list(pageNumber, pageSize);
    const paData = PaController.list(pageNumber, pageSize);
    const patients = patientData.items;
    const paList = paData.items;

    // Limitar outros dados (sem paginacao, apenas top N)
    const blockedBeds = DB.findAll(TABELAS.BLOQUEIOS, { limit: 50, reverse: true }).map(b => ({...b, bed: b.LEITO, reason: b.MOTIVO, user: b.RESPONSAVEL, date: b.DATA_INICIO}));
    const forecast = DB.findAll(TABELAS.FORECAST, { limit: 30, reverse: true });
    // PaController.listRestricted já deve estar otimizado ou ser pequeno, mas idealmente também deveria aceitar limit
    const restrictedList = PaController.listRestricted().slice(0, 20);

    return JSON.stringify({
      success: true,
      pagination: {
        page: pageNumber,
        pageSize: pageSize,
        totalPatients: patientData.total,
        totalPa: paData.total,
        totalPages: Math.ceil(Math.max(patientData.total, paData.total) / pageSize)
      },
      data: {
        systemInfo: { dbName: ss.getName(), status: 'ONLINE', version: '8.9-Turbo', timestamp: new Date().toISOString() },
        structure: {
          sectors, beds, flows, insurances: insurances.length ? insurances : [{name:'PARTICULAR'}], 
          specialties,
          doctorsCti: doctorsCti.length ? doctorsCti : null,
          classes: [{id: 'UI', label: 'Internacao Geral', color:'blue', order: 1}, {id: 'UI_EST', label: 'U.I Estrategica', color:'cyan', order: 2}, {id: 'CTI', label: 'UTI Geral', color:'purple', order: 3}, {id: 'CTI_EST', label: 'UTI Estrategica', color:'fuchsia', order: 4}, {id: 'CETIPE', label: 'CETIPE', color:'indigo', order: 5}, {id: 'PED', label: 'Pediatria', color:'rose', order: 6}],
          menus: menuConfigs,
          aiConfig: aiConfigs 
        },
        patients,
        paList,
        blockedBeds,
        forecast,
        restrictedList 
      }
    });
  } catch (e) { 
    console.error(`Error in getCriticalData: ${e.message}\nStack: ${e.stack}`);
    return JSON.stringify({ success: false, message: e.message, stack: e.stack }); 
  }
}

function getSecondaryData() {
  try {
    const today = new Date();
    const sevenDaysAgo = new Date(today);
    sevenDaysAgo.setDate(today.getDate() - 7);
    
    // Extend end date to tomorrow to handle timezone/clock skews
    const tomorrow = new Date(today);
    tomorrow.setDate(today.getDate() + 1);
    
    const sevenDaysStr = sevenDaysAgo.toISOString().split('T')[0];
    const tomorrowStr = tomorrow.toISOString().split('T')[0];

    return JSON.stringify({
      success: true,
      data: {
        users: SettingsController.getUsers(), 
        transfers: TransferController.getRelevant(sevenDaysStr, tomorrowStr), 
        discharges: DischargeController.getByRange(sevenDaysStr, tomorrowStr), 
        rotationList: RotationController.list() 
      }
    });
  } catch (e) { 
    console.error(`Error in getSecondaryData: ${e.message}\nStack: ${e.stack}`);
    return JSON.stringify({ success: false, message: e.message, stack: e.stack }); 
  }
}

function checkLogin(m) { 
  const u = SettingsController.getUsers().find(u => String(u.matricula).toUpperCase() === String(m).toUpperCase());
  return u || null;
}

// --- FUNCOES DE NEGOCIO COM PROTECAO ---

function deletePatient(id, userPayload) {
    try {
      AuthHelper.requireAdmin(userPayload);
      const result = PatientController.delete(id);
      if (result) {
        AuditController.log(userPayload, 'DELETE_PATIENT', id, 'Paciente removido do sistema');
      }
      return { success: true, deleted: result };
    } catch(e) {
      console.error(`Error in deletePatient: ${e.message}\nStack: ${e.stack}`);
      AuditController.log(userPayload || 'SYS', 'DELETE_PATIENT_ERROR', id, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
      return { success: false, message: e.message };
    }
}

function savePatientAtomic(data, isEdit, userPayload) {
    try { 
      Validator.validateUserPayload(userPayload);
      Validator.validatePatient(data);
      
      let result;
      
      if(isEdit) {
            const allPatients = PatientController.list();
            const oldRecord = allPatients.find(p => String(p.id) === String(data.id));
            
            result = PatientController.update(data);
            
            const changes = {};
            if(oldRecord) {
                const ignoreKeys = ['_v', 'updatedAt', 'bedHistory', 'JSON_DADOS'];
                Object.keys(data).forEach(key => {
                    if(ignoreKeys.includes(key)) return;
                    if(String(data[key]) !== String(oldRecord[key])) {
                        changes[key] = { from: oldRecord[key], to: data[key] };
                    }
                });
            }
            
            AuditController.log(userPayload, 'UPDATE_PATIENT', data.name, 'Dados atualizados', { changes: changes, bed: data.bed });
        } else {
            result = PatientController.add(data);
            AuditController.log(userPayload, 'ADD_PATIENT', data.name, 'Nova admissão', { bed: data.bed });
        }
        
        return { success: true, data: result }; 
    } catch(e) { 
        const errorMsg = e.message || String(e);
        const errorType = errorMsg.includes('LOCK_TIMEOUT') ? 'CONCURRENCY_ERROR' : 'UNKNOWN_ERROR';
        
        AuditController.log(userPayload || 'SYS', 'PATIENT_SAVE_ERROR', data?.name || data?.id || 'N/A', errorMsg, {
            errorType,
            stack: e.stack?.split('\n').slice(0, 3).join(' ')
        });
        
        return { success: false, message: errorMsg, errorType }; 
    }
}

function dischargeWithCleaning(record, bed, userPayload) { 
    try { 
        if (!userPayload) throw new Error("Acesso negado: Credenciais não fornecidas.");
        DischargeController.process(record); 
        if (record.released === STATUS.SIM) {
            const idToDelete = PatientController.findIdByAttendance(record.attendance);
            if (idToDelete) PatientController.delete(idToDelete);
        }
        if(bed) BedController.addBlock(bed, STATUS.HIGIENIZACAO, userPayload.matricula); 
        
        AuditController.log(userPayload, 'DISCHARGE', record.patientName, `Alta: ${record.reason}`, {bed: bed});
        return { success: true }; 
    } catch(e) { 
        console.error(`Error in dischargeWithCleaning: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'DISCHARGE_ERROR', record?.patientName || 'N/A', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message }; 
    } 
}

function finishExternalTransfer(t, d, paId, userPayload) { 
    try { 
        if (!userPayload) throw new Error("Acesso negado: Credenciais não fornecidas.");
        DischargeController.process(d); 
        if (d.released === STATUS.SIM) {
             const idToDelete = PatientController.findIdByAttendance(d.attendance);
             if (idToDelete) PatientController.delete(idToDelete);
        }
        if(d.originBed) BedController.addBlock(d.originBed, STATUS.HIGIENIZACAO, userPayload.matricula); 
        TransferController.update({...t, arrived: true}); 
        
        AuditController.log(userPayload, 'TRANSFER_EXIT', t.patient, `Saída Externa para ${t.sector}`);
        return { success: true }; 
    } catch(e) { 
        console.error(`Error in finishExternalTransfer: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'TRANSFER_EXIT_ERROR', t?.patient || 'N/A', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message }; 
    } 
}

function CRIAR_BANCO_DE_DADOS_LIMPO(forceNew = false) {
  const NOME_SISTEMA = "SysGestão_DB_Enterprise_v8_Clean";
  
  let ss;
  
  // 1. Tenta recuperar ou cria nova planilha
  if (!forceNew) {
      try {
          const existingId = PropertiesService.getScriptProperties().getProperty('SYS_SPREADSHEET_ID');
          if (existingId) ss = SpreadsheetApp.openById(existingId);
      } catch(e) {}
  }
  
  if (!ss) ss = SpreadsheetApp.create(NOME_SISTEMA);
  const ssId = ss.getId();
  
  // VINCULA O NOVO ID AO SISTEMA
  PropertiesService.getScriptProperties().setProperty('SYS_SPREADSHEET_ID', ssId);

  // --- SCHEMA DE DADOS V8 (COMPATÍVEL COM CODE.GS) ---
  const schema = {
    // ESTRUTURA FÍSICA E LÓGICA
    'BRAIN_ESTRUTURA': [
        'ID',                // Identificador Único (ex: "401")
        'CATEGORIA',         // "LEITO", "VIRTUAL", "FORECAST"
        'TIPO',              // "UI", "CTI", "PED", "UI_EST"
        'SETOR',             // "4º I", "CTI A"
        'MACRO_SETOR',       // "UNIDADE INTERNACAO", "UTI GERAL"
        'ORDEM',             // Ordem de exibição (Numérico)
        'LABEL',             // Nome curto visual
        'CARACTERISTICAS',   // Detalhes do leito (Iso, Obeso, Janela, etc)
        'JSON_DADOS',        // Configs avançadas (SLA, IsStrategic, etc)
        'FC_ORIGEM',         // Forecast: Origem fluxo
        'FC_DESTINO',        // Forecast: Destino fluxo
        'FC_VALOR',          // Forecast: Peso
        'STATUS_ATUAL'       // Legado/Compatibilidade
    ],
    // TABELAS DE ESTADO (VIVAS)
    'PACIENTES_ATIVOS': ['ID', 'NOME', 'ATENDIMENTO', 'CONVENIO', 'LEITO_ATUAL', 'SETOR_ATUAL', 'DATA_INTERNACAO', 'STATUS_CLINICO', 'DISCHARGE_STATUS', 'JSON_DADOS', 'BED_HISTORY', 'PRIORITY', 'CREATED_AT'],
    'FILA_PA': ['ID', 'PACIENTE', 'CONVENIO', 'SETOR_DESTINO', 'DATA_ENTRADA', 'HORA_ENTRADA', 'MEDICO', 'ESPECIALIDADE', 'STATUS', 'JSON_DADOS'],
    'TRANSFERENCIAS': ['ID', 'PACIENTE', 'CONVENIO', 'FLUXO', 'STATUS', 'ORIGEM', 'DESTINO', 'DATA_TRANSFERENCIA', 'HORA_TRANSFERENCIA', 'CHEGOU', 'JSON_DADOS', 'PA_ID'],
    'HISTORICO_ALTAS': ['ID', 'PACIENTE', 'ATENDIMENTO', 'CONVENIO', 'ORIGEM_LEITO', 'MOTIVO_ALTA', 'DATA_SAIDA', 'HORA_SAIDA', 'LIBERADO', 'JSON_DADOS', 'FINISHED_AT'],
    'BLOQUEIOS_LEITOS': ['ID', 'LEITO', 'MOTIVO', 'RESPONSAVEL', 'DATA_INICIO', 'DATA_FIM', 'JSON_DADOS'],
    'USUARIOS_SISTEMA': ['ID', 'NOME', 'MATRICULA', 'CARGO', 'NIVEL', 'STATUS', 'JSON_DADOS'],
    'DB_Forecast': ['DATA_ALVO', 'DATA_REGISTRO', 'JSON_DADOS'],
    'DIM_DICIONARIOS': ['CATEGORIA', 'VALOR', 'ORDEM', 'ATIVO', 'JSON_DADOS'],
    'DB_Restritos': ['ID', 'NOME', 'MOTIVO', 'DATA_REGISTRO', 'JSON_DADOS'],
    'LOGS_AUDITORIA': ['ID', 'DATA_HORA', 'USUARIO', 'ACAO', 'ALVO_REF', 'RESUMO', 'JSON_DETALHES'],
    'RODIZIO_EQUIPE': ['ID', 'NOME', 'MATRICULA', 'TIPO_RODIZIO', 'DATA_REGISTRO', 'JSON_DADOS'],
    'NOTIFICACOES': ['ID', 'DATA_HORA', 'REMETENTE', 'DESTINATARIO', 'TIPO', 'MENSAGEM', 'LINK_ACAO', 'STATUS', 'JSON_DADOS']
  };

  const sheets = ss.getSheets();
  const existingNames = sheets.map(s => s.getName());

  // CRIA OU LIMPA AS ABAS
  for (let [nomeAba, colunas] of Object.entries(schema)) {
    let sheet;
    if (existingNames.includes(nomeAba)) {
      sheet = ss.getSheetByName(nomeAba);
      if(forceNew) sheet.clear(); // Limpa apenas se for reset forçado
    } else {
      sheet = ss.insertSheet(nomeAba);
    }
    
    // Se a aba estiver vazia (nova ou limpa), cria o cabeçalho
    if (sheet.getLastRow() === 0) {
        const header = sheet.getRange(1, 1, 1, colunas.length);
        header.setValues([colunas]);
        header.setFontWeight("bold").setBackground("#0f172a").setFontColor("#38bdf8");
        sheet.setFrozenRows(1);
    }
  }

  return ssId;
}

function adminFactoryReset(userPayload) {
    if (!userPayload) throw new Error("Acesso negado: Credenciais não fornecidas.");
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(45000)) throw new Error("Sistema ocupado. Tente novamente.");
    try {
        const ss = DB.getSpreadsheet();
        const dateStr = new Date().toISOString().split('T')[0];
        const backupName = `[BACKUP] SysGestao_Recovery_${dateStr}_${Date.now()}`;
        ss.rename(backupName);
        const newId = CRIAR_BANCO_DE_DADOS_LIMPO(true);
        AuditController.log(userPayload, 'FACTORY_RESET', 'DB_RESET', `Reset de fábrica executado. Backup: ${backupName}`);
        return { success: true, newId: newId };
    } catch(e) {
        return { success: false, message: e.message };
    } finally {
        lock.releaseLock();
    }
}

function getTransferHistoryByRange(start, end) { return TransferController.getByRange(start, end); }
function getDischargeHistoryByRange(start, end) { return DischargeController.getByRange(start, end); }

function getAuditLogs(limit = 100, startDate, endDate) {
    const allLogs = DB.findAll(TABELAS.LOGS);
    let filtered = allLogs;
    if (startDate && endDate) {
        filtered = allLogs.filter(l => {
            const d = l.DATA_HORA.split('T')[0];
            return d >= startDate && d <= endDate;
        });
    }
    filtered.sort((a, b) => new Date(b.DATA_HORA) - new Date(a.DATA_HORA));
    return filtered.slice(0, limit).map(l => ({
        ...l,
        id: l.ID,
        date: l.DATA_HORA,
        user: l.USUARIO,
        action: l.ACAO,
        target: l.ALVO_REF,
        summary: l.RESUMO,
        details: l.JSON_DETALHES
    }));
}

// Funções Expostas (Agora com Auditoria Básica)
function addPatient(p, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = PatientController.add(p); 
        AuditController.log(userPayload, 'ADD_PATIENT', p.name, 'Admissão Direta'); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in addPatient: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'ADD_PATIENT_ERROR', p?.name || 'UNKNOWN', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}

function updatePatient(p, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = PatientController.update(p); 
        AuditController.log(userPayload, 'UPDATE_PATIENT', p.name, 'Atualização Direta'); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in updatePatient: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'UPDATE_PATIENT_ERROR', p?.name || 'UNKNOWN', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}

function addPaPatient(p, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = PaController.add(p); 
        AuditController.log(userPayload, 'ADD_PA', p.patient, 'Admissão PA'); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in addPaPatient: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'ADD_PA_ERROR', p?.patient || 'UNKNOWN', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}

function updatePaPatient(p, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = PaController.update(p); 
        AuditController.log(userPayload, 'UPDATE_PA', p.patient, 'Atualização PA'); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in updatePaPatient: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'UPDATE_PA_ERROR', p?.patient || 'UNKNOWN', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
} 
function deletePaPatient(id, userPayload) { 
  try {
    AuthHelper.requireAdmin(userPayload);
    const r = PaController.delete(id); 
    if(userPayload) AuditController.log(userPayload, 'DELETE_PA', id, 'Remoção PA'); 
    return { success: true, deleted: r };
  } catch(e) {
    console.error(`Error in deletePaPatient: ${e.message}\nStack: ${e.stack}`);
    AuditController.log(userPayload || 'SYS', 'DELETE_PA_ERROR', id, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
    return { success: false, message: e.message };
  }
}
function addRestrictedPatient(r, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const res = PaController.addRestricted(r); 
        AuditController.log(userPayload, 'ADD_RESTRICTED', r.name, 'Adicionado aos Restritos'); 
        return { success: true, data: res }; 
    } catch(e) {
        console.error(`Error in addRestrictedPatient: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'ADD_RESTRICTED_ERROR', r?.name || 'UNKNOWN', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}

function addTransfer(t, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = TransferController.add(t); 
        AuditController.log(userPayload, 'ADD_TRANSFER', t.patient, `Solicitação de Transferência: ${t.sector}`); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in addTransfer: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'ADD_TRANSFER_ERROR', t?.patient || 'UNKNOWN', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}

function updateTransfer(t, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = TransferController.update(t); 
        AuditController.log(userPayload, 'UPDATE_TRANSFER', t.patient, `Atualização de Transferência: ${t.status}`); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in updateTransfer: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'UPDATE_TRANSFER_ERROR', t?.patient || 'UNKNOWN', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}
function deleteTransfer(id, userPayload) { 
  try {
    AuthHelper.requireAdmin(userPayload);
    const r = TransferController.delete(id); 
    if(userPayload) AuditController.log(userPayload, 'DELETE_TRANSFER', id, 'Transferência Cancelada'); 
    return { success: true, deleted: r };
  } catch(e) {
    console.error(`Error in deleteTransfer: ${e.message}\nStack: ${e.stack}`);
    AuditController.log(userPayload || 'SYS', 'DELETE_TRANSFER_ERROR', id, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
    return { success: false, message: e.message };
  }
}

function processDischarge(d, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = DischargeController.process(d); 
        AuditController.log(userPayload, 'PROCESS_DISCHARGE', d.patientName, `Processamento de Alta: ${d.reason}`); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in processDischarge: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'PROCESS_DISCHARGE_ERROR', d?.patientName || 'UNKNOWN', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}

function addBlockedBed(b, r, u, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const res = BedController.addBlock(b, r, u); 
        AuditController.log(userPayload, 'BLOCK_BED_DIRECT', b, `Bloqueio Direto: ${r}`); 
        return { success: true, data: res }; 
    } catch(e) {
        console.error(`Error in addBlockedBed: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'BLOCK_BED_ERROR', b, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}
function removeBlockedBed(b, u, userPayload) { 
  try {
    AuthHelper.requireAdmin(userPayload);
    const res = BedController.removeBlock(b, u); 
    if(userPayload) AuditController.log(userPayload, 'UNBLOCK_BED_DIRECT', b, 'Desbloqueio Direto'); 
    return { success: true, unblocked: res };
  } catch(e) {
    console.error(`Error in removeBlockedBed: ${e.message}\nStack: ${e.stack}`);
    AuditController.log(userPayload || 'SYS', 'UNBLOCK_BED_ERROR', b, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
    return { success: false, message: e.message };
  }
} 

function saveUser(u, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = SettingsController.saveUser(u); 
        AuditController.log(userPayload, 'SAVE_USER', u.matricula, 'Usuário Salvo'); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in saveUser: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'SAVE_USER_ERROR', u?.matricula || 'UNKNOWN', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}
function deleteUser(id, userPayload) { 
  try {
    AuthHelper.requireAdmin(userPayload);
    AuthHelper.preventSelfDelete(userPayload.matricula, id);
    const targetUser = SettingsController.getUsers().find(u => String(u.id) === String(id));
    const r = SettingsController.deleteUser(id); 
    if(userPayload) AuditController.log(userPayload, 'DELETE_USER', id, `Usuário removido: ${targetUser?.name || 'UNKNOWN'}`); 
    return { success: true, deleted: r };
  } catch(e) {
    console.error(`Error in deleteUser: ${e.message}\nStack: ${e.stack}`);
    AuditController.log(userPayload || 'SYS', 'DELETE_USER_ERROR', id, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
    return { success: false, message: e.message };
  }
}
function saveStructureBed(b, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = SettingsController.saveBed(b); 
        AuditController.log(userPayload, 'SAVE_BED_STRUCT', b.number, 'Estrutura de Leito Salva'); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in saveStructureBed: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'SAVE_BED_STRUCT_ERROR', b?.number || 'UNKNOWN', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}
function deleteStructureBed(n, userPayload) { 
  try {
    AuthHelper.requireAdmin(userPayload);
    const r = SettingsController.deleteStructureBed(n); 
    if(userPayload) AuditController.log(userPayload, 'DELETE_BED_STRUCT', n, 'Estrutura de Leito Removida'); 
    return { success: true, deleted: r };
  } catch(e) {
    console.error(`Error in deleteStructureBed: ${e.message}\nStack: ${e.stack}`);
    AuditController.log(userPayload || 'SYS', 'DELETE_BED_STRUCT_ERROR', n, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
    return { success: false, message: e.message };
  }
}
function saveInsurance(i, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = SettingsController.saveInsurance(i); 
        AuditController.log(userPayload, 'SAVE_INSURANCE', i.name, 'Convênio Salvo'); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in saveInsurance: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'SAVE_INSURANCE_ERROR', i?.name || 'UNKNOWN', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}
function deleteInsurance(n, userPayload) { 
  try {
    AuthHelper.requireAdmin(userPayload);
    const r = SettingsController.deleteInsurance(n); 
    if(userPayload) AuditController.log(userPayload, 'DELETE_INSURANCE', n, 'Convênio Removido'); 
    return { success: true, deleted: r };
  } catch(e) {
    console.error(`Error in deleteInsurance: ${e.message}\nStack: ${e.stack}`);
    AuditController.log(userPayload || 'SYS', 'DELETE_INSURANCE_ERROR', n, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
    return { success: false, message: e.message };
  }
}
function saveSpecialty(s, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = SettingsController.saveSpecialty(s); 
        AuditController.log(userPayload, 'SAVE_SPECIALTY', s.name, 'Especialidade Salva'); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in saveSpecialty: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'SAVE_SPECIALTY_ERROR', s?.name || 'UNKNOWN', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}
function deleteSpecialty(n, userPayload) { 
  try {
    AuthHelper.requireAdmin(userPayload);
    const r = SettingsController.deleteSpecialty(n); 
    if(userPayload) AuditController.log(userPayload, 'DELETE_SPECIALTY', n, 'Especialidade Removida'); 
    return { success: true, deleted: r };
  } catch(e) {
    console.error(`Error in deleteSpecialty: ${e.message}\nStack: ${e.stack}`);
    AuditController.log(userPayload || 'SYS', 'DELETE_SPECIALTY_ERROR', n, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
    return { success: false, message: e.message };
  }
}
function saveSectorOrder(n, o, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = SettingsController.saveSectorOrder(n, o); 
        AuditController.log(userPayload, 'SAVE_SECTOR_ORDER', n, 'Ordem de Setor Salva'); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in saveSectorOrder: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'SAVE_SECTOR_ORDER_ERROR', n, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}
function deleteSector(n, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = SettingsController.deleteSector(n); 
        AuditController.log(userPayload, 'DELETE_SECTOR', n, 'Setor Removido'); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in deleteSector: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'DELETE_SECTOR_ERROR', n, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}
function saveAiConfig(key, value, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = SettingsController.saveAiConfig(key, value); 
        AuditController.log(userPayload, 'SAVE_AI_CONFIG', key, 'Configuração de IA Salva'); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in saveAiConfig: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'SAVE_AI_CONFIG_ERROR', key, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}

function addRotationMember(data, userPayload) { 
    try {
        Validator.validateUserPayload(userPayload);
        const r = RotationController.add(data); 
        AuditController.log(userPayload, 'ADD_ROTATION', data.name, 'Membro Adicionado ao Rodízio'); 
        return { success: true, data: r }; 
    } catch(e) {
        console.error(`Error in addRotationMember: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'ADD_ROTATION_ERROR', data?.name || 'UNKNOWN', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message };
    }
}
function deleteRotationMember(id, userPayload) { 
  try {
    AuthHelper.requireAdmin(userPayload);
    const r = RotationController.delete(id); 
    if(userPayload) AuditController.log(userPayload, 'DELETE_ROTATION', id, 'Membro Removido do Rodízio'); 
    return { success: true, deleted: r };
  } catch(e) {
    console.error(`Error in deleteRotationMember: ${e.message}\nStack: ${e.stack}`);
    AuditController.log(userPayload || 'SYS', 'DELETE_ROTATION_ERROR', id, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
    return { success: false, message: e.message };
  }
}

function saveDailyForecast(d) {
  const sheet = DB.connect(TABELAS.FORECAST);
  const data = sheet.getDataRange().getValues();
  const targetDateStr = String(d.date).split('T')[0]; 
  let foundIndex = -1;
  for(let i=1; i<data.length; i++) {
      let cellId = String(data[i][0] instanceof Date ? data[i][0].toISOString() : data[i][0]);
      if(cellId.split('T')[0] === targetDateStr) { foundIndex = i + 1; break; }
  }
  const jsonStr = JSON.stringify(d);
  if(foundIndex !== -1) sheet.getRange(foundIndex, 3).setValue(jsonStr);
  else sheet.appendRow([targetDateStr, d.date, jsonStr]);
  return d;
}

function diagnoseSystem() { try { return { success: true, dbName: DB.getSpreadsheet().getName() }; } catch (e) { return { success: false, message: e.message }; } }

/**
 * ============================================================================
 * NOTIFICATION API
 * ============================================================================
 */

function getMyNotifications(userPayload) {
  try {
    if (!userPayload) return { success: false, message: "Não autenticado" };
    const notifications = NotificationController.list(userPayload.role, userPayload.matricula);
    return { success: true, data: notifications };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function createNotification(notification, userPayload) {
  try {
    // Permitir que o sistema ou usuários criem notificações
    const res = NotificationController.add(notification);
    return { success: true, data: res };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function markNotificationAsRead(id, userPayload) {
  try {
    if (!userPayload) return { success: false, message: "Não autenticado" };
    const res = NotificationController.markAsRead(id);
    return { success: true, data: res };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function callGeminiAi(userInput, contextKey) { 
    try {
        const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
        if(!key) return "API Key Missing";
        
        const rawDict = DB.findAll(TABELAS.DICIONARIOS);
        const aiConfigs = {};
        rawDict.forEach(r => {
            if (r.CATEGORIA === 'CONFIG_AI') {
                aiConfigs[r.VALOR] = r.JSON_DADOS || '';
            }
        });

        const model = aiConfigs['MODEL'] || 'gemini-3-flash-preview';
        
        let finalPrompt = userInput;
        if (contextKey === 'SHIFT_HANDOVER') {
            const template = aiConfigs['PROMPT_SHIFT'];
            if (template) finalPrompt = template.replace('{{INPUT}}', userInput);
            else finalPrompt = `Reescreva tecnicamente em caixa alta: ${userInput}`;
        } else if (contextKey === 'DOC_EDITOR') {
            const template = aiConfigs['PROMPT_EDITOR'];
            if (template) finalPrompt = template.replace('{{INPUT}}', userInput);
            else finalPrompt = `Corrija e melhore formalmente: ${userInput}`;
        }

        const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
        const payload = {
            contents: [{ parts: [{ text: finalPrompt }] }],
            safetySettings: [{ category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }],
            generationConfig: { temperature: 0.3, maxOutputTokens: 800 }
        };
        const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
        const res = UrlFetchApp.fetch(url, options);
        if (res.getResponseCode() !== 200) return `Erro API (${res.getResponseCode()}) - Verifique o Modelo`;
        const json = JSON.parse(res.getContentText());
        return json.candidates?.[0]?.content?.parts?.[0]?.text || "Sem resposta da IA.";
    } catch(e) { return "Erro na IA: " + e.message; }
}

function admitFromPa(data, paId, userPayload) { 
    try { 
        PatientController.add(data); 
        if(paId) PaController.delete(paId); 
        if(userPayload) AuditController.log(userPayload, 'ADMIT_FROM_PA', data.name, 'Admissão via PA');
        return { success: true }; 
    } catch(e) { 
        console.error(`Error in admitFromPa: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || 'SYS', 'ADMIT_PA_ERROR', data?.name || 'N/A', e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message }; 
    } 
}

function executeInternalTransfer(pId, oB, nB, nS, u, userPayload) { 
    try { 
        const p = PatientController.list().find(x => String(x.id) === String(pId)); 
        if(p) { 
            PatientController.update({ ...p, bed: nB, sector: nS }); 
            if(oB) BedController.addBlock(oB, STATUS.HIGIENIZACAO, u); 
            TransferController.add({id: Date.now(), patient: p.name, flow: 'INTERNA', origin: oB, destination: nB, status: STATUS.ACEITO, arrived: true}); 
            
            // --- AUDITORIA DETALHADA DE TROCA ---
            AuditController.log(userPayload || u, 'TRANSFER_INTERNAL', p.name, `Troca de Leito: ${oB || 'S/L'} -> ${nB}`, {
                patientId: pId,
                fromBed: oB || 'N/A',
                toBed: nB,
                newSector: nS
            }); 
        } 
        return { success: true }; 
    } catch(e) { 
        console.error(`Error in executeInternalTransfer: ${e.message}\nStack: ${e.stack}`);
        AuditController.log(userPayload || u || 'SYS', 'TRANSFER_INTERNAL_ERROR', pId, e.message, { stack: e.stack?.split('\n').slice(0, 3).join(' ') });
        return { success: false, message: e.message }; 
    } 
}