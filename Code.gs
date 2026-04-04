/**
 * SYSGESTAO v2.0 - Sistema de Gestão Hospitalar
 * Backend Google Apps Script
 * 
 * Autor: Assistente de IA
 * Versão: 2.0.0
 */

// --- CONFIGURAÇÕES GERAIS ---
const CONFIG = {
  SS_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
  SHEETS: {
    PACIENTES: 'Pacientes',
    LEITOS: 'Leitos',
    MOVIMENTACAO: 'Movimentacao',
    USUARIOS: 'Usuarios',
    SETORES: 'Setores',
    EQUIPES: 'Equipes',
    LOGS: 'Logs_Auditoria'
  },
  VERSAO: '2.0.0'
};

// --- INICIALIZAÇÃO E MENU ---

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏥 SYSGESTAO v2')
    .addItem('🚀 Abrir Sistema', 'abrirSistema')
    .addItem('⚙️ Instalar/Atualizar DB', 'instalarBancoDados')
    .addItem('📖 Sobre', 'mostrarSobre')
    .addToUi();
  
  // Opcional: Rodar instalação automática na primeira abertura se desejar
  // instalarBancoDados(); 
}

function abrirSistema() {
  try {
    const template = HtmlService.createTemplateFromFile('index');
    const htmlOutput = template.evaluate()
      .setTitle('SYSGESTAO v2.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Carregando SYSGESTAO...');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro ao carregar sistema: ' + e.toString());
  }
}

function mostrarSobre() {
  SpreadsheetApp.getUi().alert('SYSGESTAO v2.0\nSistema de Gestão Hospitalar\nVersão: ' + CONFIG.VERSAO + '\nDesenvolvido com Google Apps Script');
}

// --- GERENCIAMENTO DE DADOS (DATABASE) ---

function instalarBancoDados() {
  const ss = SpreadsheetApp.openById(CONFIG.SS_ID);
  const sheetsToCreate = Object.values(CONFIG.SHEETS);
  
  sheetsToCreate.forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      setupCabecalhos(sheet, sheetName);
    }
  });
  
  garantirAdminPadrao();
  SpreadsheetApp.getUi().alert('✅ Banco de dados verificado e atualizado com sucesso!\nUsuário Admin padrão:\nMatrícula: 1\nSenha: 1234');
}

function setupCabecalhos(sheet, nome) {
  const headers = [];
  switch(nome) {
    case CONFIG.SHEETS.PACIENTES:
      headers.push(['ID_Paciente', 'Nome', 'DataNascimento', 'Genero', 'CPF', 'Convenio', 'Status', 'DataEntrada', 'LeitoID', 'Observacoes']);
      break;
    case CONFIG.SHEETS.LEITOS:
      headers.push(['ID_Leito', 'Nome', 'Setor', 'Tipo', 'Status', 'PacienteID', 'UltimaAtualizacao']);
      break;
    case CONFIG.SHEETS.MOVIMENTACAO:
      headers.push(['ID_Mov', 'DataHora', 'Tipo', 'PacienteID', 'Origem', 'Destino', 'Usuario', 'Observacao']);
      break;
    case CONFIG.SHEETS.USUARIOS:
      headers.push(['Matricula', 'Nome', 'Senha', 'Perfil', 'Ativo']);
      break;
    case CONFIG.SHEETS.SETORES:
      headers.push(['ID_Setor', 'Nome', 'Descricao', 'Ordem']);
      break;
    case CONFIG.SHEETS.EQUIPES:
      headers.push(['ID_Equipe', 'Nome', 'Membros', 'Escala']);
      break;
    case CONFIG.SHEETS.LOGS:
      headers.push(['Timestamp', 'Usuario', 'Acao', 'Detalhes']);
      break;
  }
  
  if (headers.length > 0 && sheet.getLastRow() === 0) {
    sheet.appendRow(headers[0]);
    sheet.getRange(1, 1, 1, headers[0].length).setFontWeight('bold').setBackground('#e8f0fe');
  }
}

function garantirAdminPadrao() {
  const sheet = getSheet(CONFIG.SHEETS.USUARIOS);
  // Verifica se já existe algum usuário
  if (sheet.getLastRow() < 2) {
    sheet.appendRow(['1', 'Administrador', '1234', 'ADMIN', 'TRUE']);
  }
}

// --- FUNÇÕES AUXILIARES DE PLANILHA ---

function getSheet(name) {
  const ss = SpreadsheetApp.openById(CONFIG.SS_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error(`Planilha "${name}" não encontrada. Execute "Instalar DB" no menu.`);
  }
  return sheet;
}

function logAuditoria(usuario, acao, detalhes) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LOGS);
    sheet.appendRow([new Date(), usuario, acao, detalhes]);
  } catch (e) {
    console.error('Erro ao gravar log:', e);
  }
}

// --- AUTENTICAÇÃO ---

function login(matricula, senha) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USUARIOS);
    const data = sheet.getDataRange().getValues();
    
    // Pula cabeçalho
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // [Matricula, Nome, Senha, Perfil, Ativo]
      if (String(row[0]) === String(matricula) && String(row[2]) === String(senha)) {
        if (row[4] === true || row[4] === 'TRUE' || row[4] === 1) {
          logAuditoria(row[1], 'LOGIN', 'Login realizado com sucesso');
          return {
            success: true,
            user: {
              matricula: row[0],
              nome: row[1],
              perfil: row[3]
            }
          };
        } else {
          return { success: false, message: 'Usuário inativo.' };
        }
      }
    }
    return { success: false, message: 'Matrícula ou senha inválidos.' };
  } catch (e) {
    return { success: false, message: 'Erro no servidor: ' + e.toString() };
  }
}

// --- API DE DADOS (CRUD) ---

function getDados(tabela) {
  try {
    const sheet = getSheet(tabela);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; // Apenas cabeçalho ou vazio
    
    const headers = data[0];
    const rows = data.slice(1);
    
    return rows.map((row, index) => {
      let obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i];
      });
      obj._id = row[0]; // ID padrão primeira coluna
      obj._rowIndex = index + 2; // Índice da linha na planilha (1-based + header)
      return obj;
    });
  } catch (e) {
    throw new Error('Erro ao buscar dados de ' + tabela + ': ' + e.toString());
  }
}

function salvarDados(tabela, dados) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: 'Sistema ocupado. Tente novamente.' };
  }
  
  try {
    const sheet = getSheet(tabela);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Se for novo registro (sem ID ou ID vazio)
    if (!dados._id || dados._id === '') {
      const newRow = headers.map(h => {
        // Gera ID automático se for a coluna de ID
        if (h.startsWith('ID_')) {
           return tabela.substring(0,1).toUpperCase() + '-' + new Date().getTime().toString().slice(-6);
        }
        // Campos de data/hora automáticos
        if (h === 'DataEntrada' || h === 'DataHora' || h === 'UltimaAtualizacao') {
          return new Date();
        }
        if (h === 'Status' && tabela === CONFIG.SHEETS.LEITOS) {
          return 'Livre';
        }
        if (h === 'Status' && tabela === CONFIG.SHEETS.PACIENTES) {
          return 'Internado';
        }
        return dados[h] || '';
      });
      
      sheet.appendRow(newRow);
      logAuditoria(Session.getActiveUser().getEmail(), 'INSERT', `Novo registro em ${tabela}: ${newRow[0]}`);
      return { success: true, message: 'Registro criado com sucesso!', id: newRow[0] };
      
    } else {
      // Atualizar existente
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      
      // Encontra linha pelo ID (primeira coluna)
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(dados._id)) {
          rowIndex = i + 1;
          break;
        }
      }
      
      if (rowIndex > 0) {
        const newRow = headers.map(h => {
          if (h === 'UltimaAtualizacao') return new Date();
          return dados[h] !== undefined ? dados[h] : '';
        });
        
        sheet.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
        logAuditoria(Session.getActiveUser().getEmail(), 'UPDATE', `Atualizado registro ${dados._id} em ${tabela}`);
        return { success: true, message: 'Registro atualizado com sucesso!' };
      } else {
        return { success: false, message: 'Registro não encontrado para atualização.' };
      }
    }
  } catch (e) {
    return { success: false, message: 'Erro ao salvar: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function excluirDados(tabela, id) {
  // Soft delete ou apenas log por segurança
  logAuditoria(Session.getActiveUser().getEmail(), 'DELETE_REQUEST', `Tentativa de exclusão ID ${id} em ${tabela}`);
  return { success: true, message: 'Solicitação registrada. Exclusão física desabilitada.' };
}

// --- FUNÇÕES ESPECÍFICAS DO NEGÓCIO ---

function getMapaLeitos() {
  const leitos = getDados(CONFIG.SHEETS.LEITOS);
  const pacientes = getDados(CONFIG.SHEETS.PACIENTES);
  
  return leitos.map(leito => {
    const paciente = leito.PacienteID ? pacientes.find(p => String(p._id) === String(leito.PacienteID)) : null;
    return {
      ...leito,
      pacienteInfo: paciente || null
    };
  });
}

function realizarTransferencia(dados) {
  // Dados esperados: { pacienteId, leitoOrigemId, leitoDestinoId, observacao }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: 'Sistema ocupado.' };

  try {
    const user = Session.getActiveUser().getEmail();
    const dataHora = new Date();

    // 1. Registrar Movimentação
    const mov = {
      _id: '',
      DataHora: dataHora,
      Tipo: 'TRANSFERENCIA',
      PacienteID: dados.pacienteId,
      Origem: dados.leitoOrigemId,
      Destino: dados.leitoDestinoId,
      Usuario: user,
      Observacao: dados.observacao || ''
    };
    salvarDados(CONFIG.SHEETS.MOVIMENTACAO, mov);

    // 2. Liberar Leito Origem
    const leitoOrigem = getDados(CONFIG.SHEETS.LEITOS).find(l => String(l._id) === String(dados.leitoOrigemId));
    if (leitoOrigem) {
      leitoOrigem.Status = 'Livre';
      leitoOrigem.PacienteID = '';
      leitoOrigem.UltimaAtualizacao = dataHora;
      salvarDados(CONFIG.SHEETS.LEITOS, leitoOrigem);
    }

    // 3. Ocupar Leito Destino
    const leitoDestino = getDados(CONFIG.SHEETS.LEITOS).find(l => String(l._id) === String(dados.leitoDestinoId));
    if (leitoDestino) {
      leitoDestino.Status = 'Ocupado';
      leitoDestino.PacienteID = dados.pacienteId;
      leitoDestino.UltimaAtualizacao = dataHora;
      salvarDados(CONFIG.SHEETS.LEITOS, leitoDestino);
    }

    // 4. Atualizar Paciente
    const pacientes = getDados(CONFIG.SHEETS.PACIENTES);
    const paciente = pacientes.find(p => String(p._id) === String(dados.pacienteId));
    if (paciente) {
      paciente.LeitoID = dados.leitoDestinoId;
      salvarDados(CONFIG.SHEETS.PACIENTES, paciente);
    }

    logAuditoria(user, 'TRANSFERENCIA', `Paciente ${dados.pacienteId} movido de ${dados.leitoOrigemId} para ${dados.leitoDestinoId}`);
    return { success: true, message: 'Transferência realizada com sucesso!' };

  } catch (e) {
    return { success: false, message: 'Erro na transferência: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function darAlta(pacienteId, observacao) {
  try {
    const pacientes = getDados(CONFIG.SHEETS.PACIENTES);
    const paciente = pacientes.find(p => String(p._id) === String(pacienteId));
    
    if (!paciente) return { success: false, message: 'Paciente não encontrado' };
    
    // Atualiza Paciente
    paciente.Status = 'ALTA';
    paciente.Observacoes = (paciente.Observacoes || '') + ' | Alta em: ' + new Date() + '. Motivo: ' + observacao;
    if (paciente.LeitoID) {
       // Se tiver leito, libera ele também
       const leitos = getDados(CONFIG.SHEETS.LEITOS);
       const leito = leitos.find(l => String(l._id) === String(paciente.LeitoID));
       if (leito) {
         leito.Status = 'Livre';
         leito.PacienteID = '';
         leito.UltimaAtualizacao = new Date();
         salvarDados(CONFIG.SHEETS.LEITOS, leito);
       }
       paciente.LeitoID = ''; // Desvincula leito
    }
    
    const res = salvarDados(CONFIG.SHEETS.PACIENTES, paciente);
    logAuditoria(Session.getActiveUser().getEmail(), 'ALTA', `Paciente ${paciente.Nome} recebeu alta`);
    
    return res;
  } catch (e) {
    return { success: false, message: 'Erro ao dar alta: ' + e.toString() };
  }
}

// Include handler para carregar CSS e JS separadamente no HTML
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
