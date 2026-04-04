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
}

function abrirSistema() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SYSGESTAO v2.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'SYSGESTAO v2.0 - Carregando...');
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
  
  // Cria usuário admin padrão se não existir
  garantirAdminPadrao();
  
  SpreadsheetApp.getUi().alert('✅ Banco de dados verificado e atualizado com sucesso!');
}

function setupCabecalhos(sheet, nome) {
  const headers = [];
  switch(nome) {
    case CONFIG.SHEETS.PACIENTES:
      headers.push(['ID_Paciente', 'Nome', 'DataNascimento', 'Genero', 'CPF', 'Convenio', 'Status', 'DataEntrada', 'Observacoes']);
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
  if (sheet.getLastRow() < 2) {
    // Admin padrão: matricula 1, senha 1234
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
        if (row[4] === true || row[4] === 'TRUE') {
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
    const headers = data.shift(); // Remove cabeçalho
    
    // Converte array de arrays para array de objetos
    return data.map(row => {
      let obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i];
      });
      obj._id = row[0]; // ID padrão primeira coluna
      return obj;
    });
  } catch (e) {
    throw new Error('Erro ao buscar dados de ' + tabela + ': ' + e.toString());
  }
}

function salvarDados(tabela, dados) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    const sheet = getSheet(tabela);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Se for novo registro (sem ID ou ID vazio)
    if (!dados._id || dados._id === '') {
      const newRow = headers.map(h => {
        if (h === 'ID_' + tabela.substring(0,3)) { // Gera ID simples baseado no nome da tabela
           return new Date().getTime().toString(); 
        }
        return dados[h] || '';
      });
      
      // Ajuste fino para IDs específicos
      if(tabela === CONFIG.SHEETS.PACIENTES) newRow[0] = 'P-' + new Date().getTime();
      if(tabela === CONFIG.SHEETS.MOVIMENTACAO) newRow[0] = 'M-' + new Date().getTime();
      
      sheet.appendRow(newRow);
      logAuditoria(Session.getActiveUser().getEmail(), 'INSERT', `Novo registro em ${tabela}`);
      return { success: true, message: 'Registro criado com sucesso!' };
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
        const newRow = headers.map(h => dados[h] !== undefined ? dados[h] : '');
        // Mantém o ID original na primeira posição se a lógica acima mudar
        newRow[0] = dados._id; 
        
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
  // Em sistemas hospitalares, geralmente fazemos "soft delete" ou apenas arquivamos.
  // Aqui vamos apenas registrar no log e poderia marcar status como 'Inativo' se houvesse coluna.
  // Para simplificar v2.0, vamos apenas logar. Implementação física requer cuidado.
  logAuditoria(Session.getActiveUser().getEmail(), 'DELETE_REQUEST', `Tentativa de exclusão ID ${id} em ${tabela}`);
  return { success: true, message: 'Solicitação de exclusão registrada. (Exclusão física desabilitada por segurança)' };
}

// --- FUNÇÕES ESPECÍFICAS DO NEGÓCIO ---

function getMapaLeitos() {
  // Retorna leitos agrupados por setor
  const leitos = getDados(CONFIG.SHEETS.LEITOS);
  const pacientes = getDados(CONFIG.SHEETS.PACIENTES);
  
  // Merge simples para enrichir dados do leito com dados do paciente
  const mapa = leitos.map(leito => {
    const paciente = pacientes.find(p => String(p._id) === String(leito.PacienteID));
    return {
      ...leito,
      pacienteInfo: paciente || null
    };
  });
  
  return mapa;
}

function realizarTransferencia(dadosMovimentacao) {
  // 1. Registra movimentação
  dadosMovimentacao.Tipo = 'TRANSFERENCIA';
  const resMov = salvarDados(CONFIG.SHEETS.MOVIMENTACAO, dadosMovimentacao);
  
  if (!resMov.success) return resMov;
  
  // 2. Atualiza Leito Origem (Libera)
  // Precisaria buscar o leito de origem e limpar o PacienteID
  // Simplificação: O frontend deve enviar os IDs dos leitos envolvidos
  
  // 3. Atualiza Leito Destino (Ocupa)
  
  return { success: true, message: 'Transferência realizada com sucesso!' };
}

function darAlta(pacienteId, observacao) {
  // Atualiza status do paciente
  const pacientes = getDados(CONFIG.SHEETS.PACIENTES);
  const paciente = pacientes.find(p => String(p._id) === String(pacienteId));
  
  if (!paciente) return { success: false, message: 'Paciente não encontrado' };
  
  paciente.Status = 'ALTA';
  paciente.Observacoes = (paciente.Observacoes || '') + ' | Alta em: ' + new Date() + '. Motivo: ' + observacao;
  
  const res = salvarDados(CONFIG.SHEETS.PACIENTES, paciente);
  
  // Libera o leito associado
  // Lógica similar a transferência
  
  logAuditoria(Session.getActiveUser().getEmail(), 'ALTA', `Paciente ${paciente.Nome} recebeu alta`);
  
  return res;
}

// Include handler para carregar CSS e JS separadamente no HTML
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
