/**
 * SYSGESTAO v2.0 - Sistema de Gestão Hospitalar
 * Backend Completo e Otimizado para Google Apps Script
 * 
 * Instruções:
 * 1. Copie este código para o arquivo Code.gs
 * 2. Execute a função 'instalarBancoDados' uma vez pelo menu ou manualmente
 * 3. Login padrão: Matrícula '1', Senha '1234'
 */

// ================= CONFIGURAÇÕES GERAIS =================
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
  VERSAO: '2.0.0-FINAL'
};

// ================= INICIALIZAÇÃO E MENU =================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏥 SYSGESTAO v2')
    .addItem('🚀 Abrir Sistema', 'abrirSistema')
    .addItem('⚙️ Instalar/Atualizar DB', 'instalarBancoDados')
    .addItem('📖 Sobre', 'mostrarSobre')
    .addToUi();
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
    SpreadsheetApp.getUi().alert('Erro crítico ao carregar: ' + e.toString());
  }
}

function mostrarSobre() {
  SpreadsheetApp.getUi().alert('SYSGESTAO v2.0\nSistema de Gestão Hospitalar\nVersão: ' + CONFIG.VERSAO + '\nDesenvolvido com Google Apps Script');
}

// ================= GERENCIAMENTO DE BANCO DE DADOS =================

function instalarBancoDados() {
  const ss = SpreadsheetApp.openById(CONFIG.SS_ID);
  const sheetsToCreate = Object.values(CONFIG.SHEETS);
  
  sheetsToCreate.forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      setupCabecalhos(sheet, sheetName);
      Logger.log('Planilha criada: ' + sheetName);
    }
  });
  
  garantirAdminPadrao();
  preencherSetoresPadrao();
  
  SpreadsheetApp.getUi().alert('✅ Banco de dados instalado/atualizado com sucesso!\n\nUsuário Admin:\nMatrícula: 1\nSenha: 1234');
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
    sheet.getRange(1, 1, 1, headers[0].length)
      .setFontWeight('bold')
      .setBackground('#e8f0fe')
      .setBorder(null, null, true, null, null, null);
  }
}

function garantirAdminPadrao() {
  const sheet = getSheet(CONFIG.SHEETS.USUARIOS);
  if (sheet.getLastRow() < 2) {
    sheet.appendRow(['1', 'Administrador', '1234', 'ADMIN', 'TRUE']);
  }
}

function preencherSetoresPadrao() {
  const sheet = getSheet(CONFIG.SHEETS.SETORES);
  if (sheet.getLastRow() < 2) {
    const setores = [
      ['S-UTI', 'UTI Adulto', 'Unidade de Terapia Intensiva', 1],
      ['S-ENF', 'Enfermaria', 'Enfermaria Geral', 2],
      ['S-ISO', 'Isolamento', 'Quartos de Isolamento', 3],
      ['S-OBS', 'Observação', 'Pronto Atendimento', 4]
    ];
    setores.forEach(s => sheet.appendRow(s));
  }
}

// ================= FUNÇÕES AUXILIARES =================

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

function gerarID(prefixo) {
  return prefixo + '-' + new Date().getTime().toString().slice(-6) + Math.floor(Math.random() * 100).toString().padStart(2, '0');
}

// ================= AUTENTICAÇÃO =================

function login(matricula, senha) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USUARIOS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // [Matricula, Nome, Senha, Perfil, Ativo]
      if (String(row[0]).trim() === String(matricula).trim() && String(row[2]) === String(senha)) {
        if (row[4] === true || row[4] === 'TRUE' || row[4] === 1) {
          logAuditoria(row[1], 'LOGIN', 'Login realizado com sucesso via WebApp');
          return {
            success: true,
            user: {
              matricula: row[0],
              nome: row[1],
              perfil: row[3]
            }
          };
        } else {
          return { success: false, message: 'Usuário inativo no sistema.' };
        }
      }
    }
    return { success: false, message: 'Matrícula ou senha inválidos.' };
  } catch (e) {
    return { success: false, message: 'Erro no servidor: ' + e.toString() };
  }
}

// ================= API DE DADOS (CRUD GENÉRICO) =================

function getDados(tabela) {
  try {
    const sheet = getSheet(tabela);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const headers = data[0];
    const rows = data.slice(1);
    
    return rows.map((row, index) => {
      let obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i];
      });
      obj._id = row[0];
      obj._rowIndex = index + 2;
      return obj;
    });
  } catch (e) {
    throw new Error('Erro ao buscar dados de ' + tabela + ': ' + e.toString());
  }
}

function salvarDados(tabela, dados) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return { success: false, message: 'Sistema ocupado. Tente novamente em alguns segundos.' };
  }
  
  try {
    const sheet = getSheet(tabela);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // INSERÇÃO
    if (!dados._id || dados._id === '' || dados._id === null) {
      const newRow = headers.map(h => {
        if (h.startsWith('ID_')) return gerarID(h.split('_')[1]);
        if (['DataEntrada', 'DataHora', 'UltimaAtualizacao', 'Timestamp'].includes(h)) return new Date();
        if (h === 'Status' && tabela === CONFIG.SHEETS.LEITOS) return 'Livre';
        if (h === 'Status' && tabela === CONFIG.SHEETS.PACIENTES) return 'Internado';
        if (h === 'Ativo' && tabela === CONFIG.SHEETS.USUARIOS) return 'TRUE';
        return dados[h] || '';
      });
      
      sheet.appendRow(newRow);
      logAuditoria(Session.getActiveUser().getEmail(), 'INSERT', `Novo registro em ${tabela}: ${newRow[0]}`);
      return { success: true, message: 'Registro criado!', id: newRow[0] };
      
    } 
    // ATUALIZAÇÃO
    else {
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(dados._id)) {
          rowIndex = i + 1;
          break;
        }
      }
      
      if (rowIndex > 0) {
        const newRow = headers.map(h => {
          if (h === 'UltimaAtualizacao') return new Date();
          return dados.hasOwnProperty(h) ? dados[h] : data[rowIndex-1][headers.indexOf(h)];
        });
        
        sheet.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
        logAuditoria(Session.getActiveUser().getEmail(), 'UPDATE', `Atualizado ${dados._id} em ${tabela}`);
        return { success: true, message: 'Registro atualizado!' };
      } else {
        return { success: false, message: 'Registro não encontrado.' };
      }
    }
  } catch (e) {
    return { success: false, message: 'Erro ao salvar: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ================= REGRAS DE NEGÓCIO ESPECÍFICAS =================

function getDashboardData() {
  try {
    const leitos = getDados(CONFIG.SHEETS.LEITOS);
    const pacientes = getDados(CONFIG.SHEETS.PACIENTES);
    const logs = getDados(CONFIG.SHEETS.LOGS);
    
    const total = leitos.length;
    const ocupados = leitos.filter(l => l.Status === 'Ocupado').length;
    const livres = total - ocupados;
    
    // Últimos 5 logs
    const recentLogs = logs.slice(-5).reverse().map(l => ({
      time: l.Timestamp instanceof Date ? l.Timestamp.toLocaleTimeString() : l.Timestamp,
      user: l.Usuario,
      action: l.Acao,
      details: l.Detalhes
    }));
    
    return {
      success: true,
      stats: { total, ocupados, livres, filaPA: 0 }, // Fila PA seria outra lógica
      recentLogs: recentLogs
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getMapaLeitosCompleto() {
  const leitos = getDados(CONFIG.SHEETS.LEITOS);
  const pacientes = getDados(CONFIG.SHEETS.PACIENTES);
  
  return leitos.map(leito => {
    const paciente = leito.PacienteID ? pacientes.find(p => String(p._id) === String(leito.PacienteID)) : null;
    return { ...leito, pacienteInfo: paciente };
  });
}

function realizarAdmissao(dadosPaciente, idLeito) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Sistema ocupado.' };

  try {
    // 1. Verificar se leito está livre
    const leitos = getDados(CONFIG.SHEETS.LEITOS);
    const leito = leitos.find(l => String(l._id) === String(idLeito));
    
    if (!leito) return { success: false, message: 'Leito não encontrado.' };
    if (leito.Status !== 'Livre') return { success: false, message: 'Leito já está ocupado.' };

    // 2. Criar Paciente
    dadosPaciente.LeitoID = idLeito;
    dadosPaciente.Status = 'Internado';
    const resPaciente = salvarDados(CONFIG.SHEETS.PACIENTES, dadosPaciente);
    
    if (!resPaciente.success) throw new Error(resPaciente.message);
    const pacienteId = resPaciente.id;

    // 3. Ocupar Leito
    leito.Status = 'Ocupado';
    leito.PacienteID = pacienteId;
    leito.UltimaAtualizacao = new Date();
    // Remover _id gerado automaticamente se houver conflito na atualização
    delete leito._id; 
    // Precisamos passar o ID original para a função saber que é update
    // A função salvarDados usa o campo _id do objeto passado.
    // Vamos reconstruir o objeto garantindo o ID correto.
    const leitoUpdate = {
      _id: idLeito,
      Nome: leito.Nome,
      Setor: leito.Setor,
      Tipo: leito.Tipo,
      Status: 'Ocupado',
      PacienteID: pacienteId,
      UltimaAtualizacao: new Date()
    };
    
    const resLeito = salvarDados(CONFIG.SHEETS.LEITOS, leitoUpdate);
    if (!resLeito.success) throw new Error(resLeito.message);

    // 4. Log de Movimentação
    salvarDados(CONFIG.SHEETS.MOVIMENTACAO, {
      _id: '',
      Tipo: 'ADMISSAO',
      PacienteID: pacienteId,
      Destino: idLeito,
      Observacao: 'Admissão inicial'
    });

    logAuditoria(Session.getActiveUser().getEmail(), 'ADMISSAO', `Paciente ${pacienteId} admitido no leito ${idLeito}`);
    
    return { success: true, message: 'Paciente admitido com sucesso!' };

  } catch (e) {
    return { success: false, message: 'Erro na admissão: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function realizarTransferencia(pacienteId, leitoOrigemId, leitoDestinoId, observacao) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Sistema ocupado.' };

  try {
    const user = Session.getActiveUser().getEmail();
    const now = new Date();

    // Validar leito destino
    const leitoDestinoRaw = getDados(CONFIG.SHEETS.LEITOS).find(l => String(l._id) === String(leitoDestinoId));
    if (!leitoDestinoRaw || leitoDestinoRaw.Status !== 'Livre') {
      return { success: false, message: 'Leito de destino indisponível.' };
    }

    // 1. Liberar Origem
    const leitoOrigemUpdate = { _id: leitoOrigemId, Status: 'Livre', PacienteID: '' };
    salvarDados(CONFIG.SHEETS.LEITOS, leitoOrigemUpdate);

    // 2. Ocupar Destino
    const leitoDestinoUpdate = { _id: leitoDestinoId, Status: 'Ocupado', PacienteID: pacienteId };
    salvarDados(CONFIG.SHEETS.LEITOS, leitoDestinoUpdate);

    // 3. Atualizar Paciente
    const pacientes = getDados(CONFIG.SHEETS.PACIENTES);
    const paciente = pacientes.find(p => String(p._id) === String(pacienteId));
    if (paciente) {
      salvarDados(CONFIG.SHEETS.PACIENTES, { _id: pacienteId, LeitoID: leitoDestinoId });
    }

    // 4. Log
    salvarDados(CONFIG.SHEETS.MOVIMENTACAO, {
      _id: '', DataHora: now, Tipo: 'TRANSFERENCIA',
      PacienteID: pacienteId, Origem: leitoOrigemId, Destino: leitoDestinoId,
      Usuario: user, Observacao: observacao
    });

    logAuditoria(user, 'TRANSFERENCIA', `${pacienteId}: ${leitoOrigemId} -> ${leitoDestinoId}`);
    return { success: true, message: 'Transferência realizada!' };

  } catch (e) {
    return { success: false, message: 'Erro: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function realizarAlta(pacienteId, tipo, observacao) {
  try {
    const pacientes = getDados(CONFIG.SHEETS.PACIENTES);
    const paciente = pacientes.find(p => String(p._id) === String(pacienteId));
    
    if (!paciente) return { success: false, message: 'Paciente não encontrado.' };

    // Liberar leito se existir
    if (paciente.LeitoID) {
      const leitoUpdate = { _id: paciente.LeitoID, Status: 'Livre', PacienteID: '' };
      salvarDados(CONFIG.SHEETS.LEITOS, leitoUpdate);
    }

    // Atualizar Paciente
    const statusFinal = tipo === 'OBITO' ? 'ÓBITO' : 'ALTA';
    salvarDados(CONFIG.SHEETS.PACIENTES, {
      _id: pacienteId,
      Status: statusFinal,
      Observacoes: (paciente.Observacoes || '') + `\n[${tipo}] ${new Date().toLocaleString()}: ${observacao}`
    });

    // Log
    salvarDados(CONFIG.SHEETS.MOVIMENTACAO, {
      _id: '', DataHora: new Date(), Tipo: tipo,
      PacienteID: pacienteId, Observacao: observacao
    });

    logAuditoria(Session.getActiveUser().getEmail(), tipo, `Paciente ${paciente.Nome} (${pacienteId})`);
    return { success: true, message: `${tipo} registrada com sucesso!` };

  } catch (e) {
    return { success: false, message: 'Erro: ' + e.toString() };
  }
}

// ================= UTILITÁRIO DE INCLUDE =================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
