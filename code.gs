/**
 * ============================================================
 * ESCOLA DE MÚSICA - WEB APP (API) - v2.3
 * ============================================================
 * CORREÇÕES CORS:
 *  - doPost aceita text/plain (não dispara preflight)
 *  - doGet também processa ?data=<json> como fallback
 *  - doOptions presente (fallback p/ proxies)
 *  - Toda resposta é JSON puro
 *
 * IMPLANTAÇÃO (obrigatório):
 *  1. Implantar → Gerenciar implantações → Editar
 *  2. Executar como: "Eu"
 *  3. Quem tem acesso: "Qualquer pessoa" (NÃO "com conta Google")
 *  4. Implantar → copiar URL /exec → colar em GAS_URL no HTML
 * ============================================================
 */

const SHEET_ID = '1FRtD7cbV2tvK2og4NDkyWi8V_hziyERNZoSh20gG980';

/**
 * ID opcional da pasta do Google Drive onde os arquivos serão salvos.
 * - Vazio ('') → cria/usa "Escola de Música - Arquivos" na raiz do Drive.
 * - Para pasta específica, cole o ID dela aqui.
 */
const PASTA_DRIVE_ID = '';

const ABAS = {
  alunos: 'ALUNOS',
  matriculas: 'MATRICULAS',
  professores: 'PROFESSORES',
  avaliacoes: 'AVALIACOES',
  presencas: 'PRESENCAS',
  conteudos: 'CONTEUDOS',
  observacoes: 'OBSERVACOES',
  relatorios: 'RELATORIOS',
  logs: 'LOGS',
  usuarios: 'USUARIOS',
  arquivos: 'ARQUIVOS',
  mensagens: 'MENSAGENS',
  notificacoes: 'NOTIFICACOES'
};

// ============================================================
// ENTRY POINTS (CORS-safe)
// ============================================================
function doOptions(e) {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    let dados = null;

    if (e && e.postData && e.postData.contents) {
      const raw = e.postData.contents;
      try { dados = JSON.parse(raw); }
      catch (p) {
        if (e.parameter && e.parameter.data) {
          try { dados = JSON.parse(e.parameter.data); }
          catch (e2) { return respostaJSON({ success: false, error: 'JSON inválido no campo data' }); }
        } else {
          return respostaJSON({ success: false, error: 'JSON inválido: ' + p.message });
        }
      }
    } else if (e && e.parameter && e.parameter.data) {
      try { dados = JSON.parse(e.parameter.data); }
      catch (p) { return respostaJSON({ success: false, error: 'JSON inválido: ' + p.message }); }
    }

    if (!dados || !dados.acao) {
      return respostaJSON({ success: false, error: 'Ação não especificada' });
    }
    return processarAcao(dados.acao, dados.payload || {});
  } catch (err) {
    logErro('doPost', err);
    return respostaJSON({ success: false, error: err.message });
  }
}

function doGet(e) {
  try {
    if (!e || !e.parameter || !e.parameter.data) {
      return respostaJSON({
        success: true,
        message: '🎵 API Escola de Música online (v2.3)',
        timestamp: new Date().toISOString(),
        acoes: listarAcoes()
      });
    }
    let dados;
    try { dados = JSON.parse(e.parameter.data); }
    catch (p) { return respostaJSON({ success: false, error: 'JSON inválido: ' + p.message }); }
    return processarAcao(dados.acao, dados.payload || {});
  } catch (err) {
    logErro('doGet', err);
    return respostaJSON({ success: false, error: err.message });
  }
}

// ============================================================
// ROTEADOR
// ============================================================
function processarAcao(acao, payload) {
  Logger.log('📥 Ação: ' + acao + ' | Payload: ' + JSON.stringify(payload));

  const acoes = {
    // LOGIN / USUÁRIOS
    'login': loginUsuario,
    'logout': logoutUsuario,
    'criarUsuario': criarUsuario,
    'listarUsuarios': listarUsuarios,
    'atualizarSenha': atualizarSenha,
    'criarUsuarioAluno': criarUsuarioAluno,
    'buscarUsuarioPorNome': buscarUsuarioPorNome,

    // Alunos
    'salvarAluno': salvarAluno,
    'atualizarAluno': atualizarAluno,
    'removerAluno': removerAluno,
    'listarAlunos': listarAlunos,

    // Professores
    'salvarProfessor': salvarProfessor,
    'atualizarProfessor': atualizarProfessor,
    'removerProfessor': removerProfessor,
    'listarProfessores': listarProfessores,
    'salvarValorProfessor': salvarValorProfessor,

    // Matrículas
    'salvarMatricula': salvarMatricula,
    'atualizarMatricula': atualizarMatricula,
    'removerMatricula': removerMatricula,
    'listarMatriculas': listarMatriculas,

    // Avaliações
    'salvarAvaliacoes': salvarAvaliacoes,
    'salvarNotas': salvarAvaliacoes,
    'listarAvaliacoes': listarAvaliacoes,

    // Presenças
    'salvarPresenca': salvarPresenca,
    'listarPresencas': listarPresencas,

    // Conteúdos
    'salvarConteudo': salvarConteudo,
    'listarConteudos': listarConteudos,

    // Observações
    'salvarObservacao': salvarObservacao,
    'listarObservacoes': listarObservacoes,

    // Relatórios
    'salvarRelatorio': salvarRelatorio,
    'enviarRelatorio': salvarRelatorio,
    'marcarRelatorioVisto': marcarRelatorioVisto,
    'listarRelatorios': listarRelatorios,

    // Atribuição
    'atribuirCurso': atribuirCurso,
    'removerCurso': removerCurso,
    'solicitarRelatorio': solicitarRelatorio,

    // LOGS
    'registrarLogin': registrarLogin,
    'listarLogs': listarLogs,
    'garantirAbaLogs': garantirAbaLogsAPI,

    // ALUNO
    'dadosAlunoCompleto': dadosAlunoCompleto,
    'garantirAbaAluno': garantirAbaAlunoAPI,

    // COMUNICAÇÃO
    'enviarMensagemAluno': enviarMensagemAluno,
    'uploadArquivo': uploadArquivoGAS,
    'listarMensagens': listarMensagens,
    'listarArquivos': listarArquivos,
    'listarNotificacoes': listarNotificacoes,
    'marcarNotificacoesLidas': marcarNotificacoesLidas,

    // Utilitários
    'ping': function() { return respostaJSON({ success: true, message: 'pong' }); },
    'listarCursos': listarCursosAPI,
    'sincronizar': sincronizarTudo,
    'limparAba': limparAba,
    'criarAbasComunicacao': criarAbasComunicacaoAPI
  };

  const fn = acoes[acao];
  if (!fn) {
    return respostaJSON({ success: false, error: 'Ação desconhecida: ' + acao, acoes: Object.keys(acoes) });
  }
  try { return fn(payload); }
  catch (err) {
    logErro('processarAcao:' + acao, err);
    return respostaJSON({ success: false, error: err.message, acao: acao });
  }
}

// ============================================================
// HELPERS BÁSICOS
// ============================================================
function respostaJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function logErro(ctx, err) {
  Logger.log('❌ [' + ctx + '] ' + err.message + ' | Stack: ' + (err.stack || 'sem stack'));
}

function getSS() { return SpreadsheetApp.openById(SHEET_ID); }

function getSheet(chave) {
  const sheet = getSS().getSheetByName(ABAS[chave]);
  if (!sheet) throw new Error('Aba não encontrada: ' + ABAS[chave]);
  return sheet;
}

function listarAcoes() {
  return Object.keys({
    login:1, logout:1, criarUsuario:1, criarUsuarioAluno:1, listarUsuarios:1, atualizarSenha:1,
    salvarAluno:1, salvarProfessor:1, salvarMatricula:1, salvarAvaliacoes:1,
    salvarPresenca:1, salvarObservacao:1, salvarRelatorio:1, registrarLogin:1,
    dadosAlunoCompleto:1, enviarMensagemAluno:1, uploadArquivo:1,
    listarNotificacoes:1, marcarNotificacoesLidas:1, ping:1
  });
}

function gerarID(prefixo) {
  return prefixo + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 10000);
}

function formatarData(data) {
  if (!data) return '';
  if (typeof data === 'string') {
    if (/^\d{2}-\d{2}-\d{4}$/.test(data)) return data;
    if (/^\d{4}-\d{2}-\d{2}/.test(data)) {
      const p = data.split('T')[0].split('-');
      return p[2] + '-' + p[1] + '-' + p[0];
    }
    const d = new Date(data);
    if (!isNaN(d.getTime())) {
      return String(d.getDate()).padStart(2,'0') + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + d.getFullYear();
    }
    return data;
  }
  if (data instanceof Date) {
    return String(data.getDate()).padStart(2,'0') + '-' + String(data.getMonth()+1).padStart(2,'0') + '-' + data.getFullYear();
  }
  return '';
}

function dataHoje() { return formatarData(new Date()); }

function dataHoraAgora() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
}

function getHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
}

function getRegistros(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const headers = getHeaders(sheet);
  const valores = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return valores.map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function encontrarLinha(sheet, criterios) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const headers = getHeaders(sheet);
  const valores = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const colIndex = {};
  headers.forEach(function(h, i) { colIndex[h] = i; });
  for (let r = 0; r < valores.length; r++) {
    let match = true;
    const campos = Object.keys(criterios);
    for (let k = 0; k < campos.length; k++) {
      const campo = campos[k];
      const idx = colIndex[campo];
      if (idx === undefined) { match = false; break; }
      if (String(valores[r][idx]).trim() !== String(criterios[campo]).trim()) { match = false; break; }
    }
    if (match) return r + 2;
  }
  return -1;
}

function appendRegistro(sheet, valores) { sheet.appendRow(valores); return sheet.getLastRow(); }
function atualizarLinha(sheet, linha, valores) { sheet.getRange(linha, 1, 1, valores.length).setValues([valores]); }
function removerLinha(sheet, linha) { sheet.deleteRow(linha); }

function s(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }

// ============================================================
// HASH DE SENHA
// ============================================================
function hashSenha(senha, salt) {
  const raw = String(salt || '') + '|' + String(senha || '');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digest.map(function(b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function gerarSalt() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 12);
}

// ============================================================
// GARANTIR ABAS
// ============================================================
function garantirAbaLogs() {
  const ss = getSS();
  let sheet = ss.getSheetByName(ABAS.logs);
  if (!sheet) {
    sheet = ss.insertSheet(ABAS.logs);
    sheet.appendRow(['ID_Log', 'Nome', 'Perfil', 'Data_Hora', 'Acao']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, 5).setValues([['ID_Log', 'Nome', 'Perfil', 'Data_Hora', 'Acao']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function garantirAbaUsuarios() {
  const ss = getSS();
  let sheet = ss.getSheetByName(ABAS.usuarios);
  const headersDesejados = ['ID_Usuario', 'Nome', 'Perfil', 'Usuario', 'Senha_Hash', 'Salt', 'Status', 'Data_Criacao', 'Aluno_Vinculado'];
  if (!sheet) {
    sheet = ss.insertSheet(ABAS.usuarios);
    sheet.appendRow(headersDesejados);
    sheet.getRange(1, 1, 1, headersDesejados.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }
  const headers = getHeaders(sheet);
  if (headers.length === 0) {
    sheet.getRange(1, 1, 1, headersDesejados.length).setValues([headersDesejados]);
    sheet.getRange(1, 1, 1, headersDesejados.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function garantirAbaArquivos() {
  const ss = getSS();
  let sheet = ss.getSheetByName(ABAS.arquivos);
  const headersDesejados = ['ID','Aluno','Curso','Professor','Nome_Arquivo','URL','File_ID','Mensagem','Data'];
  if (!sheet) {
    sheet = ss.insertSheet(ABAS.arquivos);
    sheet.appendRow(headersDesejados);
    sheet.getRange(1, 1, 1, headersDesejados.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }
  const headers = getHeaders(sheet);
  if (headers.length === 0) {
    sheet.getRange(1, 1, 1, headersDesejados.length).setValues([headersDesejados]);
    sheet.getRange(1, 1, 1, headersDesejados.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function garantirAbaMensagens() {
  const ss = getSS();
  let sheet = ss.getSheetByName(ABAS.mensagens);
  const headersDesejados = ['ID','Aluno','Curso','Professor','Assunto','Mensagem','Lida','Data'];
  if (!sheet) {
    sheet = ss.insertSheet(ABAS.mensagens);
    sheet.appendRow(headersDesejados);
    sheet.getRange(1, 1, 1, headersDesejados.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }
  const headers = getHeaders(sheet);
  if (headers.length === 0) {
    sheet.getRange(1, 1, 1, headersDesejados.length).setValues([headersDesejados]);
    sheet.getRange(1, 1, 1, headersDesejados.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function garantirAbaNotificacoes() {
  const ss = getSS();
  let sheet = ss.getSheetByName(ABAS.notificacoes);
  const headersDesejados = ['ID','Aluno','Tipo','Titulo','Descricao','Curso','Professor','Lida','Data'];
  if (!sheet) {
    sheet = ss.insertSheet(ABAS.notificacoes);
    sheet.appendRow(headersDesejados);
    sheet.getRange(1, 1, 1, headersDesejados.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }
  const headers = getHeaders(sheet);
  if (headers.length === 0) {
    sheet.getRange(1, 1, 1, headersDesejados.length).setValues([headersDesejados]);
    sheet.getRange(1, 1, 1, headersDesejados.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function garantirAbaAluno(nomeAluno) {
  const nome = s(nomeAluno).toUpperCase();
  if (!nome) throw new Error('Nome do aluno é obrigatório');

  const ss = getSS();
  const nomeAba = ('ALUNO_' + nome).substring(0, 100).replace(/[\\/?*\[\]:]/g, '');
  let sheet = ss.getSheetByName(nomeAba);

  if (!sheet) {
    sheet = ss.insertSheet(nomeAba);
    sheet.getRange(1, 1).setValue('📊 FICHA DO ALUNO');
    sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
    sheet.getRange(3, 1, 1, 2).setValues([['Nome', nome]]).setFontWeight('bold');
    sheet.getRange(4, 1, 1, 2).setValues([['Criado_Em', dataHoje()]]);
    sheet.getRange(6, 1, 1, 4).setValues([['Curso', 'Professor', 'Status', 'Data_Matrícula']]).setFontWeight('bold');
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 140);
  }
  return sheet;
}

function sincronizarAbaAluno(nomeAluno) {
  const nome = s(nomeAluno).toUpperCase();
  const sheet = garantirAbaAluno(nome);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 6) sheet.getRange(6, 1, lastRow - 5, 8).clearContent();

  sheet.getRange(6, 1, 1, 4).setValues([['Curso', 'Professor', 'Status', 'Data_Matrícula']]).setFontWeight('bold');
  const matriculas = getRegistros(getSheet('matriculas')).filter(function(m) { return s(m.Aluno).toUpperCase() === nome; });
  let row = 7;
  matriculas.forEach(function(m) {
    sheet.getRange(row, 1, 1, 4).setValues([[s(m.Curso), s(m.Professor), s(m.Status), s(m.Data_Matricula)]]);
    row++;
  });

  row += 1;
  sheet.getRange(row, 1, 1, 5).setValues([['Curso','Bimestre','Nota','Professor','Data_Lançamento']]).setFontWeight('bold');
  row++;
  const avals = getRegistros(getSheet('avaliacoes')).filter(function(a) { return s(a.Aluno).toUpperCase() === nome; });
  avals.forEach(function(a) {
    sheet.getRange(row, 1, 1, 5).setValues([[s(a.Curso), s(a.Bimestre), a.Nota, s(a.Professor), s(a.Data_Lancamento)]]);
    row++;
  });

  row += 1;
  sheet.getRange(row, 1, 1, 4).setValues([['Curso','Data','Presente','Professor']]).setFontWeight('bold');
  row++;
  getRegistros(getSheet('presencas')).forEach(function(p) {
    const presentes = s(p.Presentes).toUpperCase().split(',').map(function(x) { return x.trim(); });
    if (presentes.indexOf(nome) !== -1) {
      sheet.getRange(row, 1, 1, 4).setValues([[s(p.Curso), s(p.Data), 'SIM', s(p.Professor)]]);
      row++;
    }
  });

  row += 1;
  sheet.getRange(row, 1, 1, 4).setValues([['Curso','Bimestre','Data','Relatório']]).setFontWeight('bold');
  row++;
  getRegistros(getSheet('relatorios')).filter(function(r) { return s(r.Aluno).toUpperCase() === nome; }).forEach(function(r) {
    sheet.getRange(row, 1, 1, 4).setValues([[s(r.Curso), s(r.Bimestre), s(r.Data), s(r.Relatorio).substring(0, 500)]]);
    row++;
  });

  return { success: true, aba: sheet.getName(), matriculas: matriculas.length };
}

// ============================================================
// LOGIN / USUÁRIOS
// ============================================================
function criarUsuario(payload) {
  const nome = payload.nome, perfil = payload.perfil, usuario = payload.usuario, senha = payload.senha, alunoVinculado = payload.alunoVinculado;
  if (!nome || !perfil || !usuario || !senha) {
    return respostaJSON({ success: false, error: 'Nome, perfil, usuário e senha são obrigatórios' });
  }
  if (String(senha).length < 4) {
    return respostaJSON({ success: false, error: 'Senha deve ter ao menos 4 caracteres' });
  }

  const sheet = garantirAbaUsuarios();
  const headers = getHeaders(sheet);
  const usuarioLimpo = s(usuario).toLowerCase();
  if (encontrarLinha(sheet, { 'Usuario': usuarioLimpo }) > -1) {
    return respostaJSON({ success: false, error: 'Usuário já existe: ' + usuarioLimpo });
  }
  const salt = gerarSalt();
  const hash = hashSenha(senha, salt);
  const idUser = gerarID('USR');
  const dataCriacao = dataHoje();

  const valores = headers.map(function(h) {
    switch (h) {
      case 'ID_Usuario': return idUser;
      case 'Nome': return s(nome).toUpperCase();
      case 'Perfil': return s(perfil).toLowerCase();
      case 'Usuario': return usuarioLimpo;
      case 'Senha_Hash': return hash;
      case 'Salt': return salt;
      case 'Status': return 'ATIVO';
      case 'Data_Criacao': return dataCriacao;
      case 'Aluno_Vinculado': return s(alunoVinculado).toUpperCase();
      default: return '';
    }
  });
  appendRegistro(sheet, valores);
  return respostaJSON({ success: true, message: 'Usuário criado: ' + usuarioLimpo, id: idUser, usuario: usuarioLimpo, perfil: s(perfil).toLowerCase() });
}

function criarUsuarioAluno(payload) {
  const nome = payload.nome, usuario = payload.usuario, senha = payload.senha;
  if (!nome || !usuario || !senha) {
    return respostaJSON({ success: false, error: 'Nome, usuário e senha são obrigatórios' });
  }
  if (String(senha).length < 4) {
    return respostaJSON({ success: false, error: 'Senha deve ter ao menos 4 caracteres' });
  }

  const alunoSheet = getSheet('alunos');
  const alunoLinha = encontrarLinha(alunoSheet, { 'Nome': s(nome).toUpperCase() });
  if (alunoLinha < 0) {
    return respostaJSON({ success: false, error: 'Aluno não encontrado na aba ALUNOS: ' + nome });
  }

  const sheet = garantirAbaUsuarios();
  const usuarioLimpo = s(usuario).toLowerCase();
  if (encontrarLinha(sheet, { 'Usuario': usuarioLimpo }) > -1) {
    return respostaJSON({ success: false, error: 'Usuário já existe: ' + usuarioLimpo });
  }

  const headers = getHeaders(sheet);
  const salt = gerarSalt();
  const hash = hashSenha(senha, salt);
  const idUser = gerarID('USR');
  const dataCriacao = dataHoje();

  const valores = headers.map(function(h) {
    switch (h) {
      case 'ID_Usuario': return idUser;
      case 'Nome': return s(nome).toUpperCase();
      case 'Perfil': return 'aluno';
      case 'Usuario': return usuarioLimpo;
      case 'Senha_Hash': return hash;
      case 'Salt': return salt;
      case 'Status': return 'ATIVO';
      case 'Data_Criacao': return dataCriacao;
      case 'Aluno_Vinculado': return s(nome).toUpperCase();
      default: return '';
    }
  });
  appendRegistro(sheet, valores);

  let abaAluno = '';
  try {
    const info = sincronizarAbaAluno(nome);
    abaAluno = info.aba;
  } catch (e) {
    Logger.log('⚠️ Falha ao criar aba individual: ' + e.message);
  }
  return respostaJSON({ success: true, message: 'Usuário de aluno criado: ' + usuarioLimpo, id: idUser, usuario: usuarioLimpo, abaAluno: abaAluno });
}

function listarUsuarios(payload) {
  const sheet = garantirAbaUsuarios();
  const registros = getRegistros(sheet).map(function(r) {
    return {
      id: s(r.ID_Usuario), nome: s(r.Nome), perfil: s(r.Perfil), usuario: s(r.Usuario),
      status: s(r.Status), dataCriacao: s(r.Data_Criacao), alunoVinculado: s(r.Aluno_Vinculado)
    };
  });
  return respostaJSON({ success: true, total: registros.length, usuarios: registros });
}

function atualizarSenha(payload) {
  const usuario = payload.usuario, senhaNova = payload.senhaNova;
  if (!usuario || !senhaNova) return respostaJSON({ success: false, error: 'Usuário e nova senha são obrigatórios' });
  const sheet = garantirAbaUsuarios();
  const headers = getHeaders(sheet);
  const linha = encontrarLinha(sheet, { 'Usuario': s(usuario).toLowerCase() });
  if (linha < 0) return respostaJSON({ success: false, error: 'Usuário não encontrado' });

  const idxHash = headers.indexOf('Senha_Hash');
  const idxSalt = headers.indexOf('Salt');
  const novoSalt = gerarSalt();
  const novoHash = hashSenha(senhaNova, novoSalt);

  sheet.getRange(linha, idxSalt + 1).setValue(novoSalt);
  sheet.getRange(linha, idxHash + 1).setValue(novoHash);
  return respostaJSON({ success: true, message: 'Senha atualizada' });
}

function buscarUsuarioPorNome(payload) {
  const nome = payload.nome;
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  const sheet = garantirAbaUsuarios();
  const linha = encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() });
  if (linha < 0) return respostaJSON({ success: true, encontrado: false });
  const headers = getHeaders(sheet);
  const valores = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const obj = {};
  headers.forEach(function(h, i) { obj[h] = valores[i]; });
  return respostaJSON({
    success: true, encontrado: true,
    usuario: {
      id: s(obj.ID_Usuario), nome: s(obj.Nome), perfil: s(obj.Perfil),
      usuario: s(obj.Usuario), status: s(obj.Status), alunoVinculado: s(obj.Aluno_Vinculado)
    }
  });
}

function loginUsuario(payload) {
  try {
    const usuario = payload.usuario, senha = payload.senha, perfil = payload.perfil;
    if (!usuario || !senha) return respostaJSON({ success: false, error: 'Usuário e senha são obrigatórios' });

    const sheet = garantirAbaUsuarios();
    const headers = getHeaders(sheet);
    const linha = encontrarLinha(sheet, { 'Usuario': s(usuario).toLowerCase() });
    if (linha < 0) return respostaJSON({ success: false, error: 'Usuário ou senha inválidos' });

    const valores = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = valores[i]; });

    if (s(obj.Status).toUpperCase() !== 'ATIVO') {
      return respostaJSON({ success: false, error: 'Usuário inativo. Contate a secretaria.' });
    }

    const salt = s(obj.Salt);
    const hashArmazenado = s(obj.Senha_Hash);
    const hashTentativa = hashSenha(senha, salt);
    if (hashTentativa !== hashArmazenado) {
      return respostaJSON({ success: false, error: 'Usuário ou senha inválidos' });
    }

    const perfilSalvo = s(obj.Perfil).toLowerCase();
    if (perfil && perfilSalvo !== s(perfil).toLowerCase()) {
      return respostaJSON({ success: false, error: 'Perfil não corresponde. Este usuário é: ' + perfilSalvo });
    }

    return respostaJSON({
      success: true, message: 'Login OK',
      token: gerarID('TOKEN'),
      nome: s(obj.Nome).toUpperCase(),
      perfil: perfilSalvo,
      usuario: s(obj.Usuario),
      alunoVinculado: s(obj.Aluno_Vinculado),
      expiraEm: Date.now() + 60 * 60 * 1000
    });
  } catch (e) {
    logErro('loginUsuario', e);
    return respostaJSON({ success: false, error: e.message });
  }
}

function logoutUsuario(payload) {
  return respostaJSON({ success: true, message: 'Sessão encerrada' });
}

// ============================================================
// COMUNICAÇÃO: UPLOAD, MENSAGENS, NOTIFICAÇÕES
// ============================================================
function obterPastaDrive() {
  if (PASTA_DRIVE_ID) {
    try { return DriveApp.getFolderById(PASTA_DRIVE_ID); }
    catch (e) { Logger.log('⚠️ PASTA_DRIVE_ID inválido: ' + e.message); }
  }
  const nomePasta = 'Escola de Música - Arquivos';
  const it = DriveApp.getFoldersByName(nomePasta);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(nomePasta);
}

function uploadArquivoGAS(payload) {
  try {
    const p = payload || {};
    const nomeArquivo = s(p.nomeArquivo) || 'arquivo';
    const mimeType = s(p.mimeType) || 'application/octet-stream';
    const base64 = s(p.file);
    if (!base64) return respostaJSON({ success: false, error: 'Conteúdo do arquivo vazio.' });

    const professor = s(p.professor).toUpperCase();
    const curso = s(p.curso);
    const mensagem = s(p.mensagem);

    let destinatarios = [];
    try {
      const parsed = typeof p.destinatarios === 'string' ? JSON.parse(p.destinatarios) : p.destinatarios;
      destinatarios = Array.isArray(parsed) ? parsed : [];
    } catch (e) {
      destinatarios = String(p.destinatarios || '').split(',').map(function(x) { return x.trim(); }).filter(Boolean);
    }
    if (!destinatarios.length) return respostaJSON({ success: false, error: 'Nenhum destinatário informado.' });

    const pasta = obterPastaDrive();
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, nomeArquivo);
    const file = pasta.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
    catch (err) { Logger.log('⚠️ Não foi possível compartilhar publicamente: ' + err.message); }
    const url = file.getUrl();
    const fileId = file.getId();
    const dataHora = dataHoraAgora();

    const arqSheet = garantirAbaArquivos();
    const arqHeaders = getHeaders(arqSheet);
    const notifSheet = garantirAbaNotificacoes();
    const notifHeaders = getHeaders(notifSheet);

    destinatarios.forEach(function(alunoRaw) {
      const aluno = s(alunoRaw).toUpperCase();
      if (!aluno) return;

      const idArq = gerarID('ARQ');
      const valoresArq = arqHeaders.map(function(h) {
        switch (h) {
          case 'ID': return idArq;
          case 'Aluno': return aluno;
          case 'Curso': return curso;
          case 'Professor': return professor;
          case 'Nome_Arquivo': return nomeArquivo;
          case 'URL': return url;
          case 'File_ID': return fileId;
          case 'Mensagem': return mensagem;
          case 'Data': return dataHora;
          default: return '';
        }
      });
      appendRegistro(arqSheet, valoresArq);

      const idNotif = gerarID('NTF');
      const valoresNotif = notifHeaders.map(function(h) {
        switch (h) {
          case 'ID': return idNotif;
          case 'Aluno': return aluno;
          case 'Tipo': return 'arquivo';
          case 'Titulo': return 'Arquivo recebido: ' + nomeArquivo;
          case 'Descricao': return mensagem ? mensagem : ('Enviado pelo professor ' + professor);
          case 'Curso': return curso;
          case 'Professor': return professor;
          case 'Lida': return 'NAO';
          case 'Data': return dataHora;
          default: return '';
        }
      });
      appendRegistro(notifSheet, valoresNotif);
    });

    return respostaJSON({
      success: true,
      message: 'Arquivo enviado para ' + destinatarios.length + ' aluno(s).',
      url: url, fileId: fileId, nome: nomeArquivo
    });
  } catch (err) {
    logErro('uploadArquivoGAS', err);
    return respostaJSON({ success: false, error: err.message });
  }
}

function enviarMensagemAluno(payload) {
  try {
    const curso = payload.curso, professor = payload.professor, destinatarios = payload.destinatarios,
          assunto = payload.assunto, mensagem = payload.mensagem;
    if (!mensagem) return respostaJSON({ success: false, error: 'Mensagem vazia' });
    if (!destinatarios || !destinatarios.length) {
      return respostaJSON({ success: false, error: 'Nenhum destinatário informado' });
    }
    const dataHora = dataHoraAgora();
    const msgSheet = garantirAbaMensagens();
    const msgHeaders = getHeaders(msgSheet);
    const notifSheet = garantirAbaNotificacoes();
    const notifHeaders = getHeaders(notifSheet);

    let enviados = 0;
    destinatarios.forEach(function(alunoRaw) {
      const aluno = s(alunoRaw).toUpperCase();
      if (!aluno) return;

      const idMsg = gerarID('MSG');
      const valoresMsg = msgHeaders.map(function(h) {
        switch (h) {
          case 'ID': return idMsg;
          case 'Aluno': return aluno;
          case 'Curso': return s(curso);
          case 'Professor': return s(professor).toUpperCase();
          case 'Assunto': return s(assunto);
          case 'Mensagem': return s(mensagem);
          case 'Lida': return 'NAO';
          case 'Data': return dataHora;
          default: return '';
        }
      });
      appendRegistro(msgSheet, valoresMsg);

      const idNotif = gerarID('NTF');
      const valoresNotif = notifHeaders.map(function(h) {
        switch (h) {
          case 'ID': return idNotif;
          case 'Aluno': return aluno;
          case 'Tipo': return 'mensagem';
          case 'Titulo': return 'Mensagem: ' + (s(assunto) || '(sem assunto)');
          case 'Descricao': return s(mensagem).substring(0, 300);
          case 'Curso': return s(curso);
          case 'Professor': return s(professor).toUpperCase();
          case 'Lida': return 'NAO';
          case 'Data': return dataHora;
          default: return '';
        }
      });
      appendRegistro(notifSheet, valoresNotif);
      enviados++;
    });

    return respostaJSON({ success: true, message: 'Mensagem enviada para ' + enviados + ' aluno(s).', enviados: enviados });
  } catch (e) {
    logErro('enviarMensagemAluno', e);
    return respostaJSON({ success: false, error: e.message });
  }
}

function listarMensagens(payload) {
  const aluno = payload.aluno, curso = payload.curso, professor = payload.professor;
  let regs = getRegistros(garantirAbaMensagens());
  if (aluno) regs = regs.filter(function(r) { return s(r.Aluno).toUpperCase() === s(aluno).toUpperCase(); });
  if (curso) regs = regs.filter(function(r) { return s(r.Curso) === s(curso); });
  if (professor) regs = regs.filter(function(r) { return s(r.Professor).toUpperCase() === s(professor).toUpperCase(); });
  return respostaJSON({
    success: true, total: regs.length,
    mensagens: regs.map(function(r) {
      return {
        id: s(r.ID), aluno: s(r.Aluno), curso: s(r.Curso), professor: s(r.Professor),
        assunto: s(r.Assunto), mensagem: s(r.Mensagem), lida: s(r.Lida), data: s(r.Data)
      };
    })
  });
}

function listarArquivos(payload) {
  const aluno = payload.aluno, curso = payload.curso, professor = payload.professor;
  let regs = getRegistros(garantirAbaArquivos());
  if (aluno) regs = regs.filter(function(r) { return s(r.Aluno).toUpperCase() === s(aluno).toUpperCase(); });
  if (curso) regs = regs.filter(function(r) { return s(r.Curso) === s(curso); });
  if (professor) regs = regs.filter(function(r) { return s(r.Professor).toUpperCase() === s(professor).toUpperCase(); });
  return respostaJSON({
    success: true, total: regs.length,
    arquivos: regs.map(function(r) {
      return {
        id: s(r.ID), aluno: s(r.Aluno), curso: s(r.Curso), professor: s(r.Professor),
        nome: s(r.Nome_Arquivo), url: s(r.URL), fileId: s(r.File_ID),
        mensagem: s(r.Mensagem), data: s(r.Data)
      };
    })
  });
}

function listarNotificacoes(payload) {
  const aluno = payload.aluno, apenasNaoLidas = payload.apenasNaoLidas;
  if (!aluno) return respostaJSON({ success: false, error: 'Aluno obrigatório' });
  let regs = getRegistros(garantirAbaNotificacoes())
    .filter(function(r) { return s(r.Aluno).toUpperCase() === s(aluno).toUpperCase(); });
  if (apenasNaoLidas) regs = regs.filter(function(r) { return s(r.Lida).toUpperCase() !== 'SIM'; });
  return respostaJSON({
    success: true, total: regs.length,
    notificacoes: regs.map(function(r) {
      return {
        id: s(r.ID), aluno: s(r.Aluno), tipo: s(r.Tipo), titulo: s(r.Titulo),
        descricao: s(r.Descricao), curso: s(r.Curso), professor: s(r.Professor),
        lida: s(r.Lida), data: s(r.Data)
      };
    })
  });
}

function marcarNotificacoesLidas(payload) {
  const aluno = payload.aluno;
  if (!aluno) return respostaJSON({ success: false, error: 'Aluno obrigatório' });
  const sheet = garantirAbaNotificacoes();
  const headers = getHeaders(sheet);
  const idxLida = headers.indexOf('Lida');
  const idxAluno = headers.indexOf('Aluno');
  if (idxLida < 0 || idxAluno < 0) return respostaJSON({ success: false, error: 'Colunas ausentes' });
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return respostaJSON({ success: true, atualizadas: 0 });

  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  let atualizadas = 0;
  for (let i = 0; i < values.length; i++) {
    if (s(values[i][idxAluno]).toUpperCase() === s(aluno).toUpperCase()
        && s(values[i][idxLida]).toUpperCase() !== 'SIM') {
      sheet.getRange(i + 2, idxLida + 1).setValue('SIM');
      atualizadas++;
    }
  }
  return respostaJSON({ success: true, atualizadas: atualizadas });
}

// ============================================================
// DADOS DO ALUNO
// ============================================================
function dadosAlunoCompleto(payload) {
  const nome = payload.nome;
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  const nomeUp = s(nome).toUpperCase();

  const alunosSheet = getSheet('alunos');
  const linha = encontrarLinha(alunosSheet, { 'Nome': nomeUp });
  if (linha < 0) return respostaJSON({ success: false, error: 'Aluno não encontrado' });
  const alunosHeaders = getHeaders(alunosSheet);
  const alunoVals = alunosSheet.getRange(linha, 1, 1, alunosHeaders.length).getValues()[0];
  const aluno = {};
  alunosHeaders.forEach(function(h, i) { aluno[h] = alunoVals[i]; });

  const matriculas = getRegistros(getSheet('matriculas')).filter(function(m) { return s(m.Aluno).toUpperCase() === nomeUp; });
  const avaliacoes = getRegistros(getSheet('avaliacoes')).filter(function(a) { return s(a.Aluno).toUpperCase() === nomeUp; });

  const todasPres = getRegistros(getSheet('presencas'));
  const presencas = [];
  todasPres.forEach(function(p) {
    const presentes = s(p.Presentes).toUpperCase().split(',').map(function(x) { return x.trim(); });
    const faltantes = s(p.Faltantes).toUpperCase().split(',').map(function(x) { return x.trim(); });
    if (presentes.indexOf(nomeUp) !== -1) {
      presencas.push({ curso: s(p.Curso), data: s(p.Data), status: 'PRESENTE', conteudo: s(p.Conteudo) || s(p.Observacao), professor: s(p.Professor) });
    } else if (faltantes.indexOf(nomeUp) !== -1) {
      presencas.push({ curso: s(p.Curso), data: s(p.Data), status: 'FALTA', conteudo: s(p.Conteudo) || s(p.Observacao), professor: s(p.Professor) });
    }
  });

  const conteudos = getRegistros(getSheet('conteudos')).map(function(c) {
    return { curso: s(c.Curso), data: s(c.Data), conteudo: s(c.Conteudo), objetivos: s(c.Objetivos), professor: s(c.Professor) };
  });

  const relatorios = getRegistros(getSheet('relatorios'))
    .filter(function(r) { return s(r.Aluno).toUpperCase() === nomeUp; })
    .map(function(r) {
      return {
        curso: s(r.Curso), bimestre: s(r.Bimestre), professor: s(r.Professor),
        relatorio: s(r.Relatorio), data: s(r.Data), visto: s(r.Visto)
      };
    });

  let arquivos = [];
  let mensagens = [];
  try {
    arquivos = getRegistros(garantirAbaArquivos())
      .filter(function(a) { return s(a.Aluno).toUpperCase() === nomeUp; })
      .map(function(a) {
        return {
          id: s(a.ID), curso: s(a.Curso), professor: s(a.Professor),
          nome: s(a.Nome_Arquivo), url: s(a.URL), fileId: s(a.File_ID),
          mensagem: s(a.Mensagem), data: s(a.Data)
        };
      });
  } catch (e) { Logger.log('⚠️ Erro ao ler arquivos: ' + e.message); }

  try {
    mensagens = getRegistros(garantirAbaMensagens())
      .filter(function(m) { return s(m.Aluno).toUpperCase() === nomeUp; })
      .map(function(m) {
        return {
          id: s(m.ID), curso: s(m.Curso), professor: s(m.Professor),
          assunto: s(m.Assunto), mensagem: s(m.Mensagem), data: s(m.Data),
          lida: s(m.Lida).toUpperCase() === 'SIM'
        };
      });
  } catch (e) { Logger.log('⚠️ Erro ao ler mensagens: ' + e.message); }

  return respostaJSON({
    success: true,
    aluno: {
      nome: s(aluno.Nome),
      dataNascimento: s(aluno.Data_Nascimento),
      responsavel: s(aluno.Responsavel),
      contato: s(aluno.Contato),
      status: s(aluno.Status)
    },
    matriculas: matriculas.map(function(m) {
      return {
        curso: s(m.Curso), professor: s(m.Professor),
        status: s(m.Status), dataMatricula: s(m.Data_Matricula), anoLetivo: s(m.Ano_Letivo)
      };
    }),
    avaliacoes: avaliacoes.map(function(a) {
      return {
        curso: s(a.Curso), bimestre: Number(a.Bimestre) || 1, nota: Number(a.Nota) || 0,
        professor: s(a.Professor), dataLancamento: s(a.Data_Lancamento)
      };
    }),
    presencas: presencas,
    conteudos: conteudos,
    relatorios: relatorios,
    arquivos: arquivos,
    mensagens: mensagens
  });
}

function garantirAbaAlunoAPI(payload) {
  const nome = payload.nome;
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  try {
    const info = sincronizarAbaAluno(nome);
    return respostaJSON({ success: true, message: 'Aba sincronizada', aba: info.aba, matriculas: info.matriculas });
  } catch (e) {
    return respostaJSON({ success: false, error: e.message });
  }
}

// ============================================================
// DEMAIS AÇÕES
// ============================================================
function salvarAluno(payload) {
  const nome = payload.nome, dataNascimento = payload.dataNascimento, responsavel = payload.responsavel,
        contato = payload.contato, status = payload.status;
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  const sheet = getSheet('alunos');
  const headers = getHeaders(sheet);
  if (encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() }) > -1) {
    return respostaJSON({ success: false, error: 'Aluno já cadastrado: ' + nome });
  }
  const valores = headers.map(function(h) {
    switch (h) {
      case 'Nome': return s(nome).toUpperCase();
      case 'Data_Nascimento': return formatarData(dataNascimento);
      case 'Responsavel': return s(responsavel).toUpperCase();
      case 'Contato': return s(contato);
      case 'Status': return s(status || 'ATIVO').toUpperCase();
      case 'Data_Cadastro': return dataHoje();
      default: return '';
    }
  });
  const linha = appendRegistro(sheet, valores);
  return respostaJSON({ success: true, message: 'Aluno salvo', id: linha });
}

function atualizarAluno(payload) {
  const nome = payload.nome, nomeOriginal = payload.nomeOriginal, dataNascimento = payload.dataNascimento,
        responsavel = payload.responsavel, contato = payload.contato, status = payload.status;
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  const sheet = getSheet('alunos');
  const headers = getHeaders(sheet);
  const linha = encontrarLinha(sheet, { 'Nome': s(nomeOriginal || nome).toUpperCase() });
  if (linha < 0) return respostaJSON({ success: false, error: 'Aluno não encontrado' });
  const antigos = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const novos = headers.map(function(h, i) {
    switch (h) {
      case 'Nome': return s(nome).toUpperCase();
      case 'Data_Nascimento': return dataNascimento !== undefined ? formatarData(dataNascimento) : antigos[i];
      case 'Responsavel': return responsavel !== undefined ? s(responsavel).toUpperCase() : antigos[i];
      case 'Contato': return contato !== undefined ? s(contato) : antigos[i];
      case 'Status': return status !== undefined ? s(status).toUpperCase() : antigos[i];
      default: return antigos[i];
    }
  });
  atualizarLinha(sheet, linha, novos);
  return respostaJSON({ success: true, message: 'Aluno atualizado' });
}

function removerAluno(payload) {
  const nome = payload.nome;
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  const sheet = getSheet('alunos');
  const linha = encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() });
  if (linha < 0) return respostaJSON({ success: false, error: 'Aluno não encontrado' });
  removerLinha(sheet, linha);
  return respostaJSON({ success: true, message: 'Aluno removido' });
}

function listarAlunos() {
  const registros = getRegistros(getSheet('alunos'));
  return respostaJSON({ success: true, total: registros.length, alunos: registros });
}

function salvarProfessor(payload) {
  const id = payload.id, nome = payload.nome, tipo = payload.tipo, curso = payload.curso,
        telefone = payload.telefone, email = payload.email, valorHora = payload.valorHora;
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  const sheet = getSheet('professores');
  garantirColunaValorHora(sheet);
  const headers = getHeaders(sheet);
  const idFinal = s(id) || gerarID('PROF');
  if (encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() }) > -1) {
    return respostaJSON({ success: false, error: 'Professor já cadastrado' });
  }
  const valores = headers.map(function(h) {
    switch (h) {
      case 'ID_Professor': return idFinal;
      case 'Nome': return s(nome).toUpperCase();
      case 'Tipo': return s(tipo);
      case 'Curso': return s(curso);
      case 'Telefone': return s(telefone);
      case 'Email': return s(email);
      case 'Valor_Hora': return Number(valorHora) || 0;
      case 'Status': return 'ATIVO';
      case 'Data_Cadastro': return dataHoje();
      default: return '';
    }
  });
  appendRegistro(sheet, valores);
  return respostaJSON({ success: true, message: 'Professor salvo', id: idFinal });
}

function atualizarProfessor(payload) {
  const id = payload.id, nome = payload.nome, tipo = payload.tipo, curso = payload.curso,
        telefone = payload.telefone, email = payload.email, status = payload.status, valorHora = payload.valorHora;
  if (!id && !nome) return respostaJSON({ success: false, error: 'ID ou Nome obrigatório' });
  const sheet = getSheet('professores');
  garantirColunaValorHora(sheet);
  const headers = getHeaders(sheet);
  let linha = -1;
  if (id) linha = encontrarLinha(sheet, { 'ID_Professor': s(id) });
  if (linha < 0 && nome) linha = encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() });
  if (linha < 0) return respostaJSON({ success: false, error: 'Professor não encontrado' });
  const antigos = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const novos = headers.map(function(h, i) {
    switch (h) {
      case 'Nome': return nome !== undefined ? s(nome).toUpperCase() : antigos[i];
      case 'Tipo': return tipo !== undefined ? s(tipo) : antigos[i];
      case 'Curso': return curso !== undefined ? s(curso) : antigos[i];
      case 'Telefone': return telefone !== undefined ? s(telefone) : antigos[i];
      case 'Email': return email !== undefined ? s(email) : antigos[i];
      case 'Status': return status !== undefined ? s(status).toUpperCase() : antigos[i];
      case 'Valor_Hora': return valorHora !== undefined ? (Number(valorHora) || 0) : antigos[i];
      default: return antigos[i];
    }
  });
  atualizarLinha(sheet, linha, novos);
  return respostaJSON({ success: true, message: 'Professor atualizado' });
}

function salvarValorProfessor(payload) {
  const id = payload.id, valorHora = payload.valorHora;
  if (!id) return respostaJSON({ success: false, error: 'ID obrigatório' });
  const sheet = getSheet('professores');
  garantirColunaValorHora(sheet);
  const headers = getHeaders(sheet);
  const idxValor = headers.indexOf('Valor_Hora');
  const linha = encontrarLinha(sheet, { 'ID_Professor': s(id) });
  if (linha < 0) return respostaJSON({ success: false, error: 'Professor não encontrado' });
  const valor = Number(valorHora) || 0;
  sheet.getRange(linha, idxValor + 1).setValue(valor);
  return respostaJSON({ success: true, message: 'Valor atualizado', valorHora: valor });
}

function garantirColunaValorHora(sheet) {
  const headers = getHeaders(sheet);
  if (headers.indexOf('Valor_Hora') < 0) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Valor_Hora');
  }
}

function removerProfessor(payload) {
  const id = payload.id, nome = payload.nome;
  if (!id && !nome) return respostaJSON({ success: false, error: 'ID ou Nome obrigatório' });
  const sheet = getSheet('professores');
  let linha = -1;
  if (id) linha = encontrarLinha(sheet, { 'ID_Professor': s(id) });
  if (linha < 0 && nome) linha = encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() });
  if (linha < 0) return respostaJSON({ success: false, error: 'Professor não encontrado' });
  removerLinha(sheet, linha);
  return respostaJSON({ success: true, message: 'Professor removido' });
}

function listarProfessores() {
  const registros = getRegistros(getSheet('professores'));
  return respostaJSON({ success: true, total: registros.length, professores: registros });
}

function salvarMatricula(payload) {
  const aluno = payload.aluno, curso = payload.curso, professor = payload.professor,
        dataMatricula = payload.dataMatricula, anoLetivo = payload.anoLetivo, status = payload.status;
  if (!aluno || !curso) return respostaJSON({ success: false, error: 'Aluno e Curso obrigatórios' });
  const sheet = getSheet('matriculas');
  const headers = getHeaders(sheet);
  const existente = encontrarLinha(sheet, { 'Aluno': s(aluno).toUpperCase(), 'Curso': s(curso) });
  if (existente > -1) {
    const vals = sheet.getRange(existente, 1, 1, headers.length).getValues()[0];
    const idxStatus = headers.indexOf('Status');
    if (idxStatus >= 0 && String(vals[idxStatus]).toUpperCase() === 'ATIVO') {
      return respostaJSON({ success: false, error: 'Aluno já possui matrícula ativa neste curso' });
    }
  }
  const id = gerarID('MAT');
  const valores = headers.map(function(h) {
    switch (h) {
      case 'ID_Matricula': return id;
      case 'Aluno': return s(aluno).toUpperCase();
      case 'Curso': return s(curso);
      case 'Professor': return s(professor);
      case 'Data_Matricula': return formatarData(dataMatricula) || dataHoje();
      case 'Ano_Letivo': return s(anoLetivo) || String(new Date().getFullYear());
      case 'Status': return s(status || 'ATIVO').toUpperCase();
      default: return '';
    }
  });
  appendRegistro(sheet, valores);
  return respostaJSON({ success: true, message: 'Matrícula registrada', id: id });
}

function atualizarMatricula(payload) {
  const id = payload.id, aluno = payload.aluno, curso = payload.curso, professor = payload.professor,
        dataMatricula = payload.dataMatricula, anoLetivo = payload.anoLetivo, status = payload.status;
  const sheet = getSheet('matriculas');
  const headers = getHeaders(sheet);
  let linha = -1;
  if (id) linha = encontrarLinha(sheet, { 'ID_Matricula': s(id) });
  if (linha < 0 && aluno && curso) linha = encontrarLinha(sheet, { 'Aluno': s(aluno).toUpperCase(), 'Curso': s(curso) });
  if (linha < 0) return respostaJSON({ success: false, error: 'Matrícula não encontrada' });
  const antigos = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const novos = headers.map(function(h, i) {
    switch (h) {
      case 'Aluno': return aluno !== undefined ? s(aluno).toUpperCase() : antigos[i];
      case 'Curso': return curso !== undefined ? s(curso) : antigos[i];
      case 'Professor': return professor !== undefined ? s(professor) : antigos[i];
      case 'Data_Matricula': return dataMatricula !== undefined ? formatarData(dataMatricula) : antigos[i];
      case 'Ano_Letivo': return anoLetivo !== undefined ? s(anoLetivo) : antigos[i];
      case 'Status': return status !== undefined ? s(status).toUpperCase() : antigos[i];
      default: return antigos[i];
    }
  });
  atualizarLinha(sheet, linha, novos);
  return respostaJSON({ success: true, message: 'Matrícula atualizada' });
}

function removerMatricula(payload) {
  const id = payload.id, aluno = payload.aluno, curso = payload.curso;
  const sheet = getSheet('matriculas');
  let linha = -1;
  if (id) linha = encontrarLinha(sheet, { 'ID_Matricula': s(id) });
  if (linha < 0 && aluno && curso) linha = encontrarLinha(sheet, { 'Aluno': s(aluno).toUpperCase(), 'Curso': s(curso) });
  if (linha < 0) return respostaJSON({ success: false, error: 'Matrícula não encontrada' });
  removerLinha(sheet, linha);
  return respostaJSON({ success: true, message: 'Matrícula removida' });
}

function listarMatriculas(payload) {
  let registros = getRegistros(getSheet('matriculas'));
  if (payload) {
    if (payload.aluno) registros = registros.filter(function(r) { return s(r.Aluno).toUpperCase() === s(payload.aluno).toUpperCase(); });
    if (payload.curso) registros = registros.filter(function(r) { return s(r.Curso) === s(payload.curso); });
    if (payload.status) registros = registros.filter(function(r) { return s(r.Status).toUpperCase() === s(payload.status).toUpperCase(); });
  }
  return respostaJSON({ success: true, total: registros.length, matriculas: registros });
}

function salvarAvaliacoes(payload) {
  const lista = payload.avaliacoes || payload.notas || [];
  if (!Array.isArray(lista) || lista.length === 0) return respostaJSON({ success: false, error: 'Sem avaliações' });
  const sheet = getSheet('avaliacoes');
  const headers = getHeaders(sheet);
  const res = { inseridos: 0, atualizados: 0, erros: [] };

  let notifSheet = null, notifHeaders = null;
  try { notifSheet = garantirAbaNotificacoes(); notifHeaders = getHeaders(notifSheet); }
  catch (e) { Logger.log('⚠️ Não foi possível preparar NOTIFICACOES: ' + e.message); }

  lista.forEach(function(av) {
    try {
      const aluno = s(av.aluno).toUpperCase();
      const curso = s(av.curso);
      const bimestre = String(av.bimestre || '1');
      const nota = Number(av.nota) || 0;
      const professor = s(av.professor || '');
      const linha = encontrarLinha(sheet, { 'Aluno': aluno, 'Curso': curso, 'Bimestre': bimestre });
      if (linha > -1) {
        const vals = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
        const novos = headers.map(function(h, i) {
          switch (h) {
            case 'Nota': return nota;
            case 'Data_Lancamento': return dataHoje();
            case 'Professor': return professor || vals[i];
            default: return vals[i];
          }
        });
        atualizarLinha(sheet, linha, novos);
        res.atualizados++;
      } else {
        const id = gerarID('AVA');
        const valores = headers.map(function(h) {
          switch (h) {
            case 'ID_Avaliacao': return id;
            case 'Aluno': return aluno;
            case 'Curso': return curso;
            case 'Bimestre': return bimestre;
            case 'Nota': return nota;
            case 'Data_Lancamento': return dataHoje();
            case 'Professor': return professor;
            default: return '';
          }
        });
        appendRegistro(sheet, valores);
        res.inseridos++;
      }

      if (notifSheet && notifHeaders) {
        const dataHora = dataHoraAgora();
        const idNotif = gerarID('NTF');
        const valoresNotif = notifHeaders.map(function(h) {
          switch (h) {
            case 'ID': return idNotif;
            case 'Aluno': return aluno;
            case 'Tipo': return 'nota';
            case 'Titulo': return 'Nova nota em ' + curso;
            case 'Descricao': return bimestre + 'º Bimestre: ' + nota.toFixed(1) + ' — Prof. ' + professor;
            case 'Curso': return curso;
            case 'Professor': return professor;
            case 'Lida': return 'NAO';
            case 'Data': return dataHora;
            default: return '';
          }
        });
        appendRegistro(notifSheet, valoresNotif);
      }
    } catch (e) { res.erros.push(e.message); }
  });

  return respostaJSON({ success: res.erros.length === 0, message: res.inseridos + ' inserida(s), ' + res.atualizados + ' atualizada(s)', detalhes: res });
}

function listarAvaliacoes(payload) {
  let registros = getRegistros(getSheet('avaliacoes'));
  if (payload) {
    if (payload.aluno) registros = registros.filter(function(r) { return s(r.Aluno).toUpperCase() === s(payload.aluno).toUpperCase(); });
    if (payload.curso) registros = registros.filter(function(r) { return s(r.Curso) === s(payload.curso); });
    if (payload.bimestre) registros = registros.filter(function(r) { return String(r.Bimestre) === String(payload.bimestre); });
  }
  return respostaJSON({ success: true, total: registros.length, avaliacoes: registros });
}

function salvarPresenca(payload) {
  const curso = payload.curso, data = payload.data, presentes = payload.presentes, faltantes = payload.faltantes,
        observacao = payload.observacao, conteudo = payload.conteudo, professor = payload.professor;
  if (!curso || !data) return respostaJSON({ success: false, error: 'Curso e Data obrigatórios' });
  const sheet = getSheet('presencas');
  const headers = getHeaders(sheet);
  const listaP = Array.isArray(presentes) ? presentes.join(', ') : s(presentes);
  const listaF = Array.isArray(faltantes) ? faltantes.join(', ') : s(faltantes);
  const conteudoFinal = s(conteudo || observacao || '');
  const dataF = formatarData(data);
  const id = gerarID('PRE');
  const valores = headers.map(function(h) {
    switch (h) {
      case 'ID_Presenca': return id;
      case 'Curso': return s(curso);
      case 'Data': return dataF;
      case 'Presentes': return listaP;
      case 'Faltantes': return listaF;
      case 'Conteudo': return conteudoFinal;
      case 'Observacao': return conteudoFinal;
      case 'Professor': return s(professor);
      default: return '';
    }
  });
  if (!headers.some(function(h) { return h === 'Conteudo' || h === 'Observacao'; }) && headers.length >= 6) {
    valores[5] = conteudoFinal;
  }
  appendRegistro(sheet, valores);
  return respostaJSON({ success: true, message: 'Presença registrada', id: id });
}

function listarPresencas(payload) {
  let registros = getRegistros(getSheet('presencas'));
  if (payload) {
    if (payload.curso) registros = registros.filter(function(r) { return s(r.Curso) === s(payload.curso); });
    if (payload.data) registros = registros.filter(function(r) { return s(r.Data) === formatarData(payload.data); });
  }
  return respostaJSON({ success: true, total: registros.length, presencas: registros });
}

function salvarConteudo(payload) {
  const curso = payload.curso, data = payload.data, conteudo = payload.conteudo,
        objetivos = payload.objetivos, professor = payload.professor;
  if (!curso || !conteudo) return respostaJSON({ success: false, error: 'Curso e Conteúdo obrigatórios' });
  const sheet = getSheet('conteudos');
  const headers = getHeaders(sheet);
  const id = gerarID('CON');
  const dataF = formatarData(data) || dataHoje();
  const valores = headers.map(function(h) {
    switch (h) {
      case 'ID_Conteudo': return id;
      case 'Curso': return s(curso);
      case 'Data': return dataF;
      case 'Conteudo': return s(conteudo);
      case 'Objetivos': return s(objetivos);
      case 'Professor': return s(professor);
      default: return '';
    }
  });
  appendRegistro(sheet, valores);
  return respostaJSON({ success: true, message: 'Conteúdo salvo', id: id });
}

function listarConteudos(payload) {
  let registros = getRegistros(getSheet('conteudos'));
  if (payload && payload.curso) registros = registros.filter(function(r) { return s(r.Curso) === s(payload.curso); });
  return respostaJSON({ success: true, total: registros.length, conteudos: registros });
}

function salvarObservacao(payload) {
  const curso = payload.curso, data = payload.data, aluno = payload.aluno,
        observacao = payload.observacao, tipo = payload.tipo, professor = payload.professor;
  if (!curso || !aluno || !observacao) return respostaJSON({ success: false, error: 'Campos obrigatórios' });
  const sheet = getSheet('observacoes');
  const headers = getHeaders(sheet);
  const id = gerarID('OBS');
  const dataF = formatarData(data) || dataHoje();
  const valores = headers.map(function(h) {
    switch (h) {
      case 'ID_Observacao': return id;
      case 'Curso': return s(curso);
      case 'Data': return dataF;
      case 'Aluno': return s(aluno).toUpperCase();
      case 'Observacao': return s(observacao);
      case 'Tipo': return s(tipo || 'Neutra');
      case 'Professor': return s(professor);
      default: return '';
    }
  });
  appendRegistro(sheet, valores);
  return respostaJSON({ success: true, message: 'Observação salva', id: id });
}

function listarObservacoes(payload) {
  let registros = getRegistros(getSheet('observacoes'));
  if (payload) {
    if (payload.curso) registros = registros.filter(function(r) { return s(r.Curso) === s(payload.curso); });
    if (payload.aluno) registros = registros.filter(function(r) { return s(r.Aluno).toUpperCase() === s(payload.aluno).toUpperCase(); });
  }
  return respostaJSON({ success: true, total: registros.length, observacoes: registros });
}

function salvarRelatorio(payload) {
  const aluno = payload.aluno, curso = payload.curso, bimestre = payload.bimestre,
        professor = payload.professor, relatorio = payload.relatorio, data = payload.data;
  if (!aluno || !curso || !relatorio) return respostaJSON({ success: false, error: 'Campos obrigatórios' });
  const sheet = getSheet('relatorios');
  const headers = getHeaders(sheet);
  const id = gerarID('REL');
  const dataF = formatarData(data) || dataHoje();
  const valores = headers.map(function(h) {
    switch (h) {
      case 'ID_Relatorio': return id;
      case 'Aluno': return s(aluno).toUpperCase();
      case 'Curso': return s(curso);
      case 'Bimestre': return String(bimestre || '1');
      case 'Professor': return s(professor);
      case 'Relatorio': return s(relatorio);
      case 'Data': return dataF;
      case 'Visto': return 'NÃO';
      default: return '';
    }
  });
  appendRegistro(sheet, valores);
  return respostaJSON({ success: true, message: 'Relatório salvo', id: id });
}

function marcarRelatorioVisto(payload) {
  const id = payload.id;
  if (!id) return respostaJSON({ success: false, error: 'ID obrigatório' });
  const sheet = getSheet('relatorios');
  const headers = getHeaders(sheet);
  const linha = encontrarLinha(sheet, { 'ID_Relatorio': s(id) });
  if (linha < 0) return respostaJSON({ success: false, error: 'Relatório não encontrado' });
  const vals = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const idx = headers.indexOf('Visto');
  const idxData = headers.indexOf('Data_Visto');
  if (idx >= 0) vals[idx] = 'SIM';
  if (idxData >= 0) vals[idxData] = dataHoje();
  atualizarLinha(sheet, linha, vals);
  return respostaJSON({ success: true, message: 'Relatório marcado como visto' });
}

function listarRelatorios(payload) {
  let registros = getRegistros(getSheet('relatorios'));
  if (payload) {
    if (payload.aluno) registros = registros.filter(function(r) { return s(r.Aluno).toUpperCase() === s(payload.aluno).toUpperCase(); });
    if (payload.curso) registros = registros.filter(function(r) { return s(r.Curso) === s(payload.curso); });
    if (payload.professor) registros = registros.filter(function(r) { return s(r.Professor) === s(payload.professor); });
  }
  return respostaJSON({ success: true, total: registros.length, relatorios: registros });
}

function atribuirCurso(payload) {
  const professorId = payload.professorId, professorNome = payload.professorNome, curso = payload.curso;
  if (!curso) return respostaJSON({ success: false, error: 'Curso obrigatório' });
  const sheet = getSheet('professores');
  const headers = getHeaders(sheet);
  const idxCurso = headers.indexOf('Curso');
  let linha = -1;
  if (professorId) linha = encontrarLinha(sheet, { 'ID_Professor': s(professorId) });
  if (linha < 0 && professorNome) linha = encontrarLinha(sheet, { 'Nome': s(professorNome).toUpperCase() });
  if (linha < 0) return respostaJSON({ success: false, error: 'Professor não encontrado' });
  const vals = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const cursos = s(vals[idxCurso]).split(',').map(function(c) { return c.trim(); }).filter(Boolean);
  if (cursos.indexOf(curso) !== -1) return respostaJSON({ success: true, message: 'Curso já atribuído', jaExistia: true });
  cursos.push(curso);
  vals[idxCurso] = cursos.join(',');
  atualizarLinha(sheet, linha, vals);
  return respostaJSON({ success: true, message: 'Curso atribuído', cursos: cursos });
}

function removerCurso(payload) {
  const professorId = payload.professorId, professorNome = payload.professorNome, curso = payload.curso;
  if (!curso) return respostaJSON({ success: false, error: 'Curso obrigatório' });
  const sheet = getSheet('professores');
  const headers = getHeaders(sheet);
  const idxCurso = headers.indexOf('Curso');
  let linha = -1;
  if (professorId) linha = encontrarLinha(sheet, { 'ID_Professor': s(professorId) });
  if (linha < 0 && professorNome) linha = encontrarLinha(sheet, { 'Nome': s(professorNome).toUpperCase() });
  if (linha < 0) return respostaJSON({ success: false, error: 'Professor não encontrado' });
  const vals = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  let cursos = s(vals[idxCurso]).split(',').map(function(c) { return c.trim(); }).filter(Boolean);
  cursos = cursos.filter(function(c) { return c !== curso; });
  vals[idxCurso] = cursos.join(',');
  atualizarLinha(sheet, linha, vals);
  return respostaJSON({ success: true, message: 'Curso removido', cursos: cursos });
}

function solicitarRelatorio(payload) {
  const aluno = payload.aluno, curso = payload.curso, bimestre = payload.bimestre,
        tipo = payload.tipo, professor = payload.professor;
  if (!aluno || !curso) return respostaJSON({ success: false, error: 'Aluno e Curso obrigatórios' });
  const sheet = getSheet('relatorios');
  const headers = getHeaders(sheet);
  const id = gerarID('SOL');
  const dataF = dataHoje();
  const texto = '[SOLICITADO EM ' + dataF + ']\n[TIPO: ' + s(tipo || 'pedagogico') + ']\nAguardando preenchimento.';
  const valores = headers.map(function(h) {
    switch (h) {
      case 'ID_Relatorio': return id;
      case 'Aluno': return s(aluno).toUpperCase();
      case 'Curso': return s(curso);
      case 'Bimestre': return String(bimestre || 'anual');
      case 'Professor': return s(professor);
      case 'Relatorio': return texto;
      case 'Data': return dataF;
      case 'Visto': return 'NÃO';
      default: return '';
    }
  });
  appendRegistro(sheet, valores);
  return respostaJSON({ success: true, message: 'Solicitação registrada', id: id });
}

function garantirAbaLogsAPI() {
  try {
    const sheet = garantirAbaLogs();
    return respostaJSON({ success: true, message: 'Aba LOGS pronta', total: sheet.getLastRow() });
  } catch (e) {
    return respostaJSON({ success: false, error: e.message });
  }
}

function registrarLogin(payload) {
  const nome = payload.nome, perfil = payload.perfil, acao = payload.acao;
  if (!nome || !perfil) return respostaJSON({ success: false, error: 'Nome e perfil obrigatórios' });
  try {
    const sheet = garantirAbaLogs();
    const headers = getHeaders(sheet);
    const id = gerarID('LOG');
    const dataHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss');
    const valores = headers.map(function(h) {
      switch (h) {
        case 'ID_Log': return id;
        case 'Nome': return s(nome).toUpperCase();
        case 'Perfil': return s(perfil);
        case 'Data_Hora': return dataHora;
        case 'Acao': return s(acao || 'LOGIN');
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    return respostaJSON({ success: true, id: id, dataHora: dataHora });
  } catch (e) {
    return respostaJSON({ success: false, error: e.message });
  }
}

function listarLogs(payload) {
  let registros = getRegistros(garantirAbaLogs());
  if (payload) {
    if (payload.perfil) registros = registros.filter(function(r) { return s(r.Perfil).toLowerCase() === s(payload.perfil).toLowerCase(); });
    if (payload.nome) registros = registros.filter(function(r) { return s(r.Nome).toUpperCase() === s(payload.nome).toUpperCase(); });
  }
  return respostaJSON({ success: true, total: registros.length, logs: registros });
}

function listarCursosAPI() {
  const CURSOS = [
    { id: 'piano', nome: 'Piano', tipo: 'Instrumento' },
    { id: 'teclado', nome: 'Teclado', tipo: 'Instrumento' },
    { id: 'orgao', nome: 'Órgão', tipo: 'Instrumento' },
    { id: 'violao', nome: 'Violão', tipo: 'Instrumento' },
    { id: 'guitarra', nome: 'Guitarra', tipo: 'Instrumento' },
    { id: 'violino', nome: 'Violino', tipo: 'Instrumento' },
    { id: 'violoncelo', nome: 'Violoncelo', tipo: 'Instrumento' },
    { id: 'viola-erudita', nome: 'Viola Erudita', tipo: 'Instrumento' },
    { id: 'contrabaixo-acustico', nome: 'Contrabaixo Acústico', tipo: 'Instrumento' },
    { id: 'bateria', nome: 'Bateria', tipo: 'Instrumento' },
    { id: 'trompete', nome: 'Trompete', tipo: 'Instrumento' },
    { id: 'trombone', nome: 'Trombone', tipo: 'Instrumento' },
    { id: 'euphonio', nome: 'Eufônio', tipo: 'Instrumento' },
    { id: 'tuba', nome: 'Tuba', tipo: 'Instrumento' },
    { id: 'saxofone', nome: 'Saxofone', tipo: 'Instrumento' },
    { id: 'flauta', nome: 'Flauta', tipo: 'Instrumento' },
    { id: 'clarinete', nome: 'Clarinete', tipo: 'Instrumento' },
    { id: 'canto', nome: 'Canto', tipo: 'Instrumento' },
    { id: 'teoria-musical1', nome: 'Teoria Musical 1', tipo: 'Teoria Musical' },
    { id: 'teoria-musical2', nome: 'Teoria Musical 2', tipo: 'Teoria Musical' },
    { id: 'musicalizacao-infantil', nome: 'Musicalização Infantil', tipo: 'Musicalização' }
  ];
  return respostaJSON({ success: true, total: CURSOS.length, cursos: CURSOS });
}

function sincronizarTudo() {
  const res = {};
  Object.keys(ABAS).forEach(function(chave) {
    try {
      if (chave === 'logs') garantirAbaLogs();
      if (chave === 'usuarios') garantirAbaUsuarios();
      if (chave === 'arquivos') garantirAbaArquivos();
      if (chave === 'mensagens') garantirAbaMensagens();
      if (chave === 'notificacoes') garantirAbaNotificacoes();
      const sheet = getSheet(chave);
      res[chave] = getRegistros(sheet);
    } catch (e) {
      res[chave] = [];
      res[chave + '_erro'] = e.message;
    }
  });
  return respostaJSON({ success: true, timestamp: new Date().toISOString(), dados: res });
}

function limparAba(payload) {
  const aba = payload.aba;
  if (!aba) return respostaJSON({ success: false, error: 'Aba obrigatória' });
  const chave = Object.keys(ABAS).find(function(k) { return ABAS[k] === aba || k === aba; });
  if (!chave) return respostaJSON({ success: false, error: 'Aba não reconhecida' });
  const sheet = getSheet(chave);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  return respostaJSON({ success: true, message: 'Aba limpa' });
}

function criarAbasComunicacaoAPI() {
  try {
    garantirAbaArquivos();
    garantirAbaMensagens();
    garantirAbaNotificacoes();
    return respostaJSON({ success: true, message: 'Abas ARQUIVOS, MENSAGENS e NOTIFICACOES prontas.' });
  } catch (e) {
    return respostaJSON({ success: false, error: e.message });
  }
}

// ============================================================
// MENU E TESTES
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎵 Escola de Música')
    .addItem('🧪 Testar API', 'testarAPI')
    .addItem('📋 Criar aba USUARIOS', 'menuCriarUsuarios')
    .addItem('📋 Criar aba LOGS', 'menuCriarLogs')
    .addItem('📋 Criar abas de COMUNICAÇÃO', 'menuCriarComunicacao')
    .addItem('📋 Verificar coluna Valor_Hora', 'menuGarantirValorHora')
    .addSeparator()
    .addItem('🔗 URL do Web App', 'mostrarURLWebApp')
    .addToUi();
}

function menuCriarUsuarios() {
  garantirAbaUsuarios();
  SpreadsheetApp.getUi().alert('✅ Aba USUARIOS criada/verificada.');
}

function menuCriarLogs() {
  garantirAbaLogs();
  SpreadsheetApp.getUi().alert('✅ Aba LOGS criada/verificada.');
}

function menuCriarComunicacao() {
  garantirAbaArquivos();
  garantirAbaMensagens();
  garantirAbaNotificacoes();
  SpreadsheetApp.getUi().alert('✅ Abas ARQUIVOS, MENSAGENS e NOTIFICACOES criadas/verificadas.');
}

function menuGarantirValorHora() {
  const sheet = getSheet('professores');
  garantirColunaValorHora(sheet);
  SpreadsheetApp.getUi().alert('✅ Coluna Valor_Hora verificada.');
}

function testarAPI() {
  Logger.log('🧪 Testando...');
  Logger.log('ping: ' + processarAcao('ping', {}).getContent());
  Logger.log('garantirAbaUsuarios: ' + (garantirAbaUsuarios() ? 'ok' : 'falhou'));
  Logger.log('garantirAbaLogs: ' + (garantirAbaLogs() ? 'ok' : 'falhou'));
  Logger.log('garantirAbaArquivos: ' + (garantirAbaArquivos() ? 'ok' : 'falhou'));
  Logger.log('garantirAbaMensagens: ' + (garantirAbaMensagens() ? 'ok' : 'falhou'));
  Logger.log('garantirAbaNotificacoes: ' + (garantirAbaNotificacoes() ? 'ok' : 'falhou'));
  Logger.log('listarCursos: ' + processarAcao('listarCursos', {}).getContent().substring(0, 200));
  Logger.log('✅ Concluído');
}

function mostrarURLWebApp() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('🔗 URL do Web App:\n\n' + (url || '(vazio — implante primeiro)'));
}
