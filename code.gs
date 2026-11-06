/**
 * ============================================================
 * ESCOLA DE MÚSICA - WEB APP (API) - v2.0
 * ============================================================
 * Ações: login, cadastro de usuários, aba individual por aluno.
 *
 * CONFIGURAÇÃO:
 * 1. Implantar como Web App
 * 2. Executar como: "Eu"
 * 3. Quem tem acesso: "Qualquer pessoa"
 * 4. Copiar URL em GAS_URL no HTML
 * ============================================================
 */

const SHEET_ID = '1FRtD7cbV2tvK2og4NDkyWi8V_hziyERNZoSh20gG980';

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
  usuarios: 'USUARIOS'   // NOVO
};

// ============================================================
// ENTRY POINTS
// ============================================================
function doGet(e) {
  try {
    if (!e || !e.parameter || !e.parameter.data) {
      return respostaJSON({
        success: true,
        message: '🎵 API Escola de Música online (v2.0)',
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

function doPost(e) {
  try {
    let dados = null;
    if (e && e.postData && e.postData.contents) {
      try { dados = JSON.parse(e.postData.contents); }
      catch (p) {
        if (e.parameter && e.parameter.data) {
          try { dados = JSON.parse(e.parameter.data); }
          catch (e2) { return respostaJSON({ success: false, error: 'Formato inválido' }); }
        } else {
          return respostaJSON({ success: false, error: 'JSON inválido: ' + p.message });
        }
      }
    } else if (e && e.parameter && e.parameter.data) {
      dados = JSON.parse(e.parameter.data);
    }
    if (!dados || !dados.acao) return respostaJSON({ success: false, error: 'Ação não especificada' });
    return processarAcao(dados.acao, dados.payload || {});
  } catch (err) {
    logErro('doPost', err);
    return respostaJSON({ success: false, error: err.message });
  }
}

// ============================================================
// ROTEADOR
// ============================================================
function processarAcao(acao, payload) {
  Logger.log('📥 Ação: ' + acao + ' | Payload: ' + JSON.stringify(payload));

  const acoes = {
    // LOGIN / USUÁRIOS (NOVO)
    'login': loginUsuario,
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

    // ALUNO (NOVO)
    'dadosAlunoCompleto': dadosAlunoCompleto,
    'garantirAbaAluno': garantirAbaAluno,

    // Utilitários
    'ping': () => respostaJSON({ success: true, message: 'pong' }),
    'listarCursos': listarCursosAPI,
    'sincronizar': sincronizarTudo,
    'limparAba': limparAba
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

function listarAcoes() { return Object.keys({
  login:1, criarUsuario:1, criarUsuarioAluno:1, listarUsuarios:1, atualizarSenha:1,
  salvarAluno:1, salvarProfessor:1, salvarMatricula:1, salvarAvaliacoes:1,
  salvarPresenca:1, salvarObservacao:1, salvarRelatorio:1, registrarLogin:1,
  dadosAlunoCompleto:1, ping:1
}); }

function gerarID(prefixo) {
  return prefixo + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 10000);
}

function formatarData(data) {
  if (!data) return '';
  if (typeof data === 'string') {
    if (/^\d{2}-\d{2}-\d{4}$/.test(data)) return data;
    if (/^\d{4}-\d{2}-\d{2}/.test(data)) {
      const p = data.split('T')[0].split('-');
      return `${p[2]}-${p[1]}-${p[0]}`;
    }
    const d = new Date(data);
    if (!isNaN(d.getTime())) {
      return `${String(d.getDate()).padStart(2,'0')}-${String(d.getMonth()+1).padStart(2,'0')}-${d.getFullYear()}`;
    }
    return data;
  }
  if (data instanceof Date) {
    return `${String(data.getDate()).padStart(2,'0')}-${String(data.getMonth()+1).padStart(2,'0')}-${data.getFullYear()}`;
  }
  return '';
}

function dataHoje() { return formatarData(new Date()); }

function getHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
}

function getRegistros(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const headers = getHeaders(sheet);
  const valores = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return valores.map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function encontrarLinha(sheet, criterios) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const headers = getHeaders(sheet);
  const valores = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const colIndex = {};
  headers.forEach((h, i) => { colIndex[h] = i; });
  for (let r = 0; r < valores.length; r++) {
    let match = true;
    for (const [campo, valor] of Object.entries(criterios)) {
      const idx = colIndex[campo];
      if (idx === undefined) { match = false; break; }
      if (String(valores[r][idx]).trim() !== String(valor).trim()) { match = false; break; }
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
// HASH DE SENHA (SHA-256 + salt simples)
// ============================================================
function hashSenha(senha, salt) {
  const raw = String(salt || '') + '|' + String(senha || '');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digest.map(b => {
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
    Logger.log('✅ Aba LOGS criada');
    return sheet;
  }
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, 5).setValues([['ID_Log', 'Nome', 'Perfil', 'Data_Hora', 'Acao']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Garante a aba USUARIOS.
 * Colunas: ID_Usuario | Nome | Perfil | Usuario | Senha_Hash | Salt | Status | Data_Criacao | Aluno_Vinculado
 */
function garantirAbaUsuarios() {
  const ss = getSS();
  let sheet = ss.getSheetByName(ABAS.usuarios);
  if (!sheet) {
    sheet = ss.insertSheet(ABAS.usuarios);
    sheet.appendRow(['ID_Usuario', 'Nome', 'Perfil', 'Usuario', 'Senha_Hash', 'Salt', 'Status', 'Data_Criacao', 'Aluno_Vinculado']);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log('✅ Aba USUARIOS criada');
    return sheet;
  }
  const headers = getHeaders(sheet);
  if (headers.length === 0) {
    sheet.getRange(1, 1, 1, 9).setValues([['ID_Usuario','Nome','Perfil','Usuario','Senha_Hash','Salt','Status','Data_Criacao','Aluno_Vinculado']]);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Garante a existência de uma aba individual para o aluno.
 * A aba contém o resumo do aluno (dados pessoais, cursos, notas, frequência).
 */
function garantirAbaAluno(nomeAluno) {
  const nome = s(nomeAluno).toUpperCase();
  if (!nome) throw new Error('Nome do aluno é obrigatório');

  const ss = getSS();
  // Nome da aba seguro: até 30 chars, sem caracteres inválidos
  let nomeAba = ('ALUNO_' + nome).substring(0, 100).replace(/[\\/?*\[\]:]/g, '');
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
    Logger.log('✅ Aba do aluno criada: ' + nomeAba);
  }
  return sheet;
}

/**
 * Sincroniza dados do aluno (matrículas, notas, presenças, relatórios)
 * na aba individual dele.
 */
function sincronizarAbaAluno(nomeAluno) {
  const nome = s(nomeAluno).toUpperCase();
  const sheet = garantirAbaAluno(nome);

  // Limpa conteúdo dinâmico (a partir da linha 6)
  const lastRow = sheet.getLastRow();
  if (lastRow >= 6) sheet.getRange(6, 1, lastRow - 5, 8).clearContent();

  // Cabeçalho de matrículas
  sheet.getRange(6, 1, 1, 4).setValues([['Curso', 'Professor', 'Status', 'Data_Matrícula']]).setFontWeight('bold');

  // Matrículas do aluno
  const matSheet = getSheet('matriculas');
  const matriculas = getRegistros(matSheet).filter(m => s(m.Aluno).toUpperCase() === nome);

  let row = 7;
  matriculas.forEach(m => {
    sheet.getRange(row, 1, 1, 4).setValues([[s(m.Curso), s(m.Professor), s(m.Status), s(m.Data_Matricula)]]);
    row++;
  });

  // Bloco de notas
  row += 1;
  sheet.getRange(row, 1, 1, 5).setValues([['Curso','Bimestre','Nota','Professor','Data_Lançamento']]).setFontWeight('bold');
  row++;
  const avalSheet = getSheet('avaliacoes');
  const avals = getRegistros(avalSheet).filter(a => s(a.Aluno).toUpperCase() === nome);
  avals.forEach(a => {
    sheet.getRange(row, 1, 1, 5).setValues([[s(a.Curso), s(a.Bimestre), a.Nota, s(a.Professor), s(a.Data_Lancamento)]]);
    row++;
  });

  // Bloco de presenças
  row += 1;
  sheet.getRange(row, 1, 1, 4).setValues([['Curso','Data','Presente','Professor']]).setFontWeight('bold');
  row++;
  const presSheet = getSheet('presencas');
  const presencas = getRegistros(presSheet);
  presencas.forEach(p => {
    const presentes = s(p.Presentes).toUpperCase().split(',').map(x => x.trim());
    if (presentes.includes(nome)) {
      sheet.getRange(row, 1, 1, 4).setValues([[s(p.Curso), s(p.Data), 'SIM', s(p.Professor)]]);
      row++;
    }
  });

  // Bloco de relatórios
  row += 1;
  sheet.getRange(row, 1, 1, 4).setValues([['Curso','Bimestre','Data','Relatório']]).setFontWeight('bold');
  row++;
  const relSheet = getSheet('relatorios');
  const rels = getRegistros(relSheet).filter(r => s(r.Aluno).toUpperCase() === nome);
  rels.forEach(r => {
    sheet.getRange(row, 1, 1, 4).setValues([[s(r.Curso), s(r.Bimestre), s(r.Data), s(r.Relatorio).substring(0, 500)]]);
    row++;
  });

  return { success: true, aba: sheet.getName(), matriculas: matriculas.length };
}

// ============================================================
// LOGIN / USUÁRIOS (NOVO)
// ============================================================

/**
 * Cria um usuário novo (secretaria, coordenação, professor, direção, financeiro).
 * payload: { nome, perfil, usuario, senha, alunoVinculado (opcional) }
 */
function criarUsuario(payload) {
  const { nome, perfil, usuario, senha, alunoVinculado } = payload || {};
  if (!nome || !perfil || !usuario || !senha) {
    return respostaJSON({ success: false, error: 'Nome, perfil, usuário e senha são obrigatórios' });
  }
  if (String(senha).length < 4) {
    return respostaJSON({ success: false, error: 'Senha deve ter ao menos 4 caracteres' });
  }

  const sheet = garantirAbaUsuarios();
  const headers = getHeaders(sheet);
  const usuarioLimpo = s(usuario).toLowerCase();

  // Verifica duplicidade de usuário
  const existente = encontrarLinha(sheet, { 'Usuario': usuarioLimpo });
  if (existente > -1) {
    return respostaJSON({ success: false, error: 'Usuário já existe: ' + usuarioLimpo });
  }

  const salt = gerarSalt();
  const hash = hashSenha(senha, salt);
  const idUser = gerarID('USR');
  const dataCriacao = dataHoje();

  const valores = headers.map(h => {
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

  return respostaJSON({
    success: true,
    message: 'Usuário criado: ' + usuarioLimpo,
    id: idUser,
    usuario: usuarioLimpo,
    perfil: s(perfil).toLowerCase()
  });
}

/**
 * Cria um usuário do tipo aluno. Também cria a aba individual do aluno.
 * payload: { nome (nome completo do aluno), usuario, senha }
 */
function criarUsuarioAluno(payload) {
  const { nome, usuario, senha } = payload || {};
  if (!nome || !usuario || !senha) {
    return respostaJSON({ success: false, error: 'Nome, usuário e senha são obrigatórios' });
  }
  if (String(senha).length < 4) {
    return respostaJSON({ success: false, error: 'Senha deve ter ao menos 4 caracteres' });
  }

  // Verifica se o aluno existe em ALUNOS
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

  const valores = headers.map(h => {
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

  // Cria/sincroniza a aba individual do aluno
  let abaAluno = '';
  try {
    const info = sincronizarAbaAluno(nome);
    abaAluno = info.aba;
  } catch (e) {
    Logger.log('⚠️ Falha ao criar aba individual: ' + e.message);
  }

  return respostaJSON({
    success: true,
    message: 'Usuário de aluno criado: ' + usuarioLimpo,
    id: idUser,
    usuario: usuarioLimpo,
    abaAluno: abaAluno
  });
}

function listarUsuarios(payload) {
  const sheet = garantirAbaUsuarios();
  const registros = getRegistros(sheet).map(r => ({
    id: s(r.ID_Usuario),
    nome: s(r.Nome),
    perfil: s(r.Perfil),
    usuario: s(r.Usuario),
    status: s(r.Status),
    dataCriacao: s(r.Data_Criacao),
    alunoVinculado: s(r.Aluno_Vinculado)
    // hash e salt omitidos
  }));
  return respostaJSON({ success: true, total: registros.length, usuarios: registros });
}

function atualizarSenha(payload) {
  const { usuario, senhaNova } = payload || {};
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
  const { nome } = payload || {};
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  const sheet = garantirAbaUsuarios();
  const linha = encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() });
  if (linha < 0) return respostaJSON({ success: true, encontrado: false });
  const headers = getHeaders(sheet);
  const valores = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const obj = {};
  headers.forEach((h, i) => { obj[h] = valores[i]; });
  return respostaJSON({
    success: true,
    encontrado: true,
    usuario: {
      id: s(obj.ID_Usuario),
      nome: s(obj.Nome),
      perfil: s(obj.Perfil),
      usuario: s(obj.Usuario),
      status: s(obj.Status),
      alunoVinculado: s(obj.Aluno_Vinculado)
    }
  });
}

/**
 * Ação principal de login.
 * payload: { usuario, senha, perfil (opcional para validação cruzada) }
 */
function loginUsuario(payload) {
  try {
    const { usuario, senha, perfil } = payload || {};
    if (!usuario || !senha) {
      return respostaJSON({ success: false, error: 'Usuário e senha são obrigatórios' });
    }

    const sheet = garantirAbaUsuarios();
    const headers = getHeaders(sheet);
    const linha = encontrarLinha(sheet, { 'Usuario': s(usuario).toLowerCase() });

    if (linha < 0) {
      return respostaJSON({ success: false, error: 'Usuário ou senha inválidos' });
    }

    const valores = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
    const obj = {};
    headers.forEach((h, i) => { obj[h] = valores[i]; });

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
      return respostaJSON({
        success: false,
        error: 'Perfil não corresponde. Este usuário é: ' + perfilSalvo
      });
    }

    return respostaJSON({
      success: true,
      message: 'Login OK',
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

// ============================================================
// AÇÕES - ALUNO (NOVO)
// ============================================================

/**
 * Retorna dados consolidados do aluno para a área do aluno.
 * payload: { nome }
 */
function dadosAlunoCompleto(payload) {
  const { nome } = payload || {};
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });

  const nomeUp = s(nome).toUpperCase();

  // Dados pessoais
  const alunosSheet = getSheet('alunos');
  const linha = encontrarLinha(alunosSheet, { 'Nome': nomeUp });
  if (linha < 0) return respostaJSON({ success: false, error: 'Aluno não encontrado' });
  const alunosHeaders = getHeaders(alunosSheet);
  const alunoVals = alunosSheet.getRange(linha, 1, 1, alunosHeaders.length).getValues()[0];
  const aluno = {};
  alunosHeaders.forEach((h, i) => { aluno[h] = alunoVals[i]; });

  // Matrículas
  const matriculas = getRegistros(getSheet('matriculas'))
    .filter(m => s(m.Aluno).toUpperCase() === nomeUp);

  // Avaliações
  const avaliacoes = getRegistros(getSheet('avaliacoes'))
    .filter(a => s(a.Aluno).toUpperCase() === nomeUp);

  // Presenças — filtra apenas as aulas em que o aluno estava
  const todasPres = getRegistros(getSheet('presencas'));
  const presencas = [];
  todasPres.forEach(p => {
    const presentes = s(p.Presentes).toUpperCase().split(',').map(x => x.trim());
    const faltantes = s(p.Faltantes).toUpperCase().split(',').map(x => x.trim());
    if (presentes.includes(nomeUp)) {
      presencas.push({ curso: s(p.Curso), data: s(p.Data), status: 'PRESENTE', conteudo: s(p.Conteudo) || s(p.Observacao), professor: s(p.Professor) });
    } else if (faltantes.includes(nomeUp)) {
      presencas.push({ curso: s(p.Curso), data: s(p.Data), status: 'FALTA', conteudo: s(p.Conteudo) || s(p.Observacao), professor: s(p.Professor) });
    }
  });

  // Conteúdos por curso (para o aluno ver o que foi dado)
  const conteudos = getRegistros(getSheet('conteudos')).map(c => ({
    curso: s(c.Curso), data: s(c.Data), conteudo: s(c.Conteudo), objetivos: s(c.Objetivos), professor: s(c.Professor)
  }));

  // Relatórios
  const relatorios = getRegistros(getSheet('relatorios'))
    .filter(r => s(r.Aluno).toUpperCase() === nomeUp)
    .map(r => ({
      curso: s(r.Curso), bimestre: s(r.Bimestre), professor: s(r.Professor),
      relatorio: s(r.Relatorio), data: s(r.Data), visto: s(r.Visto)
    }));

  return respostaJSON({
    success: true,
    aluno: {
      nome: s(aluno.Nome),
      dataNascimento: s(aluno.Data_Nascimento),
      responsavel: s(aluno.Responsavel),
      contato: s(aluno.Contato),
      status: s(aluno.Status)
    },
    matriculas: matriculas.map(m => ({
      curso: s(m.Curso), professor: s(m.Professor),
      status: s(m.Status), dataMatricula: s(m.Data_Matricula), anoLetivo: s(m.Ano_Letivo)
    })),
    avaliacoes: avaliacoes.map(a => ({
      curso: s(a.Curso), bimestre: Number(a.Bimestre) || 1, nota: Number(a.Nota) || 0,
      professor: s(a.Professor), dataLancamento: s(a.Data_Lancamento)
    })),
    presencas: presencas,
    conteudos: conteudos,
    relatorios: relatorios
  });
}

function garantirAbaAlunoAPI(payload) {
  const { nome } = payload || {};
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  try {
    const info = sincronizarAbaAluno(nome);
    return respostaJSON({ success: true, message: 'Aba sincronizada', ...info });
  } catch (e) {
    return respostaJSON({ success: false, error: e.message });
  }
}

// ============================================================
// DEMAIS AÇÕES (mantidas do original, com pequenos ajustes)
// ============================================================

function salvarAluno(payload) {
  const { nome, dataNascimento, responsavel, contato, status } = payload;
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  const sheet = getSheet('alunos');
  const headers = getHeaders(sheet);
  if (encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() }) > -1) {
    return respostaJSON({ success: false, error: 'Aluno já cadastrado: ' + nome });
  }
  const valores = headers.map(h => {
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
  const { nome, nomeOriginal, dataNascimento, responsavel, contato, status } = payload;
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  const sheet = getSheet('alunos');
  const headers = getHeaders(sheet);
  const linha = encontrarLinha(sheet, { 'Nome': s(nomeOriginal || nome).toUpperCase() });
  if (linha < 0) return respostaJSON({ success: false, error: 'Aluno não encontrado' });
  const antigos = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const novos = headers.map((h, i) => {
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
  const { nome } = payload;
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
  const { id, nome, tipo, curso, telefone, email, valorHora } = payload;
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  const sheet = getSheet('professores');
  garantirColunaValorHora(sheet);
  const headers = getHeaders(sheet);
  const idFinal = s(id) || gerarID('PROF');
  if (encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() }) > -1) {
    return respostaJSON({ success: false, error: 'Professor já cadastrado' });
  }
  const valores = headers.map(h => {
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
  const { id, nome, tipo, curso, telefone, email, status, valorHora } = payload;
  if (!id && !nome) return respostaJSON({ success: false, error: 'ID ou Nome obrigatório' });
  const sheet = getSheet('professores');
  garantirColunaValorHora(sheet);
  const headers = getHeaders(sheet);
  let linha = -1;
  if (id) linha = encontrarLinha(sheet, { 'ID_Professor': s(id) });
  if (linha < 0 && nome) linha = encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() });
  if (linha < 0) return respostaJSON({ success: false, error: 'Professor não encontrado' });
  const antigos = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const novos = headers.map((h, i) => {
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
  const { id, valorHora } = payload;
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
  const { id, nome } = payload;
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
  const { aluno, curso, professor, dataMatricula, anoLetivo, status } = payload;
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
  const valores = headers.map(h => {
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
  return respostaJSON({ success: true, message: 'Matrícula registrada', id });
}

function atualizarMatricula(payload) {
  const { id, aluno, curso, professor, dataMatricula, anoLetivo, status } = payload;
  const sheet = getSheet('matriculas');
  const headers = getHeaders(sheet);
  let linha = -1;
  if (id) linha = encontrarLinha(sheet, { 'ID_Matricula': s(id) });
  if (linha < 0 && aluno && curso) linha = encontrarLinha(sheet, { 'Aluno': s(aluno).toUpperCase(), 'Curso': s(curso) });
  if (linha < 0) return respostaJSON({ success: false, error: 'Matrícula não encontrada' });
  const antigos = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const novos = headers.map((h, i) => {
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
  const { id, aluno, curso } = payload;
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
    if (payload.aluno) registros = registros.filter(r => s(r.Aluno).toUpperCase() === s(payload.aluno).toUpperCase());
    if (payload.curso) registros = registros.filter(r => s(r.Curso) === s(payload.curso));
    if (payload.status) registros = registros.filter(r => s(r.Status).toUpperCase() === s(payload.status).toUpperCase());
  }
  return respostaJSON({ success: true, total: registros.length, matriculas: registros });
}

function salvarAvaliacoes(payload) {
  const lista = payload.avaliacoes || payload.notas || [];
  if (!Array.isArray(lista) || lista.length === 0) return respostaJSON({ success: false, error: 'Sem avaliações' });
  const sheet = getSheet('avaliacoes');
  const headers = getHeaders(sheet);
  const res = { inseridos: 0, atualizados: 0, erros: [] };
  lista.forEach(av => {
    try {
      const aluno = s(av.aluno).toUpperCase();
      const curso = s(av.curso);
      const bimestre = String(av.bimestre || '1');
      const nota = Number(av.nota) || 0;
      const professor = s(av.professor || '');
      const linha = encontrarLinha(sheet, { 'Aluno': aluno, 'Curso': curso, 'Bimestre': bimestre });
      if (linha > -1) {
        const vals = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
        const novos = headers.map((h, i) => {
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
        const valores = headers.map(h => {
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
    } catch (e) { res.erros.push(e.message); }
  });
  return respostaJSON({ success: res.erros.length === 0, message: `${res.inseridos} inserida(s), ${res.atualizados} atualizada(s)`, detalhes: res });
}

function listarAvaliacoes(payload) {
  let registros = getRegistros(getSheet('avaliacoes'));
  if (payload) {
    if (payload.aluno) registros = registros.filter(r => s(r.Aluno).toUpperCase() === s(payload.aluno).toUpperCase());
    if (payload.curso) registros = registros.filter(r => s(r.Curso) === s(payload.curso));
    if (payload.bimestre) registros = registros.filter(r => String(r.Bimestre) === String(payload.bimestre));
  }
  return respostaJSON({ success: true, total: registros.length, avaliacoes: registros });
}

function salvarPresenca(payload) {
  const { curso, data, presentes, faltantes, observacao, conteudo, professor } = payload;
  if (!curso || !data) return respostaJSON({ success: false, error: 'Curso e Data obrigatórios' });
  const sheet = getSheet('presencas');
  const headers = getHeaders(sheet);
  let listaP = Array.isArray(presentes) ? presentes.join(', ') : s(presentes);
  let listaF = Array.isArray(faltantes) ? faltantes.join(', ') : s(faltantes);
  const conteudoFinal = s(conteudo || observacao || '');
  const dataF = formatarData(data);
  const id = gerarID('PRE');
  const valores = headers.map(h => {
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
  if (!headers.some(h => h === 'Conteudo' || h === 'Observacao') && headers.length >= 6) {
    valores[5] = conteudoFinal;
  }
  appendRegistro(sheet, valores);
  return respostaJSON({ success: true, message: 'Presença registrada', id });
}

function listarPresencas(payload) {
  let registros = getRegistros(getSheet('presencas'));
  if (payload) {
    if (payload.curso) registros = registros.filter(r => s(r.Curso) === s(payload.curso));
    if (payload.data) registros = registros.filter(r => s(r.Data) === formatarData(payload.data));
  }
  return respostaJSON({ success: true, total: registros.length, presencas: registros });
}

function salvarConteudo(payload) {
  const { curso, data, conteudo, objetivos, professor } = payload;
  if (!curso || !conteudo) return respostaJSON({ success: false, error: 'Curso e Conteúdo obrigatórios' });
  const sheet = getSheet('conteudos');
  const headers = getHeaders(sheet);
  const id = gerarID('CON');
  const dataF = formatarData(data) || dataHoje();
  const valores = headers.map(h => {
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
  return respostaJSON({ success: true, message: 'Conteúdo salvo', id });
}

function listarConteudos(payload) {
  let registros = getRegistros(getSheet('conteudos'));
  if (payload && payload.curso) registros = registros.filter(r => s(r.Curso) === s(payload.curso));
  return respostaJSON({ success: true, total: registros.length, conteudos: registros });
}

function salvarObservacao(payload) {
  const { curso, data, aluno, observacao, tipo, professor } = payload;
  if (!curso || !aluno || !observacao) return respostaJSON({ success: false, error: 'Campos obrigatórios' });
  const sheet = getSheet('observacoes');
  const headers = getHeaders(sheet);
  const id = gerarID('OBS');
  const dataF = formatarData(data) || dataHoje();
  const valores = headers.map(h => {
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
  return respostaJSON({ success: true, message: 'Observação salva', id });
}

function listarObservacoes(payload) {
  let registros = getRegistros(getSheet('observacoes'));
  if (payload) {
    if (payload.curso) registros = registros.filter(r => s(r.Curso) === s(payload.curso));
    if (payload.aluno) registros = registros.filter(r => s(r.Aluno).toUpperCase() === s(payload.aluno).toUpperCase());
  }
  return respostaJSON({ success: true, total: registros.length, observacoes: registros });
}

function salvarRelatorio(payload) {
  const { aluno, curso, bimestre, professor, relatorio, data } = payload;
  if (!aluno || !curso || !relatorio) return respostaJSON({ success: false, error: 'Campos obrigatórios' });
  const sheet = getSheet('relatorios');
  const headers = getHeaders(sheet);
  const id = gerarID('REL');
  const dataF = formatarData(data) || dataHoje();
  const valores = headers.map(h => {
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
  return respostaJSON({ success: true, message: 'Relatório salvo', id });
}

function marcarRelatorioVisto(payload) {
  const { id } = payload;
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
    if (payload.aluno) registros = registros.filter(r => s(r.Aluno).toUpperCase() === s(payload.aluno).toUpperCase());
    if (payload.curso) registros = registros.filter(r => s(r.Curso) === s(payload.curso));
    if (payload.professor) registros = registros.filter(r => s(r.Professor) === s(payload.professor));
  }
  return respostaJSON({ success: true, total: registros.length, relatorios: registros });
}

function atribuirCurso(payload) {
  const { professorId, professorNome, curso } = payload;
  if (!curso) return respostaJSON({ success: false, error: 'Curso obrigatório' });
  const sheet = getSheet('professores');
  const headers = getHeaders(sheet);
  const idxCurso = headers.indexOf('Curso');
  let linha = -1;
  if (professorId) linha = encontrarLinha(sheet, { 'ID_Professor': s(professorId) });
  if (linha < 0 && professorNome) linha = encontrarLinha(sheet, { 'Nome': s(professorNome).toUpperCase() });
  if (linha < 0) return respostaJSON({ success: false, error: 'Professor não encontrado' });
  const vals = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const cursos = s(vals[idxCurso]).split(',').map(c => c.trim()).filter(Boolean);
  if (cursos.includes(curso)) return respostaJSON({ success: true, message: 'Curso já atribuído', jaExistia: true });
  cursos.push(curso);
  vals[idxCurso] = cursos.join(',');
  atualizarLinha(sheet, linha, vals);
  return respostaJSON({ success: true, message: 'Curso atribuído', cursos });
}

function removerCurso(payload) {
  const { professorId, professorNome, curso } = payload;
  if (!curso) return respostaJSON({ success: false, error: 'Curso obrigatório' });
  const sheet = getSheet('professores');
  const headers = getHeaders(sheet);
  const idxCurso = headers.indexOf('Curso');
  let linha = -1;
  if (professorId) linha = encontrarLinha(sheet, { 'ID_Professor': s(professorId) });
  if (linha < 0 && professorNome) linha = encontrarLinha(sheet, { 'Nome': s(professorNome).toUpperCase() });
  if (linha < 0) return respostaJSON({ success: false, error: 'Professor não encontrado' });
  const vals = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  let cursos = s(vals[idxCurso]).split(',').map(c => c.trim()).filter(Boolean);
  cursos = cursos.filter(c => c !== curso);
  vals[idxCurso] = cursos.join(',');
  atualizarLinha(sheet, linha, vals);
  return respostaJSON({ success: true, message: 'Curso removido', cursos });
}

function solicitarRelatorio(payload) {
  const { aluno, curso, bimestre, tipo, professor } = payload;
  if (!aluno || !curso) return respostaJSON({ success: false, error: 'Aluno e Curso obrigatórios' });
  const sheet = getSheet('relatorios');
  const headers = getHeaders(sheet);
  const id = gerarID('SOL');
  const dataF = dataHoje();
  const texto = '[SOLICITADO EM ' + dataF + ']\n[TIPO: ' + s(tipo || 'pedagogico') + ']\nAguardando preenchimento.';
  const valores = headers.map(h => {
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
  return respostaJSON({ success: true, message: 'Solicitação registrada', id });
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
  const { nome, perfil, acao } = payload || {};
  if (!nome || !perfil) return respostaJSON({ success: false, error: 'Nome e perfil obrigatórios' });
  try {
    const sheet = garantirAbaLogs();
    const headers = getHeaders(sheet);
    const id = gerarID('LOG');
    const dataHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss');
    const valores = headers.map(h => {
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
    return respostaJSON({ success: true, id, dataHora });
  } catch (e) {
    return respostaJSON({ success: false, error: e.message });
  }
}

function listarLogs(payload) {
  let registros = getRegistros(garantirAbaLogs());
  if (payload) {
    if (payload.perfil) registros = registros.filter(r => s(r.Perfil).toLowerCase() === s(payload.perfil).toLowerCase());
    if (payload.nome) registros = registros.filter(r => s(r.Nome).toUpperCase() === s(payload.nome).toUpperCase());
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
  Object.entries(ABAS).forEach(([chave, nomeAba]) => {
    try {
      if (chave === 'logs') garantirAbaLogs();
      if (chave === 'usuarios') garantirAbaUsuarios();
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
  const { aba } = payload;
  if (!aba) return respostaJSON({ success: false, error: 'Aba obrigatória' });
  const chave = Object.keys(ABAS).find(k => ABAS[k] === aba || k === aba);
  if (!chave) return respostaJSON({ success: false, error: 'Aba não reconhecida' });
  const sheet = getSheet(chave);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  return respostaJSON({ success: true, message: 'Aba limpa' });
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
  Logger.log('listarCursos: ' + processarAcao('listarCursos', {}).getContent().substring(0, 200));
  Logger.log('✅ Concluído');
}

function mostrarURLWebApp() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('🔗 URL do Web App:\n\n' + (url || '(vazio — implante primeiro)'));
}
