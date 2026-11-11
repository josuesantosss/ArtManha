/**
 * ============================================================
 * GAS — Sistema de Gestão Escola de Música
 * Versão corrigida: login tolerante + admin padrão + auto-migração
 * ============================================================
 */

const SHEET_ID = '1FRtD7cbV2tvK2og4NDkyWi8V_hziyERNZoSh20gG980';
const PASTA_DRIVE_ID = ''; // opcional

const ABA_ALUNOS = 'ALUNOS';
const ABA_MATRICULAS = 'MATRICULAS';
const ABA_PROFESSORES = 'PROFESSORES';
const ABA_AVALIACOES = 'AVALIACOES';
const ABA_PRESENCAS = 'PRESENCAS';
const ABA_CONTEUDOS = 'CONTEUDOS';
const ABA_OBSERVACOES = 'OBSERVACOES';
const ABA_RELATORIOS = 'RELATORIOS';
const ABA_NOTIFICACOES = 'NOTIFICACOES';
const ABA_ARQUIVOS = 'ARQUIVOS';
const ABA_MENSAGENS = 'MENSAGENS';
const ABA_PAGAMENTOS = 'PAGAMENTOS';
const ABA_LOGS = 'LOGS';
const ABA_USUARIOS = 'USUARIOS';

// Credenciais do admin padrão (criado automaticamente)
const ADMIN_PADRAO = {
  nome: 'ADMINISTRADOR',
  usuario: 'admin',
  senha: 'admin123',
  perfil: 'admin'
};

// ============================================================
// UTILITÁRIOS
// ============================================================
function s(v) { return v === null || v === undefined ? '' : String(v).trim(); }

function respostaJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function logErro(funcao, err) {
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID);
    let aba = sh.getSheetByName(ABA_LOGS);
    if (!aba) {
      aba = sh.insertSheet(ABA_LOGS);
      aba.appendRow(['Data', 'Usuario', 'Perfil', 'Acao', 'Detalhe']);
    }
    aba.appendRow([new Date(), 'SISTEMA', 'erro', funcao, s(err && err.message ? err.message : err)]);
  } catch (e) {}
}

function registrarLog(usuario, perfil, acao, detalhe) {
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID);
    let aba = sh.getSheetByName(ABA_LOGS);
    if (!aba) {
      aba = sh.insertSheet(ABA_LOGS);
      aba.appendRow(['Data', 'Usuario', 'Perfil', 'Acao', 'Detalhe']);
    }
    aba.appendRow([new Date(), s(usuario), s(perfil), s(acao), s(detalhe)]);
  } catch (e) {}
}

function dataHoje() {
  const d = new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MM-yyyy');
}

function dataHoraAgora() {
  const d = new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
}

function gerarID(prefixo) {
  return (prefixo || 'ID') + '_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
}

function gerarSalt() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 16);
}

function hashSenha(senha, salt) {
  const raw = s(senha) + '::' + s(salt);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digest.map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function normalizarChave(str) {
  return s(str).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}

// ============================================================
// ABAS
// ============================================================
function getSheet(nome) {
  const sh = SpreadsheetApp.openById(SHEET_ID);
  const aba = sh.getSheetByName(nome);
  if (!aba) throw new Error('Aba não encontrada: ' + nome);
  return aba;
}

function getHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return s(h); });
}

function garantirAba(nome, headersPadrao) {
  const sh = SpreadsheetApp.openById(SHEET_ID);
  let aba = sh.getSheetByName(nome);
  if (!aba) {
    aba = sh.insertSheet(nome);
    if (headersPadrao && headersPadrao.length) aba.appendRow(headersPadrao);
  }
  return aba;
}

function garantirAbaUsuarios() {
  return garantirAba(ABA_USUARIOS, ['ID_Usuario', 'Nome', 'Perfil', 'Usuario', 'Senha_Hash', 'Salt', 'Status', 'Data_Criacao', 'Aluno_Vinculado']);
}

function appendRegistro(sheet, valores) {
  const headers = getHeaders(sheet);
  const linha = headers.map(function(h, i) { return valores[i] !== undefined ? valores[i] : ''; });
  sheet.appendRow(linha);
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
    for (const campo in criterios) {
      const idx = colIndex[campo];
      if (idx === undefined) { match = false; break; }
      if (normalizarChave(valores[r][idx]) !== normalizarChave(criterios[campo])) { match = false; break; }
    }
    if (match) return r + 2;
  }
  return -1;
}

function encontrarLinhasFlex(sheet, criterios) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const headers = getHeaders(sheet);
  const valores = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const colIndex = {};
  headers.forEach(function(h, i) { colIndex[h] = i; });
  const out = [];
  for (let r = 0; r < valores.length; r++) {
    let match = true;
    for (const campo in criterios) {
      const idx = colIndex[campo];
      if (idx === undefined) { match = false; break; }
      if (normalizarChave(valores[r][idx]) !== normalizarChave(criterios[campo])) { match = false; break; }
    }
    if (match) {
      const obj = {};
      headers.forEach(function(h, i) { obj[h] = valores[r][i]; });
      out.push({ linha: r + 2, valores: valores[r], obj: obj });
    }
  }
  return out;
}

function linhaUsuarioPerfil(sheet, usuario, perfil) {
  const alvoU = normalizarChave(usuario);
  const alvoP = normalizarChave(perfil);
  const linhas = encontrarLinhasFlex(sheet, {});
  for (let i = 0; i < linhas.length; i++) {
    const u = normalizarChave(linhas[i].obj.Usuario);
    const p = normalizarChave(linhas[i].obj.Perfil);
    if (u === alvoU && p === alvoP) return linhas[i].linha;
  }
  return -1;
}

// ============================================================
// ✅ ADMIN PADRÃO — garante que exista sempre
// ============================================================
/**
 * Cria (ou recria) o usuário admin padrão se não existir.
 * Retorna { criado: bool, usuario, senha }.
 */
function garantirUsuarioPadrao() {
  try {
    const sheet = garantirAbaUsuarios();
    const headers = getHeaders(sheet);

    // Verifica se já existe admin com login "admin"
    const existentes = encontrarLinhasFlex(sheet, {}).filter(function(l) {
      return normalizarChave(l.obj.Usuario) === normalizarChave(ADMIN_PADRAO.usuario) &&
             normalizarChave(l.obj.Perfil) === normalizarChave(ADMIN_PADRAO.perfil);
    });

    if (existentes.length > 0) {
      // Já existe. Verifica se a senha bate; se não, corrige o hash.
      const l = existentes[0];
      const saltAtual = s(l.obj.Salt);
      const hashAtual = s(l.obj.Senha_Hash);
      const hashEsperado = hashSenha(ADMIN_PADRAO.senha, saltAtual);
      if (hashAtual !== hashEsperado) {
        // Recalcula o hash com o salt existente (ou novo)
        const novoSalt = saltAtual || gerarSalt();
        const novoHash = hashSenha(ADMIN_PADRAO.senha, novoSalt);
        const idxHash = headers.indexOf('Senha_Hash');
        const idxSalt = headers.indexOf('Salt');
        if (idxSalt >= 0) sheet.getRange(l.linha, idxSalt + 1).setValue(novoSalt);
        if (idxHash >= 0) sheet.getRange(l.linha, idxHash + 1).setValue(novoHash);
        // Garante status ATIVO
        const idxStatus = headers.indexOf('Status');
        if (idxStatus >= 0) sheet.getRange(l.linha, idxStatus + 1).setValue('ATIVO');
        return { criado: false, corrigido: true, usuario: ADMIN_PADRAO.usuario, senha: ADMIN_PADRAO.senha };
      }
      return { criado: false, usuario: ADMIN_PADRAO.usuario, senha: ADMIN_PADRAO.senha };
    }

    // Cria do zero
    const salt = gerarSalt();
    const hash = hashSenha(ADMIN_PADRAO.senha, salt);
    const valores = headers.map(function(h) {
      switch (h) {
        case 'ID_Usuario': return gerarID('USR');
        case 'Nome': return ADMIN_PADRAO.nome;
        case 'Perfil': return ADMIN_PADRAO.perfil;
        case 'Usuario': return ADMIN_PADRAO.usuario;
        case 'Senha_Hash': return hash;
        case 'Salt': return salt;
        case 'Status': return 'ATIVO';
        case 'Data_Criacao': return dataHoje();
        case 'Aluno_Vinculado': return '';
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    return { criado: true, usuario: ADMIN_PADRAO.usuario, senha: ADMIN_PADRAO.senha };
  } catch (e) {
    logErro('garantirUsuarioPadrao', e);
    return { criado: false, erro: e.message };
  }
}

/**
 * Função para rodar UMA VEZ no editor do Apps Script
 * (Run > criarUsuarioPadraoAdmin). Cria/atualiza o admin.
 */
function criarUsuarioPadraoAdmin() {
  const r = garantirUsuarioPadrao();
  Logger.log(JSON.stringify(r));
  return r;
}

/**
 * Diagnóstico: lista todos os usuários da planilha (para debug).
 * Rode no editor do Apps Script.
 */
function debugUsuarios() {
  const sheet = garantirAbaUsuarios();
  const linhas = encontrarLinhasFlex(sheet, {});
  const out = linhas.map(function(l) {
    return {
      linha: l.linha,
      id: s(l.obj.ID_Usuario),
      nome: s(l.obj.Nome),
      perfil: s(l.obj.Perfil),
      usuario: s(l.obj.Usuario),
      status: s(l.obj.Status),
      alunoVinculado: s(l.obj.Aluno_Vinculado),
      temHash: s(l.obj.Senha_Hash).length > 0,
      temSalt: s(l.obj.Salt).length > 0
    };
  });
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

/**
 * Reset total (CUIDADO): apaga todos os usuários e recria o admin padrão.
 * Rode no editor do Apps Script apenas se necessário.
 */
function resetarTodosUsuarios() {
  const sh = SpreadsheetApp.openById(SHEET_ID);
  const aba = sh.getSheetByName(ABA_USUARIOS);
  if (!aba) {
    garantirAbaUsuarios();
  } else {
    const lastRow = aba.getLastRow();
    if (lastRow > 1) {
      aba.getRange(2, 1, lastRow - 1, aba.getLastColumn()).clearContent();
    }
  }
  const r = garantirUsuarioPadrao();
  Logger.log('Reset concluído. Admin: ' + JSON.stringify(r));
  return r;
}

// ============================================================
// HANDLERS PRINCIPAIS
// ============================================================
function doGet(e) {
  // Garante admin padrão no primeiro acesso
  try { garantirUsuarioPadrao(); } catch (err) {}
  return respostaJSON({ success: true, message: 'GAS Escola de Música ativo', timestamp: new Date().toISOString() });
}

function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    const acao = dados.acao;
    const payload = dados.payload || {};
    const resultado = rotear(acao, payload);
    return respostaJSON(resultado);
  } catch (err) {
    logErro('doPost', err);
    return respostaJSON({ success: false, error: err.message });
  }
}

function rotear(acao, payload) {
  switch (acao) {
    case 'ping': return { success: true, message: 'pong', timestamp: new Date().toISOString() };
    case 'login': return loginUsuario(payload);
    case 'registrarLogin': return { success: true };
    case 'logout': return { success: true };
    case 'criarUsuario': return criarUsuario(payload);
    case 'criarUsuarioAluno': return criarUsuarioAluno(payload);
    case 'listarUsuarios': return listarUsuarios(payload);
    case 'listarUsuariosPorLogin': return listarUsuariosPorLogin(payload);
    case 'atualizarSenha': return atualizarSenha(payload);
    case 'adminVerSenha': return adminVerSenha(payload);
    case 'salvarAluno': return salvarAluno(payload);
    case 'atualizarAluno': return atualizarAluno(payload);
    case 'salvarMatricula': return salvarMatricula(payload);
    case 'atualizarMatricula': return atualizarMatricula(payload);
    case 'salvarProfessor': return salvarProfessor(payload);
    case 'atualizarProfessor': return atualizarProfessor(payload);
    case 'salvarValorProfessor': return salvarValorProfessor(payload);
    case 'salvarAvaliacoes': return salvarAvaliacoes(payload);
    case 'salvarPresenca': return salvarPresenca(payload);
    case 'salvarObservacao': return salvarObservacao(payload);
    case 'salvarRelatorio': return salvarRelatorio(payload);
    case 'salvarPagamento': return salvarPagamento(payload);
    case 'listarPagamentos': return listarPagamentos(payload);
    case 'notificarAtraso': return notificarAtraso(payload);
    case 'notificarPagamento': return notificarPagamento(payload);
    case 'listarNotificacoes': return listarNotificacoes(payload);
    case 'listarMensagens': return listarMensagens(payload);
    case 'listarArquivos': return listarArquivos(payload);
    case 'marcarNotificacoesLidas': return marcarNotificacoesLidas(payload);
    case 'enviarMensagemAluno': return enviarMensagemAluno(payload);
    case 'uploadArquivo': return uploadArquivo(payload);
    case 'dadosAlunoCompleto': return dadosAlunoCompleto(payload);
    case 'adminListarLogs': return adminListarLogs(payload);
    case 'adminListarBugs': return adminListarBugs(payload);
    case 'garantirAdminPadrao': {
      try { return { success: true, ...garantirUsuarioPadrao() }; }
      catch (err) { return { success: false, error: err.message }; }
    }
    default: return { success: false, error: 'Ação desconhecida: ' + acao };
  }
}

// ============================================================
// ✅ LOGIN — tolerante: hash, texto puro e migração automática
// ============================================================
function loginUsuario(payload) {
  try {
    const usuario = s(payload.usuario).toLowerCase();
    const senha = s(payload.senha);
    const perfilSolicitado = s(payload.perfil).toLowerCase();

    if (!usuario || !senha) {
      return { success: false, error: 'Usuário e senha são obrigatórios' };
    }

    // Garante que o admin padrão exista (auto-cria se necessário)
    try { garantirUsuarioPadrao(); } catch (err) {}

    const sheet = garantirAbaUsuarios();
    const headers = getHeaders(sheet);
    const idxHash = headers.indexOf('Senha_Hash');
    const idxSalt = headers.indexOf('Salt');

    const linhas = encontrarLinhasFlex(sheet, {}).filter(function(l) {
      return normalizarChave(l.obj.Usuario) === normalizarChave(usuario);
    });
    if (linhas.length === 0) {
      return { success: false, error: 'Usuário ou senha inválidos' };
    }

    const validos = [];
    linhas.forEach(function(l) {
      if (s(l.obj.Status).toUpperCase() !== 'ATIVO') return;
      const salt = s(l.obj.Salt);
      const hashArmazenado = s(l.obj.Senha_Hash);
      if (!hashArmazenado) return;

      // 1) Tenta hash normal
      const hashCalculado = hashSenha(senha, salt);
      if (hashCalculado === hashArmazenado) {
        validos.push(l);
        return;
      }

      // 2) Fallback: se o hash armazenado estiver em texto puro (legado/manual)
      //    e a senha bater literalmente, aceita e MIGRA para hash.
      if (hashArmazenado === senha) {
        try {
          const novoSalt = salt || gerarSalt();
          const novoHash = hashSenha(senha, novoSalt);
          if (idxSalt >= 0) sheet.getRange(l.linha, idxSalt + 1).setValue(novoSalt);
          if (idxHash >= 0) sheet.getRange(l.linha, idxHash + 1).setValue(novoHash);
          Logger.log('🔐 Senha em texto puro migrada para hash: ' + usuario + ' (' + s(l.obj.Perfil) + ')');
        } catch (err) { /* ignora falha de migração */ }
        validos.push(l);
        return;
      }

      // 3) Fallback: hash sem salt (alguns sistemas usam hash puro de senha)
      if (!salt) {
        const hashPuro = hashSenha(senha, '');
        if (hashPuro === hashArmazenado) {
          // migra para hash com salt
          try {
            const novoSalt = gerarSalt();
            const novoHash = hashSenha(senha, novoSalt);
            if (idxSalt >= 0) sheet.getRange(l.linha, idxSalt + 1).setValue(novoSalt);
            if (idxHash >= 0) sheet.getRange(l.linha, idxHash + 1).setValue(novoHash);
            Logger.log('🔐 Senha sem salt migrada: ' + usuario);
          } catch (err) {}
          validos.push(l);
          return;
        }
      }
    });

    if (validos.length === 0) {
      return { success: false, error: 'Usuário ou senha inválidos' };
    }

    // 1) Perfil exato
    const doPerfil = validos.find(function(l) {
      return normalizarChave(l.obj.Perfil) === normalizarChave(perfilSolicitado);
    });
    if (doPerfil) {
      return {
        success: true,
        message: 'Login OK',
        token: gerarID('TOKEN'),
        nome: s(doPerfil.obj.Nome).toUpperCase(),
        perfil: s(doPerfil.obj.Perfil).toLowerCase(),
        perfilEfetivo: s(doPerfil.obj.Perfil).toLowerCase(),
        usuario: s(doPerfil.obj.Usuario),
        alunoVinculado: s(doPerfil.obj.Aluno_Vinculado),
        expiraEm: Date.now() + 60 * 60 * 1000
      };
    }

    // 2) Admin override
    const adminValido = validos.find(function(l) {
      return normalizarChave(l.obj.Perfil) === 'admin';
    });
    if (adminValido) {
      return {
        success: true,
        message: 'Login OK (admin em outro perfil)',
        token: gerarID('TOKEN'),
        nome: s(adminValido.obj.Nome).toUpperCase(),
        perfil: 'admin',
        perfilEfetivo: perfilSolicitado || 'admin',
        usuario: s(adminValido.obj.Usuario),
        alunoVinculado: '',
        expiraEm: Date.now() + 60 * 60 * 1000,
        adminOverride: true
      };
    }

    // 3) Perfil não bate
    const perfisDisponiveis = validos.map(function(l) { return s(l.obj.Perfil); });
    return {
      success: false,
      error: 'Este login não possui o perfil "' + perfilSolicitado + '". Perfis disponíveis: ' + perfisDisponiveis.join(', '),
      perfisDisponiveis: perfisDisponiveis
    };
  } catch (e) {
    logErro('loginUsuario', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// CRIAR USUÁRIO
// ============================================================
function criarUsuario(payload) {
  const nome = s(payload.nome);
  const perfil = s(payload.perfil).toLowerCase();
  const usuario = s(payload.usuario).toLowerCase();
  const senha = s(payload.senha);
  const alunoVinculado = s(payload.alunoVinculado);

  if (!nome || !perfil || !usuario || !senha) {
    return { success: false, error: 'Nome, perfil, usuário e senha são obrigatórios' };
  }
  if (senha.length < 4) {
    return { success: false, error: 'Senha deve ter ao menos 4 caracteres' };
  }

  const sheet = garantirAbaUsuarios();
  const headers = getHeaders(sheet);

  if (linhaUsuarioPerfil(sheet, usuario, perfil) > -1) {
    return {
      success: false,
      error: 'Já existe um usuário "' + usuario + '" com perfil "' + perfil + '". Escolha outro perfil ou outro login.'
    };
  }

  const outros = [];
  encontrarLinhasFlex(sheet, {}).forEach(function(l) {
    if (normalizarChave(l.obj.Usuario) === normalizarChave(usuario)) {
      outros.push(s(l.obj.Perfil));
    }
  });

  const salt = gerarSalt();
  const hash = hashSenha(senha, salt);
  const idUser = gerarID('USR');
  const dataCriacao = dataHoje();

  const valores = headers.map(function(h) {
    switch (h) {
      case 'ID_Usuario': return idUser;
      case 'Nome': return s(nome).toUpperCase();
      case 'Perfil': return perfil;
      case 'Usuario': return usuario;
      case 'Senha_Hash': return hash;
      case 'Salt': return salt;
      case 'Status': return 'ATIVO';
      case 'Data_Criacao': return dataCriacao;
      case 'Aluno_Vinculado': return s(alunoVinculado).toUpperCase();
      default: return '';
    }
  });
  appendRegistro(sheet, valores);

  return {
    success: true,
    message: 'Usuário criado: ' + usuario + ' (' + perfil + ')',
    id: idUser,
    usuario: usuario,
    perfil: perfil,
    outrosPerfis: outros
  };
}

// ============================================================
// CRIAR USUÁRIO ALUNO
// ============================================================
function criarUsuarioAluno(payload) {
  const nome = s(payload.nome);
  const usuario = s(payload.usuario).toLowerCase();
  const senha = s(payload.senha);

  if (!nome || !usuario || !senha) {
    return { success: false, error: 'Nome, usuário e senha são obrigatórios' };
  }
  if (senha.length < 4) {
    return { success: false, error: 'Senha deve ter ao menos 4 caracteres' };
  }

  const alunoSheet = getSheet(ABA_ALUNOS);
  const alunoLinha = encontrarLinha(alunoSheet, { 'Nome': s(nome).toUpperCase() });
  if (alunoLinha < 0) {
    return { success: false, error: 'Aluno não encontrado na aba ALUNOS: ' + nome };
  }

  const sheet = garantirAbaUsuarios();
  const headers = getHeaders(sheet);

  if (linhaUsuarioPerfil(sheet, usuario, 'aluno') > -1) {
    return {
      success: false,
      error: 'Já existe um aluno com o login "' + usuario + '". Escolha outro.'
    };
  }

  const salt = gerarSalt();
  const hash = hashSenha(senha, salt);
  const idUser = gerarID('USR');
  const dataCriacao = dataHoje();

  const valores = headers.map(function(h) {
    switch (h) {
      case 'ID_Usuario': return idUser;
      case 'Nome': return s(nome).toUpperCase();
      case 'Perfil': return 'aluno';
      case 'Usuario': return usuario;
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
    abaAluno = sincronizarAbaAluno(nome);
  } catch (e) {
    Logger.log('⚠️ Falha ao criar aba individual: ' + e.message);
  }

  const outros = [];
  encontrarLinhasFlex(sheet, {}).forEach(function(l) {
    if (normalizarChave(l.obj.Usuario) === normalizarChave(usuario) && normalizarChave(l.obj.Perfil) !== 'aluno') {
      outros.push(s(l.obj.Perfil));
    }
  });

  return {
    success: true,
    message: 'Usuário de aluno criado: ' + usuario,
    id: idUser,
    usuario: usuario,
    abaAluno: abaAluno,
    outrosPerfis: outros
  };
}

// ============================================================
// LISTAR USUÁRIOS
// ============================================================
function listarUsuarios(payload) {
  try {
    const sheet = garantirAbaUsuarios();
    const linhas = encontrarLinhasFlex(sheet, {});
    const usuarios = linhas.map(function(l) {
      return {
        id: s(l.obj.ID_Usuario),
        nome: s(l.obj.Nome),
        perfil: s(l.obj.Perfil),
        usuario: s(l.obj.Usuario),
        status: s(l.obj.Status),
        alunoVinculado: s(l.obj.Aluno_Vinculado),
        dataCriacao: s(l.obj.Data_Criacao)
      };
    });
    return { success: true, total: usuarios.length, usuarios: usuarios };
  } catch (e) {
    logErro('listarUsuarios', e);
    return { success: false, error: e.message };
  }
}

function listarUsuariosPorLogin(payload) {
  try {
    const usuario = s(payload.usuario).toLowerCase();
    if (!usuario) return { success: false, error: 'Usuário obrigatório' };
    const sheet = garantirAbaUsuarios();
    const linhas = encontrarLinhasFlex(sheet, {}).filter(function(l) {
      return normalizarChave(l.obj.Usuario) === normalizarChave(usuario);
    });
    const registros = linhas.map(function(l) {
      return {
        id: s(l.obj.ID_Usuario),
        nome: s(l.obj.Nome),
        perfil: s(l.obj.Perfil),
        usuario: s(l.obj.Usuario),
        status: s(l.obj.Status),
        alunoVinculado: s(l.obj.Aluno_Vinculado)
      };
    });
    return {
      success: true,
      total: registros.length,
      usuarios: registros,
      perfis: registros.map(function(r) { return r.perfil; })
    };
  } catch (e) {
    logErro('listarUsuariosPorLogin', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ATUALIZAR SENHA
// ============================================================
function atualizarSenha(payload) {
  const usuario = s(payload.usuario).toLowerCase();
  const senhaNova = s(payload.senhaNova);
  const perfil = s(payload.perfil).toLowerCase();

  if (!usuario || !senhaNova) {
    return { success: false, error: 'Usuário e nova senha são obrigatórios' };
  }
  if (senhaNova.length < 4) {
    return { success: false, error: 'Senha deve ter ao menos 4 caracteres' };
  }

  const sheet = garantirAbaUsuarios();
  const headers = getHeaders(sheet);
  const linhasTodas = encontrarLinhasFlex(sheet, {}).filter(function(l) {
    return normalizarChave(l.obj.Usuario) === normalizarChave(usuario);
  });
  if (linhasTodas.length === 0) {
    return { success: false, error: 'Usuário não encontrado: ' + usuario };
  }

  if (!perfil && linhasTodas.length > 1) {
    return {
      success: false,
      error: 'Este login existe em ' + linhasTodas.length + ' perfis. Especifique o perfil.',
      perfis: linhasTodas.map(function(l) { return s(l.obj.Perfil); })
    };
  }

  const alvo = perfil
    ? linhasTodas.filter(function(l) { return normalizarChave(l.obj.Perfil) === normalizarChave(perfil); })
    : linhasTodas;

  if (alvo.length === 0) {
    return { success: false, error: 'Nenhum registro encontrado para o perfil "' + perfil + '"' };
  }

  const idxHash = headers.indexOf('Senha_Hash');
  const idxSalt = headers.indexOf('Salt');
  alvo.forEach(function(l) {
    const novoSalt = gerarSalt();
    const novoHash = hashSenha(senhaNova, novoSalt);
    if (idxSalt >= 0) sheet.getRange(l.linha, idxSalt + 1).setValue(novoSalt);
    if (idxHash >= 0) sheet.getRange(l.linha, idxHash + 1).setValue(novoHash);
  });

  return {
    success: true,
    message: 'Senha redefinida para ' + alvo.length + ' registro(s)',
    perfis: alvo.map(function(l) { return s(l.obj.Perfil); })
  };
}

// ============================================================
// ADMIN VER SENHA
// ============================================================
function adminVerSenha(payload) {
  try {
    const usuario = s(payload.usuario).toLowerCase();
    const perfil = s(payload.perfil).toLowerCase();
    if (!usuario) return { success: false, error: 'Usuário obrigatório' };

    const sheet = garantirAbaUsuarios();
    const linhas = encontrarLinhasFlex(sheet, {}).filter(function(l) {
      return normalizarChave(l.obj.Usuario) === normalizarChave(usuario);
    });
    if (linhas.length === 0) return { success: false, error: 'Usuário não encontrado' };

    let alvo = linhas;
    if (perfil) {
      alvo = linhas.filter(function(l) {
        return normalizarChave(l.obj.Perfil) === normalizarChave(perfil);
      });
      if (alvo.length === 0) return { success: false, error: 'Perfil não encontrado para este usuário' };
    }
    if (alvo.length > 1) {
      return {
        success: false,
        error: 'Múltiplos perfis. Especifique o perfil.',
        perfis: alvo.map(function(l) { return s(l.obj.Perfil); })
      };
    }

    const senhaHash = s(alvo[0].obj.Senha_Hash);
    return {
      success: true,
      usuario: usuario,
      perfil: s(alvo[0].obj.Perfil),
      senha: '(hash) ' + senhaHash.substring(0, 16) + '...',
      nota: 'A senha é armazenada como hash. Não é possível recuperar a senha original.'
    };
  } catch (e) {
    logErro('adminVerSenha', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ALUNOS
// ============================================================
function salvarAluno(payload) {
  try {
    const sheet = getSheet(ABA_ALUNOS);
    const headers = getHeaders(sheet);
    const nome = s(payload.nome).toUpperCase();
    if (!nome) return { success: false, error: 'Nome obrigatório' };

    const linhaExistente = encontrarLinha(sheet, { 'Nome': nome });
    const valores = headers.map(function(h) {
      switch (h) {
        case 'Nome': return nome;
        case 'Data_Nascimento': return s(payload.dataNascimento);
        case 'Responsavel': return s(payload.responsavel).toUpperCase();
        case 'Contato': return s(payload.contato);
        case 'Status': return s(payload.status || 'ATIVO').toUpperCase();
        case 'Mensalidade': return Number(payload.mensalidade) || 0;
        case 'Vencimento': return Number(payload.vencimento) || 10;
        default: return '';
      }
    });

    if (linhaExistente > 0) {
      sheet.getRange(linhaExistente, 1, 1, headers.length).setValues([valores]);
      return { success: true, message: 'Aluno atualizado: ' + nome };
    }
    appendRegistro(sheet, valores);
    return { success: true, message: 'Aluno criado: ' + nome };
  } catch (e) {
    logErro('salvarAluno', e);
    return { success: false, error: e.message };
  }
}

function atualizarAluno(payload) {
  try {
    const sheet = getSheet(ABA_ALUNOS);
    const headers = getHeaders(sheet);
    const nome = s(payload.nome).toUpperCase();
    const linha = encontrarLinha(sheet, { 'Nome': nome });
    if (linha < 0) return { success: false, error: 'Aluno não encontrado' };

    const idxMens = headers.indexOf('Mensalidade');
    const idxVenc = headers.indexOf('Vencimento');
    if (idxMens >= 0 && payload.mensalidade !== undefined) {
      sheet.getRange(linha, idxMens + 1).setValue(Number(payload.mensalidade) || 0);
    }
    if (idxVenc >= 0 && payload.vencimento !== undefined) {
      sheet.getRange(linha, idxVenc + 1).setValue(Number(payload.vencimento) || 10);
    }
    return { success: true, message: 'Aluno atualizado' };
  } catch (e) {
    logErro('atualizarAluno', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// MATRÍCULAS
// ============================================================
function salvarMatricula(payload) {
  try {
    const sheet = getSheet(ABA_MATRICULAS);
    const headers = getHeaders(sheet);
    const aluno = s(payload.aluno).toUpperCase();
    const curso = s(payload.curso).toLowerCase();
    if (!aluno || !curso) return { success: false, error: 'Aluno e curso obrigatórios' };

    const valores = headers.map(function(h) {
      switch (h) {
        case 'ID_Matricula': return gerarID('MAT');
        case 'Aluno': return aluno;
        case 'Curso': return curso;
        case 'Professor': return s(payload.professor).toUpperCase();
        case 'Data_Matricula': return s(payload.dataMatricula) || dataHoje();
        case 'Ano_Letivo': return s(payload.anoLetivo);
        case 'Status': return s(payload.status || 'ATIVO').toUpperCase();
        case 'Mensalidade': return Number(payload.mensalidade) || 0;
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    return { success: true, message: 'Matrícula criada' };
  } catch (e) {
    logErro('salvarMatricula', e);
    return { success: false, error: e.message };
  }
}

function atualizarMatricula(payload) {
  try {
    const sheet = getSheet(ABA_MATRICULAS);
    const headers = getHeaders(sheet);
    const aluno = s(payload.aluno).toUpperCase();
    const curso = s(payload.curso).toLowerCase();
    const linhas = encontrarLinhasFlex(sheet, { 'Aluno': aluno, 'Curso': curso });
    if (linhas.length === 0) return { success: false, error: 'Matrícula não encontrada' };

    const idxStatus = headers.indexOf('Status');
    linhas.forEach(function(l) {
      sheet.getRange(l.linha, idxStatus + 1).setValue(s(payload.status || 'ATIVO').toUpperCase());
    });
    return { success: true, message: 'Matrícula atualizada' };
  } catch (e) {
    logErro('atualizarMatricula', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// PROFESSORES
// ============================================================
function salvarProfessor(payload) {
  try {
    const sheet = getSheet(ABA_PROFESSORES);
    const headers = getHeaders(sheet);
    const nome = s(payload.nome).toUpperCase();
    if (!nome) return { success: false, error: 'Nome obrigatório' };

    const valores = headers.map(function(h) {
      switch (h) {
        case 'ID_Professor': return s(payload.id) || gerarID('PROF');
        case 'Nome': return nome;
        case 'Tipo': return s(payload.tipo);
        case 'Curso': return s(payload.curso).toLowerCase();
        case 'Telefone': return s(payload.telefone);
        case 'Email': return s(payload.email);
        case 'Status': return 'ATIVO';
        case 'Valor_Hora': return Number(payload.valorHora) || 0;
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    return { success: true, message: 'Professor criado' };
  } catch (e) {
    logErro('salvarProfessor', e);
    return { success: false, error: e.message };
  }
}

function atualizarProfessor(payload) {
  try {
    const sheet = getSheet(ABA_PROFESSORES);
    const headers = getHeaders(sheet);
    const id = s(payload.id);
    const linha = encontrarLinha(sheet, { 'ID_Professor': id });
    if (linha < 0) return { success: false, error: 'Professor não encontrado' };

    if (payload.curso !== undefined) {
      const idx = headers.indexOf('Curso');
      if (idx >= 0) sheet.getRange(linha, idx + 1).setValue(s(payload.curso).toLowerCase());
    }
    return { success: true, message: 'Professor atualizado' };
  } catch (e) {
    logErro('atualizarProfessor', e);
    return { success: false, error: e.message };
  }
}

function salvarValorProfessor(payload) {
  try {
    const sheet = getSheet(ABA_PROFESSORES);
    const headers = getHeaders(sheet);
    const id = s(payload.id);
    const linha = encontrarLinha(sheet, { 'ID_Professor': id });
    if (linha < 0) return { success: false, error: 'Professor não encontrado' };
    const idx = headers.indexOf('Valor_Hora');
    if (idx >= 0) sheet.getRange(linha, idx + 1).setValue(Number(payload.valorHora) || 0);
    return { success: true, message: 'Valor atualizado' };
  } catch (e) {
    logErro('salvarValorProfessor', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// AVALIAÇÕES / PRESENÇAS / OBSERVAÇÕES / RELATÓRIOS
// ============================================================
function salvarAvaliacoes(payload) {
  try {
    const avals = payload.avaliacoes || [];
    const sheet = getSheet(ABA_AVALIACOES);
    const headers = getHeaders(sheet);
    avals.forEach(function(av) {
      const valores = headers.map(function(h) {
        switch (h) {
          case 'Aluno': return s(av.aluno).toUpperCase();
          case 'Curso': return s(av.curso).toLowerCase();
          case 'Bimestre': return Number(av.bimestre) || 1;
          case 'Nota': return Number(av.nota) || 0;
          case 'Professor': return s(av.professor).toUpperCase();
          case 'Data_Lancamento': return dataHoje();
          default: return '';
        }
      });
      appendRegistro(sheet, valores);
    });
    return { success: true, message: 'Avaliações salvas' };
  } catch (e) {
    logErro('salvarAvaliacoes', e);
    return { success: false, error: e.message };
  }
}

function salvarPresenca(payload) {
  try {
    const sheet = getSheet(ABA_PRESENCAS);
    const headers = getHeaders(sheet);
    const presentes = payload.presentes || [];
    const valores = headers.map(function(h) {
      switch (h) {
        case 'Data': return s(payload.data);
        case 'Curso': return s(payload.curso).toLowerCase();
        case 'Presentes': return presentes.join(', ');
        case 'Faltantes': return '';
        case 'Professor': return s(payload.professor).toUpperCase();
        case 'Conteudo': return s(payload.conteudo);
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    return { success: true, message: 'Presença salva' };
  } catch (e) {
    logErro('salvarPresenca', e);
    return { success: false, error: e.message };
  }
}

function salvarObservacao(payload) {
  try {
    const sheet = getSheet(ABA_OBSERVACOES);
    const headers = getHeaders(sheet);
    const valores = headers.map(function(h) {
      switch (h) {
        case 'Data': return s(payload.data);
        case 'Curso': return s(payload.curso).toLowerCase();
        case 'Aluno': return s(payload.aluno).toUpperCase();
        case 'Observacao': return s(payload.observacao);
        case 'Tipo': return s(payload.tipo || 'Neutra');
        case 'Professor': return s(payload.professor).toUpperCase();
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    return { success: true, message: 'Observação salva' };
  } catch (e) {
    logErro('salvarObservacao', e);
    return { success: false, error: e.message };
  }
}

function salvarRelatorio(payload) {
  try {
    const sheet = getSheet(ABA_RELATORIOS);
    const headers = getHeaders(sheet);
    const valores = headers.map(function(h) {
      switch (h) {
        case 'Aluno': return s(payload.aluno).toUpperCase();
        case 'Curso': return s(payload.curso).toLowerCase();
        case 'Bimestre': return s(payload.bimestre);
        case 'Professor': return s(payload.professor).toUpperCase();
        case 'Relatorio': return s(payload.relatorio);
        case 'Data': return s(payload.data) || dataHoje();
        case 'Visto': return '';
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    return { success: true, message: 'Relatório salvo' };
  } catch (e) {
    logErro('salvarRelatorio', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// PAGAMENTOS
// ============================================================
function salvarPagamento(payload) {
  try {
    const sheet = getSheet(ABA_PAGAMENTOS);
    const headers = getHeaders(sheet);
    const aluno = s(payload.aluno).toUpperCase();
    if (!aluno) return { success: false, error: 'Aluno obrigatório' };

    const ref = s(payload.referencia);
    const linhas = encontrarLinhasFlex(sheet, { 'Aluno': aluno, 'Referencia': ref });

    const valores = headers.map(function(h) {
      switch (h) {
        case 'Aluno': return aluno;
        case 'Valor': return Number(payload.valor) || 0;
        case 'Vencimento': return Number(payload.vencimento) || 10;
        case 'Status': return s(payload.status || 'PENDENTE').toUpperCase();
        case 'Data_Pagamento': return s(payload.dataPagamento);
        case 'Referencia': return ref;
        default: return '';
      }
    });

    if (linhas.length > 0) {
      sheet.getRange(linhas[0].linha, 1, 1, headers.length).setValues([valores]);
      return { success: true, message: 'Pagamento atualizado' };
    }
    appendRegistro(sheet, valores);
    return { success: true, message: 'Pagamento criado' };
  } catch (e) {
    logErro('salvarPagamento', e);
    return { success: false, error: e.message };
  }
}

function listarPagamentos(payload) {
  try {
    const sheet = getSheet(ABA_PAGAMENTOS);
    const linhas = encontrarLinhasFlex(sheet, {});
    const pagamentos = linhas.map(function(l) {
      return {
        aluno: s(l.obj.Aluno),
        valor: Number(l.obj.Valor) || 0,
        vencimento: Number(l.obj.Vencimento) || 10,
        status: s(l.obj.Status),
        dataPagamento: s(l.obj.Data_Pagamento),
        referencia: s(l.obj.Referencia)
      };
    });
    return { success: true, total: pagamentos.length, pagamentos: pagamentos };
  } catch (e) {
    logErro('listarPagamentos', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// NOTIFICAÇÕES
// ============================================================
function notificarAtraso(payload) {
  try {
    const sheet = getSheet(ABA_NOTIFICACOES);
    const headers = getHeaders(sheet);
    const aluno = s(payload.aluno).toUpperCase();
    const ref = s(payload.referencia);
    const valores = headers.map(function(h) {
      switch (h) {
        case 'ID': return gerarID('NOTIF');
        case 'Aluno': return aluno;
        case 'Tipo': return 'atraso';
        case 'Titulo': return '⚠️ Mensalidade em atraso — ' + ref;
        case 'Descricao': return 'Sua mensalidade de ' + ref + ' está em atraso.';
        case 'Curso': return '';
        case 'Professor': return 'SISTEMA';
        case 'Lida': return 'NAO';
        case 'Data': return dataHoraAgora();
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    return { success: true };
  } catch (e) {
    logErro('notificarAtraso', e);
    return { success: false, error: e.message };
  }
}

function notificarPagamento(payload) {
  try {
    const sheet = getSheet(ABA_NOTIFICACOES);
    const headers = getHeaders(sheet);
    const aluno = s(payload.aluno).toUpperCase();
    const ref = s(payload.referencia);
    const valores = headers.map(function(h) {
      switch (h) {
        case 'ID': return gerarID('NOTIF');
        case 'Aluno': return aluno;
        case 'Tipo': return 'pagamento';
        case 'Titulo': return '💰 Pagamento confirmado — ' + ref;
        case 'Descricao': return 'Seu pagamento de R$ ' + Number(payload.valor || 0).toFixed(2) + ' foi confirmado.';
        case 'Curso': return '';
        case 'Professor': return 'SISTEMA';
        case 'Lida': return 'NAO';
        case 'Data': return dataHoraAgora();
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    return { success: true };
  } catch (e) {
    logErro('notificarPagamento', e);
    return { success: false, error: e.message };
  }
}

function listarNotificacoes(payload) {
  try {
    const aluno = s(payload.aluno).toUpperCase();
    const sheet = getSheet(ABA_NOTIFICACOES);
    const linhas = encontrarLinhasFlex(sheet, aluno ? { 'Aluno': aluno } : {});
    const notificacoes = linhas.map(function(l) {
      return {
        id: s(l.obj.ID),
        aluno: s(l.obj.Aluno),
        tipo: s(l.obj.Tipo),
        titulo: s(l.obj.Titulo),
        descricao: s(l.obj.Descricao),
        curso: s(l.obj.Curso),
        professor: s(l.obj.Professor),
        lida: s(l.obj.Lida).toUpperCase() === 'SIM',
        data: s(l.obj.Data)
      };
    });
    return { success: true, total: notificacoes.length, notificacoes: notificacoes };
  } catch (e) {
    logErro('listarNotificacoes', e);
    return { success: false, error: e.message };
  }
}

function marcarNotificacoesLidas(payload) {
  try {
    const aluno = s(payload.aluno).toUpperCase();
    if (!aluno) return { success: false, error: 'Aluno obrigatório' };
    const sheet = getSheet(ABA_NOTIFICACOES);
    const headers = getHeaders(sheet);
    const idxLida = headers.indexOf('Lida');
    const linhas = encontrarLinhasFlex(sheet, { 'Aluno': aluno });
    linhas.forEach(function(l) {
      sheet.getRange(l.linha, idxLida + 1).setValue('SIM');
    });
    return { success: true };
  } catch (e) {
    logErro('marcarNotificacoesLidas', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// MENSAGENS
// ============================================================
function listarMensagens(payload) {
  try {
    const aluno = s(payload.aluno).toUpperCase();
    const sheet = getSheet(ABA_MENSAGENS);
    const linhas = encontrarLinhasFlex(sheet, aluno ? { 'Aluno': aluno } : {});
    const mensagens = linhas.map(function(l) {
      return {
        id: s(l.obj.ID),
        aluno: s(l.obj.Aluno),
        curso: s(l.obj.Curso),
        professor: s(l.obj.Professor),
        assunto: s(l.obj.Assunto),
        mensagem: s(l.obj.Mensagem),
        lida: s(l.obj.Lida).toUpperCase() === 'SIM',
        data: s(l.obj.Data)
      };
    });
    return { success: true, total: mensagens.length, mensagens: mensagens };
  } catch (e) {
    logErro('listarMensagens', e);
    return { success: false, error: e.message };
  }
}

function enviarMensagemAluno(payload) {
  try {
    const sheet = getSheet(ABA_MENSAGENS);
    const headers = getHeaders(sheet);
    const destinatarios = payload.destinatarios || [];
    destinatarios.forEach(function(aluno) {
      const valores = headers.map(function(h) {
        switch (h) {
          case 'ID': return gerarID('MSG');
          case 'Aluno': return s(aluno).toUpperCase();
          case 'Curso': return s(payload.curso).toLowerCase();
          case 'Professor': return s(payload.professor).toUpperCase();
          case 'Assunto': return s(payload.assunto);
          case 'Mensagem': return s(payload.mensagem);
          case 'Lida': return 'NAO';
          case 'Data': return dataHoraAgora();
          default: return '';
        }
      });
      appendRegistro(sheet, valores);
    });
    return { success: true, message: 'Mensagem enviada para ' + destinatarios.length + ' aluno(s)' };
  } catch (e) {
    logErro('enviarMensagemAluno', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ARQUIVOS
// ============================================================
function listarArquivos(payload) {
  try {
    const aluno = s(payload.aluno).toUpperCase();
    const sheet = getSheet(ABA_ARQUIVOS);
    const linhas = encontrarLinhasFlex(sheet, aluno ? { 'Aluno': aluno } : {});
    const arquivos = linhas.map(function(l) {
      return {
        id: s(l.obj.ID),
        aluno: s(l.obj.Aluno),
        curso: s(l.obj.Curso),
        professor: s(l.obj.Professor),
        nome: s(l.obj.Nome_Arquivo),
        url: s(l.obj.URL),
        fileId: s(l.obj.File_ID),
        mensagem: s(l.obj.Mensagem),
        data: s(l.obj.Data)
      };
    });
    return { success: true, total: arquivos.length, arquivos: arquivos };
  } catch (e) {
    logErro('listarArquivos', e);
    return { success: false, error: e.message };
  }
}

function uploadArquivo(payload) {
  try {
    const nomeArquivo = s(payload.nomeArquivo);
    const base64 = s(payload.file);
    const mimeType = s(payload.mimeType) || 'application/octet-stream';
    const destinatarios = JSON.parse(s(payload.destinatarios) || '[]');
    if (!nomeArquivo || !base64) return { success: false, error: 'Arquivo inválido' };

    const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, nomeArquivo);
    let pasta;
    if (PASTA_DRIVE_ID) {
      try { pasta = DriveApp.getFolderById(PASTA_DRIVE_ID); } catch (e) { pasta = DriveApp.getRootFolder(); }
    } else {
      pasta = DriveApp.getRootFolder();
    }
    const file = pasta.createFile(blob);
    const url = file.getUrl();
    const fileId = file.getId();

    const sheet = getSheet(ABA_ARQUIVOS);
    const headers = getHeaders(sheet);
    destinatarios.forEach(function(aluno) {
      const valores = headers.map(function(h) {
        switch (h) {
          case 'ID': return gerarID('ARQ');
          case 'Aluno': return s(aluno).toUpperCase();
          case 'Curso': return s(payload.curso).toLowerCase();
          case 'Professor': return s(payload.professor).toUpperCase();
          case 'Nome_Arquivo': return nomeArquivo;
          case 'URL': return url;
          case 'File_ID': return fileId;
          case 'Mensagem': return s(payload.mensagem);
          case 'Data': return dataHoraAgora();
          default: return '';
        }
      });
      appendRegistro(sheet, valores);
    });

    return { success: true, url: url, fileId: fileId };
  } catch (e) {
    logErro('uploadArquivo', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// DADOS COMPLETOS DO ALUNO
// ============================================================
function dadosAlunoCompleto(payload) {
  try {
    const nome = s(payload.nome).toUpperCase();
    if (!nome) return { success: false, error: 'Nome obrigatório' };

    const alunoSheet = getSheet(ABA_ALUNOS);
    const alunoLinha = encontrarLinha(alunoSheet, { 'Nome': nome });
    if (alunoLinha < 0) return { success: false, error: 'Aluno não encontrado' };

    const alunoHeaders = getHeaders(alunoSheet);
    const alunoValores = alunoSheet.getRange(alunoLinha, 1, 1, alunoHeaders.length).getValues()[0];
    const alunoObj = {};
    alunoHeaders.forEach(function(h, i) { alunoObj[h] = alunoValores[i]; });

    const matSheet = getSheet(ABA_MATRICULAS);
    const matLinhas = encontrarLinhasFlex(matSheet, { 'Aluno': nome });
    const matriculas = matLinhas.map(function(l) {
      return {
        curso: s(l.obj.Curso),
        professor: s(l.obj.Professor),
        status: s(l.obj.Status),
        dataMatricula: s(l.obj.Data_Matricula),
        anoLetivo: s(l.obj.Ano_Letivo),
        mensalidade: Number(l.obj.Mensalidade) || 0
      };
    });

    const avSheet = getSheet(ABA_AVALIACOES);
    const avLinhas = encontrarLinhasFlex(avSheet, { 'Aluno': nome });
    const avaliacoes = avLinhas.map(function(l) {
      return {
        curso: s(l.obj.Curso),
        bimestre: Number(l.obj.Bimestre) || 1,
        nota: Number(l.obj.Nota) || 0,
        professor: s(l.obj.Professor),
        dataLancamento: s(l.obj.Data_Lancamento)
      };
    });

    const presSheet = getSheet(ABA_PRESENCAS);
    const presTodas = encontrarLinhasFlex(presSheet, {});
    const presencas = [];
    presTodas.forEach(function(l) {
      const listaPres = s(l.obj.Presentes).split(',').map(function(n) { return n.trim().toUpperCase(); });
      const listaFalt = s(l.obj.Faltantes).split(',').map(function(n) { return n.trim().toUpperCase(); });
      if (listaPres.indexOf(nome) > -1) {
        presencas.push({ curso: s(l.obj.Curso), data: s(l.obj.Data), status: 'PRESENTE', conteudo: s(l.obj.Conteudo), professor: s(l.obj.Professor) });
      } else if (listaFalt.indexOf(nome) > -1) {
        presencas.push({ curso: s(l.obj.Curso), data: s(l.obj.Data), status: 'FALTA', conteudo: s(l.obj.Conteudo), professor: s(l.obj.Professor) });
      }
    });

    const contSheet = getSheet(ABA_CONTEUDOS);
    const contLinhas = encontrarLinhasFlex(contSheet, {});
    const conteudos = contLinhas.map(function(l) {
      return { curso: s(l.obj.Curso), data: s(l.obj.Data), conteudo: s(l.obj.Conteudo), objetivos: s(l.obj.Objetivos), professor: s(l.obj.Professor) };
    });

    const relSheet = getSheet(ABA_RELATORIOS);
    const relLinhas = encontrarLinhasFlex(relSheet, { 'Aluno': nome });
    const relatorios = relLinhas.map(function(l) {
      return { curso: s(l.obj.Curso), bimestre: s(l.obj.Bimestre), professor: s(l.obj.Professor), relatorio: s(l.obj.Relatorio), data: s(l.obj.Data), visto: s(l.obj.Visto) };
    });

    const arqSheet = getSheet(ABA_ARQUIVOS);
    const arqLinhas = encontrarLinhasFlex(arqSheet, { 'Aluno': nome });
    const arquivos = arqLinhas.map(function(l) {
      return { id: s(l.obj.ID), curso: s(l.obj.Curso), professor: s(l.obj.Professor), nome: s(l.obj.Nome_Arquivo), url: s(l.obj.URL), fileId: s(l.obj.File_ID), mensagem: s(l.obj.Mensagem), data: s(l.obj.Data) };
    });

    const msgSheet = getSheet(ABA_MENSAGENS);
    const msgLinhas = encontrarLinhasFlex(msgSheet, { 'Aluno': nome });
    const mensagens = msgLinhas.map(function(l) {
      return { id: s(l.obj.ID), curso: s(l.obj.Curso), professor: s(l.obj.Professor), assunto: s(l.obj.Assunto), mensagem: s(l.obj.Mensagem), lida: s(l.obj.Lida).toUpperCase() === 'SIM', data: s(l.obj.Data) };
    });

    return {
      success: true,
      aluno: {
        nome: s(alunoObj.Nome),
        dataNascimento: s(alunoObj.Data_Nascimento),
        responsavel: s(alunoObj.Responsavel),
        contato: s(alunoObj.Contato),
        status: s(alunoObj.Status),
        mensalidade: Number(alunoObj.Mensalidade) || 0,
        vencimento: Number(alunoObj.Vencimento) || 10
      },
      matriculas: matriculas,
      avaliacoes: avaliacoes,
      presencas: presencas,
      conteudos: conteudos,
      relatorios: relatorios,
      arquivos: arquivos,
      mensagens: mensagens
    };
  } catch (e) {
    logErro('dadosAlunoCompleto', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// SINC ABAS INDIVIDUAIS (OPCIONAL)
// ============================================================
function sincronizarAbaAluno(nome) {
  const sh = SpreadsheetApp.openById(SHEET_ID);
  const nomeAba = 'ALUNO_' + s(nome).toUpperCase().replace(/[^A-Z0-9 ]/g, '').replace(/\s+/g, '_').substring(0, 30);
  let aba = sh.getSheetByName(nomeAba);
  if (!aba) {
    aba = sh.insertSheet(nomeAba);
    aba.appendRow(['Informação', 'Valor']);
    aba.appendRow(['Nome', s(nome).toUpperCase()]);
    aba.appendRow(['Criado em', dataHoraAgora()]);
  }
  return nomeAba;
}

// ============================================================
// ADMIN — LOGS E BUGS
// ============================================================
function adminListarLogs(payload) {
  try {
    const limite = Number(payload.limite) || 100;
    const sheet = getSheet(ABA_LOGS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, total: 0, logs: [] };
    const headers = getHeaders(sheet);
    const start = Math.max(2, lastRow - limite + 1);
    const valores = sheet.getRange(start, 1, lastRow - start + 1, headers.length).getValues();
    const logs = valores.map(function(row) {
      const obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return {
        data: obj.Data instanceof Date ? Utilities.formatDate(obj.Data, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm') : s(obj.Data),
        usuario: s(obj.Usuario),
        perfil: s(obj.Perfil),
        acao: s(obj.Acao),
        detalhe: s(obj.Detalhe)
      };
    }).reverse();
    return { success: true, total: logs.length, logs: logs };
  } catch (e) {
    logErro('adminListarLogs', e);
    return { success: false, error: e.message };
  }
}

function adminListarBugs(payload) {
  try {
    const sheet = getSheet(ABA_LOGS);
    const linhas = encontrarLinhasFlex(sheet, {});
    const bugs = linhas.filter(function(l) {
      return normalizarChave(l.obj.Perfil) === 'erro';
    }).map(function(l) {
      return {
        titulo: s(l.obj.Acao),
        descricao: s(l.obj.Detalhe),
        data: l.obj.Data instanceof Date ? Utilities.formatDate(l.obj.Data, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm') : s(l.obj.Data)
      };
    });
    return { success: true, total: bugs.length, bugs: bugs };
  } catch (e) {
    logErro('adminListarBugs', e);
    return { success: false, error: e.message };
  }
}
