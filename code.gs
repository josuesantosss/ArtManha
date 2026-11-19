/**
 * ============================================================
 * GAS — Sistema de Gestão Escola de Música
 * Versão: alinhada 100% aos cabeçalhos reais da planilha
 *         + LOGS FILTRADOS POR PERFIL
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
const ABA_BUGS = 'BUGS';
const ABA_CURSOS = 'CURSOS';

const ADMIN_PADRAO = {
  nome: 'ADMINISTRADOR',
  usuario: 'admin',
  senha: 'admin123',
  perfil: 'admin'
};

// ✅ PERFIS CUJAS AÇÕES SERÃO REGISTRADAS EM LOGS
const PERFIS_LOGADOS = new Set(['direcao', 'aluno', 'professor', 'coordenacao', 'admin']);

// ============================================================
// UTILITÁRIOS
// ============================================================
function s(v) { return v === null || v === undefined ? '' : String(v).trim(); }

function respostaJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function dataHoje() {
  const d = new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MM-yyyy');
}

function dataHoraAgora() {
  const d = new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss');
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
// LOGS — formato: ID_Log | Nome | Perfil | Data_Hora | Acao
// ✅ Só registra se o perfil estiver em PERFIS_LOGADOS
// ============================================================
function garantirAbaLogs() {
  const sh = SpreadsheetApp.openById(SHEET_ID);
  let aba = sh.getSheetByName(ABA_LOGS);
  if (!aba) {
    aba = sh.insertSheet(ABA_LOGS);
    aba.appendRow(['ID_Log', 'Nome', 'Perfil', 'Data_Hora', 'Acao']);
  }
  return aba;
}

/**
 * Registra log SOMENTE se o perfil estiver na lista de perfis logados.
 * Perfis como 'sistema', 'anônimo', 'desconhecido', 'secretaria',
 * 'financeiro', 'erro' NÃO geram linha em LOGS.
 * Exceto: perfil === 'erro' → SEMPRE grava (para o admin ver falhas).
 */
function registrarLog(nome, perfil, acao) {
  try {
    const p = normalizarChave(perfil);
    const permitido = (p === 'erro') || PERFIS_LOGADOS.has(p);
    if (!permitido) return; // ❌ perfil não logado
    const aba = garantirAbaLogs();
    aba.appendRow([gerarID('LOG'), s(nome) || '—', p || '—', dataHoraAgora(), s(acao)]);
  } catch (e) {
    try { Logger.log('Falha registrarLog: ' + e.message); } catch (e2) {}
  }
}

function logErro(funcao, err) {
  const msg = (err && err.message) ? err.message : String(err);
  // Erros sempre gravam em LOGS (perfil 'erro') E em BUGS
  try { registrarLog('SISTEMA', 'erro', 'ERRO: ' + funcao + ' — ' + msg); } catch (e) {}
  try { registrarBug(funcao, msg, '', 'SISTEMA'); } catch (e) {}
}

function registrarBug(titulo, descricao, contexto, usuario) {
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID);
    let aba = sh.getSheetByName(ABA_BUGS);
    if (!aba) {
      aba = sh.insertSheet(ABA_BUGS);
      aba.appendRow(['ID_Bug', 'Titulo', 'Descricao', 'Contexto', 'Usuario', 'Data', 'Status']);
    }
    aba.appendRow([
      gerarID('BUG'),
      s(titulo) || 'Bug sem título',
      s(descricao),
      s(contexto),
      s(usuario) || 'SISTEMA',
      dataHoraAgora(),
      'ABERTO'
    ]);
  } catch (e) {
    try { Logger.log('Falha registrarBug: ' + e.message); } catch (e2) {}
  }
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
// ADMIN PADRÃO
// ============================================================
function garantirUsuarioPadrao() {
  try {
    const sheet = garantirAbaUsuarios();
    const headers = getHeaders(sheet);

    const existentes = encontrarLinhasFlex(sheet, {}).filter(function(l) {
      return normalizarChave(l.obj.Usuario) === normalizarChave(ADMIN_PADRAO.usuario) &&
             normalizarChave(l.obj.Perfil) === normalizarChave(ADMIN_PADRAO.perfil);
    });

    if (existentes.length > 0) {
      const l = existentes[0];
      const saltAtual = s(l.obj.Salt);
      const hashAtual = s(l.obj.Senha_Hash);
      const hashEsperado = hashSenha(ADMIN_PADRAO.senha, saltAtual);
      if (hashAtual !== hashEsperado) {
        const novoSalt = saltAtual || gerarSalt();
        const novoHash = hashSenha(ADMIN_PADRAO.senha, novoSalt);
        const idxHash = headers.indexOf('Senha_Hash');
        const idxSalt = headers.indexOf('Salt');
        if (idxSalt >= 0) sheet.getRange(l.linha, idxSalt + 1).setValue(novoSalt);
        if (idxHash >= 0) sheet.getRange(l.linha, idxHash + 1).setValue(novoHash);
        const idxStatus = headers.indexOf('Status');
        if (idxStatus >= 0) sheet.getRange(l.linha, idxStatus + 1).setValue('ATIVO');
        return { criado: false, corrigido: true, usuario: ADMIN_PADRAO.usuario, senha: ADMIN_PADRAO.senha };
      }
      return { criado: false, usuario: ADMIN_PADRAO.usuario, senha: ADMIN_PADRAO.senha };
    }

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

function criarUsuarioPadraoAdmin() {
  const r = garantirUsuarioPadrao();
  Logger.log(JSON.stringify(r));
  return r;
}

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

function resetarTodosUsuarios() {
  const sh = SpreadsheetApp.openById(SHEET_ID);
  const aba = sh.getSheetByName(ABA_USUARIOS);
  if (aba) {
    const lastRow = aba.getLastRow();
    if (lastRow > 1) aba.getRange(2, 1, lastRow - 1, aba.getLastColumn()).clearContent();
  } else {
    garantirAbaUsuarios();
  }
  const r = garantirUsuarioPadrao();
  Logger.log('Reset concluído: ' + JSON.stringify(r));
  return r;
}

// ============================================================
// HANDLERS PRINCIPAIS
// ============================================================
function doGet(e) {
  try { garantirUsuarioPadrao(); } catch (err) {}
  try { garantirAbaLogs(); } catch (err) {}
  // NÃO registra log de health-check (perfil 'sistema' é ignorado)
  return respostaJSON({ success: true, message: 'GAS Escola de Música ativo', timestamp: new Date().toISOString() });
}

function doPost(e) {
  const inicio = Date.now();
  let acao = 'desconhecida';
  let payload = {};
  try {
    const dados = JSON.parse(e.postData.contents);
    acao = dados.acao || 'desconhecida';
    payload = dados.payload || {};
    const resultado = rotear(acao, payload);
    try { registrarLogAcao(acao, payload, resultado, Date.now() - inicio); } catch (err) {}
    return respostaJSON(resultado);
  } catch (err) {
    try { logErro('doPost:' + acao, err); } catch (err2) {}
    return respostaJSON({ success: false, error: err.message });
  }
}

/**
 * Registra log da ação SOMENTE se o perfil do usuário for um dos logados.
 * Para "login", o perfil do usuário ainda não foi resolvido no payload,
 * mas o handler `loginUsuario` já registra o log específico de sucesso/falha.
 * Então, aqui, só registramos se `payload.perfil` for um perfil logado
 * (o que acontece após login, nas ações seguintes).
 */
function registrarLogAcao(acao, payload, resultado, duracaoMs) {
  try {
    // Para 'login', pulamos: o loginUsuario já registra o log certo.
    if (acao === 'login') return;

    const perfil = normalizarChave(payload && payload._logPerfil || payload && payload.perfil || '');
    if (!PERFIS_LOGADOS.has(perfil)) return; // ❌ perfil não logado

    const nome = s(payload && (payload._logNome || payload.nome || payload.usuario || payload.professor)) || '—';
    let status = (resultado && resultado.success === false) ? 'FALHA' : 'OK';
    let detalhe = '';
    if (resultado && resultado.error) detalhe = ' | Erro: ' + resultado.error;
    else if (resultado && resultado.message) detalhe = ' | ' + resultado.message;
    const tempo = duracaoMs !== undefined ? ' | ' + duracaoMs + 'ms' : '';
    registrarLog(nome, perfil, acao.toUpperCase() + ': ' + extrairResumoPayload(acao, payload) + ' ' + status + detalhe + tempo);
  } catch (e) {}
}

function extrairResumoPayload(acao, payload) {
  if (!payload) return '';
  const p = payload;
  const partes = [];
  switch (acao) {
    case 'login': partes.push('usuario=' + s(p.usuario)); break;
    case 'criarUsuario':
    case 'criarUsuarioAluno':
      partes.push('usuario=' + s(p.usuario) + ' perfil=' + s(p.perfil || 'aluno') + ' nome=' + s(p.nome));
      break;
    case 'atualizarSenha': partes.push('usuario=' + s(p.usuario) + ' perfil=' + s(p.perfil || '-')); break;
    case 'salvarAluno': partes.push('aluno=' + s(p.nome)); break;
    case 'atualizarAluno': partes.push('aluno=' + s(p.nome) + ' mensalidade=' + s(p.mensalidade)); break;
    case 'salvarMatricula': partes.push('aluno=' + s(p.aluno) + ' curso=' + s(p.curso) + ' prof=' + s(p.professor)); break;
    case 'atualizarMatricula': partes.push('aluno=' + s(p.aluno) + ' curso=' + s(p.curso) + ' status=' + s(p.status)); break;
    case 'salvarProfessor': partes.push('prof=' + s(p.nome) + ' curso=' + s(p.curso)); break;
    case 'atualizarProfessor': partes.push('id=' + s(p.id) + ' curso=' + s(p.curso)); break;
    case 'salvarValorProfessor': partes.push('id=' + s(p.id) + ' valorHora=' + s(p.valorHora)); break;
    case 'salvarAvaliacoes': {
      const avs = p.avaliacoes || [];
      partes.push('qtd=' + avs.length);
      if (avs.length > 0) partes.push('curso=' + s(avs[0].curso) + ' bim=' + s(avs[0].bimestre) + ' prof=' + s(avs[0].professor));
      break;
    }
    case 'salvarPresenca': partes.push('curso=' + s(p.curso) + ' data=' + s(p.data) + ' presentes=' + ((p.presentes||[]).length) + ' prof=' + s(p.professor)); break;
    case 'salvarObservacao': partes.push('aluno=' + s(p.aluno) + ' curso=' + s(p.curso) + ' tipo=' + s(p.tipo)); break;
    case 'salvarRelatorio': partes.push('aluno=' + s(p.aluno) + ' curso=' + s(p.curso) + ' bim=' + s(p.bimestre) + ' prof=' + s(p.professor)); break;
    case 'salvarPagamento': partes.push('aluno=' + s(p.aluno) + ' ref=' + s(p.referencia) + ' valor=' + s(p.valor) + ' status=' + s(p.status)); break;
    case 'notificarAtraso': partes.push('aluno=' + s(p.aluno) + ' ref=' + s(p.referencia)); break;
    case 'notificarPagamento': partes.push('aluno=' + s(p.aluno) + ' ref=' + s(p.referencia) + ' valor=' + s(p.valor)); break;
    case 'enviarMensagemAluno': partes.push('curso=' + s(p.curso) + ' dest=' + ((p.destinatarios||[]).length) + ' assunto=' + s(p.assunto)); break;
    case 'uploadArquivo': partes.push('curso=' + s(p.curso) + ' prof=' + s(p.professor) + ' arquivo=' + s(p.nomeArquivo) + ' dest=' + s(p.destinatarios).substring(0, 80)); break;
    case 'marcarNotificacoesLidas': partes.push('aluno=' + s(p.aluno)); break;
    case 'dadosAlunoCompleto': partes.push('aluno=' + s(p.nome)); break;
    case 'listarUsuarios':
    case 'listarPagamentos':
    case 'adminListarLogs':
    case 'adminListarBugs':
      partes.push('consulta');
      break;
    default:
      partes.push('payload_keys=' + Object.keys(p).slice(0, 6).join(','));
  }
  return partes.join(' ');
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
    case 'listarCursos': return listarCursos(payload);
    case 'salvarCurso': return salvarCurso(payload);
    case 'carregarContatosChat': return ChatMod.carregarContatosChat(payload);
    case 'carregarMensagensChat': return ChatMod.carregarMensagensChat(payload);
    case 'enviarMensagemChat': return ChatMod.enviarMensagemChat(payload);
    case 'garantirAdminPadrao': {
      try { return { success: true, ...garantirUsuarioPadrao() }; }
      catch (err) { return { success: false, error: err.message }; }
    }
    default: return { success: false, error: 'Ação desconhecida: ' + acao };
  }
}
// ============================================================
// LOGIN — registra log apenas para perfis logados
// ============================================================
function loginUsuario(payload) {
  try {
    const usuario = s(payload.usuario).toLowerCase();
    const senha = s(payload.senha);
    const perfilSolicitado = s(payload.perfil).toLowerCase();

    if (!usuario || !senha) {
      // ❌ não loga: perfil 'desconhecido' não está em PERFIS_LOGADOS
      return { success: false, error: 'Usuário e senha são obrigatórios' };
    }

    try { garantirUsuarioPadrao(); } catch (err) {}

    const sheet = garantirAbaUsuarios();
    const headers = getHeaders(sheet);
    const idxHash = headers.indexOf('Senha_Hash');
    const idxSalt = headers.indexOf('Salt');

    const linhas = encontrarLinhasFlex(sheet, {}).filter(function(l) {
      return normalizarChave(l.obj.Usuario) === normalizarChave(usuario);
    });
    if (linhas.length === 0) {
      // ❌ usuário não encontrado → perfil desconhecido → não loga
      return { success: false, error: 'Usuário ou senha inválidos' };
    }

    const validos = [];
    linhas.forEach(function(l) {
      if (s(l.obj.Status).toUpperCase() !== 'ATIVO') return;
      const salt = s(l.obj.Salt);
      const hashArmazenado = s(l.obj.Senha_Hash);
      if (!hashArmazenado) return;

      const hashCalculado = hashSenha(senha, salt);
      if (hashCalculado === hashArmazenado) { validos.push(l); return; }

      // migração de senha em texto puro
      if (hashArmazenado === senha) {
        try {
          const novoSalt = salt || gerarSalt();
          const novoHash = hashSenha(senha, novoSalt);
          if (idxSalt >= 0) sheet.getRange(l.linha, idxSalt + 1).setValue(novoSalt);
          if (idxHash >= 0) sheet.getRange(l.linha, idxHash + 1).setValue(novoHash);
        } catch (err) {}
        validos.push(l);
        return;
      }

      // migração de hash sem salt
      if (!salt) {
        const hashPuro = hashSenha(senha, '');
        if (hashPuro === hashArmazenado) {
          try {
            const novoSalt = gerarSalt();
            const novoHash = hashSenha(senha, novoSalt);
            if (idxSalt >= 0) sheet.getRange(l.linha, idxSalt + 1).setValue(novoSalt);
            if (idxHash >= 0) sheet.getRange(l.linha, idxHash + 1).setValue(novoHash);
          } catch (err) {}
          validos.push(l);
          return;
        }
      }
    });

    if (validos.length === 0) {
      // ❌ senha incorreta → não loga (perfil ainda não confirmado)
      return { success: false, error: 'Usuário ou senha inválidos' };
    }

    // perfil exato → LOGA
    const doPerfil = validos.find(function(l) {
      return normalizarChave(l.obj.Perfil) === normalizarChave(perfilSolicitado);
    });
    if (doPerfil) {
      const perfilReal = normalizarChave(doPerfil.obj.Perfil);
      if (PERFIS_LOGADOS.has(perfilReal)) {
        registrarLog(s(doPerfil.obj.Nome), perfilReal, 'LOGIN: OK usuario=' + usuario + ' perfil=' + perfilReal);
      }
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

    // admin override → LOGA
    const adminValido = validos.find(function(l) {
      return normalizarChave(l.obj.Perfil) === 'admin';
    });
    if (adminValido) {
      registrarLog(s(adminValido.obj.Nome), 'admin', 'LOGIN: OK (admin override) → perfil solicitado=' + (perfilSolicitado || 'admin'));
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

    const perfisDisponiveis = validos.map(function(l) { return s(l.obj.Perfil); });
    // ❌ perfil não bate → não loga
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

  if (!nome || !perfil || !usuario || !senha) return { success: false, error: 'Nome, perfil, usuário e senha são obrigatórios' };
  if (senha.length < 4) return { success: false, error: 'Senha deve ter ao menos 4 caracteres' };

  const sheet = garantirAbaUsuarios();
  const headers = getHeaders(sheet);

  if (linhaUsuarioPerfil(sheet, usuario, perfil) > -1) {
    if (PERFIS_LOGADOS.has(perfil)) {
      registrarLog(nome, perfil, 'CRIAR_USUARIO: FALHA — já existe usuario=' + usuario + ' perfil=' + perfil);
    }
    return { success: false, error: 'Já existe um usuário "' + usuario + '" com perfil "' + perfil + '". Escolha outro perfil ou outro login.' };
  }

  const outros = [];
  encontrarLinhasFlex(sheet, {}).forEach(function(l) {
    if (normalizarChave(l.obj.Usuario) === normalizarChave(usuario)) outros.push(s(l.obj.Perfil));
  });

  const salt = gerarSalt();
  const hash = hashSenha(senha, salt);
  const idUser = gerarID('USR');

  const valores = headers.map(function(h) {
    switch (h) {
      case 'ID_Usuario': return idUser;
      case 'Nome': return s(nome).toUpperCase();
      case 'Perfil': return perfil;
      case 'Usuario': return usuario;
      case 'Senha_Hash': return hash;
      case 'Salt': return salt;
      case 'Status': return 'ATIVO';
      case 'Data_Criacao': return dataHoje();
      case 'Aluno_Vinculado': return s(alunoVinculado).toUpperCase();
      default: return '';
    }
  });
  appendRegistro(sheet, valores);

  if (PERFIS_LOGADOS.has(perfil)) {
    registrarLog(nome, perfil, 'CRIAR_USUARIO: OK usuario=' + usuario + ' perfil=' + perfil + (outros.length ? ' outrosPerfis=' + outros.join(',') : ''));
  }

  return { success: true, message: 'Usuário criado: ' + usuario + ' (' + perfil + ')', id: idUser, usuario: usuario, perfil: perfil, outrosPerfis: outros };
}

function criarUsuarioAluno(payload) {
  const nome = s(payload.nome);
  const usuario = s(payload.usuario).toLowerCase();
  const senha = s(payload.senha);

  if (!nome || !usuario || !senha) return { success: false, error: 'Nome, usuário e senha são obrigatórios' };
  if (senha.length < 4) return { success: false, error: 'Senha deve ter ao menos 4 caracteres' };

  const alunoSheet = getSheet(ABA_ALUNOS);
  const alunoLinha = encontrarLinha(alunoSheet, { 'Nome': s(nome).toUpperCase() });
  if (alunoLinha < 0) return { success: false, error: 'Aluno não encontrado na aba ALUNOS: ' + nome };

  const sheet = garantirAbaUsuarios();
  const headers = getHeaders(sheet);

  if (linhaUsuarioPerfil(sheet, usuario, 'aluno') > -1) {
    registrarLog(nome, 'aluno', 'CRIAR_USUARIO_ALUNO: FALHA — login aluno já existe: ' + usuario);
    return { success: false, error: 'Já existe um aluno com o login "' + usuario + '". Escolha outro.' };
  }

  const salt = gerarSalt();
  const hash = hashSenha(senha, salt);
  const idUser = gerarID('USR');

  const valores = headers.map(function(h) {
    switch (h) {
      case 'ID_Usuario': return idUser;
      case 'Nome': return s(nome).toUpperCase();
      case 'Perfil': return 'aluno';
      case 'Usuario': return usuario;
      case 'Senha_Hash': return hash;
      case 'Salt': return salt;
      case 'Status': return 'ATIVO';
      case 'Data_Criacao': return dataHoje();
      case 'Aluno_Vinculado': return s(nome).toUpperCase();
      default: return '';
    }
  });
  appendRegistro(sheet, valores);

  let abaAluno = '';
  try { abaAluno = sincronizarAbaAluno(nome); } catch (e) {}

  const outros = [];
  encontrarLinhasFlex(sheet, {}).forEach(function(l) {
    if (normalizarChave(l.obj.Usuario) === normalizarChave(usuario) && normalizarChave(l.obj.Perfil) !== 'aluno') outros.push(s(l.obj.Perfil));
  });

  registrarLog(nome, 'aluno', 'CRIAR_USUARIO_ALUNO: OK usuario=' + usuario + (outros.length ? ' outrosPerfis=' + outros.join(',') : ''));

  return { success: true, message: 'Usuário de aluno criado: ' + usuario, id: idUser, usuario: usuario, abaAluno: abaAluno, outrosPerfis: outros };
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
    return { success: true, total: registros.length, usuarios: registros, perfis: registros.map(function(r) { return r.perfil; }) };
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

  if (!usuario || !senhaNova) return { success: false, error: 'Usuário e nova senha são obrigatórios' };
  if (senhaNova.length < 4) return { success: false, error: 'Senha deve ter ao menos 4 caracteres' };

  const sheet = garantirAbaUsuarios();
  const headers = getHeaders(sheet);
  const linhasTodas = encontrarLinhasFlex(sheet, {}).filter(function(l) {
    return normalizarChave(l.obj.Usuario) === normalizarChave(usuario);
  });
  if (linhasTodas.length === 0) return { success: false, error: 'Usuário não encontrado: ' + usuario };

  if (!perfil && linhasTodas.length > 1) {
    return { success: false, error: 'Este login existe em ' + linhasTodas.length + ' perfis. Especifique o perfil.', perfis: linhasTodas.map(function(l) { return s(l.obj.Perfil); }) };
  }

  const alvo = perfil ? linhasTodas.filter(function(l) { return normalizarChave(l.obj.Perfil) === normalizarChave(perfil); }) : linhasTodas;
  if (alvo.length === 0) return { success: false, error: 'Nenhum registro para o perfil "' + perfil + '"' };

  const idxHash = headers.indexOf('Senha_Hash');
  const idxSalt = headers.indexOf('Salt');
  alvo.forEach(function(l) {
    const novoSalt = gerarSalt();
    const novoHash = hashSenha(senhaNova, novoSalt);
    if (idxSalt >= 0) sheet.getRange(l.linha, idxSalt + 1).setValue(novoSalt);
    if (idxHash >= 0) sheet.getRange(l.linha, idxHash + 1).setValue(novoHash);
  });

  const perfisAlvo = alvo.map(function(l) { return normalizarChave(l.obj.Perfil); });
  // Só loga se algum perfil alterado for um perfil logado
  if (perfisAlvo.some(function(p) { return PERFIS_LOGADOS.has(p); })) {
    registrarLog('ADMIN', 'admin', 'ATUALIZAR_SENHA: usuario=' + usuario + ' perfis=' + perfisAlvo.join(','));
  }

  return { success: true, message: 'Senha redefinida para ' + alvo.length + ' registro(s)', perfis: alvo.map(function(l) { return s(l.obj.Perfil); }) };
}

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
      alvo = linhas.filter(function(l) { return normalizarChave(l.obj.Perfil) === normalizarChave(perfil); });
      if (alvo.length === 0) return { success: false, error: 'Perfil não encontrado para este usuário' };
    }
    if (alvo.length > 1) return { success: false, error: 'Múltiplos perfis. Especifique o perfil.', perfis: alvo.map(function(l) { return s(l.obj.Perfil); }) };

    registrarLog('ADMIN', 'admin', 'ADMIN_VER_SENHA: usuario=' + usuario + ' perfil=' + s(alvo[0].obj.Perfil));

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
        case 'Data_Cadastro': return linhaExistente > 0 ? '' : dataHoje();
        case 'Mensalidade': return Number(payload.mensalidade) || 0;
        case 'Vencimento': return Number(payload.vencimento) || 10;
        default: return '';
      }
    });

    if (linhaExistente > 0) {
      const headersAntigos = getHeaders(sheet);
      const valoresAntigos = sheet.getRange(linhaExistente, 1, 1, headersAntigos.length).getValues()[0];
      const idxData = headers.indexOf('Data_Cadastro');
      if (idxData >= 0) valores[idxData] = valoresAntigos[idxData] || dataHoje();
      sheet.getRange(linhaExistente, 1, 1, headers.length).setValues([valores]);
      registrarLog(nome, 'secretaria', 'SALVAR_ALUNO: ATUALIZADO');
      return { success: true, message: 'Aluno atualizado: ' + nome };
    }
    appendRegistro(sheet, valores);
    registrarLog(nome, 'secretaria', 'SALVAR_ALUNO: CRIADO mensalidade=' + (Number(payload.mensalidade) || 0));
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
    if (idxMens >= 0 && payload.mensalidade !== undefined) sheet.getRange(linha, idxMens + 1).setValue(Number(payload.mensalidade) || 0);
    if (idxVenc >= 0 && payload.vencimento !== undefined) sheet.getRange(linha, idxVenc + 1).setValue(Number(payload.vencimento) || 10);

    registrarLog(nome, 'secretaria', 'ATUALIZAR_ALUNO: mensalidade=' + s(payload.mensalidade) + ' venc=' + s(payload.vencimento));
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
        case 'Ano_Letivo': return s(payload.anoLetivo) || String(new Date().getFullYear());
        case 'Status': return s(payload.status || 'ATIVO').toUpperCase();
        case 'Observacoes': return s(payload.observacoes);
        case 'Mensalidade': return Number(payload.mensalidade) || 0;
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    registrarLog(aluno, 'secretaria', 'SALVAR_MATRICULA: curso=' + curso + ' prof=' + s(payload.professor));
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
    const idxObs = headers.indexOf('Observacoes');
    linhas.forEach(function(l) {
      if (idxStatus >= 0) sheet.getRange(l.linha, idxStatus + 1).setValue(s(payload.status || 'ATIVO').toUpperCase());
      if (idxObs >= 0 && payload.observacoes !== undefined) sheet.getRange(l.linha, idxObs + 1).setValue(s(payload.observacoes));
    });
    registrarLog(aluno, 'secretaria', 'ATUALIZAR_MATRICULA: curso=' + curso + ' status=' + s(payload.status));
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
        case 'Status': return s(payload.status || 'ATIVO').toUpperCase();
        case 'Data_Cadastro': return dataHoje();
        case 'Valor_Hora': return Number(payload.valorHora) || 0;
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    registrarLog(nome, 'coordenacao', 'SALVAR_PROFESSOR: curso=' + s(payload.curso) + ' tipo=' + s(payload.tipo));
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
    registrarLog('COORDENAÇÃO', 'coordenacao', 'ATUALIZAR_PROFESSOR: id=' + id + ' curso=' + s(payload.curso));
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
    registrarLog('FINANCEIRO', 'financeiro', 'SALVAR_VALOR_PROFESSOR: id=' + id + ' valorHora=' + (Number(payload.valorHora) || 0));
    return { success: true, message: 'Valor atualizado' };
  } catch (e) {
    logErro('salvarValorProfessor', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// AVALIAÇÕES
// ============================================================
function salvarAvaliacoes(payload) {
  try {
    const avals = payload.avaliacoes || [];
    const sheet = getSheet(ABA_AVALIACOES);
    const headers = getHeaders(sheet);
    avals.forEach(function(av) {
      const valores = headers.map(function(h) {
        switch (h) {
          case 'ID_Avaliacao': return gerarID('AVA');
          case 'Aluno': return s(av.aluno).toUpperCase();
          case 'Curso': return s(av.curso).toLowerCase();
          case 'Bimestre': return Number(av.bimestre) || 1;
          case 'Nota': return Number(av.nota) || 0;
          case 'Data_Lancamento': return dataHoje();
          case 'Professor': return s(av.professor).toUpperCase();
          default: return '';
        }
      });
      appendRegistro(sheet, valores);
      registrarLog(s(av.professor) || 'PROFESSOR', 'professor', 'SALVAR_NOTA: aluno=' + s(av.aluno) + ' curso=' + s(av.curso) + ' bim=' + (Number(av.bimestre) || 1) + ' nota=' + (Number(av.nota) || 0));
    });
    return { success: true, message: 'Avaliações salvas' };
  } catch (e) {
    logErro('salvarAvaliacoes', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// PRESENÇAS
// ============================================================
function salvarPresenca(payload) {
  try {
    const sheet = getSheet(ABA_PRESENCAS);
    const headers = getHeaders(sheet);
    const presentes = payload.presentes || [];
    const faltantes = payload.faltantes || [];
    const valores = headers.map(function(h) {
      switch (h) {
        case 'ID_Presenca': return gerarID('PRES');
        case 'Curso': return s(payload.curso).toLowerCase();
        case 'Data': return s(payload.data);
        case 'Presentes': return presentes.join(', ');
        case 'Faltantes': return faltantes.join(', ');
        case 'Observacao': return s(payload.observacao || payload.conteudo);
        case 'Professor': return s(payload.professor).toUpperCase();
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    registrarLog(s(payload.professor) || 'PROFESSOR', 'professor', 'SALVAR_PRESENCA: curso=' + s(payload.curso) + ' data=' + s(payload.data) + ' presentes=' + presentes.length);
    return { success: true, message: 'Presença salva' };
  } catch (e) {
    logErro('salvarPresenca', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// OBSERVAÇÕES
// ============================================================
function salvarObservacao(payload) {
  try {
    const sheet = getSheet(ABA_OBSERVACOES);
    const headers = getHeaders(sheet);
    const valores = headers.map(function(h) {
      switch (h) {
        case 'ID_Observacao': return gerarID('OBS');
        case 'Curso': return s(payload.curso).toLowerCase();
        case 'Data': return s(payload.data);
        case 'Aluno': return s(payload.aluno).toUpperCase();
        case 'Observacao': return s(payload.observacao);
        case 'Tipo': return s(payload.tipo || 'Neutra');
        case 'Professor': return s(payload.professor).toUpperCase();
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    registrarLog(s(payload.professor) || 'PROFESSOR', 'professor', 'SALVAR_OBSERVACAO: aluno=' + s(payload.aluno) + ' curso=' + s(payload.curso) + ' tipo=' + s(payload.tipo));
    return { success: true, message: 'Observação salva' };
  } catch (e) {
    logErro('salvarObservacao', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// RELATÓRIOS
// ============================================================
function salvarRelatorio(payload) {
  try {
    const sheet = getSheet(ABA_RELATORIOS);
    const headers = getHeaders(sheet);
    const valores = headers.map(function(h) {
      switch (h) {
        case 'ID_Relatorio': return gerarID('REL');
        case 'Aluno': return s(payload.aluno).toUpperCase();
        case 'Curso': return s(payload.curso).toLowerCase();
        case 'Bimestre': return s(payload.bimestre);
        case 'Professor': return s(payload.professor).toUpperCase();
        case 'Relatorio': return s(payload.relatorio);
        case 'Data': return s(payload.data) || dataHoje();
        case 'Visto': return '';
        case 'Data_Visto': return '';
        case 'Observacoes_Coordenacao': return '';
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    registrarLog(s(payload.professor) || 'PROFESSOR', 'professor', 'SALVAR_RELATORIO: aluno=' + s(payload.aluno) + ' curso=' + s(payload.curso) + ' bim=' + s(payload.bimestre));
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
      registrarLog('FINANCEIRO', 'financeiro', 'SALVAR_PAGAMENTO: ATUALIZADO aluno=' + aluno + ' ref=' + ref + ' status=' + s(payload.status));
      return { success: true, message: 'Pagamento atualizado' };
    }
    appendRegistro(sheet, valores);
    registrarLog('FINANCEIRO', 'financeiro', 'SALVAR_PAGAMENTO: CRIADO aluno=' + aluno + ' ref=' + ref + ' valor=' + (Number(payload.valor) || 0));
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
    // Não loga: perfil 'financeiro' não está em PERFIS_LOGADOS
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
    // Não loga: perfil 'financeiro' não está em PERFIS_LOGADOS
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
    registrarLog(aluno, 'aluno', 'MARCAR_NOTIFICACOES_LIDAS: total=' + linhas.length);
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
    registrarLog(s(payload.professor) || 'PROFESSOR', 'professor', 'ENVIAR_MENSAGEM: curso=' + s(payload.curso) + ' dest=' + destinatarios.length + ' assunto=' + s(payload.assunto));
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

    registrarLog(s(payload.professor) || 'PROFESSOR', 'professor', 'UPLOAD_ARQUIVO: curso=' + s(payload.curso) + ' arquivo=' + nomeArquivo + ' dest=' + destinatarios.length);

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
        presencas.push({ curso: s(l.obj.Curso), data: s(l.obj.Data), status: 'PRESENTE', conteudo: s(l.obj.Observacao), professor: s(l.obj.Professor) });
      } else if (listaFalt.indexOf(nome) > -1) {
        presencas.push({ curso: s(l.obj.Curso), data: s(l.obj.Data), status: 'FALTA', conteudo: s(l.obj.Observacao), professor: s(l.obj.Professor) });
      }
    });

    const contSheet = sh_getSheetSafe(ABA_CONTEUDOS);
    let conteudos = [];
    if (contSheet) {
      const contLinhas = encontrarLinhasFlex(contSheet, {});
      conteudos = contLinhas.map(function(l) {
        return { curso: s(l.obj.Curso), data: s(l.obj.Data), conteudo: s(l.obj.Conteudo), objetivos: s(l.obj.Objetivos), professor: s(l.obj.Professor) };
      });
    }

    const relSheet = getSheet(ABA_RELATORIOS);
    const relLinhas = encontrarLinhasFlex(relSheet, { 'Aluno': nome });
    const relatorios = relLinhas.map(function(l) {
      return { curso: s(l.obj.Curso), bimestre: s(l.obj.Bimestre), professor: s(l.obj.Professor), relatorio: s(l.obj.Relatorio), data: s(l.obj.Data), visto: s(l.obj.Visto), observacoesCoordenacao: s(l.obj.Observacoes_Coordenacao) };
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

    registrarLog(nome, 'aluno', 'DADOS_ALUNO_COMPLETO: consulta');

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

function sh_getSheetSafe(nome) {
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID);
    return sh.getSheetByName(nome);
  } catch (e) { return null; }
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
// CURSOS
// ============================================================
function listarCursos(payload) {
  try {
    const sheet = sh_getSheetSafe(ABA_CURSOS);
    if (!sheet) return { success: true, total: 0, cursos: [] };
    const linhas = encontrarLinhasFlex(sheet, {});
    const cursos = linhas.map(function(l) {
      return {
        id: s(l.obj.ID),
        nome: s(l.obj.Nome),
        tipo: s(l.obj.Tipo),
        ativo: s(l.obj.Ativo).toUpperCase() === 'SIM' || s(l.obj.Ativo).toUpperCase() === 'TRUE'
      };
    });
    return { success: true, total: cursos.length, cursos: cursos };
  } catch (e) {
    logErro('listarCursos', e);
    return { success: false, error: e.message };
  }
}

function salvarCurso(payload) {
  try {
    const sheet = sh_getSheetSafe(ABA_CURSOS);
    if (!sheet) return { success: false, error: 'Aba CURSOS não encontrada' };
    const headers = getHeaders(sheet);
    const valores = headers.map(function(h) {
      switch (h) {
        case 'ID': return s(payload.id) || gerarID('CUR');
        case 'Nome': return s(payload.nome);
        case 'Tipo': return s(payload.tipo);
        case 'Ativo': return payload.ativo === false ? 'NAO' : 'SIM';
        default: return '';
      }
    });
    appendRegistro(sheet, valores);
    registrarLog('COORDENAÇÃO', 'coordenacao', 'SALVAR_CURSO: nome=' + s(payload.nome) + ' tipo=' + s(payload.tipo));
    return { success: true, message: 'Curso salvo' };
  } catch (e) {
    logErro('salvarCurso', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ADMIN — LOGS E BUGS
// ============================================================
function adminListarLogs(payload) {
  try {
    const limite = Number(payload.limite) || 200;
    const sheet = garantirAbaLogs();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, total: 0, logs: [] };
    const headers = getHeaders(sheet);
    const start = Math.max(2, lastRow - limite + 1);
    const valores = sheet.getRange(start, 1, lastRow - start + 1, headers.length).getValues();
    const logs = valores.map(function(row) {
      const obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return {
        id: s(obj.ID_Log),
        data: s(obj.Data_Hora),
        usuario: s(obj.Nome),
        perfil: s(obj.Perfil),
        acao: s(obj.Acao),
        detalhe: s(obj.Acao)
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
    const sheet = sh_getSheetSafe(ABA_BUGS);
    if (!sheet) return { success: true, total: 0, bugs: [] };
    const linhas = encontrarLinhasFlex(sheet, {});
    const bugs = linhas.map(function(l) {
      return {
        id: s(l.obj.ID_Bug),
        titulo: s(l.obj.Titulo),
        descricao: s(l.obj.Descricao),
        contexto: s(l.obj.Contexto),
        usuario: s(l.obj.Usuario),
        data: s(l.obj.Data),
        status: s(l.obj.Status)
      };
    }).reverse();
    return { success: true, total: bugs.length, bugs: bugs };
  } catch (e) {
    logErro('adminListarBugs', e);
    return { success: false, error: e.message };
  }
}

// =========================================================
// CHAT INTERNO - BACKEND (ÚNICO E CORRIGIDO)
// =========================================================
var ChatMod = (function () {
  var SHEET_NAME = 'Chat';
  var HEADERS = ['ID','Timestamp','DeUsuario','DePerfil','DeNome',
                 'ParaUsuario','ParaPerfil','ParaNome','Mensagem','Lida','LidaEm'];

  function getSpreadsheet() {
    return SpreadsheetApp.openById(SHEET_ID);
  }

  function getChatSheet() {
    var ss = getSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheetByName('CHAT');
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(HEADERS);
      sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setColumnWidth(9, 400);
    } else if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  }

  function getUsuariosSheet() {
    var ss = getSpreadsheet();
    var abas = ss.getSheets();
    for (var i = 0; i < abas.length; i++) {
      var nomeNorm = normalizarChave(abas[i].getName());
      if (nomeNorm === 'usuarios' || nomeNorm === 'usuario' || nomeNorm === 'logins') {
        return abas[i];
      }
    }
    return garantirAbaUsuarios();
  }

  function normPerfil(p) {
    return String(p == null ? '' : p).trim().toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  }

  function isLida(v) {
    if (v === true) return true;
    if (v === false) return false;
    var str = String(v == null ? '' : v).trim().toLowerCase();
    return str === 'true' || str === 'sim' || str === '1' || str === 'yes';
  }

  function podeConversar(pa, pb) {
    var a = normPerfil(pa), b = normPerfil(pb);
    if (a === 'aluno' && b === 'aluno') return false;
    return true;
  }

  function getUsuariosChat() {
    var sh = getUsuariosSheet();
    if (!sh) return [];
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return [];
    var headers = data[0].map(function(h){ return normalizarChave(h); });

    function idx(names) {
      for (var k = 0; k < names.length; k++) {
        var i = headers.indexOf(names[k]);
        if (i >= 0) return i;
      }
      return -1;
    }

    var iUsuario = idx(['usuario','login','user','username']);
    var iNome    = idx(['nome','nome completo','name','nome_usuario']);
    var iPerfil  = idx(['perfil','role','tipo','cargo']);
    var iStatus  = idx(['status','ativo','situacao']);
    if (iUsuario < 0 || iPerfil < 0) return [];

    var prioridade = { admin:10, direcao:9, coordenacao:8, secretaria:7,
                       financeiro:6, professor:5, aluno:1 };
    var porUsuario = {};

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var usuario = String(row[iUsuario] == null ? '' : row[iUsuario]).trim();
      if (!usuario) continue;

      if (iStatus >= 0) {
        var st = normalizarChave(row[iStatus]);
        if (st && ['inativo','inactive','nao','0','false','bloqueado','desativado','cancelado'].indexOf(st) >= 0) continue;
      }

      var perfil = iPerfil >= 0 ? normPerfil(row[iPerfil]) : '';
      var nome   = iNome   >= 0 ? String(row[iNome] || '').trim() || usuario : usuario;
      var key    = usuario.toLowerCase();
      var peso   = prioridade[perfil] || 0;

      if (!porUsuario[key] || peso > (porUsuario[key]._peso || 0)) {
        porUsuario[key] = { usuario: usuario.toLowerCase(), nome: nome, perfil: perfil, _peso: peso };
      }
    }

    var out = [];
    for (var k in porUsuario) {
      var u = porUsuario[k];
      delete u._peso;
      out.push(u);
    }
    return out;
  }

  function carregarContatosChat(payload) {
    try {
      var p = payload || {};
      var usuario = String(p.usuario || '').trim();
      var perfil  = String(p.perfil  || '').trim();
      if (!usuario) return { success: false, ok: false, erro: 'Usuário não informado.' };

      var todos = getUsuariosChat();
      var contatos = todos.filter(function(u) {
        if (u.usuario.toLowerCase() === usuario.toLowerCase()) return false;
        return podeConversar(perfil, u.perfil);
      });

      var sh = getChatSheet();
      var data = sh.getDataRange().getValues();
      var uLow = usuario.toLowerCase();
      var unread = {}, total = 0;

      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        var paraU = String(row[5] == null ? '' : row[5]).toLowerCase();
        if (paraU === uLow && !isLida(row[9])) {
          var deU = String(row[2] == null ? '' : row[2]).toLowerCase();
          unread[deU] = (unread[deU] || 0) + 1;
          total++;
        }
      }

      contatos = contatos.map(function(c){
        return {
          usuario: c.usuario,
          nome: c.nome,
          perfil: c.perfil,
          naoLidas: unread[c.usuario.toLowerCase()] || 0
        };
      }).sort(function(a,b){
        return String(a.nome).localeCompare(String(b.nome), 'pt-BR', {sensitivity:'base'});
      });

      return { success: true, ok: true, contatos: contatos, totalNaoLidas: total };
    } catch (e) {
      return { success: false, ok: false, erro: String(e && e.message ? e.message : e) };
    }
  }

  function carregarMensagensChat(payload) {
    try {
      var p = payload || {};
      var usuario        = String(p.usuario        || '').trim();
      var perfil         = String(p.perfil         || '').trim();
      var contatoUsuario = String(p.contatoUsuario || '').trim();
      var contatoPerfil  = String(p.contatoPerfil  || '').trim();
      var marcarLidas    = !!p.marcarLidas;

      if (!usuario || !contatoUsuario) return { success: false, ok: false, erro: 'Parâmetros inválidos.' };
      if (!podeConversar(perfil, contatoPerfil)) {
        return { success: false, ok: false, erro: 'Conversa não permitida entre estes perfis.' };
      }

      var sh = getChatSheet();
      var data = sh.getDataRange().getValues();
      var u1 = usuario.toLowerCase(), u2 = contatoUsuario.toLowerCase();

      var mensagens = [], updates = [];

      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        var deU   = String(row[2] == null ? '' : row[2]).toLowerCase();
        var paraU = String(row[5] == null ? '' : row[5]).toLowerCase();

        var matchA = (deU === u1 && paraU === u2);
        var matchB = (deU === u2 && paraU === u1);
        if (!matchA && !matchB) continue;

        var ts = row[1];
        var tsISO = (ts instanceof Date) ? ts.toISOString() : String(ts || '');

        mensagens.push({
          id: row[0], timestamp: tsISO,
          deUsuario: row[2], dePerfil: row[3], deNome: row[4],
          paraUsuario: row[5], paraPerfil: row[6], paraNome: row[7],
          mensagem: row[8], lida: isLida(row[9]), minha: matchA
        });

        if (marcarLidas && matchB && !isLida(row[9])) updates.push(r + 1);
      }

      if (updates.length) {
        var now = new Date();
        for (var k = 0; k < updates.length; k++) {
          sh.getRange(updates[k], 10).setValue(true);
          sh.getRange(updates[k], 11).setValue(now);
        }
      }

      mensagens.sort(function(a,b){
        return new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime();
      });

      return { success: true, ok: true, mensagens: mensagens };
    } catch (e) {
      return { success: false, ok: false, erro: String(e && e.message ? e.message : e) };
    }
  }

  function enviarMensagemChat(payload) {
    try {
      var p = payload || {};
      var deUsuario   = String(p.deUsuario   || '').trim();
      var dePerfil    = String(p.dePerfil    || '').trim();
      var deNome      = String(p.deNome      || deUsuario).trim();
      var paraUsuario = String(p.paraUsuario || '').trim();
      var paraPerfil  = String(p.paraPerfil  || '').trim();
      var paraNome    = String(p.paraNome    || paraUsuario).trim();
      var mensagem    = String(p.mensagem    || '').trim();

      if (!deUsuario || !paraUsuario || !mensagem)
        return { success: false, ok: false, erro: 'Dados incompletos.' };
      if (deUsuario.toLowerCase() === paraUsuario.toLowerCase())
        return { success: false, ok: false, erro: 'Não é possível enviar mensagem para si mesmo.' };
      if (!podeConversar(dePerfil, paraPerfil))
        return { success: false, ok: false, erro: 'Conversa não permitida entre estes perfis.' };
      if (mensagem.length > 2000)
        return { success: false, ok: false, erro: 'Mensagem muito longa (máx. 2000 caracteres).' };

      var sh = getChatSheet();
      var id = Utilities.getUuid();
      var agora = new Date();
      sh.appendRow([id, agora, deUsuario, dePerfil, deNome,
                    paraUsuario, paraPerfil, paraNome, mensagem, false, '']);

      return { success: true, ok: true, id: id, timestamp: agora.toISOString() };
    } catch (e) {
      return { success: false, ok: false, erro: String(e && e.message ? e.message : e) };
    }
  }

  return {
    carregarContatosChat: carregarContatosChat,
    carregarMensagensChat: carregarMensagensChat,
    enviarMensagemChat: enviarMensagemChat
  };
})();

/* =========================================================
   CHAT INTERNO - BACKEND (versão encapsulada / idempotente)
   Pode colar 1x ou Nx — não gera "already been declared".
   ========================================================= */
var ChatMod = (function () {
  var SHEET_NAME = 'Chat';
  var HEADERS = ['ID','Timestamp','DeUsuario','DePerfil','DeNome',
                 'ParaUsuario','ParaPerfil','ParaNome','Mensagem','Lida','LidaEm'];

  function getChatSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(HEADERS);
      sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setColumnWidth(9, 400);
    } else if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  }

  function getUsuariosSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var nomes = ['Usuarios','Usuários','usuarios','USUARIOS','Login','Logins',
                 'Autenticacao','Autenticação','UsuariosSenhas'];
    for (var i = 0; i < nomes.length; i++) {
      var s = ss.getSheetByName(nomes[i]);
      if (s) return s;
    }
    return null;
  }

  function normPerfil(p) {
    return String(p == null ? '' : p).trim().toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  }

  function isLida(v) {
    if (v === true) return true;
    if (v === false) return false;
    var s = String(v == null ? '' : v).trim().toLowerCase();
    return s === 'true' || s === 'sim' || s === '1' || s === 'yes';
  }

  function podeConversar(pa, pb) {
    var a = normPerfil(pa), b = normPerfil(pb);
    if (a === 'aluno' && b === 'aluno') return false;
    return true;
  }

  function getUsuariosChat_() {
  var sh = getUsuariosSheet_();
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h){ return String(h).trim().toLowerCase(); });

  function idx(names) {
    for (var k = 0; k < names.length; k++) {
      var i = headers.indexOf(names[k]);
      if (i >= 0) return i;
    }
    return -1;
  }

  var iUsuario = idx(['usuario','usuário','login','user','username']);
  var iNome    = idx(['nome','nome completo','name']);
  var iPerfil  = idx(['perfil','role','tipo']);
  var iStatus  = idx(['status','ativo','ativo?']);
  if (iUsuario < 0 || iPerfil < 0) return [];

  // 🔑 Prioridade de perfil: mantém o de maior privilégio por usuário
  var prioridade = { admin:10, direcao:9, coordenacao:8, secretaria:7,
                     financeiro:6, professor:5, aluno:1 };

  var porUsuario = {}; // usuário.toLowerCase() → { usuario, nome, perfil }

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var usuario = String(row[iUsuario] == null ? '' : row[iUsuario]).trim();
    if (!usuario) continue;

    if (iStatus >= 0) {
      var st = String(row[iStatus] == null ? '' : row[iStatus]).trim().toLowerCase();
      if (st && ['inativo','inactive','nao','não','0','false','bloqueado','desativado'].indexOf(st) >= 0) continue;
    }

    var perfil = iPerfil >= 0 ? String(row[iPerfil] || '').trim() : '';
    var nome   = iNome   >= 0 ? String(row[iNome]   || '').trim() || usuario : usuario;
    var key    = usuario.toLowerCase();
    var peso   = prioridade[perfil.toLowerCase()] || 0;

    // Guarda só o registro de maior privilégio por usuário
    if (!porUsuario[key] || peso > (porUsuario[key]._peso || 0)) {
      porUsuario[key] = { usuario: usuario, nome: nome, perfil: perfil, _peso: peso };
    }
  }

  var out = [];
  for (var k in porUsuario) {
    var u = porUsuario[k];
    delete u._peso;
    out.push(u);
  }
  return out;
}

  function carregarContatosChat(payload) {
    try {
      var p = payload || {};
      var usuario = String(p.usuario || '').trim();
      var perfil  = String(p.perfil  || '').trim();
      if (!usuario) return { ok:false, erro:'Usuário não informado.' };

      var todos = getUsuariosChat();
      var contatos = todos.filter(function(u) {
        if (u.usuario.toLowerCase() === usuario.toLowerCase()) return false;
        return podeConversar(perfil, u.perfil);
      });

      var sh = getChatSheet();
      var data = sh.getDataRange().getValues();
      var uLow = usuario.toLowerCase();
      var pLow = normPerfil(perfil);
      var unread = {}, total = 0;

      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        var paraU = String(row[5] == null ? '' : row[5]).toLowerCase();
        var paraP = normPerfil(row[6]);
        if (paraU === uLow && paraP === pLow && !isLida(row[9])) {
          var deU = String(row[2] == null ? '' : row[2]).toLowerCase();
          unread[deU] = (unread[deU] || 0) + 1;
          total++;
        }
      }

      contatos = contatos.map(function(c){
        return {
          usuario: c.usuario,
          nome: c.nome,
          perfil: c.perfil,
          naoLidas: unread[c.usuario.toLowerCase()] || 0
        };
      }).sort(function(a,b){
        return String(a.nome).localeCompare(String(b.nome), 'pt-BR', {sensitivity:'base'});
      });

      return { ok:true, contatos: contatos, totalNaoLidas: total };
    } catch (e) {
      return { ok:false, erro: String(e && e.message ? e.message : e) };
    }
  }

  function carregarMensagensChat(payload) {
    try {
      var p = payload || {};
      var usuario        = String(p.usuario        || '').trim();
      var perfil         = String(p.perfil         || '').trim();
      var contatoUsuario = String(p.contatoUsuario || '').trim();
      var contatoPerfil  = String(p.contatoPerfil  || '').trim();
      var marcarLidas    = !!p.marcarLidas;

      if (!usuario || !contatoUsuario) return { ok:false, erro:'Parâmetros inválidos.' };
      if (!podeConversar(perfil, contatoPerfil)) {
        return { ok:false, erro:'Conversa não permitida entre estes perfis.' };
      }

      var sh = getChatSheet();
      var data = sh.getDataRange().getValues();
      var u1 = usuario.toLowerCase(),     u2 = contatoUsuario.toLowerCase();
      var p1 = normPerfil(perfil),        p2 = normPerfil(contatoPerfil);

      var mensagens = [], updates = [];

      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        var deU   = String(row[2] == null ? '' : row[2]).toLowerCase();
        var deP   = normPerfil(row[3]);
        var paraU = String(row[5] == null ? '' : row[5]).toLowerCase();
        var paraP = normPerfil(row[6]);

        var matchA = (deU === u1 && deP === p1 && paraU === u2 && paraP === p2);
        var matchB = (deU === u2 && deP === p2 && paraU === u1 && paraP === p1);
        if (!matchA && !matchB) continue;

        var ts = row[1];
        var tsISO = (ts instanceof Date) ? ts.toISOString() : String(ts || '');

        mensagens.push({
          id: row[0], timestamp: tsISO,
          deUsuario: row[2], dePerfil: row[3], deNome: row[4],
          paraUsuario: row[5], paraPerfil: row[6], paraNome: row[7],
          mensagem: row[8], lida: isLida(row[9]), minha: matchA
        });

        if (marcarLidas && matchB && !isLida(row[9])) updates.push(r + 1);
      }

      if (updates.length) {
        var now = new Date();
        for (var k = 0; k < updates.length; k++) {
          sh.getRange(updates[k], 10).setValue(true);
          sh.getRange(updates[k], 11).setValue(now);
        }
      }

      mensagens.sort(function(a,b){
        return new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime();
      });

      return { ok:true, mensagens: mensagens };
    } catch (e) {
      return { ok:false, erro: String(e && e.message ? e.message : e) };
    }
  }

  function enviarMensagemChat(payload) {
    try {
      var p = payload || {};
      var deUsuario   = String(p.deUsuario   || '').trim();
      var dePerfil    = String(p.dePerfil    || '').trim();
      var deNome      = String(p.deNome      || deUsuario).trim();
      var paraUsuario = String(p.paraUsuario || '').trim();
      var paraPerfil  = String(p.paraPerfil  || '').trim();
      var paraNome    = String(p.paraNome    || paraUsuario).trim();
      var mensagem    = String(p.mensagem    || '').trim();

      if (!deUsuario || !paraUsuario || !mensagem)
        return { ok:false, erro:'Dados incompletos.' };
      if (deUsuario.toLowerCase() === paraUsuario.toLowerCase() &&
          normPerfil(dePerfil) === normPerfil(paraPerfil))
        return { ok:false, erro:'Não é possível enviar mensagem para si mesmo.' };
      if (!podeConversar(dePerfil, paraPerfil))
        return { ok:false, erro:'Conversa não permitida entre estes perfis.' };
      if (mensagem.length > 2000)
        return { ok:false, erro:'Mensagem muito longa (máx. 2000 caracteres).' };

      var sh = getChatSheet();
      var id = Utilities.getUuid();
      var agora = new Date();
      sh.appendRow([id, agora, deUsuario, dePerfil, deNome,
                    paraUsuario, paraPerfil, paraNome, mensagem, false, '']);

      return { ok:true, id: id, timestamp: agora.toISOString() };
    } catch (e) {
      return { ok:false, erro: String(e && e.message ? e.message : e) };
    }
  }

  return {
    carregarContatosChat: carregarContatosChat,
    carregarMensagensChat: carregarMensagensChat,
    enviarMensagemChat: enviarMensagemChat
  };
})();

/* ---- Wrappers globais (obrigatórios: google.script.run só enxerga funções globais) ---- */
function carregarContatosChat(payload) { return ChatMod.carregarContatosChat(payload); }
function carregarMensagensChat(payload) { return ChatMod.carregarMensagensChat(payload); }
function enviarMensagemChat(payload)    { return ChatMod.enviarMensagemChat(payload); }
