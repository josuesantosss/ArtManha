/**
 * ============================================================
 * ESCOLA DE MÚSICA - WEB APP (API)
 * ============================================================
 * Processa todas as ações do frontend HTML via doGet/doPost
 * 
 * CONFIGURAÇÃO NECESSÁRIA:
 * 1. Implantar como Web App
 * 2. Executar como: "Eu"
 * 3. Quem tem acesso: "Qualquer pessoa"
 * 4. Copiar a URL gerada e colocar em GAS_URL no HTML
 * ============================================================
 */

// ============================================================
// CONFIGURAÇÃO
// ============================================================
const SHEET_ID = '1FRtD7cbV2tvK2og4NDkyWi8V_hziyERNZoSh20gG980';

// Mapa de abas -> nome exato na planilha
const ABAS = {
  alunos: 'ALUNOS',
  matriculas: 'MATRICULAS',
  professores: 'PROFESSORES',
  avaliacoes: 'AVALIACOES',
  presencas: 'PRESENCAS',
  conteudos: 'CONTEUDOS',
  observacoes: 'OBSERVACOES',
  relatorios: 'RELATORIOS',
  logs: 'LOGS'   // NOVO
};

// ============================================================
// ENTRY POINTS - doGet e doPost
// ============================================================

/**
 * GET - Suporte para chamadas via query string (?data=...)
 * Também retorna status quando chamado sem parâmetros
 */
function doGet(e) {
  try {
    if (!e || !e.parameter || !e.parameter.data) {
      return respostaJSON({
        success: true,
        message: '🎵 API Escola de Música online',
        timestamp: new Date().toISOString(),
        versao: '1.1.0',
        acoes_disponiveis: listarAcoes()
      });
    }

    let dados;
    try {
      dados = JSON.parse(e.parameter.data);
    } catch (parseErr) {
      return respostaJSON({
        success: false,
        error: 'JSON inválido: ' + parseErr.message
      });
    }

    return processarAcao(dados.acao, dados.payload);

  } catch (err) {
    logErro('doGet', err);
    return respostaJSON({ success: false, error: err.message });
  }
}

/**
 * POST - Entry point principal
 * Aceita JSON no body (application/json) ou FormData
 */
function doPost(e) {
  try {
    let dados = null;

    if (e && e.postData && e.postData.contents) {
      try {
        dados = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        if (e.parameter && e.parameter.data) {
          try {
            dados = JSON.parse(e.parameter.data);
          } catch (e2) {
            return respostaJSON({
              success: false,
              error: 'Formato de dados inválido'
            });
          }
        } else {
          return respostaJSON({
            success: false,
            error: 'JSON inválido: ' + parseErr.message
          });
        }
      }
    }
    else if (e && e.parameter && e.parameter.data) {
      dados = JSON.parse(e.parameter.data);
    }

    if (!dados || !dados.acao) {
      return respostaJSON({
        success: false,
        error: 'Ação não especificada'
      });
    }

    return processarAcao(dados.acao, dados.payload || {});

  } catch (err) {
    logErro('doPost', err);
    return respostaJSON({ success: false, error: err.message });
  }
}

// ============================================================
// ROTEADOR PRINCIPAL DE AÇÕES
// ============================================================
function processarAcao(acao, payload) {
  Logger.log('📥 Ação recebida: ' + acao);
  Logger.log('📦 Payload: ' + JSON.stringify(payload));

  const acoes = {
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
    'atribuirTurma': atribuirCurso,
    'atribuirCurso': atribuirCurso,
    'removerTurma': removerCurso,
    'removerCurso': removerCurso,

    // Solicitação
    'solicitarRelatorio': solicitarRelatorio,

    // LOGS (NOVO)
    'registrarLogin': registrarLogin,
    'listarLogs': listarLogs,
    'garantirAbaLogs': garantirAbaLogsAPI,

    // Utilitários
    'ping': () => respostaJSON({ success: true, message: 'pong' }),
    'listarCursos': listarCursosAPI,
    'sincronizar': sincronizarTudo,
    'limparAba': limparAba
  };

  const fn = acoes[acao];
  if (!fn) {
    return respostaJSON({
      success: false,
      error: 'Ação desconhecida: ' + acao,
      acoes_disponiveis: Object.keys(acoes)
    });
  }

  try {
    const resultado = fn(payload);
    return resultado;
  } catch (err) {
    logErro('processarAcao:' + acao, err);
    return respostaJSON({
      success: false,
      error: err.message,
      acao: acao
    });
  }
}

// ============================================================
// HELPERS
// ============================================================

function respostaJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function logErro(contexto, err) {
  Logger.log('❌ ERRO [' + contexto + ']: ' + err.message);
  Logger.log('Stack: ' + (err.stack || 'sem stack'));
}

function getSheet(chave) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const nomeAba = ABAS[chave];
  const sheet = ss.getSheetByName(nomeAba);
  if (!sheet) {
    throw new Error('Aba não encontrada: ' + nomeAba);
  }
  return sheet;
}

function listarAcoes() {
  return [
    'salvarAluno', 'atualizarAluno', 'removerAluno', 'listarAlunos',
    'salvarProfessor', 'atualizarProfessor', 'removerProfessor', 'listarProfessores', 'salvarValorProfessor',
    'salvarMatricula', 'atualizarMatricula', 'removerMatricula', 'listarMatriculas',
    'salvarAvaliacoes', 'listarAvaliacoes',
    'salvarPresenca', 'listarPresencas',
    'salvarConteudo', 'listarConteudos',
    'salvarObservacao', 'listarObservacoes',
    'salvarRelatorio', 'marcarRelatorioVisto', 'listarRelatorios',
    'atribuirCurso', 'removerCurso', 'solicitarRelatorio',
    'registrarLogin', 'listarLogs', 'garantirAbaLogs',
    'listarCursos', 'sincronizar', 'limparAba', 'ping'
  ];
}

/**
 * Gera um ID único com prefixo
 */
function gerarID(prefixo) {
  const agora = new Date();
  const ts = agora.getTime();
  const rand = Math.floor(Math.random() * 10000);
  return prefixo + '_' + ts + '_' + rand;
}

/**
 * Converte data para string DD-MM-YYYY
 */
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

function dataHoje() {
  return formatarData(new Date());
}

/**
 * Retorna headers de uma aba
 */
function getHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
}

/**
 * Retorna todos os registros de uma aba como objetos
 */
function getRegistros(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const headers = getHeaders(sheet);
  const valores = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  
  return valores.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i];
    });
    return obj;
  });
}

/**
 * Encontra a linha de um registro por critério
 * Retorna o número da linha (1-based) ou -1
 */
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
      const cellValor = String(valores[r][idx]).trim();
      const buscaValor = String(valor).trim();
      if (cellValor !== buscaValor) {
        match = false;
        break;
      }
    }
    if (match) return r + 2;
  }
  return -1;
}

/**
 * Adiciona um registro (array de valores)
 */
function appendRegistro(sheet, valores) {
  sheet.appendRow(valores);
  return sheet.getLastRow();
}

/**
 * Atualiza uma linha inteira
 */
function atualizarLinha(sheet, linha, valores) {
  sheet.getRange(linha, 1, 1, valores.length).setValues([valores]);
}

/**
 * Remove uma linha
 */
function removerLinha(sheet, linha) {
  sheet.deleteRow(linha);
}

/**
 * Garante que uma string não seja nula
 */
function s(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

// ============================================================
// AÇÕES - ALUNOS
// ============================================================

function salvarAluno(payload) {
  const { nome, dataNascimento, responsavel, contato, status } = payload;

  if (!nome) {
    return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  }

  const sheet = getSheet('alunos');
  const headers = getHeaders(sheet);

  const linhaExistente = encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() });
  if (linhaExistente > -1) {
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
  
  return respostaJSON({
    success: true,
    message: 'Aluno salvo: ' + nome,
    id: linha,
    aluno: { nome, dataNascimento, responsavel, contato, status }
  });
}

function atualizarAluno(payload) {
  const { nome, nomeOriginal, dataNascimento, responsavel, contato, status } = payload;

  if (!nome) {
    return respostaJSON({ success: false, error: 'Nome é obrigatório' });
  }

  const sheet = getSheet('alunos');
  const headers = getHeaders(sheet);
  const linha = encontrarLinha(sheet, { 'Nome': s(nomeOriginal || nome).toUpperCase() });

  if (linha < 0) {
    return respostaJSON({ success: false, error: 'Aluno não encontrado: ' + (nomeOriginal || nome) });
  }

  const valoresAntigos = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const valoresNovos = headers.map((h, i) => {
    switch (h) {
      case 'Nome': return s(nome).toUpperCase();
      case 'Data_Nascimento': return dataNascimento !== undefined ? formatarData(dataNascimento) : valoresAntigos[i];
      case 'Responsavel': return responsavel !== undefined ? s(responsavel).toUpperCase() : valoresAntigos[i];
      case 'Contato': return contato !== undefined ? s(contato) : valoresAntigos[i];
      case 'Status': return status !== undefined ? s(status).toUpperCase() : valoresAntigos[i];
      default: return valoresAntigos[i];
    }
  });

  atualizarLinha(sheet, linha, valoresNovos);
  return respostaJSON({ success: true, message: 'Aluno atualizado: ' + nome });
}

function removerAluno(payload) {
  const { nome } = payload;
  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });

  const sheet = getSheet('alunos');
  const linha = encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() });
  if (linha < 0) return respostaJSON({ success: false, error: 'Aluno não encontrado' });

  removerLinha(sheet, linha);
  return respostaJSON({ success: true, message: 'Aluno removido: ' + nome });
}

function listarAlunos(payload) {
  const sheet = getSheet('alunos');
  const registros = getRegistros(sheet);
  return respostaJSON({ success: true, total: registros.length, alunos: registros });
}

// ============================================================
// AÇÕES - PROFESSORES
// ============================================================

function salvarProfessor(payload) {
  const { id, nome, tipo, curso, telefone, email, valorHora } = payload;

  if (!nome) return respostaJSON({ success: false, error: 'Nome é obrigatório' });

  const sheet = getSheet('professores');
  const headers = getHeaders(sheet);
  const idFinal = s(id) || gerarID('PROF');

  const linhaExistente = encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() });
  if (linhaExistente > -1) {
    return respostaJSON({ success: false, error: 'Professor já cadastrado: ' + nome });
  }

  // Garante que exista a coluna Valor_Hora (cria se necessário)
  garantirColunaValorHora(sheet);

  const headersAtualizados = getHeaders(sheet);

  const valores = headersAtualizados.map(h => {
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
  return respostaJSON({
    success: true,
    message: 'Professor salvo: ' + nome,
    id: idFinal
  });
}

function atualizarProfessor(payload) {
  const { id, nome, tipo, curso, telefone, email, status, valorHora } = payload;

  if (!id && !nome) {
    return respostaJSON({ success: false, error: 'ID ou Nome do professor é obrigatório' });
  }

  const sheet = getSheet('professores');
  garantirColunaValorHora(sheet);
  const headers = getHeaders(sheet);
  
  let linha = -1;
  if (id) linha = encontrarLinha(sheet, { 'ID_Professor': s(id) });
  if (linha < 0 && nome) linha = encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() });

  if (linha < 0) {
    return respostaJSON({ success: false, error: 'Professor não encontrado' });
  }

  const valoresAntigos = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const valoresNovos = headers.map((h, i) => {
    switch (h) {
      case 'Nome': return nome !== undefined ? s(nome).toUpperCase() : valoresAntigos[i];
      case 'Tipo': return tipo !== undefined ? s(tipo) : valoresAntigos[i];
      case 'Curso': return curso !== undefined ? s(curso) : valoresAntigos[i];
      case 'Telefone': return telefone !== undefined ? s(telefone) : valoresAntigos[i];
      case 'Email': return email !== undefined ? s(email) : valoresAntigos[i];
      case 'Status': return status !== undefined ? s(status).toUpperCase() : valoresAntigos[i];
      case 'Valor_Hora': return valorHora !== undefined ? (Number(valorHora) || 0) : valoresAntigos[i];
      default: return valoresAntigos[i];
    }
  });

  atualizarLinha(sheet, linha, valoresNovos);
  return respostaJSON({ success: true, message: 'Professor atualizado' });
}

/**
 * Ação dedicada a salvar apenas o valor/hora do professor.
 * Usada pela tela do financeiro.
 */
function salvarValorProfessor(payload) {
  const { id, valorHora } = payload;
  if (!id) {
    return respostaJSON({ success: false, error: 'ID do professor é obrigatório' });
  }

  const sheet = getSheet('professores');
  garantirColunaValorHora(sheet);
  const headers = getHeaders(sheet);
  const idxValor = headers.indexOf('Valor_Hora');

  if (idxValor < 0) {
    return respostaJSON({ success: false, error: 'Coluna Valor_Hora não encontrada' });
  }

  const linha = encontrarLinha(sheet, { 'ID_Professor': s(id) });
  if (linha < 0) {
    return respostaJSON({ success: false, error: 'Professor não encontrado' });
  }

  const valor = Number(valorHora) || 0;
  sheet.getRange(linha, idxValor + 1).setValue(valor);

  return respostaJSON({ success: true, message: 'Valor hora/aula atualizado', valorHora: valor });
}

/**
 * Garante que a coluna Valor_Hora existe na aba PROFESSORES.
 * Se não existir, cria no fim do cabeçalho.
 */
function garantirColunaValorHora(sheet) {
  const headers = getHeaders(sheet);
  if (headers.indexOf('Valor_Hora') < 0) {
    const ultimaCol = sheet.getLastColumn();
    sheet.getRange(1, ultimaCol + 1).setValue('Valor_Hora');
    Logger.log('✅ Coluna Valor_Hora criada na aba PROFESSORES');
  }
}

function removerProfessor(payload) {
  const { id, nome } = payload;
  if (!id && !nome) return respostaJSON({ success: false, error: 'ID ou Nome é obrigatório' });

  const sheet = getSheet('professores');
  let linha = -1;
  if (id) linha = encontrarLinha(sheet, { 'ID_Professor': s(id) });
  if (linha < 0 && nome) linha = encontrarLinha(sheet, { 'Nome': s(nome).toUpperCase() });

  if (linha < 0) return respostaJSON({ success: false, error: 'Professor não encontrado' });

  removerLinha(sheet, linha);
  return respostaJSON({ success: true, message: 'Professor removido' });
}

function listarProfessores(payload) {
  const sheet = getSheet('professores');
  const registros = getRegistros(sheet);
  return respostaJSON({ success: true, total: registros.length, professores: registros });
}

// ============================================================
// AÇÕES - MATRÍCULAS
// ============================================================

function salvarMatricula(payload) {
  const { aluno, curso, professor, dataMatricula, anoLetivo, status, observacoes } = payload;

  if (!aluno || !curso) {
    return respostaJSON({ success: false, error: 'Aluno e Curso são obrigatórios' });
  }

  const sheet = getSheet('matriculas');
  const headers = getHeaders(sheet);

  const existente = encontrarLinha(sheet, {
    'Aluno': s(aluno).toUpperCase(),
    'Curso': s(curso)
  });

  if (existente > -1) {
    const valoresExistentes = sheet.getRange(existente, 1, 1, headers.length).getValues()[0];
    const idxStatus = headers.indexOf('Status');
    if (idxStatus >= 0 && String(valoresExistentes[idxStatus]).toUpperCase() === 'ATIVO') {
      return respostaJSON({
        success: false,
        error: 'Aluno já possui matrícula ativa neste curso'
      });
    }
  }

  const idMatricula = gerarID('MAT');
  const valores = headers.map(h => {
    switch (h) {
      case 'ID_Matricula': return idMatricula;
      case 'Aluno': return s(aluno).toUpperCase();
      case 'Curso': return s(curso);
      case 'Professor': return s(professor);
      case 'Data_Matricula': return formatarData(dataMatricula) || dataHoje();
      case 'Ano_Letivo': return s(anoLetivo) || String(new Date().getFullYear());
      case 'Status': return s(status || 'ATIVO').toUpperCase();
      case 'Observacoes': return s(observacoes);
      default: return '';
    }
  });

  appendRegistro(sheet, valores);
  return respostaJSON({
    success: true,
    message: 'Matrícula registrada: ' + aluno + ' → ' + curso,
    id: idMatricula
  });
}

function atualizarMatricula(payload) {
  const { id, aluno, curso, professor, dataMatricula, anoLetivo, status, observacoes } = payload;

  const sheet = getSheet('matriculas');
  const headers = getHeaders(sheet);

  let linha = -1;
  if (id) linha = encontrarLinha(sheet, { 'ID_Matricula': s(id) });
  if (linha < 0 && aluno && curso) linha = encontrarLinha(sheet, { 'Aluno': s(aluno).toUpperCase(), 'Curso': s(curso) });

  if (linha < 0) return respostaJSON({ success: false, error: 'Matrícula não encontrada' });

  const valoresAntigos = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const valoresNovos = headers.map((h, i) => {
    switch (h) {
      case 'Aluno': return aluno !== undefined ? s(aluno).toUpperCase() : valoresAntigos[i];
      case 'Curso': return curso !== undefined ? s(curso) : valoresAntigos[i];
      case 'Professor': return professor !== undefined ? s(professor) : valoresAntigos[i];
      case 'Data_Matricula': return dataMatricula !== undefined ? formatarData(dataMatricula) : valoresAntigos[i];
      case 'Ano_Letivo': return anoLetivo !== undefined ? s(anoLetivo) : valoresAntigos[i];
      case 'Status': return status !== undefined ? s(status).toUpperCase() : valoresAntigos[i];
      case 'Observacoes': return observacoes !== undefined ? s(observacoes) : valoresAntigos[i];
      default: return valoresAntigos[i];
    }
  });

  atualizarLinha(sheet, linha, valoresNovos);
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
  const sheet = getSheet('matriculas');
  const registros = getRegistros(sheet);
  const filtroAluno = payload && payload.aluno ? s(payload.aluno).toUpperCase() : null;
  const filtroCurso = payload && payload.curso ? s(payload.curso) : null;
  const filtroStatus = payload && payload.status ? s(payload.status).toUpperCase() : null;

  let filtrados = registros;
  if (filtroAluno) filtrados = filtrados.filter(r => s(r.Aluno).toUpperCase() === filtroAluno);
  if (filtroCurso) filtrados = filtrados.filter(r => s(r.Curso) === filtroCurso);
  if (filtroStatus) filtrados = filtrados.filter(r => s(r.Status).toUpperCase() === filtroStatus);

  return respostaJSON({ success: true, total: filtrados.length, matriculas: filtrados });
}

// ============================================================
// AÇÕES - AVALIAÇÕES
// ============================================================

function salvarAvaliacoes(payload) {
  const lista = payload.avaliacoes || payload.notas || [];
  if (!Array.isArray(lista) || lista.length === 0) {
    return respostaJSON({ success: false, error: 'Nenhuma avaliação para salvar' });
  }

  const sheet = getSheet('avaliacoes');
  const headers = getHeaders(sheet);
  const resultados = { inseridos: 0, atualizados: 0, erros: [] };

  lista.forEach(av => {
    try {
      const aluno = s(av.aluno).toUpperCase();
      const curso = s(av.curso);
      const bimestre = String(av.bimestre || '1');
      const nota = Number(av.nota) || 0;
      const professor = s(av.professor || av.professorNome || '');

      if (!aluno || !curso) {
        resultados.erros.push('Aluno ou curso vazio: ' + JSON.stringify(av));
        return;
      }

      const linhaExistente = encontrarLinha(sheet, {
        'Aluno': aluno,
        'Curso': curso,
        'Bimestre': bimestre
      });

      if (linhaExistente > -1) {
        const valoresAntigos = sheet.getRange(linhaExistente, 1, 1, headers.length).getValues()[0];
        const valoresNovos = headers.map((h, i) => {
          switch (h) {
            case 'Nota': return nota;
            case 'Data_Lancamento': return dataHoje();
            case 'Professor': return professor || valoresAntigos[i];
            default: return valoresAntigos[i];
          }
        });
        atualizarLinha(sheet, linhaExistente, valoresNovos);
        resultados.atualizados++;
      } else {
        const idAval = gerarID('AVA');
        const valores = headers.map(h => {
          switch (h) {
            case 'ID_Avaliacao': return idAval;
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
        resultados.inseridos++;
      }
    } catch (e) {
      resultados.erros.push(e.message);
    }
  });

  return respostaJSON({
    success: resultados.erros.length === 0,
    message: `${resultados.inseridos} inserida(s), ${resultados.atualizados} atualizada(s)`,
    detalhes: resultados
  });
}

function listarAvaliacoes(payload) {
  const sheet = getSheet('avaliacoes');
  let registros = getRegistros(sheet);

  if (payload) {
    if (payload.aluno) registros = registros.filter(r => s(r.Aluno).toUpperCase() === s(payload.aluno).toUpperCase());
    if (payload.curso) registros = registros.filter(r => s(r.Curso) === s(payload.curso));
    if (payload.bimestre) registros = registros.filter(r => String(r.Bimestre) === String(payload.bimestre));
  }

  return respostaJSON({ success: true, total: registros.length, avaliacoes: registros });
}

// ============================================================
// AÇÕES - PRESENÇAS
// ============================================================
/**
 * Grava presença. O conteúdo da aula vai para a coluna F da aba
 * PRESENCAS (índice 5, 0-based). Aceita os campos `conteudo` e
 * `observacao` (fallback) — o HTML envia `conteudo`.
 */
function salvarPresenca(payload) {
  const { curso, data, presentes, faltantes, observacao, conteudo, professor } = payload;

  if (!curso || !data) {
    return respostaJSON({ success: false, error: 'Curso e Data são obrigatórios' });
  }

  const sheet = getSheet('presencas');
  const headers = getHeaders(sheet);

  let listaPresentes = presentes;
  if (Array.isArray(listaPresentes)) listaPresentes = listaPresentes.join(', ');
  listaPresentes = s(listaPresentes);

  let listaFaltantes = faltantes;
  if (Array.isArray(listaFaltantes)) listaFaltantes = listaFaltantes.join(', ');
  listaFaltantes = s(listaFaltantes);

  // Conteúdo: aceita payload.conteudo OU payload.observacao
  const conteudoFinal = s(conteudo || observacao || '');

  const dataFormatada = formatarData(data);
  const idPres = gerarID('PRE');

  const valores = headers.map((h, i) => {
    switch (h) {
      case 'ID_Presenca': return idPres;
      case 'Curso': return s(curso);
      case 'Data': return dataFormatada;
      case 'Presentes': return listaPresentes;
      case 'Faltantes': return listaFaltantes;
      case 'Conteudo': return conteudoFinal;
      case 'Observacao': return conteudoFinal;
      case 'Professor': return s(professor);
      default: return '';
    }
  });

  // Garante gravação na coluna F (índice 5) se o header não for
  // explicitamente "Conteudo" nem "Observacao".
  const temHeaderConteudo = headers.some(h => h === 'Conteudo' || h === 'Observacao');
  if (!temHeaderConteudo && headers.length >= 6) {
    valores[5] = conteudoFinal;
  }

  appendRegistro(sheet, valores);
  return respostaJSON({
    success: true,
    message: 'Presença registrada para ' + curso + ' em ' + dataFormatada,
    id: idPres
  });
}

function listarPresencas(payload) {
  const sheet = getSheet('presencas');
  let registros = getRegistros(sheet);

  if (payload) {
    if (payload.curso) registros = registros.filter(r => s(r.Curso) === s(payload.curso));
    if (payload.data) registros = registros.filter(r => s(r.Data) === formatarData(payload.data));
  }

  return respostaJSON({ success: true, total: registros.length, presencas: registros });
}

// ============================================================
// AÇÕES - CONTEÚDOS
// ============================================================

function salvarConteudo(payload) {
  const { curso, data, conteudo, objetivos, professor } = payload;

  if (!curso || !conteudo) {
    return respostaJSON({ success: false, error: 'Curso e Conteúdo são obrigatórios' });
  }

  const sheet = getSheet('conteudos');
  const headers = getHeaders(sheet);
  const idCont = gerarID('CON');
  const dataFormatada = formatarData(data) || dataHoje();

  const valores = headers.map(h => {
    switch (h) {
      case 'ID_Conteudo': return idCont;
      case 'Curso': return s(curso);
      case 'Data': return dataFormatada;
      case 'Conteudo': return s(conteudo);
      case 'Objetivos': return s(objetivos);
      case 'Professor': return s(professor);
      default: return '';
    }
  });

  appendRegistro(sheet, valores);
  return respostaJSON({
    success: true,
    message: 'Conteúdo salvo',
    id: idCont
  });
}

function listarConteudos(payload) {
  const sheet = getSheet('conteudos');
  let registros = getRegistros(sheet);

  if (payload && payload.curso) {
    registros = registros.filter(r => s(r.Curso) === s(payload.curso));
  }

  return respostaJSON({ success: true, total: registros.length, conteudos: registros });
}

// ============================================================
// AÇÕES - OBSERVAÇÕES
// ============================================================

function salvarObservacao(payload) {
  const { curso, data, aluno, observacao, tipo, professor } = payload;

  if (!curso || !aluno || !observacao) {
    return respostaJSON({ success: false, error: 'Curso, Aluno e Observação são obrigatórios' });
  }

  const sheet = getSheet('observacoes');
  const headers = getHeaders(sheet);
  const idObs = gerarID('OBS');
  const dataFormatada = formatarData(data) || dataHoje();

  const valores = headers.map(h => {
    switch (h) {
      case 'ID_Observacao': return idObs;
      case 'Curso': return s(curso);
      case 'Data': return dataFormatada;
      case 'Aluno': return s(aluno).toUpperCase();
      case 'Observacao': return s(observacao);
      case 'Tipo': return s(tipo || 'Neutra');
      case 'Professor': return s(professor);
      default: return '';
    }
  });

  appendRegistro(sheet, valores);
  return respostaJSON({ success: true, message: 'Observação salva', id: idObs });
}

function listarObservacoes(payload) {
  const sheet = getSheet('observacoes');
  let registros = getRegistros(sheet);

  if (payload) {
    if (payload.curso) registros = registros.filter(r => s(r.Curso) === s(payload.curso));
    if (payload.aluno) registros = registros.filter(r => s(r.Aluno).toUpperCase() === s(payload.aluno).toUpperCase());
  }

  return respostaJSON({ success: true, total: registros.length, observacoes: registros });
}

// ============================================================
// AÇÕES - RELATÓRIOS
// ============================================================

function salvarRelatorio(payload) {
  const { aluno, curso, bimestre, professor, relatorio, data, tipo, notas, media, presenca } = payload;

  if (!aluno || !curso || !relatorio) {
    return respostaJSON({ success: false, error: 'Aluno, Curso e Relatório são obrigatórios' });
  }

  const sheet = getSheet('relatorios');
  const headers = getHeaders(sheet);
  const idRel = gerarID('REL');
  const dataFormatada = formatarData(data) || dataHoje();

  let textoFinal = s(relatorio);
  if (notas) textoFinal = '[NOTAS: ' + s(notas) + ']\n' + textoFinal;
  if (media) textoFinal = '[MÉDIA: ' + s(media) + ']\n' + textoFinal;
  if (presenca) textoFinal = '[PRESENÇA: ' + s(presenca) + '%]\n' + textoFinal;
  if (tipo) textoFinal = '[TIPO: ' + s(tipo) + ']\n' + textoFinal;

  const valores = headers.map(h => {
    switch (h) {
      case 'ID_Relatorio': return idRel;
      case 'Aluno': return s(aluno).toUpperCase();
      case 'Curso': return s(curso);
      case 'Bimestre': return String(bimestre || '1');
      case 'Professor': return s(professor);
      case 'Relatorio': return textoFinal;
      case 'Data': return dataFormatada;
      case 'Visto': return 'NÃO';
      case 'Data_Visto': return '';
      case 'Observacoes_Coordenacao': return '';
      default: return '';
    }
  });

  appendRegistro(sheet, valores);
  return respostaJSON({
    success: true,
    message: 'Relatório salvo para ' + aluno,
    id: idRel
  });
}

function marcarRelatorioVisto(payload) {
  const { id } = payload;
  if (!id) return respostaJSON({ success: false, error: 'ID do relatório é obrigatório' });

  const sheet = getSheet('relatorios');
  const headers = getHeaders(sheet);
  const linha = encontrarLinha(sheet, { 'ID_Relatorio': s(id) });

  if (linha < 0) return respostaJSON({ success: false, error: 'Relatório não encontrado' });

  const valores = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const idxVisto = headers.indexOf('Visto');
  const idxDataVisto = headers.indexOf('Data_Visto');

  if (idxVisto >= 0) valores[idxVisto] = 'SIM';
  if (idxDataVisto >= 0) valores[idxDataVisto] = dataHoje();

  atualizarLinha(sheet, linha, valores);
  return respostaJSON({ success: true, message: 'Relatório marcado como visto' });
}

function listarRelatorios(payload) {
  const sheet = getSheet('relatorios');
  let registros = getRegistros(sheet);

  if (payload) {
    if (payload.aluno) registros = registros.filter(r => s(r.Aluno).toUpperCase() === s(payload.aluno).toUpperCase());
    if (payload.curso) registros = registros.filter(r => s(r.Curso) === s(payload.curso));
    if (payload.professor) registros = registros.filter(r => s(r.Professor) === s(payload.professor));
    if (payload.visto) registros = registros.filter(r => s(r.Visto).toUpperCase() === s(payload.visto).toUpperCase());
  }

  return respostaJSON({ success: true, total: registros.length, relatorios: registros });
}

// ============================================================
// AÇÕES - ATRIBUIÇÃO DE CURSOS A PROFESSORES
// ============================================================

function atribuirCurso(payload) {
  const { professorId, professorNome, curso, turma } = payload;
  const cursoFinal = curso || turma;

  if (!cursoFinal) {
    return respostaJSON({ success: false, error: 'Curso é obrigatório' });
  }

  const sheet = getSheet('professores');
  const headers = getHeaders(sheet);
  const idxCurso = headers.indexOf('Curso');

  if (idxCurso < 0) {
    return respostaJSON({ success: false, error: 'Coluna Curso não existe na aba PROFESSORES' });
  }

  let linha = -1;
  if (professorId) linha = encontrarLinha(sheet, { 'ID_Professor': s(professorId) });
  if (linha < 0 && professorNome) linha = encontrarLinha(sheet, { 'Nome': s(professorNome).toUpperCase() });

  if (linha < 0) return respostaJSON({ success: false, error: 'Professor não encontrado' });

  const valores = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  const cursosAtuais = s(valores[idxCurso]).split(',').map(c => c.trim()).filter(Boolean);

  if (cursosAtuais.includes(cursoFinal)) {
    return respostaJSON({ success: true, message: 'Curso já atribuído', jaExistia: true });
  }

  cursosAtuais.push(cursoFinal);
  valores[idxCurso] = cursosAtuais.join(',');

  atualizarLinha(sheet, linha, valores);
  return respostaJSON({
    success: true,
    message: 'Curso ' + cursoFinal + ' atribuído',
    cursos: cursosAtuais
  });
}

function removerCurso(payload) {
  const { professorId, professorNome, curso, turma } = payload;
  const cursoFinal = curso || turma;

  if (!cursoFinal) {
    return respostaJSON({ success: false, error: 'Curso é obrigatório' });
  }

  const sheet = getSheet('professores');
  const headers = getHeaders(sheet);
  const idxCurso = headers.indexOf('Curso');

  let linha = -1;
  if (professorId) linha = encontrarLinha(sheet, { 'ID_Professor': s(professorId) });
  if (linha < 0 && professorNome) linha = encontrarLinha(sheet, { 'Nome': s(professorNome).toUpperCase() });

  if (linha < 0) return respostaJSON({ success: false, error: 'Professor não encontrado' });

  const valores = sheet.getRange(linha, 1, 1, headers.length).getValues()[0];
  let cursos = s(valores[idxCurso]).split(',').map(c => c.trim()).filter(Boolean);
  cursos = cursos.filter(c => c !== cursoFinal);

  valores[idxCurso] = cursos.join(',');
  atualizarLinha(sheet, linha, valores);

  return respostaJSON({
    success: true,
    message: 'Curso ' + cursoFinal + ' removido',
    cursos: cursos
  });
}

// ============================================================
// AÇÕES - SOLICITAÇÃO DE RELATÓRIO
// ============================================================

function solicitarRelatorio(payload) {
  const { aluno, curso, bimestre, tipo, professor, dataSolicitacao } = payload;

  if (!aluno || !curso) {
    return respostaJSON({ success: false, error: 'Aluno e Curso são obrigatórios' });
  }

  const sheet = getSheet('relatorios');
  const headers = getHeaders(sheet);
  const idRel = gerarID('SOL');
  const dataFormatada = formatarData(dataSolicitacao) || dataHoje();

  const textoInicial = '[SOLICITADO EM ' + dataFormatada + ']\n' +
                       '[TIPO: ' + s(tipo || 'pedagogico') + ']\n' +
                       'Aguardando preenchimento pelo professor.';

  const valores = headers.map(h => {
    switch (h) {
      case 'ID_Relatorio': return idRel;
      case 'Aluno': return s(aluno).toUpperCase();
      case 'Curso': return s(curso);
      case 'Bimestre': return String(bimestre || 'anual');
      case 'Professor': return s(professor);
      case 'Relatorio': return textoInicial;
      case 'Data': return dataFormatada;
      case 'Visto': return 'NÃO';
      default: return '';
    }
  });

  appendRegistro(sheet, valores);
  return respostaJSON({
    success: true,
    message: 'Solicitação registrada',
    id: idRel
  });
}

// ============================================================
// AÇÕES - LOGS (NOVO)
// ============================================================

/**
 * Garante que a aba LOGS existe e possui cabeçalho.
 * Retorna o objeto Sheet.
 */
function garantirAbaLogs() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(ABAS.logs);

  if (!sheet) {
    sheet = ss.insertSheet(ABAS.logs);
    sheet.appendRow(['ID_Log', 'Nome', 'Perfil', 'Data_Hora', 'Acao']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log('✅ Aba LOGS criada automaticamente');
    return sheet;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Cria cabeçalho se a aba estiver vazia
  if (lastRow === 0 || lastCol === 0) {
    sheet.getRange(1, 1, 1, 5).setValues([['ID_Log', 'Nome', 'Perfil', 'Data_Hora', 'Acao']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log('✅ Cabeçalho da aba LOGS criado');
  }

  return sheet;
}

/**
 * Versão da garantia exposta como ação HTTP (retorna JSON).
 */
function garantirAbaLogsAPI(payload) {
  try {
    const sheet = garantirAbaLogs();
    return respostaJSON({
      success: true,
      message: 'Aba LOGS pronta',
      total_linhas: sheet.getLastRow(),
      headers: getHeaders(sheet)
    });
  } catch (e) {
    logErro('garantirAbaLogsAPI', e);
    return respostaJSON({ success: false, error: e.message });
  }
}

/**
 * Registra um login na aba LOGS.
 * payload: { nome, perfil, acao (opcional) }
 */
function registrarLogin(payload) {
  const { nome, perfil, acao } = payload || {};

  if (!nome || !perfil) {
    return respostaJSON({
      success: false,
      error: 'Nome e perfil são obrigatórios'
    });
  }

  try {
    const sheet = garantirAbaLogs();
    const headers = getHeaders(sheet);
    const idLog = gerarID('LOG');

    const agora = new Date();
    const dataHora = Utilities.formatDate(
      agora,
      Session.getScriptTimeZone(),
      'dd-MM-yyyy HH:mm:ss'
    );

    const valores = headers.map(h => {
      switch (h) {
        case 'ID_Log':     return idLog;
        case 'Nome':       return s(nome).toUpperCase();
        case 'Perfil':     return s(perfil);
        case 'Data_Hora':  return dataHora;
        case 'Acao':       return s(acao || 'LOGIN');
        default:           return '';
      }
    });

    appendRegistro(sheet, valores);

    return respostaJSON({
      success: true,
      message: 'Login registrado',
      id: idLog,
      dataHora: dataHora
    });
  } catch (e) {
    logErro('registrarLogin', e);
    return respostaJSON({ success: false, error: e.message });
  }
}

/**
 * Lista logs. Filtros opcionais: { perfil, nome }
 */
function listarLogs(payload) {
  try {
    const sheet = garantirAbaLogs();
    let registros = getRegistros(sheet);

    if (payload) {
      if (payload.perfil) {
        registros = registros.filter(r =>
          s(r.Perfil).toLowerCase() === s(payload.perfil).toLowerCase()
        );
      }
      if (payload.nome) {
        registros = registros.filter(r =>
          s(r.Nome).toUpperCase() === s(payload.nome).toUpperCase()
        );
      }
    }

    return respostaJSON({
      success: true,
      total: registros.length,
      logs: registros
    });
  } catch (e) {
    logErro('listarLogs', e);
    return respostaJSON({ success: false, error: e.message });
  }
}

// ============================================================
// AÇÕES UTILITÁRIAS
// ============================================================

function listarCursosAPI(payload) {
  const CURSOS = [
    { id: 'piano', nome: 'Piano', tipo: 'Instrumento' },
    { id: 'teclado', nome: 'Teclado', tipo: 'Instrumento' },
    { id: 'orgao', nome: 'Órgão', tipo: 'Instrumento' },
    { id: 'violao', nome: 'Violão', tipo: 'Instrumento' },
    { id: 'guitarra', nome: 'Guitarra', tipo: 'Instrumento' },
    { id: 'violino', nome: 'Violino', tipo: 'Instrumento' },
    { id: 'violoncelo', nome: 'Violoncelo', tipo: 'Instrumento' },
    { id: 'viola', nome: 'Viola Erudita', tipo: 'Instrumento' },
    { id: 'contrabaixo', nome: 'Contrabaixo Acústico', tipo: 'Instrumento' },
    { id: 'bateria', nome: 'Bateria', tipo: 'Instrumento' },
    { id: 'trompete', nome: 'Trompete', tipo: 'Instrumento' },
    { id: 'trombone', nome: 'Trombone', tipo: 'Instrumento' },
    { id: 'euphonio', nome: 'Eufônio', tipo: 'Instrumento' },
    { id: 'tuba', nome: 'Tuba', tipo: 'Instrumento' },
    { id: 'saxofone', nome: 'Saxofone', tipo: 'Instrumento' },
    { id: 'flauta', nome: 'Flauta', tipo: 'Instrumento' },
    { id: 'clarinete', nome: 'Clarinete', tipo: 'Instrumento' },
    { id: 'canto', nome: 'Canto', tipo: 'Instrumento' },
    { id: 'teoria', nome: 'Teoria Musical', tipo: 'Teoria Musical' },
    { id: 'musicalizacao', nome: 'Musicalização Infantil', tipo: 'Musicalização' }
  ];
  return respostaJSON({ success: true, total: CURSOS.length, cursos: CURSOS });
}

function sincronizarTudo(payload) {
  const resultado = {};
  Object.entries(ABAS).forEach(([chave, nomeAba]) => {
    try {
      if (chave === 'logs') {
        // Garante a aba LOGS e carrega
        garantirAbaLogs();
      }
      const sheet = getSheet(chave);
      resultado[chave] = getRegistros(sheet);
    } catch (e) {
      resultado[chave] = [];
      resultado[chave + '_erro'] = e.message;
    }
  });

  return respostaJSON({
    success: true,
    timestamp: new Date().toISOString(),
    dados: resultado
  });
}

function limparAba(payload) {
  const { aba } = payload;
  if (!aba) return respostaJSON({ success: false, error: 'Aba é obrigatória' });

  const chave = Object.keys(ABAS).find(k => ABAS[k] === aba || k === aba);
  if (!chave) return respostaJSON({ success: false, error: 'Aba não reconhecida: ' + aba });

  const sheet = getSheet(chave);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  return respostaJSON({ success: true, message: 'Aba ' + aba + ' limpa' });
}

// ============================================================
// TESTE MANUAL
// ============================================================
function testarAPI() {
  Logger.log('🧪 Testando API...');

  const r1 = processarAcao('ping', {});
  Logger.log('Ping: ' + r1.getContent());

  const r2 = processarAcao('listarCursos', {});
  Logger.log('Cursos: ' + r2.getContent().substring(0, 200));

  // Garante e testa LOGS
  const sheetLogs = garantirAbaLogs();
  Logger.log('LOGS headers: ' + JSON.stringify(getHeaders(sheetLogs)));

  const r3 = processarAcao('registrarLogin', {
    nome: 'TESTE AUTOMÁTICO',
    perfil: 'teste',
    acao: 'TESTE_MANUAL'
  });
  Logger.log('registrarLogin: ' + r3.getContent());

  const r4 = processarAcao('sincronizar', {});
  Logger.log('Sincronizar: ' + r4.getContent().substring(0, 300));

  Logger.log('✅ Teste concluído');
}

/**
 * Cria o menu personalizado na planilha
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎵 Escola de Música')
    .addItem('🧪 Testar API', 'testarAPI')
    .addItem('📊 Ver Status da API', 'mostrarStatusAPI')
    .addSeparator()
    .addItem('📋 Criar/Verificar aba LOGS', 'garantirAbaLogs')
    .addItem('📋 Criar/Verificar coluna Valor_Hora', 'menuGarantirValorHora')
    .addSeparator()
    .addItem('🔗 Mostrar URL do Web App', 'mostrarURLWebApp')
    .addToUi();
}

function menuGarantirValorHora() {
  const sheet = getSheet('professores');
  garantirColunaValorHora(sheet);
  SpreadsheetApp.getUi().alert('✅ Coluna Valor_Hora verificada/criada na aba PROFESSORES.');
}

function mostrarStatusAPI() {
  const ui = SpreadsheetApp.getUi();
  const acoes = listarAcoes();
  ui.alert(
    '📊 STATUS DA API\n\n' +
    '✅ Web App ativo\n' +
    '📋 Ações disponíveis: ' + acoes.length + '\n\n' +
    acoes.join('\n')
  );
}

function mostrarURLWebApp() {
  const ui = SpreadsheetApp.getUi();
  const url = ScriptApp.getService().getUrl();
  ui.alert(
    '🔗 URL DO WEB APP\n\n' +
    url + '\n\n' +
    'Copie esta URL e cole em GAS_URL no HTML do sistema.\n\n' +
    '⚠️ Se a URL estiver vazia, você precisa implantar:\n' +
    'Implantar → Nova implantação → App da Web'
  );
}
