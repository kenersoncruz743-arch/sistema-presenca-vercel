// api/index.js - API UNIFICADA COMPLETA
const sheetsService = require('../lib/sheets');

// ==================== FUNÇÃO PRINCIPAL DE ROTEAMENTO ====================
module.exports = async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    const { pathname } = new URL(req.url, `http://${req.headers.host}`);
    console.log('[API] Rota:', pathname, 'Método:', req.method);

    // ==================== AUTENTICAÇÃO ====================
    if (pathname === '/api' || pathname === '/api/' || pathname === '/api/auth') {
      return await handleAuth(req, res);
    }

    // ==================== COLABORADORES ====================
    if (pathname.startsWith('/api/colaboradores')) {
      return await handleColaboradores(req, res);
    }

    // ==================== ALOCAÇÃO BOX ====================
    if (pathname.startsWith('/api/alocacaobox')) {
      return await handleAlocacaoBox(req, res);
    }

    // ==================== MAPA DE CARGA ====================
    if (pathname.startsWith('/api/mapacarga')) {
      return await handleMapaCarga(req, res);
    }

    // ==================== PRODUÇÃO ====================
    if (pathname.startsWith('/api/producao')) {
      return await handleProducao(req, res, pathname);
    }

    // ==================== QLP ====================
    if (pathname.startsWith('/api/qlp')) {
      return await handleQLP(req, res, pathname);
    }

    // ==================== TESTE DE CONEXÃO ====================
    if (pathname.startsWith('/api/test')) {
      return await handleTestConnection(req, res);
    }

    // Rota não encontrada
    return res.status(404).json({
      ok: false,
      msg: 'Rota não encontrada',
      pathname: pathname
    });

  } catch (error) {
    console.error('[API] Erro geral:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro interno do servidor',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};

// ==================== HANDLER: AUTENTICAÇÃO ====================
async function handleAuth(req, res) {
  // Aceita tanto GET quanto POST para compatibilidade
  if (req.method === 'GET') {
    return res.status(200).json({
      ok: true,
      msg: 'API de autenticação funcionando',
      endpoints: {
        login: 'POST /api/auth com { action: "login", usuario, senha }',
        test: 'POST /api/auth com { action: "test" }',
        createTestData: 'POST /api/auth com { action: "createTestData" }'
      }
    });
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ 
      ok: false, 
      msg: 'Método não permitido. Use POST.' 
    });
  }

  try {
    const { usuario, senha, action } = req.body;

    console.log('[AUTH] Action:', action);
    console.log('[AUTH] Usuario:', usuario);

    if (!action) {
      return res.status(400).json({ 
        ok: false, 
        msg: 'Ação não especificada. Use action: "login", "test" ou "createTestData"' 
      });
    }

    // LOGIN
    if (action === 'login') {
      if (!usuario || !senha) {
        return res.status(400).json({ 
          ok: false, 
          msg: 'Usuário e senha são obrigatórios' 
        });
      }

      console.log(`[AUTH] Tentativa de login: ${usuario}`);
      const result = await sheetsService.validarLogin(usuario, senha);
      console.log('[AUTH] Resultado login:', result);
      
      return res.status(200).json(result);
    }

    // TESTE DE CONEXÃO
    if (action === 'test') {
      console.log('[AUTH] Testando conexão...');
      
      try {
        const doc = await sheetsService.init();
        const sheets = Object.keys(doc.sheetsByTitle);
        
        console.log('[AUTH] Conexão OK. Planilha:', doc.title);
        
        return res.status(200).json({
          ok: true,
          msg: `Conectado à planilha: ${doc.title}`,
          sheets: sheets,
          totalSheets: sheets.length
        });
      } catch (error) {
        console.error('[AUTH] Erro no teste:', error);
        return res.status(500).json({
          ok: false,
          msg: 'Erro ao conectar com Google Sheets',
          error: error.message
        });
      }
    }

    // CRIAR DADOS DE TESTE
    if (action === 'createTestData') {
      console.log('[AUTH] Criando dados de teste...');
      
      try {
        const doc = await sheetsService.init();
        
        // Aba Usuarios
        let sheet = doc.sheetsByTitle['Usuarios'];
        if (!sheet) {
          sheet = await doc.addSheet({ 
            title: 'Usuarios',
            headerValues: ['Usuario', 'Senha', 'Aba']
          });
        }
        
        const existingUsers = await sheet.getRows();
        if (existingUsers.length === 0) {
          await sheet.addRows([
            { Usuario: 'admin', Senha: '123', Aba: 'PCP_Gestão' },
            { Usuario: 'supervisor1', Senha: '456', Aba: 'WMS TA' },
            { Usuario: 'supervisor2', Senha: '789', Aba: 'WMS TB' },
            { Usuario: 'admin', Senha: '123', Aba: 'Separação TB' }
          ]);
        }
        
        // Aba Quadro
        sheet = doc.sheetsByTitle['Quadro'];
        if (!sheet) {
          sheet = await doc.addSheet({ 
            title: 'Quadro',
            headerValues: ['Coluna 1', 'NOME', 'FUNÇÃO NO RM', 'Função que atua', 'Coluna 2']
          });
        }
        
        const existingQuadro = await sheet.getRows();
        if (existingQuadro.length === 0) {
          await sheet.addRows([
            { 'Coluna 1': '001', 'NOME': 'João Silva', 'Função que atua': 'Operador' },
            { 'Coluna 1': '002', 'NOME': 'Maria Santos', 'Função que atua': 'Supervisora' },
            { 'Coluna 1': '003', 'NOME': 'Pedro Costa', 'Função que atua': 'Operador' }
          ]);
        }
        
        if (!doc.sheetsByTitle['Lista']) {
          await doc.addSheet({ 
            title: 'Lista',
            headerValues: ['Supervisor', 'Grupo', 'matricula', 'Nome', 'Função', 'status']
          });
        }
        
        if (!doc.sheetsByTitle['Base']) {
          await doc.addSheet({ 
            title: 'Base',
            headerValues: ['Supervisor', 'Aba', 'Matricula', 'Nome', 'Função', 'Status', 'Data']
          });
        }
        
        console.log('[AUTH] Dados de teste criados com sucesso');
        
        return res.status(200).json({ 
          ok: true, 
          msg: 'Dados de teste criados com sucesso!' 
        });
        
      } catch (error) {
        console.error('[AUTH] Erro ao criar dados de teste:', error);
        return res.status(500).json({
          ok: false,
          msg: 'Erro ao criar dados de teste',
          error: error.message
        });
      }
    }

    return res.status(400).json({ 
      ok: false,
      msg: 'Ação não reconhecida: ' + action,
      validActions: ['login', 'test', 'createTestData']
    });

  } catch (error) {
    console.error('[AUTH] Erro:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro no processamento da autenticação',
      error: error.message
    });
  }
}

// ==================== HANDLER: COLABORADORES ====================
async function handleColaboradores(req, res) {
  // GET - Buscar colaboradores
  if (req.method === 'GET') {
    const { filtro } = req.query;
    console.log('[COLABORADORES] Buscando com filtro:', filtro);
    
    const colaboradores = await sheetsService.buscarColaboradores(filtro || '');
    return res.status(200).json(colaboradores);
  }

  // POST - Ações diversas
  if (req.method === 'POST') {
    const { action } = req.body;
    console.log('[COLABORADORES] Action:', action);

    switch (action) {
      case 'addBuffer': {
        const { supervisor, aba, colaborador } = req.body;
        
        if (!supervisor || !aba || !colaborador || !colaborador.matricula || !colaborador.nome) {
          return res.status(400).json({ ok: false, msg: 'Parâmetros incompletos' });
        }
        
        const result = await sheetsService.adicionarBuffer(supervisor, aba, colaborador);
        return res.status(200).json(result);
      }

      case 'getBuffer': {
        const { supervisor, aba } = req.body;
        
        if (!supervisor || !aba) {
          return res.status(400).json({ ok: false, msg: 'Supervisor e aba são obrigatórios' });
        }
        
        const buffer = await sheetsService.getBuffer(supervisor, aba);
        return res.status(200).json(buffer);
      }

      case 'removeBuffer': {
        const { supervisor, aba, matricula } = req.body;
        
        if (!supervisor || !matricula) {
          return res.status(400).json({ ok: false, msg: 'Supervisor e matrícula são obrigatórios' });
        }
        
        const chave = aba || supervisor;
        const result = await sheetsService.removerBufferPorAba(chave, matricula);
        return res.status(200).json(result);
      }

      case 'updateStatus': {
        const { supervisor, aba, matricula, status } = req.body;
        
        if (!supervisor || !matricula || status === undefined) {
          return res.status(400).json({ ok: false, msg: 'Parâmetros incompletos' });
        }
        
        const chave = aba || supervisor;
        const result = await sheetsService.atualizarStatusBufferPorAba(chave, matricula, status);
        return res.status(200).json(result);
      }

      case 'updateDesvio': {
        const { supervisor, aba, matricula, desvio } = req.body;
        
        if (!supervisor || !matricula) {
          return res.status(400).json({ ok: false, msg: 'Parâmetros incompletos' });
        }
        
        const chave = aba || supervisor;
        const result = await sheetsService.atualizarDesvioBufferPorAba(chave, matricula, desvio);
        return res.status(200).json(result);
      }

      case 'saveToBase': {
        const { dados } = req.body;
        
        if (!dados || !Array.isArray(dados) || dados.length === 0) {
          return res.status(400).json({ ok: false, msg: 'Dados inválidos' });
        }
        
        const result = await sheetsService.salvarNaBase(dados);
        return res.status(200).json(result);
      }

      default:
        return res.status(400).json({ 
          ok: false, 
          msg: 'Ação não reconhecida: ' + action
        });
    }
  }

  return res.status(405).json({ ok: false, msg: 'Método não permitido' });
}

// ==================== HANDLER: ALOCAÇÃO BOX ====================
async function handleAlocacaoBox(req, res) {
  const { action, filtros, boxNum, cargaId, dados } = req.body;

  if (!action) {
    return res.status(400).json({ ok: false, msg: 'Action é obrigatória' });
  }

  console.log('[ALOCACAO] Action:', action);

  switch (action) {
    case 'listarCargas': {
      const cargas = await sheetsService.getCargasSemBox(filtros || {});
      return res.status(200).json({ ok: true, cargas, total: cargas.length });
    }

    case 'listarBoxes': {
      const boxes = await sheetsService.getEstadoBoxes();
      return res.status(200).json({ ok: true, boxes, total: boxes.length });
    }

    case 'alocarBox': {
      if (!boxNum || !cargaId) {
        return res.status(400).json({ ok: false, msg: 'boxNum e cargaId são obrigatórios' });
      }
      
      const resultado = await sheetsService.alocarCargaBox(boxNum, cargaId);
      return res.status(200).json(resultado);
    }

    case 'liberarBox': {
      if (!boxNum) {
        return res.status(400).json({ ok: false, msg: 'boxNum é obrigatório' });
      }
      
      const resultado = await sheetsService.liberarBox(boxNum);
      return res.status(200).json(resultado);
    }

    case 'salvarAlocacoes': {
      if (!dados || !Array.isArray(dados)) {
        return res.status(400).json({ ok: false, msg: 'dados deve ser um array' });
      }
      
      const resultado = await sheetsService.salvarAlocacoes(dados);
      return res.status(200).json(resultado);
    }

    default:
      return res.status(400).json({ ok: false, msg: 'Ação inválida: ' + action });
  }
}

// ==================== HANDLER: MAPA DE CARGA ====================
async function handleMapaCarga(req, res) {
  const { action, filtros } = req.body;

  switch (action) {
    case 'listar':
      const dados = await sheetsService.getMapaCarga(filtros || {});
      return res.status(200).json({ ok: true, dados });

    case 'atualizar':
      // Implementar lógica de atualização
      return res.status(200).json({ ok: true, msg: 'Atualizado com sucesso' });

    default:
      return res.status(400).json({ ok: false, msg: 'Ação inválida' });
  }
}

// ==================== HANDLER: PRODUÇÃO ====================
async function handleProducao(req, res, pathname) {
  // /api/producao/base
  if (pathname === '/api/producao/base') {
    if (req.method !== 'GET') {
      return res.status(405).json({ ok: false, msg: 'Método não permitido' });
    }

    const doc = await sheetsService.init();
    const sheetBase = doc.sheetsByTitle['Base'];
    
    if (!sheetBase) {
      return res.status(404).json({ ok: false, msg: 'Aba Base não encontrada' });
    }
    
    const rowsBase = await sheetBase.getRows();
    const sheetQLP = doc.sheetsByTitle['QLP'];
    const mapaQLP = {};
    
    if (sheetQLP) {
      const rowsQLP = await sheetQLP.getRows();
      rowsQLP.forEach(row => {
        const chapa = String(row.get('CHAPA1') || '').trim();
        const secao = String(row.get('SECAO') || '').trim();
        const turno = String(row.get('Turno') || '').trim();
        
        if (chapa) {
          mapaQLP[chapa] = { secao, turno };
        }
      });
    }
    
    const dados = [];
    const hoje = new Date().toLocaleDateString('pt-BR');
    
    rowsBase.forEach(row => {
      const supervisor = String(row.get('Supervisor') || '').trim();
      const aba = String(row.get('Aba') || '').trim();
      const matricula = String(row.get('Matricula') || '').trim();
      const nome = String(row.get('Nome') || '').trim();
      const funcao = String(row.get('Função') || '').trim();
      const status = String(row.get('Status') || '').trim();
      const data = String(row.get('Data') || '').trim();
      
      if (data !== hoje) return;
      
      let secao = 'Sem Seção';
      let turno = 'Não definido';
      
      if (mapaQLP[matricula]) {
        secao = mapaQLP[matricula].secao || secao;
        turno = mapaQLP[matricula].turno || turno;
      }
      
      if (turno === 'Não definido' && aba) {
        const abaLower = aba.toLowerCase();
        if (abaLower.includes('ta')) turno = 'Turno A';
        else if (abaLower.includes('tb')) turno = 'Turno B';
        else if (abaLower.includes('tc')) turno = 'Turno C';
      }
      
      dados.push({ supervisor, aba, matricula, nome, funcao, status, data, secao, turno });
    });
    
    return res.status(200).json({
      ok: true,
      dados,
      total: dados.length,
      dataFiltro: hoje,
      timestamp: new Date().toISOString()
    });
  }

  // /api/producao/meta
  if (pathname === '/api/producao/meta') {
    if (req.method !== 'GET') {
      return res.status(405).json({ ok: false, msg: 'Método não permitido' });
    }

    const doc = await sheetsService.init();
    const sheetMeta = doc.sheetsByTitle['Meta'];
    
    if (!sheetMeta) {
      return res.status(404).json({ ok: false, msg: 'Aba Meta não encontrada' });
    }
    
    const rowsMeta = await sheetMeta.getRows();
    const dados = [];
    
    rowsMeta.forEach(row => {
      const data = String(row.get('Data') || '').trim();
      const meta = String(row.get('Meta') || '').trim();
      const produtividadeHora = String(row.get('Produtividade/hora') || '').trim();
      
      if (data && produtividadeHora) {
        dados.push({
          data,
          meta: parseFloat(meta.replace(',', '.')) || 0,
          produtividadeHora: parseFloat(produtividadeHora.replace(',', '.')) || 0
        });
      }
    });
    
    return res.status(200).json({
      ok: true,
      dados,
      total: dados.length,
      timestamp: new Date().toISOString()
    });
  }

  // /api/producao/produtividade
  if (pathname === '/api/producao/produtividade') {
    if (req.method !== 'GET') {
      return res.status(405).json({ ok: false, msg: 'Método não permitido' });
    }

    const doc = await sheetsService.init();
    const sheetProd = doc.sheetsByTitle['Produtividade_Hora'];
    
    if (!sheetProd) {
      return res.status(404).json({ ok: false, msg: 'Aba Produtividade_Hora não encontrada' });
    }
    
    const rowsProd = await sheetProd.getRows();
    const dados = [];
    
    rowsProd.forEach(row => {
      const funcao = String(row.get('FUNCAO') || '').trim();
      const produtividadeHora = String(row.get('Produtividade/hora') || '').trim();
      
      if (funcao && produtividadeHora) {
        dados.push({
          funcao,
          produtividadeHora: parseFloat(produtividadeHora.replace(',', '.')) || 0
        });
      }
    });
    
    return res.status(200).json({
      ok: true,
      dados,
      total: dados.length,
      timestamp: new Date().toISOString()
    });
  }

  // /api/producao/resumo-base
  if (pathname === '/api/producao/resumo-base') {
    if (req.method !== 'GET') {
      return res.status(405).json({ ok: false, msg: 'Método não permitido' });
    }

    const doc = await sheetsService.init();
    const sheetBase = doc.sheetsByTitle['Base'];
    
    if (!sheetBase) {
      return res.status(404).json({ ok: false, msg: 'Aba Base não encontrada' });
    }
    
    const rows = await sheetBase.getRows();
    
    let dataFiltro;
    if (req.query.data) {
      const [ano, mes, dia] = req.query.data.split('-');
      dataFiltro = `${dia}/${mes}/${ano}`;
    } else {
      dataFiltro = new Date().toLocaleDateString('pt-BR');
    }
    
    const resumoPorSupervisor = {};
    const resumoPorFuncao = {};
    const resumoGeral = {
      total: 0, presente: 0, ausente: 0, atestado: 0,
      ferias: 0, folga: 0, afastado: 0, desvio: 0, outros: 0
    };
    
    rows.forEach(row => {
      const dataRegistro = String(row.get('Data') || '').trim();
      if (dataRegistro !== dataFiltro) return;
      
      const supervisor = String(row.get('Supervisor') || 'Sem supervisor').trim();
      const aba = String(row.get('Aba') || '').trim();
      const funcao = String(row.get('Função') || 'Não informada').trim();
      const turno = String(row.get('Turno') || 'Não informado').trim();
      const status = String(row.get('Status') || 'Outro').trim();
      const desvio = String(row.get('Desvio') || '').trim();
      const nome = String(row.get('Nome') || '').trim();
      const matricula = String(row.get('Matricula') || '').trim();
      
      if (!nome) return;
      
      // Resumo por supervisor
      if (!resumoPorSupervisor[supervisor]) {
        resumoPorSupervisor[supervisor] = {
          supervisor, total: 0, presente: 0, ausente: 0, atestado: 0,
          ferias: 0, folga: 0, afastado: 0, desvio: 0, outros: 0,
          porFuncao: {}, colaboradores: []
        };
      }
      
      resumoPorSupervisor[supervisor].total++;
      resumoPorSupervisor[supervisor].colaboradores.push({
        nome, matricula, funcao, turno, status, desvio
      });
      
      const statusLower = status.toLowerCase();
      if (statusLower === 'presente') resumoPorSupervisor[supervisor].presente++;
      else if (statusLower === 'ausente') resumoPorSupervisor[supervisor].ausente++;
      else if (statusLower === 'atestado') resumoPorSupervisor[supervisor].atestado++;
      else if (statusLower.includes('férias')) resumoPorSupervisor[supervisor].ferias++;
      else if (statusLower === 'folga') resumoPorSupervisor[supervisor].folga++;
      else if (statusLower === 'afastado') resumoPorSupervisor[supervisor].afastado++;
      else resumoPorSupervisor[supervisor].outros++;
      
      if (desvio && desvio.toLowerCase() === 'desvio') {
        resumoPorSupervisor[supervisor].desvio++;
      }
      
      // Resumo por função
      if (!resumoPorFuncao[funcao]) {
        resumoPorFuncao[funcao] = {
          funcao, total: 0, presente: 0, ausente: 0, atestado: 0,
          ferias: 0, folga: 0, afastado: 0, desvio: 0, outros: 0,
          porSupervisor: {}, colaboradores: []
        };
      }
      
      resumoPorFuncao[funcao].total++;
      resumoPorFuncao[funcao].colaboradores.push({
        nome, matricula, supervisor, turno, status, desvio
      });
      
      if (statusLower === 'presente') resumoPorFuncao[funcao].presente++;
      else if (statusLower === 'ausente') resumoPorFuncao[funcao].ausente++;
      else if (statusLower === 'atestado') resumoPorFuncao[funcao].atestado++;
      else if (statusLower.includes('férias')) resumoPorFuncao[funcao].ferias++;
      else if (statusLower === 'folga') resumoPorFuncao[funcao].folga++;
      else if (statusLower === 'afastado') resumoPorFuncao[funcao].afastado++;
      else resumoPorFuncao[funcao].outros++;
      
      if (desvio && desvio.toLowerCase() === 'desvio') {
        resumoPorFuncao[funcao].desvio++;
      }
      
      // Resumo geral
      resumoGeral.total++;
      if (statusLower === 'presente') resumoGeral.presente++;
      else if (statusLower === 'ausente') resumoGeral.ausente++;
      else if (statusLower === 'atestado') resumoGeral.atestado++;
      else if (statusLower.includes('férias')) resumoGeral.ferias++;
      else if (statusLower === 'folga') resumoGeral.folga++;
      else if (statusLower === 'afastado') resumoGeral.afastado++;
      else resumoGeral.outros++;
      
      if (desvio && desvio.toLowerCase() === 'desvio') {
        resumoGeral.desvio++;
      }
    });
    
    const supervisores = Object.values(resumoPorSupervisor)
      .sort((a, b) => a.supervisor.localeCompare(b.supervisor));
    
    const funcoes = Object.values(resumoPorFuncao)
      .sort((a, b) => a.funcao.localeCompare(b.funcao));
    
    if (resumoGeral.total > 0) {
      resumoGeral.percentualPresente = ((resumoGeral.presente / resumoGeral.total) * 100).toFixed(1);
      resumoGeral.percentualAusente = (((resumoGeral.total - resumoGeral.presente) / resumoGeral.total) * 100).toFixed(1);
      resumoGeral.percentualDesvio = ((resumoGeral.desvio / resumoGeral.total) * 100).toFixed(1);
    }
    
    return res.status(200).json({
      ok: true,
      dataReferencia: dataFiltro,
      resumoGeral,
      porSupervisor: supervisores,
      porFuncao: funcoes,
      totais: {
        supervisores: supervisores.length,
        funcoes: funcoes.length,
        colaboradores: resumoGeral.total
      },
      timestamp: new Date().toISOString()
    });
  }

  return res.status(404).json({ ok: false, msg: 'Rota de produção não encontrada' });
}

// ==================== HANDLER: QLP ====================
async function handleQLP(req, res, pathname) {
  if (req.method !== 'GET') {
    return res.status(405).json({ ok: false, msg: 'Método não permitido' });
  }

  // /api/qlp/quadro ou /api/qlp/dados
  if (pathname === '/api/qlp/quadro' || pathname === '/api/qlp/dados') {
    const doc = await sheetsService.init();
    const sheetQLP = doc.sheetsByTitle['QLP'];
    
    if (!sheetQLP) {
      return res.status(404).json({
        ok: false,
        msg: 'Aba QLP não encontrada',
        abasDisponiveis: Object.keys(doc.sheetsByTitle)
      });
    }
    
    await sheetQLP.loadHeaderRow();
    const headers = sheetQLP.headerValues;
    const rows = await sheetQLP.getRows();
    
    if (rows.length === 0) {
      return res.status(200).json({
        ok: true,
        colaboradores: [],
        estatisticas: {
          total: 0, ativos: 0, inativos: 0,
          porSecao: {}, porTurno: {}, porSupervisor: {},
          porSituacao: {}, porFuncao: {},
          totalSecoes: 0, totalTurnos: 0, totalSupervisores: 0, totalFuncoes: 0
        },
        timestamp: new Date().toISOString()
      });
    }
    
    const colaboradores = [];
    const estatisticas = {
      total: 0, ativos: 0, inativos: 0,
      porSecao: {}, porTurno: {}, porSupervisor: {},
      porSituacao: {}, porFuncao: {}
    };
    
    rows.forEach(row => {
      const getCol = (colName) => {
        try {
          const valor = row.get(colName);
          return valor ? String(valor).trim() : '';
        } catch (e) {
          return '';
        }
      };
      
      const nome = getCol('NOME');
      if (!nome) return;
      
      const chapa = getCol('CHAPA1') || 'S/N';
      const funcao = getCol('FUNCAO');
      const secao = getCol('SECAO');
      const situacao = getCol('SITUACAO');
      const supervisor = getCol('Supervisor') || 'Sem supervisor';
      const turno = getCol('Turno') || 'Não definido';
      
      const situacaoLower = situacao.toLowerCase();
      const isAtivo = situacao === 'Ativo' || 
                     (situacaoLower.includes('ativo') && 
                      !situacaoLower.includes('não') && 
                      !situacaoLower.includes('af.'));
      
      colaboradores.push({
        filial: getCol('FILIAL'),
        bandeira: getCol('BANDEIRA'),
        chapa,
        dtAdmissao: getCol('DT_ADMISSAO'),
        nome,
        funcao,
        secao,
        situacao,
        totalGeral: getCol('Total Geral'),
        supervisor,
        turno,
        gestao: getCol('Gestão'),
        isAtivo
      });
      
      estatisticas.total++;
      if (isAtivo) estatisticas.ativos++;
      else estatisticas.inativos++;
      
      // Estatísticas por categoria
      if (secao) {
        if (!estatisticas.porSecao[secao]) {
          estatisticas.porSecao[secao] = { total: 0, ativos: 0, inativos: 0 };
        }
        estatisticas.porSecao[secao].total++;
        if (isAtivo) estatisticas.porSecao[secao].ativos++;
        else estatisticas.porSecao[secao].inativos++;
      }
      
      if (!estatisticas.porTurno[turno]) {
        estatisticas.porTurno[turno] = { total: 0, ativos: 0 };
      }
      estatisticas.porTurno[turno].total++;
      if (isAtivo) estatisticas.porTurno[turno].ativos++;
      
      if (!estatisticas.porSupervisor[supervisor]) {
        estatisticas.porSupervisor[supervisor] = { total: 0, ativos: 0 };
      }
      estatisticas.porSupervisor[supervisor].total++;
      if (isAtivo) estatisticas.porSupervisor[supervisor].ativos++;
      
      const situacaoNormalizada = situacao || 'Sem situação';
      if (!estatisticas.porSituacao[situacaoNormalizada]) {
        estatisticas.porSituacao[situacaoNormalizada] = 0;
      }
      estatisticas.porSituacao[situacaoNormalizada]++;
      
      if (funcao) {
        if (!estatisticas.porFuncao[funcao]) {
          estatisticas.porFuncao[funcao] = { total: 0, ativos: 0 };
        }
        estatisticas.porFuncao[funcao].total++;
        if (isAtivo) estatisticas.porFuncao[funcao].ativos++;
      }
    });
    
    estatisticas.totalSecoes = Object.keys(estatisticas.porSecao).length;
    estatisticas.totalTurnos = Object.keys(estatisticas.porTurno).length;
    estatisticas.totalSupervisores = Object.keys(estatisticas.porSupervisor).length;
    estatisticas.totalFuncoes = Object.keys(estatisticas.porFuncao).length;
    
    return res.status(200).json({
      ok: true,
      colaboradores,
      estatisticas,
      timestamp: new Date().toISOString()
    });
  }

  // /api/qlp/exportar
  if (pathname === '/api/qlp/exportar') {
    const { formato } = req.query;
    
    if (!formato || !['csv', 'json'].includes(formato)) {
      return res.status(400).json({
        ok: false,
        msg: 'Formato inválido. Use csv ou json'
      });
    }
    
    console.log(`[QLP/EXPORTAR] Iniciando exportação em formato: ${formato}`);
    
    const doc = await sheetsService.init();
    const sheetBase = doc.sheetsByTitle['Base'];
    
    if (!sheetBase) {
      return res.status(404).json({
        ok: false,
        msg: 'Aba Base não encontrada'
      });
    }
    
    const sheetQuadro = doc.sheetsByTitle['Quadro'];
    const rows = await sheetBase.getRows();
    
    const quadroMap = {};
    if (sheetQuadro) {
      const quadroRows = await sheetQuadro.getRows();
      quadroRows.forEach(row => {
        const matricula = String(row.get('Coluna 1') || '').trim();
        const tipoOperacional = String(row.get('Coluna 2') || '').trim();
        if (matricula) {
          quadroMap[matricula] = tipoOperacional || 'Não operacional';
        }
      });
    }
    
    const dadosExportacao = [];
    
    rows.forEach(row => {
      const supervisor = String(row.get('Supervisor') || '').trim();
      const aba = String(row.get('Aba') || '').trim();
      const matricula = String(row.get('Matricula') || '').trim();
      const nome = String(row.get('Nome') || '').trim();
      const funcao = String(row.get('Função') || '').trim();
      const status = String(row.get('Status') || '').trim();
      const data = String(row.get('Data') || '').trim();
      
      if (!nome) return;
      
      let turno = 'Não definido';
      if (aba.toLowerCase().includes('ta') || aba.toLowerCase().includes('turno a')) {
        turno = 'Turno A';
      } else if (aba.toLowerCase().includes('tb') || aba.toLowerCase().includes('turno b')) {
        turno = 'Turno B';
      } else if (aba.toLowerCase().includes('tc') || aba.toLowerCase().includes('turno c')) {
        turno = 'Turno C';
      }
      
      const tipoOperacional = quadroMap[matricula] || 'Não operacional';
      
      dadosExportacao.push({
        Matricula: matricula,
        Nome: nome,
        Funcao: funcao,
        Supervisor: supervisor,
        Aba: aba,
        Turno: turno,
        TipoOperacional: tipoOperacional,
        Status: status,
        Data: data
      });
    });
    
    if (formato === 'csv') {
      const headers = ['Matricula', 'Nome', 'Funcao', 'Supervisor', 'Aba', 'Turno', 'TipoOperacional', 'Status', 'Data'];
      let csv = headers.join(',') + '\n';
      
      dadosExportacao.forEach(item => {
        const linha = headers.map(header => {
          let valor = item[header] || '';
          if (valor.includes(',') || valor.includes('"')) {
            valor = '"' + valor.replace(/"/g, '""') + '"';
          }
          return valor;
        }).join(',');
        csv += linha + '\n';
      });
      
      return res.status(200).json({
        ok: true,
        formato: 'csv',
        csv: csv,
        totalRegistros: dadosExportacao.length
      });
      
    } else {
      return res.status(200).json({
        ok: true,
        formato: 'json',
        dados: dadosExportacao,
        totalRegistros: dadosExportacao.length
      });
    }
  }

  return res.status(404).json({ ok: false, msg: 'Rota QLP não encontrada' });
}

// ==================== HANDLER: TESTE DE CONEXÃO ====================
async function handleTestConnection(req, res) {
  try {
    console.log('[TEST] Iniciando teste de conexão...');
    
    const doc = await sheetsService.init();
    console.log(`[TEST] ✓ Conectado: ${doc.title}`);
    
    const abas = Object.keys(doc.sheetsByTitle);
    console.log('[TEST] Abas encontradas:', abas);
    
    const sheetQuadro = doc.sheetsByTitle['Quadro'];
    
    if (!sheetQuadro) {
      return res.status(200).json({
        ok: false,
        msg: 'Aba Quadro não encontrada',
        planilha: doc.title,
        abasDisponiveis: abas
      });
    }
    
    await sheetQuadro.loadHeaderRow();
    const headers = sheetQuadro.headerValues;
    
    const rows = await sheetQuadro.getRows({ limit: 3 });
    
    const amostras = rows.map((row) => {
      const dados = {};
      headers.forEach(header => {
        try {
          dados[header] = row.get(header) || '';
        } catch (e) {
          dados[header] = 'ERRO';
        }
      });
      return dados;
    });
    
    console.log('[TEST] ✓ Headers:', headers);
    console.log('[TEST] ✓ Total de linhas:', sheetQuadro.rowCount);
    
    return res.status(200).json({
      ok: true,
      planilha: doc.title,
      aba: 'Quadro',
      totalAbas: abas.length,
      abasDisponiveis: abas,
      totalLinhas: sheetQuadro.rowCount,
      totalColunas: headers.length,
      headers: headers,
      amostrasDados: amostras,
      verificacoes: {
        temCHAPA1: headers.includes('CHAPA1'),
        temNOME: headers.includes('NOME'),
        temTurno: headers.includes('Turno'),
        temSupervisor: headers.includes('Supervisor'),
        temSITUACAO: headers.includes('SITUACAO'),
        temSECAO: headers.includes('SECAO')
      }
    });
    
  } catch (error) {
    console.error('[TEST] Erro:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro no teste de conexão',
      error: error.message,
      stack: error.stack
    });
  }
}
