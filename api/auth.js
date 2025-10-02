// api/auth.js - VERSÃO FINAL CORRIGIDA
import sheetsService from '../lib/sheets';

export default async function handler(req, res) {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, msg: 'Método não permitido' });
  }

  try {
    const { usuario, senha, action } = req.body;

    if (!action) {
      return res.status(400).json({ ok: false, msg: 'Ação não especificada' });
    }

    switch (action) {
      case 'login':
        return await handleLogin(usuario, senha, res);
      case 'test':
        return await handleTest(res);
      case 'createTestData':
        return await handleCreateTestData(res);
      default:
        return res.status(400).json({
          ok: false,
          msg: `Ação não reconhecida: ${action}`,
          validActions: ['login', 'test', 'createTestData']
        });
    }
  } catch (error) {
    console.error('[AUTH] Erro geral na API:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro interno do servidor',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
}

// ---- Handlers separados ----

async function handleLogin(usuario, senha, res) {
  if (!usuario || !senha) {
    return res.status(400).json({ ok: false, msg: 'Usuário e senha são obrigatórios' });
  }

  console.log(`[AUTH] Tentativa de login: ${usuario}`);
  try {
    const result = await sheetsService.validarLogin(usuario, senha);
    console.log(`[AUTH] Resultado login:`, result);
    return res.status(200).json(result);
  } catch (err) {
    console.error('[AUTH] Erro login:', err);
    return res.status(500).json({ ok: false, msg: 'Erro ao validar credenciais' });
  }
}

async function handleTest(res) {
  console.log('[AUTH] Testando conexão...');
  try {
    const doc = await sheetsService.init();
    const sheets = Object.keys(doc.sheetsByTitle);
    return res.status(200).json({
      ok: true,
      msg: `Conectado à planilha: ${doc.title}`,
      sheets,
      totalSheets: sheets.length
    });
  } catch (err) {
    console.error('[AUTH] Erro no teste:', err);
    return res.status(500).json({ ok: false, msg: 'Erro ao conectar', details: err.message });
  }
}

async function handleCreateTestData(res) {
  console.log('[AUTH] Criando dados de teste...');
  try {
    const doc = await sheetsService.init();

    // --- Usuarios ---
    let usuariosSheet = doc.sheetsByTitle['Usuarios'];
    if (!usuariosSheet) {
      usuariosSheet = await doc.addSheet({
        title: 'Usuarios',
        headerValues: ['Usuario', 'Senha', 'Aba']
      });
    }

    const existingUsers = await usuariosSheet.getRows();
    if (existingUsers.length === 0) {
      await usuariosSheet.addRows([
        { Usuario: 'admin', Senha: '123', Aba: 'PCP_Gestão' },
        { Usuario: 'supervisor1', Senha: '456', Aba: 'WMS TA' },
        { Usuario: 'supervisor2', Senha: '789', Aba: 'WMS TB' }
      ]);
    }

    // --- Quadro ---
    let quadroSheet = doc.sheetsByTitle['Quadro'];
    if (!quadroSheet) {
      quadroSheet = await doc.addSheet({
        title: 'Quadro',
        headerValues: ['Coluna 1', 'NOME', 'FUNÇÃO NO RM', 'Função que atua', 'Coluna 2']
      });
    }

    const existingQuadro = await quadroSheet.getRows();
    if (existingQuadro.length === 0) {
      await quadroSheet.addRows([
        { 'Coluna 1': '001', NOME: 'João Silva', 'Função que atua': 'Operador' },
        { 'Coluna 1': '002', NOME: 'Maria Santos', 'Função que atua': 'Supervisora' }
      ]);
    }

    // --- Abas extras ---
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

    return res.status(200).json({ ok: true, msg: 'Dados de teste criados com sucesso!' });
  } catch (err) {
    console.error('[AUTH] Erro ao criar dados:', err);
    return res.status(500).json({ ok: false, msg: 'Erro ao criar dados', details: err.message });
  }
}
