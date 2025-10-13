// api/qlp/dados.js - API para carregar dados do QLP
const sheetsService = require('../../lib/sheets');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'GET') {
    return res.status(405).json({ 
      ok: false, 
      msg: 'Método não permitido' 
    });
  }

  try {
    console.log('[QLP/DADOS] Iniciando carregamento de dados...');
    
    // Inicializa conexão com Google Sheets
    const doc = await sheetsService.init();
    console.log(`[QLP/DADOS] Conectado à planilha: ${doc.title}`);
    
    // Busca a aba Base (onde estão os registros salvos)
    const sheetBase = doc.sheetsByTitle['Base'];
    if (!sheetBase) {
      console.error('[QLP/DADOS] Aba Base não encontrada');
      return res.status(404).json({
        ok: false,
        msg: 'Aba Base não encontrada na planilha'
      });
    }
    
    // Busca a aba Quadro (informações dos colaboradores)
    const sheetQuadro = doc.sheetsByTitle['Quadro'];
    
    // Carrega todos os registros da Base
    const rows = await sheetBase.getRows();
    console.log(`[QLP/DADOS] ${rows.length} registros encontrados na Base`);
    
    // Cria um mapa de colaboradores do Quadro (para pegar tipo operacional)
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
      console.log(`[QLP/DADOS] ${Object.keys(quadroMap).length} colaboradores carregados do Quadro`);
    }
    
    // Estrutura de dados para organizar por Função > Supervisor > Colaboradores
    const estrutura = {};
    
    rows.forEach(row => {
      const supervisor = String(row.get('Supervisor') || '').trim();
      const aba = String(row.get('Aba') || '').trim();
      const matricula = String(row.get('Matricula') || '').trim();
      const nome = String(row.get('Nome') || '').trim();
      const funcao = String(row.get('Função') || 'Sem Função').trim();
      const status = String(row.get('Status') || '').trim();
      
      // Determina o turno baseado na Aba
      let turno = 'Não definido';
      if (aba.toLowerCase().includes('ta') || aba.toLowerCase().includes('turno a')) {
        turno = 'Turno A';
      } else if (aba.toLowerCase().includes('tb') || aba.toLowerCase().includes('turno b')) {
        turno = 'Turno B';
      } else if (aba.toLowerCase().includes('tc') || aba.toLowerCase().includes('turno c')) {
        turno = 'Turno C';
      }
      
      // Busca tipo operacional do colaborador
      const tipoOperacional = quadroMap[matricula] || 'Não operacional';
      
      // Ignora linhas sem nome
      if (!nome) return;
      
      // Cria estrutura: Função > Supervisor > Colaboradores
      if (!estrutura[funcao]) {
        estrutura[funcao] = {
          supervisores: {},
          totalContadores: {
            'Turno A': 0,
            'Turno B': 0,
            'Turno C': 0,
            'Situacao': 0
          }
        };
      }
      
      if (!estrutura[funcao].supervisores[supervisor]) {
        estrutura[funcao].supervisores[supervisor] = {
          colaboradores: [],
          totalContadores: {
            'Turno A': 0,
            'Turno B': 0,
            'Turno C': 0,
            'Situacao': 0
          }
        };
      }
      
      // Adiciona colaborador
      estrutura[funcao].supervisores[supervisor].colaboradores.push({
        matricula,
        nome,
        funcao,
        status,
        turno,
        tipoOperacional,
        aba
      });
      
      // Atualiza contadores
      if (turno === 'Turno A') {
        estrutura[funcao].totalContadores['Turno A']++;
        estrutura[funcao].supervisores[supervisor].totalContadores['Turno A']++;
      } else if (turno === 'Turno B') {
        estrutura[funcao].totalContadores['Turno B']++;
        estrutura[funcao].supervisores[supervisor].totalContadores['Turno B']++;
      } else if (turno === 'Turno C') {
        estrutura[funcao].totalContadores['Turno C']++;
        estrutura[funcao].supervisores[supervisor].totalContadores['Turno C']++;
      } else {
        estrutura[funcao].totalContadores['Situacao']++;
        estrutura[funcao].supervisores[supervisor].totalContadores['Situacao']++;
      }
    });
    
    console.log(`[QLP/DADOS] Dados estruturados: ${Object.keys(estrutura).length} funções`);
    
    return res.status(200).json({
      ok: true,
      dados: estrutura,
      totalRegistros: rows.length,
      totalFuncoes: Object.keys(estrutura).length
    });
    
  } catch (error) {
    console.error('[QLP/DADOS] Erro ao carregar dados:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro ao carregar dados do QLP',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
