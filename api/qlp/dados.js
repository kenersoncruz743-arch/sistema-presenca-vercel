// api/qlp/dados.js - API atualizada para carregar dados da aba QLP
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
    console.log('[QLP/DADOS] Iniciando carregamento de dados da aba QLP...');
    
    const doc = await sheetsService.init();
    console.log(`[QLP/DADOS] Conectado à planilha: ${doc.title}`);
    
    // Busca a aba QLP
    const sheetQLP = doc.sheetsByTitle['QLP'];
    if (!sheetQLP) {
      console.error('[QLP/DADOS] Aba QLP não encontrada');
      return res.status(404).json({
        ok: false,
        msg: 'Aba QLP não encontrada na planilha. Verifique se a aba existe.'
      });
    }
    
    // Carrega os dados da aba QLP
    const rows = await sheetQLP.getRows();
    console.log(`[QLP/DADOS] ${rows.length} registros encontrados na aba QLP`);
    
    // Estrutura para organizar dados: Supervisor > Turno > Count
    const estrutura = {};
    const totalPorTurno = {
      'Turno A': 0,
      'Turno B': 0,
      'Turno C': 0,
      'Total': 0
    };
    
    rows.forEach(row => {
      const supervisor = String(row.get('Supervisor') || row.get('supervisor') || '').trim();
      const aba = String(row.get('Aba') || row.get('aba') || row.get('Grupo') || '').trim();
      const matricula = String(row.get('Matricula') || row.get('matricula') || '').trim();
      const nome = String(row.get('Nome') || row.get('nome') || '').trim();
      
      // Ignora linhas vazias
      if (!supervisor || !nome) return;
      
      // Determina o turno baseado na coluna Aba
      let turno = 'Turno C'; // Padrão
      if (aba.toLowerCase().includes('ta') || aba.toLowerCase().includes('turno a')) {
        turno = 'Turno A';
      } else if (aba.toLowerCase().includes('tb') || aba.toLowerCase().includes('turno b')) {
        turno = 'Turno B';
      } else if (aba.toLowerCase().includes('tc') || aba.toLowerCase().includes('turno c')) {
        turno = 'Turno C';
      }
      
      // Inicializa supervisor se não existe
      if (!estrutura[supervisor]) {
        estrutura[supervisor] = {
          'Turno A': new Set(),
          'Turno B': new Set(),
          'Turno C': new Set(),
          'Total': new Set()
        };
      }
      
      // Adiciona matrícula ao Set (garante contagem única)
      estrutura[supervisor][turno].add(matricula);
      estrutura[supervisor]['Total'].add(matricula);
    });
    
    // Converte Sets para contagens e calcula totais
    const dadosProcessados = {};
    
    Object.keys(estrutura).forEach(supervisor => {
      dadosProcessados[supervisor] = {
        'Turno A': estrutura[supervisor]['Turno A'].size,
        'Turno B': estrutura[supervisor]['Turno B'].size,
        'Turno C': estrutura[supervisor]['Turno C'].size,
        'Total': estrutura[supervisor]['Total'].size
      };
      
      // Atualiza totais gerais
      totalPorTurno['Turno A'] += dadosProcessados[supervisor]['Turno A'];
      totalPorTurno['Turno B'] += dadosProcessados[supervisor]['Turno B'];
      totalPorTurno['Turno C'] += dadosProcessados[supervisor]['Turno C'];
      totalPorTurno['Total'] += dadosProcessados[supervisor]['Total'];
    });
    
    console.log(`[QLP/DADOS] Dados processados: ${Object.keys(dadosProcessados).length} supervisores`);
    console.log(`[QLP/DADOS] Totais:`, totalPorTurno);
    
    return res.status(200).json({
      ok: true,
      dados: dadosProcessados,
      totais: totalPorTurno,
      totalSupervisores: Object.keys(dadosProcessados).length,
      totalRegistros: rows.length
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
