// api/producao/base.js - Busca dados da aba Base com cruzamento QLP
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
    console.log('[PRODUCAO/BASE] Iniciando busca...');
    
    const doc = await sheetsService.init();
    console.log(`[PRODUCAO/BASE] Conectado: ${doc.title}`);
    
    // Carrega Base
    const sheetBase = doc.sheetsByTitle['Base'];
    if (!sheetBase) {
      return res.status(404).json({
        ok: false,
        msg: 'Aba Base não encontrada'
      });
    }
    
    const rowsBase = await sheetBase.getRows();
    console.log(`[PRODUCAO/BASE] ${rowsBase.length} registros na Base`);
    
    // Carrega QLP para cruzamento (pegar seção de cada matrícula)
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
      console.log(`[PRODUCAO/BASE] ${Object.keys(mapaQLP).length} registros no mapa QLP`);
    }
    
    // Processa dados da Base
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
      
      // Filtra apenas registros de hoje
      if (data !== hoje) return;
      
      // Busca seção e turno do QLP
      let secao = 'Sem Seção';
      let turno = 'Não definido';
      
      if (mapaQLP[matricula]) {
        secao = mapaQLP[matricula].secao || secao;
        turno = mapaQLP[matricula].turno || turno;
      }
      
      // Determina turno pela aba se não tiver no QLP
      if (turno === 'Não definido' && aba) {
        const abaLower = aba.toLowerCase();
        if (abaLower.includes('ta') || abaLower.includes('turno a')) turno = 'Turno A';
        else if (abaLower.includes('tb') || abaLower.includes('turno b')) turno = 'Turno B';
        else if (abaLower.includes('tc') || abaLower.includes('turno c')) turno = 'Turno C';
      }
      
      dados.push({
        supervisor,
        aba,
        matricula,
        nome,
        funcao,
        status,
        data,
        secao,
        turno
      });
    });
    
    console.log(`[PRODUCAO/BASE] ${dados.length} registros processados (hoje)`);
    
    return res.status(200).json({
      ok: true,
      dados,
      total: dados.length,
      dataFiltro: hoje,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('[PRODUCAO/BASE] Erro:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro ao buscar dados da Base',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
