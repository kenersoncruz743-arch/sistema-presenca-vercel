// api/producao/base.js - 
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
    console.log('[PRODUCAO/BASE] ========== INÍCIO ==========');
    console.log('[PRODUCAO/BASE] Query params:', req.query);
    
    const doc = await sheetsService.init();
    console.log(`[PRODUCAO/BASE] ✓ Conectado: ${doc.title}`);
    
    // Carrega Base
    const sheetBase = doc.sheetsByTitle['Base'];
    if (!sheetBase) {
      return res.status(404).json({
        ok: false,
        msg: 'Aba Base não encontrada'
      });
    }
    
    const rowsBase = await sheetBase.getRows();
    console.log(`[PRODUCAO/BASE] ✓ ${rowsBase.length} registros na Base`);
    
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
      console.log(`[PRODUCAO/BASE] ✓ ${Object.keys(mapaQLP).length} registros no mapa QLP`);
    }
    
    // ===== CORREÇÃO DA DATA =====
    let dataFiltro;
    
    if (req.query.data) {
      // Se veio data do frontend (formato YYYY-MM-DD)
      const [ano, mes, dia] = req.query.data.split('-');
      dataFiltro = `${dia}/${mes}/${ano}`;
      console.log(`[PRODUCAO/BASE] Data do query: ${req.query.data} -> ${dataFiltro}`);
    } else {
      // Usa data de hoje
      const hoje = new Date();
      const dia = String(hoje.getDate()).padStart(2, '0');
      const mes = String(hoje.getMonth() + 1).padStart(2, '0');
      const ano = hoje.getFullYear();
      dataFiltro = `${dia}/${mes}/${ano}`;
      console.log(`[PRODUCAO/BASE] Data de hoje: ${dataFiltro}`);
    }
    
    console.log(`[PRODUCAO/BASE] Filtrando por data: "${dataFiltro}"`);
    
    // Processa dados da Base
    const dados = [];
    let registrosFiltrados = 0;
    let registrosIgnorados = 0;
    
    rowsBase.forEach((row, index) => {
      const supervisor = String(row.get('Supervisor') || '').trim();
      const aba = String(row.get('Aba') || '').trim();
      const matricula = String(row.get('Matricula') || '').trim();
      const nome = String(row.get('Nome') || '').trim();
      const funcao = String(row.get('Função') || '').trim();
      const status = String(row.get('Status') || '').trim();
      const data = String(row.get('Data') || '').trim();
      
      // Debug das primeiras 5 datas
      if (index < 5) {
        console.log(`[PRODUCAO/BASE] Linha ${index + 1} - Data: "${data}" | Match: ${data === dataFiltro}`);
      }
      
      // COMPARAÇÃO EXATA DE STRINGS
      if (data !== dataFiltro) {
        registrosIgnorados++;
        return;
      }
      
      registrosFiltrados++;
      
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
    
    console.log('[PRODUCAO/BASE] ========== ESTATÍSTICAS ==========');
    console.log(`[PRODUCAO/BASE] Data filtrada: ${dataFiltro}`);
    console.log(`[PRODUCAO/BASE] Registros processados: ${registrosFiltrados}`);
    console.log(`[PRODUCAO/BASE] Registros ignorados: ${registrosIgnorados}`);
    console.log(`[PRODUCAO/BASE] Total de linhas: ${rowsBase.length}`);
    console.log('[PRODUCAO/BASE] ========== FIM ==========');
    
    return res.status(200).json({
      ok: true,
      dados,
      total: dados.length,
      dataFiltro: dataFiltro,
      debug: {
        totalLinhas: rowsBase.length,
        registrosFiltrados,
        registrosIgnorados
      },
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('[PRODUCAO/BASE] ========== ERRO FATAL ==========');
    console.error('[PRODUCAO/BASE] Tipo:', error.name);
    console.error('[PRODUCAO/BASE] Mensagem:', error.message);
    console.error('[PRODUCAO/BASE] Stack:', error.stack);
    console.error('[PRODUCAO/BASE] =====================================');
    
    return res.status(500).json({
      ok: false,
      msg: 'Erro ao buscar dados da Base',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
