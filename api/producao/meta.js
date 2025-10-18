// api/producao/meta.js - Busca dados da aba Meta
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
    console.log('[PRODUCAO/META] Iniciando busca...');
    
    const doc = await sheetsService.init();
    console.log(`[PRODUCAO/META] Conectado: ${doc.title}`);
    
    const sheetMeta = doc.sheetsByTitle['Meta'];
    if (!sheetMeta) {
      console.warn('[PRODUCAO/META] Aba Meta não encontrada');
      return res.status(200).json({
        ok: true,
        dados: [],
        msg: 'Aba Meta não encontrada'
      });
    }
    
    const rows = await sheetMeta.getRows();
    console.log(`[PRODUCAO/META] ${rows.length} registros na Meta`);
    
    const dados = [];
    
    for (const row of rows) {
      const data = String(row.get('Data') || '').trim();
      const meta = String(row.get('Meta') || '0').trim();
      const produtividadeHora = String(row.get('Produtividade/hora') || '0').trim();
      
      if (data) {
        // Converte valores numéricos (remove vírgulas e converte para float)
        const metaNum = parseFloat(meta.replace(',', '.')) || 0;
        const prodHoraNum = parseFloat(produtividadeHora.replace(',', '.')) || 0;
        
        dados.push({
          data: data,
          meta: metaNum,
          produtividadeHora: prodHoraNum
        });
      }
    }
    
    console.log(`[PRODUCAO/META] ${dados.length} registros processados`);
    
    return res.status(200).json({
      ok: true,
      dados,
      total: dados.length,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('[PRODUCAO/META] Erro:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro ao buscar dados da Meta',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
