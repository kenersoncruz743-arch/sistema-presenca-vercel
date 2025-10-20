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
    
    // Carrega aba Meta
    const sheetMeta = doc.sheetsByTitle['Meta'];
    if (!sheetMeta) {
      return res.status(404).json({
        ok: false,
        msg: 'Aba Meta não encontrada'
      });
    }
    
    const rowsMeta = await sheetMeta.getRows();
    console.log(`[PRODUCAO/META] ${rowsMeta.length} registros na Meta`);
    
    const dados = [];
    
    rowsMeta.forEach(row => {
      const data = String(row.get('Data') || '').trim();
      const meta = String(row.get('Meta') || '').trim();
      const produtividadeHora = String(row.get('Produtividade/hora') || '').trim();
      
      if (data && produtividadeHora) {
        // Converte produtividade/hora para número (aceita vírgula ou ponto)
        const produtividadeNum = parseFloat(produtividadeHora.replace(',', '.')) || 0;
        const metaNum = parseFloat(meta.replace(',', '.')) || 0;
        
        dados.push({
          data: data,
          meta: metaNum,
          produtividadeHora: produtividadeNum
        });
      }
    });
    
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
