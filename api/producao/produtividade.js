// api/producao/produtividade.js - Busca dados da aba Produtividade_Hora
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
    console.log('[PRODUCAO/PRODUTIVIDADE] Iniciando busca...');
    
    const doc = await sheetsService.init();
    console.log(`[PRODUCAO/PRODUTIVIDADE] Conectado: ${doc.title}`);
    
    // Carrega aba Produtividade_Hora
    const sheetProd = doc.sheetsByTitle['Produtividade_Hora'];
    if (!sheetProd) {
      return res.status(404).json({
        ok: false,
        msg: 'Aba Produtividade_Hora não encontrada'
      });
    }
    
    const rowsProd = await sheetProd.getRows();
    console.log(`[PRODUCAO/PRODUTIVIDADE] ${rowsProd.length} registros na Produtividade_Hora`);
    
    const dados = [];
    
    rowsProd.forEach(row => {
      const funcao = String(row.get('FUNCAO') || '').trim();
      const produtividadeHora = String(row.get('Produtividade/hora') || '').trim();
      
      if (funcao && produtividadeHora) {
        // Converte produtividade/hora para número (aceita vírgula ou ponto)
        const produtividadeNum = parseFloat(produtividadeHora.replace(',', '.')) || 0;
        
        dados.push({
          funcao: funcao,
          produtividadeHora: produtividadeNum
        });
      }
    });
    
    console.log(`[PRODUCAO/PRODUTIVIDADE] ${dados.length} registros processados`);
    console.log('[PRODUCAO/PRODUTIVIDADE] Funções encontradas:', dados.map(d => d.funcao));
    
    return res.status(200).json({
      ok: true,
      dados,
      total: dados.length,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('[PRODUCAO/PRODUTIVIDADE] Erro:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro ao buscar dados da Produtividade_Hora',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
