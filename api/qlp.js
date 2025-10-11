// api/qlp.js
const sheets = require('../lib/sheets');

module.exports = async (req, res) => {
  console.log('🔍 [QLP API] Requisição recebida:', req.method, req.url);
  
  // Habilita CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Content-Type', 'application/json');
  
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }
  
  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    const pathname = url.pathname;
    console.log('📍 [QLP API] Pathname:', pathname);
    
    // Rota: GET /api/qlp/dados
    if (pathname === '/api/qlp/dados' && req.method === 'GET') {
      console.log('📊 [QLP] Buscando dados do QLP...');
      
      try {
        const resultado = await sheets.getDadosQLP();
        console.log('📦 [QLP] Resultado recebido:', resultado.ok ? 'OK' : 'ERRO');
        
        if (resultado.ok) {
          console.log('✅ [QLP] Dados carregados:', Object.keys(resultado.dados || {}).length, 'funções');
          return res.status(200).json(resultado);
        } else {
          console.error('❌ [QLP] Erro ao buscar dados:', resultado.msg);
          return res.status(500).json(resultado);
        }
      } catch (innerError) {
        console.error('❌ [QLP] Exceção ao buscar dados:', innerError);
        return res.status(500).json({ 
          ok: false, 
          msg: 'Erro ao buscar dados: ' + innerError.message 
        });
      }
    }
    
    // Rota: GET /api/qlp/exportar?formato=csv ou json
    if (pathname === '/api/qlp/exportar' && req.method === 'GET') {
      const formato = url.searchParams.get('formato') || 'json';
      console.log(`📥 [QLP] Exportando no formato: ${formato}`);
      
      try {
        const resultado = await sheets.exportarQLP(formato);
        
        if (resultado.ok) {
          console.log('✅ [QLP] Exportação concluída');
          return res.status(200).json(resultado);
        } else {
          console.error('❌ [QLP] Erro na exportação:', resultado.msg);
          return res.status(500).json(resultado);
        }
      } catch (innerError) {
        console.error('❌ [QLP] Exceção na exportação:', innerError);
        return res.status(500).json({ 
          ok: false, 
          msg: 'Erro ao exportar: ' + innerError.message 
        });
      }
    }
    
    // Rota não encontrada
    console.warn('⚠️ [QLP] Rota não encontrada:', pathname);
    return res.status(404).json({ 
      ok: false, 
      msg: 'Rota não encontrada: ' + pathname 
    });
    
  } catch (error) {
    console.error('❌ [QLP API] Erro geral:', error);
    return res.status(500).json({ 
      ok: false, 
      msg: 'Erro interno do servidor: ' + error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
