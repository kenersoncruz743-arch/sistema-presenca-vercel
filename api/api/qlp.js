// api/qlp.js
const sheets = require('../lib/sheets');

module.exports = async (req, res) => {
  // Habilita CORS se necessário
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  const { pathname } = new URL(req.url, `http://${req.headers.host}`);
  
  try {
    // Rota: GET /api/qlp/dados
    if (pathname === '/api/qlp/dados' && req.method === 'GET') {
      console.log('📊 Buscando dados do QLP...');
      const resultado = await sheets.getDadosQLP();
      
      if (resultado.ok) {
        console.log('✅ Dados QLP carregados com sucesso');
        res.status(200).json(resultado);
      } else {
        console.error('❌ Erro ao buscar dados QLP:', resultado.msg);
        res.status(500).json(resultado);
      }
      return;
    }

    // Rota: GET /api/qlp/exportar?formato=csv ou json
    if (pathname === '/api/qlp/exportar' && req.method === 'GET') {
      const url = new URL(req.url, `http://${req.headers.host}`);
      const formato = url.searchParams.get('formato') || 'json';
      
      console.log(`📥 Exportando QLP no formato: ${formato}`);
      const resultado = await sheets.exportarQLP(formato);
      
      if (resultado.ok) {
        console.log('✅ Exportação concluída');
        res.status(200).json(resultado);
      } else {
        console.error('❌ Erro na exportação:', resultado.msg);
        res.status(500).json(resultado);
      }
      return;
    }

    // Rota não encontrada
    res.status(404).json({ 
      ok: false, 
      msg: 'Rota não encontrada' 
    });

  } catch (error) {
    console.error('❌ Erro na API QLP:', error);
    res.status(500).json({ 
      ok: false, 
      msg: 'Erro interno do servidor: ' + error.message 
    });
  }
};
