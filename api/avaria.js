const sheetsAvaria = require('../lib/sheets_3');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') return res.status(200).end();
  
  console.log('[API AVARIA] Request:', { 
    method: req.method, 
    action: req.body?.action || req.query?.action 
  });
  
  try {
    const action = req.method === 'POST' ? req.body?.action : req.query?.action;
    
    if (!action) {
      return res.status(400).json({ 
        ok: false, 
        msg: 'Action é obrigatória', 
        acoesDisponiveis: [
          'obterDadosProdutos', 
          'salvarRegistro', 
          'obterHistorico', 
          'testarConexao'
        ] 
      });
    }
    
    switch (action) {
      case 'obterDadosProdutos': {
        const dados = await sheetsAvaria.obterDadosProdutos();
        return res.status(200).json({ 
          ok: true, 
          dados, 
          total: dados.length 
        });
      }
      
      case 'salvarRegistro': {
        const { usuarioLogado, codigoProduto, descricaoProduto, embalagemProduto, quantidade } = req.body;
        
        if (!usuarioLogado || !codigoProduto || !quantidade) {
          return res.status(400).json({ 
            ok: false, 
            msg: 'Campos obrigatórios: usuarioLogado, codigoProduto, quantidade' 
          });
        }
        
        const resultado = await sheetsAvaria.salvarRegistro(
          usuarioLogado, 
          codigoProduto, 
          descricaoProduto, 
          embalagemProduto, 
          quantidade
        );
        
        return res.status(200).json(resultado);
      }
      
      case 'obterHistorico': {
        const { filtros } = req.body || {};
        const resultado = await sheetsAvaria.obterHistorico(filtros || {});
        return res.status(200).json(resultado);
      }
      
      case 'testarConexao': {
        const doc = await sheetsAvaria.init();
        const sheets = Object.keys(doc.sheetsByTitle);
        return res.status(200).json({ 
          ok: true, 
          msg: `Conectado: ${doc.title}`, 
          abas: sheets 
        });
      }
      
      default:
        return res.status(400).json({ 
          ok: false, 
          msg: 'Ação inválida: ' + action,
          acoesDisponiveis: [
            'obterDadosProdutos', 
            'salvarRegistro', 
            'obterHistorico', 
            'testarConexao'
          ]
        });
    }
  } catch (error) {
    console.error('[API AVARIA] Erro:', error);
    return res.status(500).json({ 
      ok: false, 
      msg: 'Erro interno: ' + error.message, 
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined 
    });
  }
};
