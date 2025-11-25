const sheetsAvaria = require('../lib/sheets_3');

module.exports = async function handler(req, res) {
  // CORS
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
        console.log('[API AVARIA] Obtendo dados dos produtos...');
        const resultado = await sheetsAvaria.obterDadosProdutos();
        
        // ===== VALIDAÇÃO E CONVERSÃO SEGURA =====
        if (!resultado || !resultado.produtos) {
          return res.status(500).json({
            ok: false,
            msg: 'Erro ao carregar produtos: resultado inválido'
          });
        }

        // Garante que produtos é um array
        const produtos = Array.isArray(resultado.produtos) ? resultado.produtos : [];
        
        // Garante que mapaBusca é um objeto
        let mapaBuscaObj = {};
        
        if (resultado.mapaBusca) {
          if (resultado.mapaBusca instanceof Map) {
            // Se for Map, converte para objeto
            mapaBuscaObj = Object.fromEntries(resultado.mapaBusca);
          } else if (typeof resultado.mapaBusca === 'object') {
            // Se já for objeto, usa diretamente
            mapaBuscaObj = resultado.mapaBusca;
          }
        }
        
        console.log('[API AVARIA] ✓ Produtos:', produtos.length);
        console.log('[API AVARIA] ✓ Códigos indexados:', Object.keys(mapaBuscaObj).length);
        
        return res.status(200).json({ 
          ok: true, 
          dados: produtos,
          total: produtos.length,
          mapaBusca: mapaBuscaObj // Envia como objeto simples
        });
      }
      
      case 'salvarRegistro': {
        const { usuarioLogado, codigoProduto, descricaoProduto, embalagemProduto, motivo, quantidade } = req.body;
        
        if (!usuarioLogado || !codigoProduto || !motivo || !quantidade) {
          return res.status(400).json({ 
            ok: false, 
            msg: 'Campos obrigatórios: usuarioLogado, codigoProduto, motivo, quantidade' 
          });
        }
        
        const resultado = await sheetsAvaria.salvarRegistro(
          usuarioLogado, 
          codigoProduto, 
          descricaoProduto || 'Não informado', 
          embalagemProduto || 'Não informado',
          motivo, 
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
