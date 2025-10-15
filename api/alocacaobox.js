// api/alocacaobox.js - VERSÃO CORRIGIDA
const sheetsService = require('../lib/sheets');

module.exports = async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  console.log('[API ALOCACAO] Request recebido:', {
    method: req.method,
    body: req.body
  });

  try {
    const { action, filtros, boxNum, cargaId, dados } = req.body;

    // Valida action
    if (!action) {
      console.error('[API ALOCACAO] Action não fornecida');
      return res.status(400).json({ 
        ok: false, 
        msg: 'Action é obrigatória' 
      });
    }

    console.log('[API ALOCACAO] Action:', action);

    switch (action) {
      case 'listarCargas': {
        console.log('[API ALOCACAO] Buscando cargas sem BOX...');
        
        try {
          const cargas = await sheetsService.getCargasSemBox(filtros || {});
          
          console.log('[API ALOCACAO] Cargas encontradas:', cargas.length);
          
          return res.status(200).json({ 
            ok: true, 
            cargas: cargas,
            total: cargas.length
          });
        } catch (error) {
          console.error('[API ALOCACAO] Erro ao buscar cargas:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao buscar cargas: ' + error.message 
          });
        }
      }

      case 'listarBoxes': {
        console.log('[API ALOCACAO] Buscando estado dos boxes...');
        
        try {
          const boxes = await sheetsService.getEstadoBoxes();
          
          console.log('[API ALOCACAO] Boxes ocupados:', boxes.length);
          
          return res.status(200).json({ 
            ok: true, 
            boxes: boxes,
            total: boxes.length
          });
        } catch (error) {
          console.error('[API ALOCACAO] Erro ao buscar boxes:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao buscar boxes: ' + error.message 
          });
        }
      }

      case 'alocarBox': {
        console.log('[API ALOCACAO] Alocando carga:', { boxNum, cargaId });
        
        // Valida parâmetros
        if (!boxNum || !cargaId) {
          console.error('[API ALOCACAO] Parâmetros inválidos');
          return res.status(400).json({ 
            ok: false, 
            msg: 'boxNum e cargaId são obrigatórios' 
          });
        }
        
        try {
          const resultado = await sheetsService.alocarCargaBox(boxNum, cargaId);
          
          console.log('[API ALOCACAO] Resultado alocação:', resultado);
          
          return res.status(200).json(resultado);
        } catch (error) {
          console.error('[API ALOCACAO] Erro ao alocar:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao alocar carga: ' + error.message 
          });
        }
      }

      case 'liberarBox': {
        console.log('[API ALOCACAO] Liberando BOX:', boxNum);
        
        // Valida parâmetro
        if (!boxNum) {
          console.error('[API ALOCACAO] boxNum não fornecido');
          return res.status(400).json({ 
            ok: false, 
            msg: 'boxNum é obrigatório' 
          });
        }
        
        try {
          const resultadoLiberar = await sheetsService.liberarBox(boxNum);
          
          console.log('[API ALOCACAO] Resultado liberar:', resultadoLiberar);
          
          return res.status(200).json(resultadoLiberar);
        } catch (error) {
          console.error('[API ALOCACAO] Erro ao liberar:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao liberar BOX: ' + error.message 
          });
        }
      }

      case 'salvarAlocacoes': {
        console.log('[API ALOCACAO] Salvando múltiplas alocações:', dados?.length || 0);
        
        // Valida parâmetro
        if (!dados || !Array.isArray(dados)) {
          console.error('[API ALOCACAO] Dados inválidos');
          return res.status(400).json({ 
            ok: false, 
            msg: 'dados deve ser um array' 
          });
        }
        
        try {
          const resultadoSalvar = await sheetsService.salvarAlocacoes(dados);
          
          console.log('[API ALOCACAO] Resultado salvar:', resultadoSalvar);
          
          return res.status(200).json(resultadoSalvar);
        } catch (error) {
          console.error('[API ALOCACAO] Erro ao salvar:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao salvar alocações: ' + error.message 
          });
        }
      }

      default:
        console.error('[API ALOCACAO] Action inválida:', action);
        return res.status(400).json({ 
          ok: false, 
          msg: 'Ação inválida: ' + action 
        });
    }
  } catch (error) {
    console.error('[API ALOCACAO] Erro geral:', error);
    return res.status(500).json({ 
      ok: false, 
      msg: 'Erro interno do servidor: ' + error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
