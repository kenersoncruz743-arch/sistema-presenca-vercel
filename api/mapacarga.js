// api/mapacarga.js - API UNIFICADA (Mapa de Carga + Alocação BOX)
const sheetsService = require('../lib/sheets');

module.exports = async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  console.log('[API MAPACARGA] Request:', {
    method: req.method,
    body: req.body
  });

  try {
    const { action, filtros, boxNum, cargaId, dados, dadosImportacao } = req.body;

    if (!action) {
      console.error('[API MAPACARGA] Action não fornecida');
      return res.status(400).json({ 
        ok: false, 
        msg: 'Action é obrigatória' 
      });
    }

    console.log('[API MAPACARGA] Action:', action);

    switch (action) {
      // ==================== MAPA DE CARGA ====================
      
      case 'listar': {
        console.log('[API MAPACARGA] Listando cargas...');
        
        try {
          const dados = await sheetsService.getMapaCarga(filtros || {});
          
          console.log('[API MAPACARGA] Cargas encontradas:', dados.length);
          
          return res.status(200).json({ 
            ok: true, 
            dados: dados,
            total: dados.length
          });
        } catch (error) {
          console.error('[API MAPACARGA] Erro ao listar:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao listar cargas: ' + error.message 
          });
        }
      }

      case 'atualizar': {
        console.log('[API MAPACARGA] Atualizando carga...');
        
        if (!dados || !dados.carga) {
          return res.status(400).json({ 
            ok: false, 
            msg: 'Dados da carga são obrigatórios' 
          });
        }
        
        try {
          const resultado = await sheetsService.atualizarMapaCarga(dados.carga, dados.campos);
          return res.status(200).json(resultado);
        } catch (error) {
          console.error('[API MAPACARGA] Erro ao atualizar:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao atualizar carga: ' + error.message 
          });
        }
      }

      case 'importar': {
        console.log('[API MAPACARGA] Processando dados colados...');
        
        if (!dadosImportacao || !Array.isArray(dadosImportacao)) {
          return res.status(400).json({ 
            ok: false, 
            msg: 'dadosImportacao deve ser um array' 
          });
        }
        
        try {
          const resultado = await sheetsService.processarMapaCargaColado(dadosImportacao);
          
          console.log('[API MAPACARGA] Resultado processamento:', resultado);
          
          return res.status(200).json(resultado);
        } catch (error) {
          console.error('[API MAPACARGA] Erro ao processar:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao processar dados: ' + error.message 
          });
        }
      }
      
      case 'limpar': {
        console.log('[API MAPACARGA] Limpando colunas específicas...');
        
        try {
          const resultado = await sheetsService.limparColunasMapaCarga();
          
          console.log('[API MAPACARGA] Resultado limpeza:', resultado);
          
          return res.status(200).json(resultado);
        } catch (error) {
          console.error('[API MAPACARGA] Erro ao limpar:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao limpar colunas: ' + error.message 
          });
        }
      }

      // ==================== ALOCAÇÃO BOX ====================

      case 'listarCargas': {
        console.log('[API MAPACARGA] Buscando cargas sem BOX...');
        
        try {
          const cargas = await sheetsService.getCargasSemBox(filtros || {});
          
          console.log('[API MAPACARGA] Cargas sem BOX:', cargas.length);
          
          return res.status(200).json({ 
            ok: true, 
            cargas: cargas,
            total: cargas.length
          });
        } catch (error) {
          console.error('[API MAPACARGA] Erro ao buscar cargas:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao buscar cargas: ' + error.message 
          });
        }
      }

      case 'listarBoxes': {
        console.log('[API MAPACARGA] Buscando estado dos boxes...');
        
        try {
          const boxes = await sheetsService.getEstadoBoxes();
          
          console.log('[API MAPACARGA] Boxes ocupados:', boxes.length);
          
          return res.status(200).json({ 
            ok: true, 
            boxes: boxes,
            total: boxes.length
          });
        } catch (error) {
          console.error('[API MAPACARGA] Erro ao buscar boxes:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao buscar boxes: ' + error.message 
          });
        }
      }

      case 'alocarBox': {
        console.log('[API MAPACARGA] Alocando carga:', { boxNum, cargaId });
        
        if (!boxNum || !cargaId) {
          console.error('[API MAPACARGA] Parâmetros inválidos');
          return res.status(400).json({ 
            ok: false, 
            msg: 'boxNum e cargaId são obrigatórios' 
          });
        }
        
        try {
          const resultado = await sheetsService.alocarCargaBox(boxNum, cargaId);
          
          console.log('[API MAPACARGA] Resultado alocação:', resultado);
          
          return res.status(200).json(resultado);
        } catch (error) {
          console.error('[API MAPACARGA] Erro ao alocar:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao alocar carga: ' + error.message 
          });
        }
      }

      case 'liberarBox': {
        console.log('[API MAPACARGA] Liberando BOX:', boxNum);
        
        if (!boxNum) {
          console.error('[API MAPACARGA] boxNum não fornecido');
          return res.status(400).json({ 
            ok: false, 
            msg: 'boxNum é obrigatório' 
          });
        }
        
        try {
          const resultado = await sheetsService.liberarBox(boxNum);
          
          console.log('[API MAPACARGA] Resultado liberar:', resultado);
          
          return res.status(200).json(resultado);
        } catch (error) {
          console.error('[API MAPACARGA] Erro ao liberar:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao liberar BOX: ' + error.message 
          });
        }
      }

      default:
        console.error('[API MAPACARGA] Action inválida:', action);
        return res.status(400).json({ 
          ok: false, 
          msg: 'Ação inválida: ' + action 
        });
    }
  } catch (error) {
    console.error('[API MAPACARGA] Erro geral:', error);
    return res.status(500).json({ 
      ok: false, 
      msg: 'Erro interno do servidor: ' + error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
