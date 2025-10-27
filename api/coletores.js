// api/coletores.js - API para controle de coletores (CORRIGIDA)
const sheetsColetorService = require('../lib/sheets_2');

module.exports = async function handler(req, res) {
  // ===== CORS CONFIGURADO CORRETAMENTE =====
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.setHeader('Access-Control-Allow-Credentials', 'true');

  // Responde OPTIONS (preflight)
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  console.log('[API COLETORES] Request:', {
    method: req.method,
    action: req.body?.action || req.query?.action,
    body: req.body
  });

  try {
    // ===== ACEITA ACTION POR POST OU GET =====
    const action = req.method === 'POST' ? req.body?.action : req.query?.action;

    if (!action) {
      console.error('[API COLETORES] Action não fornecida');
      return res.status(400).json({ 
        ok: false, 
        msg: 'Action é obrigatória',
        acoesDisponiveis: [
          'obterDados',
          'salvarRegistro',
          'obterColetorStatus',
          'obterResumoColetores',
          'obterResumoPorSupervisor',
          'testarConexao',
          'atualizarDadosPresenca'
        ]
      });
    }

    console.log('[API COLETORES] Processando action:', action);

    switch (action) {
      // ===== BUSCAR COLABORADORES =====
      case 'obterDados': {
        console.log('[API COLETORES] Buscando dados do quadro...');
        
        try {
          const dados = await sheetsColetorService.obterDados();
          
          console.log('[API COLETORES] Colaboradores encontrados:', dados.length);
          
          return res.status(200).json({ 
            ok: true, 
            dados: dados,
            total: dados.length
          });
        } catch (error) {
          console.error('[API COLETORES] Erro ao buscar dados:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao buscar colaboradores: ' + error.message 
          });
        }
      }

      // ===== SALVAR REGISTRO =====
      case 'salvarRegistro': {
        const { chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes } = req.body;
        
        console.log('[API COLETORES] Salvando registro:', {
          chapa, nome, numeroColetor, tipoOperacao, situacoes
        });
        
        // ===== VALIDAÇÃO COMPLETA =====
        if (!chapa || !numeroColetor || !situacoes) {
          console.error('[API COLETORES] Parâmetros obrigatórios faltando');
          return res.status(400).json({ 
            ok: false, 
            msg: 'Chapa, número do coletor e situação são obrigatórios' 
          });
        }

        // Valida array de situações
        if (!Array.isArray(situacoes) || situacoes.length === 0) {
          return res.status(400).json({ 
            ok: false, 
            msg: 'Situações deve ser um array com pelo menos um item' 
          });
        }

        // Valida número do coletor
        const numColetor = parseInt(numeroColetor);
        if (isNaN(numColetor) || numColetor < 1 || numColetor > 140) {
          return res.status(400).json({ 
            ok: false, 
            msg: 'Número do coletor deve estar entre 1 e 140' 
          });
        }
        
        try {
          const resultado = await sheetsColetorService.salvarRegistro(
            chapa, 
            nome, 
            funcao, 
            numColetor, 
            tipoOperacao, 
            situacoes
          );
          
          console.log('[API COLETORES] Resultado salvamento:', resultado);
          
          if (resultado.ok) {
            return res.status(200).json(resultado);
          } else {
            return res.status(400).json(resultado);
          }
        } catch (error) {
          console.error('[API COLETORES] Erro ao salvar registro:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao salvar registro: ' + error.message 
          });
        }
      }

      // ===== STATUS DOS COLETORES =====
      case 'obterColetorStatus': {
        console.log('[API COLETORES] Obtendo status dos coletores...');
        
        try {
          const statusMap = await sheetsColetorService.obterColetorStatus();
          
          console.log('[API COLETORES] Status obtido:', Object.keys(statusMap).length, 'coletores');
          
          return res.status(200).json({ 
            ok: true, 
            statusMap: statusMap,
            total: Object.keys(statusMap).length
          });
        } catch (error) {
          console.error('[API COLETORES] Erro ao obter status:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao obter status dos coletores: ' + error.message 
          });
        }
      }

      // ===== RESUMO DE COLETORES =====
      case 'obterResumoColetores': {
        console.log('[API COLETORES] Gerando resumo de coletores...');
        
        try {
          const resumo = await sheetsColetorService.gerarResumoColetores();
          
          console.log('[API COLETORES] Resumo gerado:', resumo);
          
          return res.status(200).json({ 
            ok: true, 
            resumo: resumo
          });
        } catch (error) {
          console.error('[API COLETORES] Erro ao gerar resumo:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao gerar resumo: ' + error.message 
          });
        }
      }

      // ===== RESUMO POR SUPERVISOR =====
      case 'obterResumoPorSupervisor': {
        console.log('[API COLETORES] Gerando resumo por supervisor...');
        
        try {
          const resumoSup = await sheetsColetorService.gerarResumoPorSupervisor();
          
          console.log('[API COLETORES] Resumo por supervisor gerado');
          
          return res.status(200).json({ 
            ok: true, 
            resumoPorSupervisor: resumoSup
          });
        } catch (error) {
          console.error('[API COLETORES] Erro ao gerar resumo por supervisor:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao gerar resumo por supervisor: ' + error.message 
          });
        }
      }

      // ===== TESTE DE CONEXÃO =====
      case 'testarConexao': {
        console.log('[API COLETORES] Testando conexão...');
        
        try {
          const doc = await sheetsColetorService.init();
          const sheets = Object.keys(doc.sheetsByTitle);
          
          console.log(`[API COLETORES] ✓ Conectado: ${doc.title}`);
          console.log(`[API COLETORES] Abas encontradas:`, sheets);
          
          return res.status(200).json({
            ok: true,
            msg: `Conectado à planilha: ${doc.title}`,
            sheets: sheets,
            totalSheets: sheets.length
          });
        } catch (error) {
          console.error('[API COLETORES] Erro no teste de conexão:', error);
          return res.status(500).json({
            ok: false,
            msg: 'Erro ao conectar com planilha de coletores',
            details: error.message
          });
        }
      }

      // ===== ATUALIZAR DADOS DA PRESENÇA =====
      case 'atualizarDadosPresenca': {
        console.log('[API COLETORES] Atualizando dados da Presença...');
        
        try {
          await sheetsColetorService.atualizarDadosPresencaNaBase();
          
          console.log('[API COLETORES] ✓ Dados da Presença atualizados');
          
          return res.status(200).json({ 
            ok: true, 
            msg: 'Dados da Presença atualizados com sucesso' 
          });
        } catch (error) {
          console.error('[API COLETORES] Erro ao atualizar dados da Presença:', error);
          return res.status(500).json({ 
            ok: false, 
            msg: 'Erro ao atualizar dados da Presença: ' + error.message 
          });
        }
      }

      default:
        console.error('[API COLETORES] Action inválida:', action);
        return res.status(400).json({ 
          ok: false, 
          msg: 'Ação inválida: ' + action,
          acoesDisponiveis: [
            'obterDados',
            'salvarRegistro',
            'obterColetorStatus',
            'obterResumoColetores',
            'obterResumoPorSupervisor',
            'testarConexao',
            'atualizarDadosPresenca'
          ]
        });
    }
  } catch (error) {
    console.error('[API COLETORES] Erro geral:', error);
    return res.status(500).json({ 
      ok: false, 
      msg: 'Erro interno do servidor: ' + error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
