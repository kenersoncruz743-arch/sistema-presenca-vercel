// api/coletores.js - API para controle de coletores
const sheetsService = require('../../lib/sheets');

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  try {
    if (req.method === 'GET') {
      const { action } = req.query;
      
      switch (action) {
        case 'obterDados':
          console.log('Obtendo dados dos coletores...');
          const dados = await sheetsService.obterDadosColetores();
          return res.status(200).json(dados);
          
        case 'obterStatus':
          console.log('Obtendo status dos coletores...');
          const status = await sheetsService.obterColetorStatus();
          return res.status(200).json(status);
          
        case 'obterResumo':
          console.log('Gerando resumo dos coletores...');
          const resumo = await sheetsService.gerarResumoColetores();
          return res.status(200).json(resumo);
          
        default:
          return res.status(400).json({ 
            error: 'Ação GET não reconhecida',
            actionsDisponiveis: ['obterDados', 'obterStatus', 'obterResumo']
          });
      }
    }

    if (req.method === 'POST') {
      const { action } = req.body;
      
      switch (action) {
        case 'salvarRegistro':
          const { chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes, supervisor } = req.body;
          
          console.log('Salvando registro de coletor:', {
            chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes, supervisor
          });
          
          if (!chapa || !numeroColetor || !situacoes) {
            return res.status(400).json({
              ok: false,
              msg: 'Campos obrigatórios: chapa, numeroColetor, situacoes'
            });
          }

          if (numeroColetor < 1 || numeroColetor > 145) {
            return res.status(400).json({
              ok: false,
              msg: 'Número do coletor deve estar entre 1 e 145'
            });
          }
          
          const resultSalvar = await sheetsService.salvarRegistroColetor({
            chapa,
            nome,
            funcao,
            numeroColetor,
            tipoOperacao,
            situacoes,
            supervisor
          });
          
          return res.status(200).json(resultSalvar);
          
        case 'obterStatusCompleto':
          console.log('Obtendo status completo dos coletores...');
          const statusCompleto = await sheetsService.obterColetorStatus();
          const resumoCompleto = await sheetsService.gerarResumoColetores();
          
          return res.status(200).json({
            status: statusCompleto,
            resumo: resumoCompleto
          });
          
        default:
          return res.status(400).json({ 
            error: 'Ação POST não reconhecida',
            actionsDisponiveis: ['salvarRegistro', 'obterStatusCompleto']
          });
      }
    }

    return res.status(405).json({ error: 'Método não permitido' });

  } catch (error) {
    console.error('Erro na API de coletores:', error);
    return res.status(500).json({ 
      error: 'Erro interno do servidor',
      details: error.message 
    });
  }
}
