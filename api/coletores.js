// api/coletores.js - API para controle de coletores com tratamento de erros melhorado
const sheetsService = require('../../lib/sheets');

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Content-Type', 'application/json');
  
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  try {
    if (req.method === 'GET') {
      const { action } = req.query;
      
      console.log(`API Coletores GET: action=${action}`);
      
      switch (action) {
        case 'obterDados':
          console.log('Obtendo dados dos coletores...');
          try {
            const dados = await sheetsService.obterDadosColetores();
            console.log(`Dados obtidos com sucesso: ${dados.length} registros`);
            return res.status(200).json(dados);
          } catch (error) {
            console.error('Erro específico ao obter dados:', error);
            return res.status(500).json({ 
              error: 'Erro ao obter dados de coletores',
              details: error.message,
              suggestion: 'Verifique se a planilha de coletores existe e tem a aba "Quadro" com colunas CHAPA1, NOME, FUNCAO'
            });
          }
          
        case 'obterStatus':
          console.log('Obtendo status dos coletores...');
          try {
            const status = await sheetsService.obterColetorStatus();
            console.log(`Status obtido com sucesso: ${Object.keys(status).length} coletores`);
            return res.status(200).json(status);
          } catch (error) {
            console.error('Erro específico ao obter status:', error);
            return res.status(500).json({ 
              error: 'Erro ao obter status de coletores',
              details: error.message,
              suggestion: 'Verifique se a planilha de coletores existe e tem a aba "Base"'
            });
          }
          
        case 'obterResumo':
          console.log('Gerando resumo dos coletores...');
          try {
            const resumo = await sheetsService.gerarResumoColetores();
            console.log('Resumo gerado com sucesso:', resumo);
            return res.status(200).json(resumo);
          } catch (error) {
            console.error('Erro específico ao gerar resumo:', error);
            return res.status(500).json({ 
              error: 'Erro ao gerar resumo de coletores',
              details: error.message
            });
          }
          
        case 'obterPresenca':
          console.log('Obtendo dados de presença...');
          try {
            const presenca = await sheetsService.obterDadosPresenca();
            console.log(`Dados de presença obtidos: ${presenca.length} registros`);
            return res.status(200).json(presenca);
          } catch (error) {
            console.error('Erro específico ao obter presença:', error);
            return res.status(500).json({ 
              error: 'Erro ao obter dados de presença',
              details: error.message
            });
          }
          
        default:
          return res.status(400).json({ 
            error: 'Ação GET não reconhecida',
            action: action,
            actionsDisponiveis: ['obterDados', 'obterStatus', 'obterResumo', 'obterPresenca']
          });
      }
    }

    if (req.method === 'POST') {
      const { action } = req.body;
      
      console.log(`API Coletores POST: action=${action}`);
      
      switch (action) {
        case 'salvarRegistro':
          const { chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes, supervisor } = req.body;
          
          console.log('Salvando registro de coletor:', {
            chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes, supervisor
          });
          
          // Validações de entrada
          if (!chapa) {
            return res.status(400).json({
              ok: false,
              msg: 'Campo "chapa" é obrigatório'
            });
          }
          
          if (!numeroColetor) {
            return res.status(400).json({
              ok: false,
              msg: 'Campo "numeroColetor" é obrigatório'
            });
          }
          
          if (!situacoes || (Array.isArray(situacoes) && situacoes.length === 0)) {
            return res.status(400).json({
              ok: false,
              msg: 'Campo "situacoes" é obrigatório'
            });
          }

          if (numeroColetor < 1 || numeroColetor > 145) {
            return res.status(400).json({
              ok: false,
              msg: 'Número do coletor deve estar entre 1 e 145'
            });
          }
          
          try {
            const resultSalvar = await sheetsService.salvarRegistroColetor({
              chapa,
              nome,
              funcao,
              numeroColetor,
              tipoOperacao,
              situacoes,
              supervisor
            });
            
            console.log('Resultado do salvamento:', resultSalvar);
            return res.status(200).json(resultSalvar);
          } catch (error) {
            console.error('Erro específico ao salvar registro:', error);
            return res.status(500).json({
              ok: false,
              msg: 'Erro interno ao salvar registro',
              details: error.message,
              suggestion: 'Verifique se a planilha de coletores existe e tem permissões de escrita'
            });
          }
          
        case 'obterStatusCompleto':
          console.log('Obtendo status completo dos coletores...');
          try {
            const statusCompleto = await sheetsService.obterColetorStatus();
            const resumoCompleto = await sheetsService.gerarResumoColetores();
            
            console.log(`Status completo: ${Object.keys(statusCompleto).length} coletores, resumo:`, resumoCompleto);
            
            return res.status(200).json({
              status: statusCompleto,
              resumo: resumoCompleto
            });
          } catch (error) {
            console.error('Erro específico ao obter status completo:', error);
            return res.status(500).json({ 
              error: 'Erro ao obter status completo de coletores',
              details: error.message
            });
          }
          
        default:
          return res.status(400).json({ 
            error: 'Ação POST não reconhecida',
            action: action,
            actionsDisponiveis: ['salvarRegistro', 'obterStatusCompleto']
          });
      }
    }

    return res.status(405).json({ error: 'Método não permitido', method: req.method });

  } catch (error) {
    console.error('Erro geral na API de coletores:', error);
    
    // Tenta determinar se é erro de configuração
    if (error.message && error.message.includes('GOOGLE_SHEETS_COLETORES_ID')) {
      return res.status(500).json({ 
        error: 'Configuração inválida',
        details: 'GOOGLE_SHEETS_COLETORES_ID não configurado no arquivo .env.local',
        suggestion: 'Adicione a variável GOOGLE_SHEETS_COLETORES_ID com o ID da sua planilha de coletores'
      });
    }
    
    if (error.message && error.message.includes('Unable to parse range')) {
      return res.status(500).json({ 
        error: 'Erro na estrutura da planilha',
        details: 'Uma das abas necessárias não foi encontrada ou está com estrutura incorreta',
        suggestion: 'Verifique se a planilha de coletores tem as abas "Quadro" e "Base" com as colunas corretas'
      });
    }
    
    return res.status(500).json({ 
      error: 'Erro interno do servidor',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
}
