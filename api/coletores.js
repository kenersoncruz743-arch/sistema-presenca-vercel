// api/coletores.js - Versão CORRIGIDA para ES Module
import sheetsColetoresService from '../../lib/sheets'; // <-- CORRIGIDO

export default async function handler(req, res) {
  // Headers obrigatórios
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Content-Type', 'application/json');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {

    if (req.method === 'GET') {
      const { action } = req.query;
      console.log(`[API Coletores] GET action: ${action}`);
      
      switch (action) {
        case 'obterDados':
          try {
            console.log('[API Coletores] Iniciando obtenção de dados...');
            const dados = await sheetsColetoresService.obterDadosColaboradores();
            console.log(`[API Coletores] Dados obtidos com sucesso: ${dados.length} registros`);
            return res.status(200).json(dados);
          } catch (error) {
            console.error('[API Coletores] Erro ao obter dados:', error);
            return res.status(500).json({ 
              error: 'Erro ao obter dados de colaboradores',
              message: error.message,
              suggestion: 'Verifique se a planilha de coletores tem a aba "Quadro.1" com colunas Chapa, Nome, Funcao'
            });
          }
          
        case 'obterStatus':
          try {
            console.log('[API Coletores] Iniciando obtenção de status...');
            const status = await sheetsColetoresService.obterStatus();
            console.log(`[API Coletores] Status obtido: ${Object.keys(status).length} coletores`);
            return res.status(200).json(status);
          } catch (error) {
            console.error('[API Coletores] Erro ao obter status:', error);
            return res.status(500).json({ 
              error: 'Erro ao obter status de coletores',
              message: error.message,
              suggestion: 'Verifique se a planilha de coletores tem a aba "Historico"'
            });
          }
          
        case 'obterResumo':
          try {
            console.log('[API Coletores] Iniciando geração de resumo...');
            const resumo = await sheetsColetoresService.gerarResumo();
            console.log('[API Coletores] Resumo gerado:', resumo);
            return res.status(200).json(resumo);
          } catch (error) {
            console.error('[API Coletores] Erro ao gerar resumo:', error);
            return res.status(500).json({ 
              error: 'Erro ao gerar resumo',
              message: error.message
            });
          }
          
        case 'obterPresenca':
          try {
            console.log('[API Coletores] Iniciando obtenção de dados de presença...');
            const presenca = await sheetsColetoresService.obterDadosPresenca();
            console.log(`[API Coletores] Dados de presença obtidos: ${presenca.length} registros`);
            return res.status(200).json(presenca);
          } catch (error) {
            console.error('[API Coletores] Erro ao obter dados de presença:', error);
            return res.status(500).json({ 
              error: 'Erro ao obter dados de presença',
              message: error.message,
              suggestion: 'Verifique se a planilha de coletores tem a aba "Presenca"'
            });
          }
          
        default:
          return res.status(400).json({ 
            error: 'Ação inválida',
            action: action,
            availableActions: ['obterDados', 'obterStatus', 'obterResumo', 'obterPresenca']
          });
      }
    }

    if (req.method === 'POST') {
      const { action, registro } = req.body;
      console.log(`[API Coletores] POST action: ${action}`);
      
      switch (action) {
        case 'salvarRegistro':
          const { chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes, supervisor } = req.body;
          
          console.log('[API Coletores] Dados recebidos:', { chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes, supervisor });
          
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
          
          try {
            const resultado = await sheetsColetoresService.salvarRegistro({
              chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes, supervisor
            });
            
            console.log('[API Coletores] Resultado do salvamento:', resultado);
            return res.status(200).json(resultado);
          } catch (error) {
            console.error('[API Coletores] Erro ao salvar registro:', error);
            return res.status(500).json({
              ok: false,
              msg: 'Erro interno ao salvar registro',
              details: error.message
            });
          }
          
        case 'obterStatusCompleto':
          try {
            console.log('[API Coletores] Iniciando obtenção de status completo...');
            const [status, resumo] = await Promise.all([
              sheetsColetoresService.obterStatus(),
              sheetsColetoresService.gerarResumo()
            ]);
            
            console.log(`[API Coletores] Status completo obtido: ${Object.keys(status).length} coletores`);
            return res.status(200).json({ status, resumo });
          } catch (error) {
            console.error('[API Coletores] Erro ao obter status completo:', error);
            return res.status(500).json({ 
              error: 'Erro ao obter status completo',
              message: error.message
            });
          }
          
        default:
          return res.status(400).json({ 
            error: 'Ação POST inválida',
            action: action,
            availableActions: ['salvarRegistro', 'obterStatusCompleto']
          });
      }
    }

        return res.status(405).json({ 
          error: 'Método não permitido', 
          method: req.method,
          allowedMethods: ['GET', 'POST', 'OPTIONS']
        });
    
      } catch (error) {
        console.error('[API Coletores] Erro geral:', error);
        return res.status(500).json({ 
          error: 'Erro interno do servidor',
          message: error.message,
          stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
        });
      }
 }
