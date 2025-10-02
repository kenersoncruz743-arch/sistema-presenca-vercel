// api/coletores.js - VERSÃO FINAL CORRIGIDA (SEM NENHUM REQUIRE)
import sheetsColetoresService from '../../lib/sheets'; 

export default async function handler(req, res) {
  // Headers obrigatórios
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Content-Type', 'application/json');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    // A importação do service (sheetsColetoresService) já foi feita no topo.

    if (req.method === 'POST') {
      const { action, ...data } = req.body;
      
      switch (action) {
        case 'salvarRegistro':
          try {
            const result = await sheetsColetoresService.salvarRegistro(data);
            return res.status(200).json(result);
          } catch (error) {
            console.error('[API Coletores] Erro ao salvar registro:', error);
            return res.status(500).json({ error: 'Erro ao salvar registro', details: error.message });
          }

        case 'obterStatusCompleto':
          try {
            const [status, resumo] = await Promise.all([
              sheetsColetoresService.obterStatus(),
              sheetsColetoresService.gerarResumo()
            ]);
            return res.status(200).json({ status, resumo });
          } catch (error) {
            console.error('[API Coletores] Erro ao obter status completo:', error);
            return res.status(500).json({ error: 'Erro ao obter status completo', message: error.message });
          }
          
        default:
          return res.status(400).json({ error: 'Ação POST inválida' });
      }
    }

    if (req.method === 'GET') {
      const { action } = req.query;
      if (action === 'obterDados') {
        const dados = await sheetsColetoresService.obterDadosColaboradores();
        return res.status(200).json(dados);
      }
      return res.status(400).json({ error: 'Ação GET inválida' });
    }

    return res.status(405).json({ error: 'Método não permitido' });

  } catch (error) {
    console.error('[API Coletores] Erro geral:', error);
    return res.status(500).json({ 
      error: 'Erro interno do servidor',
      message: error.message,
    });
  }
}
