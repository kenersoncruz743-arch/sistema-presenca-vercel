// api/colaboradores.js - COM SUPORTE A VALIDAÇÃO POR ABA
const sheetsService = require('../lib/sheets');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // GET - Buscar colaboradores
  if (req.method === 'GET') {
    try {
      const { filtro } = req.query;
      const colaboradores = await sheetsService.buscarColaboradores(filtro || '');
      return res.status(200).json(colaboradores);
    } catch (error) {
      console.error('Erro ao buscar colaboradores:', error);
      return res.status(500).json({ error: 'Erro ao buscar colaboradores' });
    }
  }

  // POST - Ações diversas
  if (req.method === 'POST') {
    try {
      const { action } = req.body;

      switch (action) {
        case 'addBuffer': {
          const { supervisor, aba, colaborador } = req.body;
          const result = await sheetsService.adicionarBuffer(supervisor, aba, colaborador);
          return res.status(200).json(result);
        }

        case 'getBuffer': {
          const { supervisor, aba } = req.body;
          const buffer = await sheetsService.getBuffer(supervisor, aba);
          return res.status(200).json(buffer);
        }

        case 'removeBuffer': {
          const { supervisor, aba, matricula } = req.body;
          // Usa aba como chave principal se fornecida
          const chave = aba || supervisor;
          const result = await sheetsService.removerBufferPorAba(chave, matricula);
          return res.status(200).json(result);
        }

        case 'updateStatus': {
          const { supervisor, aba, matricula, status } = req.body;
          // Usa aba como chave principal se fornecida
          const chave = aba || supervisor;
          const result = await sheetsService.atualizarStatusBufferPorAba(chave, matricula, status);
          return res.status(200).json(result);
        }

        case 'updateDesvio': {
          const { supervisor, aba, matricula, desvio } = req.body;
          // Usa aba como chave principal se fornecida
          const chave = aba || supervisor;
          const result = await sheetsService.atualizarDesvioBufferPorAba(chave, matricula, desvio);
          return res.status(200).json(result);
        }

        case 'saveToBase': {
          const { dados } = req.body;
          const result = await sheetsService.salvarNaBase(dados);
          return res.status(200).json(result);
        }

        default:
          return res.status(400).json({ ok: false, msg: 'Ação não reconhecida' });
      }
    } catch (error) {
      console.error('Erro na API de colaboradores:', error);
      return res.status(500).json({ ok: false, msg: 'Erro interno do servidor' });
    }
  }

  return res.status(405).json({ ok: false, msg: 'Método não permitido' });
};
