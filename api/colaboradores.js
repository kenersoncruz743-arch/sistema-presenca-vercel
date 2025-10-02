import sheetsService from '../lib/sheets'; // <-- CORRIGIDO

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
      const { filtro } = req.query;
      console.log(`Buscando colaboradores com filtro: "${filtro || 'sem filtro'}"`);
      const colaboradores = await sheetsService.buscarColaboradores(filtro);
      return res.status(200).json(colaboradores);
    }

    if (req.method === 'POST') {
      const { action, supervisor, aba, matricula, colaborador, status } = req.body;

      switch (action) {
        case 'getBuffer':
          const buffer = await sheetsService.getBuffer(supervisor, aba);
          return res.status(200).json(buffer);

        case 'addBuffer':
          if (!supervisor || !aba || !colaborador) {
            return res.status(400).json({ 
              ok: false, 
              msg: 'Dados incompletos para adicionar ao buffer' 
            });
          }
          const resultAdd = await sheetsService.adicionarBuffer(supervisor, aba, colaborador);
          return res.status(200).json(resultAdd);

        case 'removeBuffer':
          if (!supervisor || !matricula) {
            return res.status(400).json({ 
              ok: false, 
              msg: 'Supervisor e matrícula são obrigatórios' 
            });
          }
          const resultRemove = await sheetsService.removerBuffer(supervisor, matricula);
          return res.status(200).json(resultRemove);

        case 'updateStatus':
          if (!supervisor || !matricula) {
            return res.status(400).json({ 
              ok: false, 
              msg: 'Supervisor e matrícula são obrigatórios' 
            });
          }
          const resultUpdate = await sheetsService.atualizarStatusBuffer(supervisor, matricula, status);
          return res.status(200).json(resultUpdate);

        case 'saveToBase':
          const { dados } = req.body;
          if (!dados || !Array.isArray(dados)) {
            return res.status(400).json({ 
              ok: false, 
              msg: 'Dados inválidos para salvar' 
            });
          }
          const resultSave = await sheetsService.salvarNaBase(dados);
          return res.status(200).json(resultSave);

        default:
          return res.status(400).json({ 
            error: 'Ação não reconhecida'
          });
      }
    }

    return res.status(405).json({ error: 'Método não permitido' });

  } catch (error) {
    console.error('Erro na API de colaboradores:', error);
    return res.status(500).json({ 
      error: 'Erro interno do servidor',
      details: error.message 
    });
  }
}
