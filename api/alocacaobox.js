// api/alocacaobox.js
const sheetsService = require('../lib/sheets');

module.exports = async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    const { action, filtros, boxNum, cargaId, dados } = req.body;

    switch (action) {
      case 'listarCargas':
        // Busca cargas do Mapa de Carga que ainda não têm BOX alocado
        const cargas = await sheetsService.getCargasSemBox(filtros || {});
        return res.status(200).json({ ok: true, cargas });

      case 'listarBoxes':
        // Busca estado atual dos boxes
        const boxes = await sheetsService.getEstadoBoxes();
        return res.status(200).json({ ok: true, boxes });

      case 'alocarBox':
        // Aloca uma carga em um BOX
        const resultado = await sheetsService.alocarCargaBox(boxNum, cargaId);
        return res.status(200).json(resultado);

      case 'liberarBox':
        // Libera um BOX
        const resultadoLiberar = await sheetsService.liberarBox(boxNum);
        return res.status(200).json(resultadoLiberar);

      case 'salvarAlocacoes':
        // Salva múltiplas alocações de uma vez
        const resultadoSalvar = await sheetsService.salvarAlocacoes(dados);
        return res.status(200).json(resultadoSalvar);

      default:
        return res.status(400).json({ ok: false, msg: 'Ação inválida' });
    }
  } catch (error) {
    console.error('Erro na API Alocação BOX:', error);
    return res.status(500).json({ ok: false, msg: error.message });
  }
};
