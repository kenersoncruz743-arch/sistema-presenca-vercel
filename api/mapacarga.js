// api/mapacarga.js
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
    const { action, filtros } = req.body;

    switch (action) {
      case 'listar':
        const dados = await sheetsService.getMapaCarga(filtros || {});
        return res.status(200).json({ ok: true, dados });

      case 'atualizar':
        // Implementar lógica de atualização
        return res.status(200).json({ ok: true, msg: 'Atualizado com sucesso' });

      default:
        return res.status(400).json({ ok: false, msg: 'Ação inválida' });
    }
  } catch (error) {
    console.error('Erro na API Mapa de Carga:', error);
    return res.status(500).json({ ok: false, msg: error.message });
  }
};
