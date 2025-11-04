const sheetsService = require('../../lib/sheets');
module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'GET') return res.status(405).json({ok: false, msg: 'Método não permitido'});
  try {
    const {tipo} = req.query;
    const doc = await sheetsService.init();
    if (tipo === 'meta') {
      const sheet = doc.sheetsByTitle['Meta'];
      if (!sheet) return res.status(404).json({ok: false, msg: 'Aba Meta não encontrada'});
      const rows = await sheet.getRows();
      const dados = rows.map(row => ({
        data: String(row.get('Data') || '').trim(),
        meta: parseFloat(String(row.get('Meta') || '0').replace(',', '.')) || 0,
        produtividadeHora: parseFloat(String(row.get('Produtividade/hora') || '0').replace(',', '.')) || 0
      })).filter(d => d.data && d.produtividadeHora);
      return res.status(200).json({ok: true, dados, total: dados.length});
    }
    if (tipo === 'produtividade') {
      const sheet = doc.sheetsByTitle['Produtividade_Hora'];
      if (!sheet) return res.status(404).json({ok: false, msg: 'Aba Produtividade_Hora não encontrada'});
      const rows = await sheet.getRows();
      const dados = rows.map(row => ({
        funcao: String(row.get('FUNCAO') || '').trim(),
        produtividadeHora: parseFloat(String(row.get('Produtividade/hora') || '0').replace(',', '.')) || 0
      })).filter(d => d.funcao && d.produtividadeHora);
      return res.status(200).json({ok: true, dados, total: dados.length});
    }
    return res.status(400).json({ok: false, msg: 'Tipo não especificado (meta ou produtividade)'});
  } catch (error) {
    return res.status(500).json({ok: false, msg: 'Erro ao buscar dados', details: error.message});
  }
};
