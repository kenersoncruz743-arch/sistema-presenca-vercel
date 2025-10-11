// api/auth.js - VERSÃO FUNCIONAL ORIGINAL
const sheetsService = require('../lib/sheets');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Método não permitido' });
  }

  try {
    const { usuario, senha, action } = req.body;

    if (action === 'login') {
      if (!usuario || !senha) {
        return res.status(400).json({ 
          ok: false, 
          msg: 'Usuário e senha são obrigatórios' 
        });
      }

      console.log(`Tentativa de login: ${usuario}`);
      const result = await sheetsService.validarLogin(usuario, senha);
      return res.status(200).json(result);
    }

    if (action === 'test') {
      const doc = await sheetsService.init();
      return res.status(200).json({
        ok: true,
        msg: `Conectado à planilha: ${doc.title}`,
        sheets: Object.keys(doc.sheetsByTitle)
      });
    }

    return res.status(400).json({ error: 'Ação não reconhecida' });

  } catch (error) {
    console.error('Erro na API de auth:', error);
    return res.status(500).json({ 
      error: 'Erro interno do servidor',
      details: error.message 
    });
  }
};
