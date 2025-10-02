// api/auth.js - Versão CORRIGIDA para ES Module
  import sheetsService from '../lib/sheets'; //
  
  export default async function handler(req, res) { 
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
    if (req.method === 'OPTIONS') {
      return res.status(200).end();
    }
  
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Método não permitido' });
  }

  try {
    const { usuario, senha, action } = req.body;

    // Validação de entrada
    if (!action) {
      return res.status(400).json({ 
        ok: false, 
        msg: 'Ação não especificada' 
      });
    }

    // LOGIN
    if (action === 'login') {
      if (!usuario || !senha) {
        return res.status(400).json({ 
          ok: false, 
          msg: 'Usuário e senha são obrigatórios' 
        });
      }

      console.log(`[AUTH] Tentativa de login: ${usuario}`);
      
      try {
        const result = await sheetsService.validarLogin(usuario, senha);
        console.log(`[AUTH] Resultado do login:`, result);
        return res.status(200).json(result);
      } catch (loginError) {
        console.error('[AUTH] Erro no processo de login:', loginError);
        return res.status(500).json({
          ok: false,
          msg: 'Erro ao validar credenciais: ' + loginError.message
        });
      }
    }



  } catch (error) {
    console.error('[AUTH] Erro geral na API:', error);
    return res.status(500).json({ 
      error: 'Erro interno do servidor',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
