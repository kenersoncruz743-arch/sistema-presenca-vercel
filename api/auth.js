import sheetsService from '../lib/sheets.js';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, msg: 'Método não permitido' });
  }

  try {
    const { usuario, senha } = req.body;

    if (!usuario || !senha) {
      return res.status(400).json({ ok: false, msg: 'Informe usuário e senha' });
    }

    const result = await sheetsService.validarLogin(usuario, senha);

    if (!result.ok) {
      return res.status(401).json(result);
    }

    const token = Buffer.from(`${usuario}:${Date.now()}`).toString('base64');

    return res.status(200).json({
      ok: true,
      usuario: result.usuario,
      abas: result.abas,
      token
    });
  } catch (error) {
    console.error('Erro no login:', error);
    return res.status(500).json({ ok: false, msg: 'Erro interno do servidor' });
  }
}
