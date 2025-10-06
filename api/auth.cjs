const sheetsService = require('../lib/sheets');

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
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

    if (action === 'createTestData') {
      try {
        const doc = await sheetsService.init();
        
        let sheet = await sheetsService.getSheet('Usuarios');
        const existingUsers = await sheet.getRows();
        
        if (existingUsers.length === 0) {
          await sheet.addRows([
            { Usuario: 'admin', Senha: '123', Aba: 'PCP_Gestão' },
            { Usuario: 'supervisor1', Senha: '456', Aba: 'WMS TA' },
            { Usuario: 'supervisor2', Senha: '789', Aba: 'WMS TB' },
            { Usuario: 'admin', Senha: '123', Aba: 'Separação TB' }
          ]);
        }
        
        sheet = await sheetsService.getSheet('Quadro');
        const existingQuadro = await sheet.getRows();
        
        if (existingQuadro.length === 0) {
          await sheet.addRows([
            { Matricula: '001', Nome: 'João Silva', Funcao: 'Operador' },
            { Matricula: '002', Nome: 'Maria Santos', Funcao: 'Supervisora' },
            { Matricula: '003', Nome: 'Pedro Costa', Funcao: 'Operador' },
            { Matricula: '004', Nome: 'Ana Oliveira', Funcao: 'Analista' },
            { Matricula: '005', Nome: 'Carlos Souza', Funcao: 'Operador' },
            { Matricula: '006', Nome: 'Julia Lima', Funcao: 'Coordenadora' },
            { Matricula: '007', Nome: 'Roberto Ferreira', Funcao: 'Técnico' },
            { Matricula: '008', Nome: 'Fernanda Alves', Funcao: 'Operadora' }
          ]);
        }
        
        await sheetsService.getSheet('Lista');
        await sheetsService.getSheet('Base');
        
        return res.status(200).json({ 
          ok: true, 
          msg: 'Dados de teste criados com sucesso!' 
        });
        
      } catch (error) {
        console.error('Erro ao criar dados de teste:', error);
        return res.status(500).json({ 
          ok: false, 
          msg: 'Erro ao criar dados de teste: ' + error.message 
        });
      }
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
}
