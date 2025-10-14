// api/test-connection.js - API SIMPLIFICADA para testar conexão
const sheetsService = require('../lib/sheets');

// Teste das variáveis de ambiente primeiro
console.log('[TEST] === VERIFICANDO VARIÁVEIS DE AMBIENTE ===');
console.log('[TEST] GOOGLE_SHEETS_ID:', process.env.GOOGLE_SHEETS_ID ? '✓ Definido' : '✗ NÃO DEFINIDO');
console.log('[TEST] GOOGLE_SHEETS_CLIENT_EMAIL:', process.env.GOOGLE_SHEETS_CLIENT_EMAIL ? '✓ Definido' : '✗ NÃO DEFINIDO');
console.log('[TEST] GOOGLE_SHEETS_PRIVATE_KEY:', process.env.GOOGLE_SHEETS_PRIVATE_KEY ? '✓ Definido (primeira parte: ' + process.env.GOOGLE_SHEETS_PRIVATE_KEY.substring(0, 50) + '...)' : '✗ NÃO DEFINIDO');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    console.log('[TEST] Iniciando teste de conexão...');
    
    // Testa inicialização
    const doc = await sheetsService.init();
    console.log(`[TEST] ✓ Conectado: ${doc.title}`);
    
    // Lista todas as abas
    const abas = Object.keys(doc.sheetsByTitle);
    console.log('[TEST] Abas encontradas:', abas);
    
    // Testa aba Quadro especificamente
    const sheetQuadro = doc.sheetsByTitle['Quadro'];
    
    if (!sheetQuadro) {
      return res.status(200).json({
        ok: false,
        msg: 'Aba Quadro não encontrada',
        planilha: doc.title,
        abasDisponiveis: abas
      });
    }
    
    // Pega informações da aba
    await sheetQuadro.loadHeaderRow();
    const headers = sheetQuadro.headerValues;
    
    // Pega primeiras 3 linhas como amostra
    const rows = await sheetQuadro.getRows({ limit: 3 });
    
    const amostras = rows.map((row, idx) => {
      const dados = {};
      headers.forEach(header => {
        try {
          dados[header] = row.get(header) || '';
        } catch (e) {
          dados[header] = 'ERRO';
        }
      });
      return dados;
    });
    
    console.log('[TEST] ✓ Headers:', headers);
    console.log('[TEST] ✓ Total de linhas:', sheetQuadro.rowCount);
    
    return res.status(200).json({
      ok: true,
      planilha: doc.title,
      aba: 'Quadro',
      totalAbas: abas.length,
      abasDisponiveis: abas,
      totalLinhas: sheetQuadro.rowCount,
      totalColunas: headers.length,
      headers: headers,
      amostrasDados: amostras,
      verificacoes: {
        temCHAPA1: headers.includes('CHAPA1'),
        temNOME: headers.includes('NOME'),
        temTurno: headers.includes('Turno'),
        temSupervisor: headers.includes('Supervisor'),
        temSITUACAO: headers.includes('SITUACAO'),
        temSECAO: headers.includes('SECAO')
      }
    });
    
  } catch (error) {
    console.error('[TEST] Erro:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro no teste de conexão',
      error: error.message,
      stack: error.stack
    });
  }
};
