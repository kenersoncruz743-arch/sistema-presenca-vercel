// api/qlp/dados.js - VERSÃO CORRIGIDA PARA ABA "QLP"
const sheetsService = require('../../lib/sheets');

module.exports = async function handler(req, res) {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'GET') {
    return res.status(405).json({ 
      ok: false, 
      msg: 'Método não permitido' 
    });
  }

  try {
    console.log('[QLP/DADOS] ========== INÍCIO DA REQUISIÇÃO ==========');
    console.log('[QLP/DADOS] Iniciando busca na aba QLP...');
    
    // Inicializa conexão com Google Sheets
    let doc;
    try {
      doc = await sheetsService.init();
      console.log(`[QLP/DADOS] ✓ Conectado à planilha: ${doc.title}`);
    } catch (initError) {
      console.error('[QLP/DADOS] ✗ Erro ao inicializar Google Sheets:', initError);
      return res.status(500).json({
        ok: false,
        msg: 'Erro ao conectar com Google Sheets',
        details: initError.message
      });
    }
    
    // Busca a aba QLP (CORRIGIDO)
    const sheetQLP = doc.sheetsByTitle['QLP'];
    if (!sheetQLP) {
      console.error('[QLP/DADOS] ✗ Aba QLP não encontrada');
      console.log('[QLP/DADOS] Abas disponíveis:', Object.keys(doc.sheetsByTitle));
      return res.status(404).json({
        ok: false,
        msg: 'Aba QLP não encontrada na planilha',
        abasDisponiveis: Object.keys(doc.sheetsByTitle)
      });
    }
    
    console.log(`[QLP/DADOS] ✓ Aba QLP encontrada`);
    
    // Carrega todos os registros
    let rows;
    try {
      rows = await sheetQLP.getRows();
      console.log(`[QLP/DADOS] ✓ ${rows.length} registros carregados`);
    } catch (rowsError) {
      console.error('[QLP/DADOS] ✗ Erro ao carregar linhas:', rowsError);
      return res.status(500).json({
        ok: false,
        msg: 'Erro ao carregar dados da planilha',
        details: rowsError.message
      });
    }
    
    if (rows.length === 0) {
      console.warn('[QLP/DADOS] ⚠ Nenhuma linha encontrada na planilha');
      return res.status(200).json({
        ok: true,
        dados: [],
        total: 0,
        timestamp: new Date().toISOString()
      });
    }

    // Debug: mostra os headers disponíveis
    const headers = rows[0]._sheet.headerValues;
    console.log('[QLP/DADOS] Headers disponíveis:', headers);
    
    // Processa os dados
    const dados = [];
    
    rows.forEach((row, index) => {
      try {
        const getCol = (colName) => {
          try {
            return String(row.get(colName) || '').trim();
          } catch (e) {
            return '';
          }
        };
        
        const nome = getCol('NOME') || getCol('Nome');
        
        if (!nome || nome === '') {
          return; // Ignora linhas sem nome
        }
        
        dados.push({
          filial: getCol('FILIAL'),
          bandeira: getCol('BANDEIRA'),
          chapa: getCol('CHAPA1') || getCol('CHAPA') || getCol('Chapa') || 'S/N',
          dtAdmissao: getCol('DT_ADMISSAO') || getCol('Data Admissão'),
          nome: nome,
          funcao: getCol('FUNCAO') || getCol('Função'),
          secao: getCol('SECAO') || getCol('Seção'),
          situacao: getCol('SITUACAO') || getCol('Situação'),
          supervisor: getCol('Supervisor') || 'Sem supervisor',
          turno: getCol('Turno') || 'Não definido',
          gestao: getCol('Gestão')
        });
        
      } catch (rowError) {
        console.error(`[QLP/DADOS] ✗ Erro ao processar linha ${index + 1}:`, rowError.message);
      }
    });
    
    console.log(`[QLP/DADOS] ✓ ${dados.length} registros processados`);
    console.log('[QLP/DADOS] ========== FIM DA REQUISIÇÃO ==========');
    
    return res.status(200).json({
      ok: true,
      dados: dados,
      total: dados.length,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('[QLP/DADOS] ========== ERRO FATAL ==========');
    console.error('[QLP/DADOS] Tipo:', error.name);
    console.error('[QLP/DADOS] Mensagem:', error.message);
    console.error('[QLP/DADOS] Stack:', error.stack);
    console.error('[QLP/DADOS] =====================================');
    
    return res.status(500).json({
      ok: false,
      msg: 'Erro ao buscar dados do QLP',
      error: error.name,
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
