// api/qlp/exportar.js - API para exportar dados do QLP
const sheetsService = require('../../lib/sheets');

module.exports = async function handler(req, res) {
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
    const { formato } = req.query;
    
    if (!formato || !['csv', 'json'].includes(formato)) {
      return res.status(400).json({
        ok: false,
        msg: 'Formato inválido. Use csv ou json'
      });
    }
    
    console.log(`[QLP/EXPORTAR] Iniciando exportação em formato: ${formato}`);
    
    // Inicializa conexão com Google Sheets
    const doc = await sheetsService.init();
    console.log(`[QLP/EXPORTAR] Conectado à planilha: ${doc.title}`);
    
    // Busca a aba Base
    const sheetBase = doc.sheetsByTitle['Base'];
    if (!sheetBase) {
      return res.status(404).json({
        ok: false,
        msg: 'Aba Base não encontrada'
      });
    }
    
    // Busca a aba Quadro
    const sheetQuadro = doc.sheetsByTitle['Quadro'];
    
    // Carrega registros
    const rows = await sheetBase.getRows();
    console.log(`[QLP/EXPORTAR] ${rows.length} registros para exportar`);
    
    // Mapa do Quadro
    const quadroMap = {};
    if (sheetQuadro) {
      const quadroRows = await sheetQuadro.getRows();
      quadroRows.forEach(row => {
        const matricula = String(row.get('Coluna 1') || '').trim();
        const tipoOperacional = String(row.get('Coluna 2') || '').trim();
        if (matricula) {
          quadroMap[matricula] = tipoOperacional || 'Não operacional';
        }
      });
    }
    
    // Processa dados
    const dadosExportacao = [];
    
    rows.forEach(row => {
      const supervisor = String(row.get('Supervisor') || '').trim();
      const aba = String(row.get('Aba') || '').trim();
      const matricula = String(row.get('Matricula') || '').trim();
      const nome = String(row.get('Nome') || '').trim();
      const funcao = String(row.get('Função') || '').trim();
      const status = String(row.get('Status') || '').trim();
      const data = String(row.get('Data') || '').trim();
      
      if (!nome) return;
      
      // Determina turno
      let turno = 'Não definido';
      if (aba.toLowerCase().includes('ta') || aba.toLowerCase().includes('turno a')) {
        turno = 'Turno A';
      } else if (aba.toLowerCase().includes('tb') || aba.toLowerCase().includes('turno b')) {
        turno = 'Turno B';
      } else if (aba.toLowerCase().includes('tc') || aba.toLowerCase().includes('turno c')) {
        turno = 'Turno C';
      }
      
      const tipoOperacional = quadroMap[matricula] || 'Não operacional';
      
      dadosExportacao.push({
        Matricula: matricula,
        Nome: nome,
        Funcao: funcao,
        Supervisor: supervisor,
        Aba: aba,
        Turno: turno,
        TipoOperacional: tipoOperacional,
        Status: status,
        Data: data
      });
    });
    
    if (formato === 'csv') {
      // Gera CSV
      const headers = ['Matricula', 'Nome', 'Funcao', 'Supervisor', 'Aba', 'Turno', 'TipoOperacional', 'Status', 'Data'];
      let csv = headers.join(',') + '\n';
      
      dadosExportacao.forEach(item => {
        const linha = headers.map(header => {
          let valor = item[header] || '';
          // Escapa vírgulas e aspas
          if (valor.includes(',') || valor.includes('"')) {
            valor = '"' + valor.replace(/"/g, '""') + '"';
          }
          return valor;
        }).join(',');
        csv += linha + '\n';
      });
      
      console.log(`[QLP/EXPORTAR] CSV gerado com ${dadosExportacao.length} linhas`);
      
      return res.status(200).json({
        ok: true,
        formato: 'csv',
        csv: csv,
        totalRegistros: dadosExportacao.length
      });
      
    } else {
      // Retorna JSON
      console.log(`[QLP/EXPORTAR] JSON gerado com ${dadosExportacao.length} registros`);
      
      return res.status(200).json({
        ok: true,
        formato: 'json',
        dados: dadosExportacao,
        totalRegistros: dadosExportacao.length
      });
    }
    
  } catch (error) {
    console.error('[QLP/EXPORTAR] Erro ao exportar:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro ao exportar dados',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
