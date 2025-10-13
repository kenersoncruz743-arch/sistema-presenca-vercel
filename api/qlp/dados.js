// api/qlp/quadro.js - API para buscar dados reais da aba Quadro
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
    console.log('[QLP/QUADRO] Iniciando busca na aba Quadro...');
    
    // Inicializa conexão com Google Sheets
    const doc = await sheetsService.init();
    console.log(`[QLP/QUADRO] Conectado à planilha: ${doc.title}`);
    
    // Busca a aba Quadro
    const sheetQuadro = doc.sheetsByTitle['Quadro'];
    if (!sheetQuadro) {
      console.error('[QLP/QUADRO] Aba Quadro não encontrada');
      return res.status(404).json({
        ok: false,
        msg: 'Aba Quadro não encontrada na planilha'
      });
    }
    
    // Carrega todos os registros
    const rows = await sheetQuadro.getRows();
    console.log(`[QLP/QUADRO] ${rows.length} registros encontrados`);
    
    // Processa os dados
    const colaboradores = [];
    const estatisticas = {
      total: 0,
      ativos: 0,
      inativos: 0,
      porSecao: {},
      porTurno: {},
      porSupervisor: {},
      porSituacao: {},
      porFuncao: {}
    };
    
    rows.forEach(row => {
      // Lê as colunas reais da planilha
      const filial = String(row.get('FILIAL') || '').trim();
      const bandeira = String(row.get('BANDEIRA') || '').trim();
      const chapa = String(row.get('CHAPA1') || '').trim();
      const dtAdmissao = String(row.get('DT_ADMISSAO') || '').trim();
      const nome = String(row.get('NOME') || '').trim();
      const funcao = String(row.get('FUNCAO') || '').trim();
      const secao = String(row.get('SECAO') || '').trim();
      const situacao = String(row.get('SITUACAO') || '').trim();
      const supervisor = String(row.get('Supervisor') || '').trim();
      const turno = String(row.get('Turno') || 'Não definido').trim();
      const gestao = String(row.get('Gestão') || '').trim();
      
      // Ignora linhas sem nome
      if (!nome) return;
      
      // Determina se está ativo
      const isAtivo = situacao.toLowerCase().includes('ativo') && 
                     !situacao.toLowerCase().includes('não ativo');
      
      // Adiciona à lista
      colaboradores.push({
        filial,
        bandeira,
        chapa,
        dtAdmissao,
        nome,
        funcao,
        secao,
        situacao,
        supervisor,
        turno,
        gestao,
        isAtivo
      });
      
      // Atualiza estatísticas
      estatisticas.total++;
      
      if (isAtivo) {
        estatisticas.ativos++;
      } else {
        estatisticas.inativos++;
      }
      
      // Por seção
      if (secao) {
        if (!estatisticas.porSecao[secao]) {
          estatisticas.porSecao[secao] = { total: 0, ativos: 0, inativos: 0 };
        }
        estatisticas.porSecao[secao].total++;
        if (isAtivo) {
          estatisticas.porSecao[secao].ativos++;
        } else {
          estatisticas.porSecao[secao].inativos++;
        }
      }
      
      // Por turno
      const turnoNormalizado = turno || 'Não definido';
      if (!estatisticas.porTurno[turnoNormalizado]) {
        estatisticas.porTurno[turnoNormalizado] = { total: 0, ativos: 0 };
      }
      estatisticas.porTurno[turnoNormalizado].total++;
      if (isAtivo) {
        estatisticas.porTurno[turnoNormalizado].ativos++;
      }
      
      // Por supervisor
      const supervisorNormalizado = supervisor || 'Sem supervisor';
      if (!estatisticas.porSupervisor[supervisorNormalizado]) {
        estatisticas.porSupervisor[supervisorNormalizado] = { total: 0, ativos: 0 };
      }
      estatisticas.porSupervisor[supervisorNormalizado].total++;
      if (isAtivo) {
        estatisticas.porSupervisor[supervisorNormalizado].ativos++;
      }
      
      // Por situação
      const situacaoNormalizada = situacao || 'Sem situação';
      if (!estatisticas.porSituacao[situacaoNormalizada]) {
        estatisticas.porSituacao[situacaoNormalizada] = 0;
      }
      estatisticas.porSituacao[situacaoNormalizada]++;
      
      // Por função
      if (funcao) {
        if (!estatisticas.porFuncao[funcao]) {
          estatisticas.porFuncao[funcao] = { total: 0, ativos: 0 };
        }
        estatisticas.porFuncao[funcao].total++;
        if (isAtivo) {
          estatisticas.porFuncao[funcao].ativos++;
        }
      }
    });
    
    // Adiciona contagens de categorias únicas
    estatisticas.totalSecoes = Object.keys(estatisticas.porSecao).length;
    estatisticas.totalTurnos = Object.keys(estatisticas.porTurno).length;
    estatisticas.totalSupervisores = Object.keys(estatisticas.porSupervisor).length;
    estatisticas.totalFuncoes = Object.keys(estatisticas.porFuncao).length;
    
    console.log(`[QLP/QUADRO] Processados ${colaboradores.length} colaboradores`);
    console.log(`[QLP/QUADRO] ${estatisticas.ativos} ativos, ${estatisticas.inativos} inativos`);
    console.log(`[QLP/QUADRO] ${estatisticas.totalSecoes} seções diferentes`);
    
    return res.status(200).json({
      ok: true,
      colaboradores: colaboradores,
      estatisticas: estatisticas,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('[QLP/QUADRO] Erro ao buscar dados:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro ao buscar dados do Quadro',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
