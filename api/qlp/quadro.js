// api/qlp/quadro.js e api/qlp/dados.js - VERSÃO CORRIGIDA COM LEITURA CORRETA DAS COLUNAS
// IMPORTANTE: Este mesmo código deve ser usado em AMBOS os arquivos
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
    
    // Debug: mostra os headers disponíveis
    if (rows.length > 0) {
      console.log('[QLP/QUADRO] Headers disponíveis:', rows[0]._sheet.headerValues);
    }
    
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
    
    rows.forEach((row, index) => {
      try {
        // Lê as colunas EXATAMENTE como estão no seu arquivo
        const filial = String(row.get('FILIAL') || '').trim();
        const bandeira = String(row.get('BANDEIRA') || '').trim();
        
        // CORREÇÃO: Usa CHAPA1 em vez de CHAPA
        const chapa = String(row.get('CHAPA1') || '').trim();
        
        const dtAdmissao = String(row.get('DT_ADMISSAO') || '').trim();
        const nome = String(row.get('NOME') || '').trim();
        const funcao = String(row.get('FUNCAO') || '').trim();
        const secao = String(row.get('SECAO') || '').trim();
        const situacao = String(row.get('SITUACAO') || '').trim();
        
        // CORREÇÃO: Lê Supervisor e Turno corretamente
        const supervisor = String(row.get('Supervisor') || '').trim();
        const turno = String(row.get('Turno') || '').trim();
        const gestao = String(row.get('Gestão') || '').trim();
        
        // Debug para primeira linha
        if (index === 0) {
          console.log('[QLP/QUADRO] Exemplo de dados lidos:', {
            chapa, nome, funcao, secao, situacao, supervisor, turno
          });
        }
        
        // Ignora linhas sem nome
        if (!nome) {
          console.log(`[QLP/QUADRO] Linha ${index + 1} ignorada: sem nome`);
          return;
        }
        
        // Determina se está ativo
        // CORREÇÃO: Melhora a lógica de detecção de ativo
        const situacaoLower = situacao.toLowerCase();
        const isAtivo = situacaoLower === 'ativo' || 
                       (situacaoLower.includes('ativo') && !situacaoLower.includes('não') && !situacaoLower.includes('af.'));
        
        // CORREÇÃO: Normaliza o turno
        let turnoNormalizado = turno || 'Não definido';
        if (!turno || turno === '' || turno.toLowerCase().includes('não')) {
          turnoNormalizado = 'Não definido';
        } else if (turno.toLowerCase().includes('turno a')) {
          turnoNormalizado = 'Turno A';
        } else if (turno.toLowerCase().includes('turno b')) {
          turnoNormalizado = 'Turno B';
        } else if (turno.toLowerCase().includes('turno c')) {
          turnoNormalizado = 'Turno C';
        }
        
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
          supervisor: supervisor || 'Sem supervisor',
          turno: turnoNormalizado,
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
      } catch (rowError) {
        console.error(`[QLP/QUADRO] Erro ao processar linha ${index + 1}:`, rowError);
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
    console.log(`[QLP/QUADRO] Distribuição por turno:`, estatisticas.porTurno);
    
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
