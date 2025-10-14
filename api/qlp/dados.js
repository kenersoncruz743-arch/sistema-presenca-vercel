// api/qlp/quadro.js e api/qlp/dados.js - MAPEAMENTO EXATO DAS COLUNAS
// USE ESTE CÓDIGO EM AMBOS OS ARQUIVOS
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
    console.log('[QLP/QUADRO] ========== INÍCIO ==========');
    
    // Inicializa conexão
    const doc = await sheetsService.init();
    console.log(`[QLP/QUADRO] ✓ Conectado: ${doc.title}`);
    
    // Busca aba Quadro
    const sheetQuadro = doc.sheetsByTitle['Quadro'];
    if (!sheetQuadro) {
      return res.status(404).json({
        ok: false,
        msg: 'Aba Quadro não encontrada',
        abasDisponiveis: Object.keys(doc.sheetsByTitle)
      });
    }
    
    // Carrega headers
    await sheetQuadro.loadHeaderRow();
    const headers = sheetQuadro.headerValues;
    console.log('[QLP/QUADRO] Headers da planilha:', headers);
    
    // Carrega linhas
    const rows = await sheetQuadro.getRows();
    console.log(`[QLP/QUADRO] ✓ ${rows.length} linhas carregadas`);
    
    if (rows.length === 0) {
      return res.status(200).json({
        ok: true,
        colaboradores: [],
        estatisticas: {
          total: 0,
          ativos: 0,
          inativos: 0,
          porSecao: {},
          porTurno: {},
          porSupervisor: {},
          porSituacao: {},
          porFuncao: {},
          totalSecoes: 0,
          totalTurnos: 0,
          totalSupervisores: 0,
          totalFuncoes: 0
        },
        timestamp: new Date().toISOString()
      });
    }
    
    // Processa dados
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
    
    let linhasProcessadas = 0;
    let linhasIgnoradas = 0;
    let linhasComErro = 0;
    
    rows.forEach((row, index) => {
      try {
        // Função helper para ler coluna com segurança
        const getCol = (colName) => {
          try {
            const valor = row.get(colName);
            return valor ? String(valor).trim() : '';
          } catch (e) {
            console.error(`[QLP/QUADRO] Erro ao ler coluna "${colName}":`, e.message);
            return '';
          }
        };
        
        // Lê as colunas COM OS NOMES EXATOS
        const filial = getCol('FILIAL');
        const bandeira = getCol('BANDEIRA');
        const chapa = getCol('CHAPA1');
        const dtAdmissao = getCol('DT_ADMISSAO');
        const nome = getCol('NOME');
        const funcao = getCol('FUNCAO');
        const secao = getCol('SECAO');
        const situacao = getCol('SITUACAO');
        const totalGeral = getCol('Total Geral');
        const supervisor = getCol('Supervisor');
        const turno = getCol('Turno');
        const gestao = getCol('Gestão');
        
        // Debug primeiras 3 linhas
        if (index < 3) {
          console.log(`[QLP/QUADRO] Linha ${index + 1}:`, {
            chapa,
            nome,
            funcao,
            secao,
            situacao,
            supervisor,
            turno
          });
        }
        
        // Ignora linhas sem nome
        if (!nome || nome === '') {
          linhasIgnoradas++;
          return;
        }
        
        // Determina se está ativo
        const situacaoLower = situacao.toLowerCase();
        const isAtivo = situacao === 'Ativo' || 
                       (situacaoLower.includes('ativo') && 
                        !situacaoLower.includes('não') && 
                        !situacaoLower.includes('af.') &&
                        !situacaoLower.includes('aviso') &&
                        !situacaoLower.includes('previdência'));
        
        // Normaliza turno
        let turnoNormalizado = 'Não definido';
        if (turno && turno !== '') {
          const turnoUpper = turno.toUpperCase().trim();
          
          if (turnoUpper === 'TURNO A') {
            turnoNormalizado = 'Turno A';
          } else if (turnoUpper === 'TURNO B') {
            turnoNormalizado = 'Turno B';
          } else if (turnoUpper === 'TURNO C') {
            turnoNormalizado = 'Turno C';
          } else if (turnoUpper.includes('NÃO DEFINIDO') || turnoUpper.includes('NAO DEFINIDO')) {
            turnoNormalizado = 'Não definido';
          } else if (turnoUpper.includes('NÃO ATIVO') || turnoUpper.includes('NAO ATIVO')) {
            turnoNormalizado = 'Não ativo';
          } else {
            // Mantém valor original se não for vazio
            turnoNormalizado = turno;
          }
        }
        
        // Normaliza supervisor
        const supervisorNormalizado = supervisor && supervisor !== '' ? supervisor : 'Sem supervisor';
        
        // Adiciona à lista de colaboradores
        colaboradores.push({
          filial,
          bandeira,
          chapa: chapa || 'S/N',
          dtAdmissao,
          nome,
          funcao,
          secao,
          situacao,
          totalGeral,
          supervisor: supervisorNormalizado,
          turno: turnoNormalizado,
          gestao,
          isAtivo
        });
        
        linhasProcessadas++;
        
        // Atualiza estatísticas
        estatisticas.total++;
        
        if (isAtivo) {
          estatisticas.ativos++;
        } else {
          estatisticas.inativos++;
        }
        
        // Estatísticas por seção
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
        
        // Estatísticas por turno
        if (!estatisticas.porTurno[turnoNormalizado]) {
          estatisticas.porTurno[turnoNormalizado] = { total: 0, ativos: 0 };
        }
        estatisticas.porTurno[turnoNormalizado].total++;
        if (isAtivo) {
          estatisticas.porTurno[turnoNormalizado].ativos++;
        }
        
        // Estatísticas por supervisor
        if (!estatisticas.porSupervisor[supervisorNormalizado]) {
          estatisticas.porSupervisor[supervisorNormalizado] = { total: 0, ativos: 0 };
        }
        estatisticas.porSupervisor[supervisorNormalizado].total++;
        if (isAtivo) {
          estatisticas.porSupervisor[supervisorNormalizado].ativos++;
        }
        
        // Estatísticas por situação
        const situacaoNormalizada = situacao || 'Sem situação';
        if (!estatisticas.porSituacao[situacaoNormalizada]) {
          estatisticas.porSituacao[situacaoNormalizada] = 0;
        }
        estatisticas.porSituacao[situacaoNormalizada]++;
        
        // Estatísticas por função
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
        linhasComErro++;
        console.error(`[QLP/QUADRO] ✗ Erro ao processar linha ${index + 1}:`, rowError.message);
      }
    });
    
    // Totais de categorias únicas
    estatisticas.totalSecoes = Object.keys(estatisticas.porSecao).length;
    estatisticas.totalTurnos = Object.keys(estatisticas.porTurno).length;
    estatisticas.totalSupervisores = Object.keys(estatisticas.porSupervisor).length;
    estatisticas.totalFuncoes = Object.keys(estatisticas.porFuncao).length;
    
    console.log('[QLP/QUADRO] ========== RESUMO ==========');
    console.log(`[QLP/QUADRO] ✓ Processados: ${linhasProcessadas} colaboradores`);
    console.log(`[QLP/QUADRO] ⚠ Ignoradas: ${linhasIgnoradas} linhas (sem nome)`);
    console.log(`[QLP/QUADRO] ✗ Com erro: ${linhasComErro} linhas`);
    console.log(`[QLP/QUADRO] Status: ${estatisticas.ativos} ativos, ${estatisticas.inativos} inativos`);
    console.log(`[QLP/QUADRO] Seções: ${estatisticas.totalSecoes} diferentes`);
    console.log(`[QLP/QUADRO] Turnos: ${estatisticas.totalTurnos} diferentes`);
    console.log('[QLP/QUADRO] Distribuição por turno:', estatisticas.porTurno);
    console.log('[QLP/QUADRO] ========== FIM ==========');
    
    return res.status(200).json({
      ok: true,
      colaboradores,
      estatisticas,
      debug: {
        linhasProcessadas,
        linhasIgnoradas,
        linhasComErro,
        totalLinhas: rows.length,
        headersEncontrados: headers
      },
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('[QLP/QUADRO] ========== ERRO FATAL ==========');
    console.error('[QLP/QUADRO] Tipo:', error.name);
    console.error('[QLP/QUADRO] Mensagem:', error.message);
    console.error('[QLP/QUADRO] Stack:', error.stack);
    console.error('[QLP/QUADRO] =====================================');
    
    return res.status(500).json({
      ok: false,
      msg: 'Erro ao buscar dados do Quadro',
      error: error.name,
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
