// api/qlp/quadro.js e api/qlp/dados.js - COM DETECÇÃO AUTOMÁTICA DE COLUNAS
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
    
    // Busca aba
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
    console.log('[QLP/QUADRO] Headers encontrados:', headers);
    
    // NOVO: Mapeia os nomes reais das colunas
    const colunaMap = {};
    headers.forEach(header => {
      const headerNorm = header.trim().toUpperCase();
      
      // Detecta CHAPA
      if (headerNorm.includes('CHAPA')) {
        colunaMap.chapa = header;
      }
      // Detecta NOME
      if (headerNorm === 'NOME') {
        colunaMap.nome = header;
      }
      // Detecta FUNÇÃO
      if (headerNorm.includes('FUNCAO') || headerNorm.includes('FUNÇÃO')) {
        colunaMap.funcao = header;
      }
      // Detecta SEÇÃO
      if (headerNorm.includes('SECAO') || headerNorm.includes('SEÇÃO')) {
        colunaMap.secao = header;
      }
      // Detecta SITUAÇÃO
      if (headerNorm.includes('SITUACAO') || headerNorm.includes('SITUAÇÃO')) {
        colunaMap.situacao = header;
      }
      // Detecta TURNO
      if (headerNorm === 'TURNO') {
        colunaMap.turno = header;
      }
      // Detecta SUPERVISOR
      if (headerNorm.includes('SUPERVISOR')) {
        colunaMap.supervisor = header;
      }
      // Detecta FILIAL
      if (headerNorm === 'FILIAL') {
        colunaMap.filial = header;
      }
      // Detecta BANDEIRA
      if (headerNorm === 'BANDEIRA') {
        colunaMap.bandeira = header;
      }
      // Detecta DATA ADMISSÃO
      if (headerNorm.includes('ADMISSAO') || headerNorm.includes('ADMISSÃO') || headerNorm.includes('DT_')) {
        colunaMap.dtAdmissao = header;
      }
      // Detecta GESTÃO
      if (headerNorm.includes('GESTAO') || headerNorm.includes('GESTÃO')) {
        colunaMap.gestao = header;
      }
    });
    
    console.log('[QLP/QUADRO] Mapeamento de colunas:', colunaMap);
    
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
        // Função helper MELHORADA
        const getCol = (colKey) => {
          try {
            const colName = colunaMap[colKey];
            if (!colName) return '';
            const valor = row.get(colName);
            return valor ? String(valor).trim() : '';
          } catch (e) {
            return '';
          }
        };
        
        // Lê dados usando o mapeamento
        const filial = getCol('filial');
        const bandeira = getCol('bandeira');
        const chapa = getCol('chapa');
        const dtAdmissao = getCol('dtAdmissao');
        const nome = getCol('nome');
        const funcao = getCol('funcao');
        const secao = getCol('secao');
        const situacao = getCol('situacao');
        const supervisor = getCol('supervisor');
        const turno = getCol('turno');
        const gestao = getCol('gestao');
        
        // Debug primeira linha
        if (index === 0) {
          console.log('[QLP/QUADRO] EXEMPLO linha 1:', {
            chapa,
            nome,
            funcao,
            secao,
            situacao,
            supervisor,
            turno
          });
        }
        
        // Ignora sem nome
        if (!nome) {
          linhasIgnoradas++;
          return;
        }
        
        // Determina ativo
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
          const turnoUpper = turno.toUpperCase();
          if (turnoUpper.includes('TURNO A') || turnoUpper === 'TURNO A') {
            turnoNormalizado = 'Turno A';
          } else if (turnoUpper.includes('TURNO B') || turnoUpper === 'TURNO B') {
            turnoNormalizado = 'Turno B';
          } else if (turnoUpper.includes('TURNO C') || turnoUpper === 'TURNO C') {
            turnoNormalizado = 'Turno C';
          } else if (!turnoUpper.includes('NÃO') && !turnoUpper.includes('DEFINIDO')) {
            turnoNormalizado = turno;
          }
        }
        
        // Normaliza supervisor
        const supervisorNormalizado = supervisor && supervisor !== '' ? supervisor : 'Sem supervisor';
        
        // Adiciona colaborador
        colaboradores.push({
          filial,
          bandeira,
          chapa: chapa || 'S/N',
          dtAdmissao,
          nome,
          funcao,
          secao,
          situacao,
          supervisor: supervisorNormalizado,
          turno: turnoNormalizado,
          gestao,
          isAtivo
        });
        
        linhasProcessadas++;
        
        // Estatísticas
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
        linhasComErro++;
        console.error(`[QLP/QUADRO] ✗ Erro linha ${index + 1}:`, rowError.message);
      }
    });
    
    // Totais
    estatisticas.totalSecoes = Object.keys(estatisticas.porSecao).length;
    estatisticas.totalTurnos = Object.keys(estatisticas.porTurno).length;
    estatisticas.totalSupervisores = Object.keys(estatisticas.porSupervisor).length;
    estatisticas.totalFuncoes = Object.keys(estatisticas.porFuncao).length;
    
    console.log('[QLP/QUADRO] ========== RESUMO ==========');
    console.log(`[QLP/QUADRO] ✓ Processados: ${linhasProcessadas}`);
    console.log(`[QLP/QUADRO] ⚠ Ignoradas: ${linhasIgnoradas}`);
    console.log(`[QLP/QUADRO] ✗ Erros: ${linhasComErro}`);
    console.log(`[QLP/QUADRO] Ativos: ${estatisticas.ativos} | Inativos: ${estatisticas.inativos}`);
    console.log('[QLP/QUADRO] Turnos:', Object.keys(estatisticas.porTurno));
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
        mapeamentoColunas: colunaMap
      },
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('[QLP/QUADRO] ========== ERRO ==========');
    console.error('[QLP/QUADRO]', error.message);
    console.error('[QLP/QUADRO]', error.stack);
    
    return res.status(500).json({
      ok: false,
      msg: 'Erro ao buscar dados',
      error: error.name,
      details: error.message
    });
  }
};
