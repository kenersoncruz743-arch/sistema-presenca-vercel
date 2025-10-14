// api/qlp/quadro.js - VERSÃO ROBUSTA COM TRATAMENTO COMPLETO DE ERROS
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
    console.log('[QLP/QUADRO] ========== INÍCIO DA REQUISIÇÃO ==========');
    console.log('[QLP/QUADRO] Iniciando busca na aba Quadro...');
    
    // Inicializa conexão com Google Sheets
    let doc;
    try {
      doc = await sheetsService.init();
      console.log(`[QLP/QUADRO] ✓ Conectado à planilha: ${doc.title}`);
    } catch (initError) {
      console.error('[QLP/QUADRO] ✗ Erro ao inicializar Google Sheets:', initError);
      return res.status(500).json({
        ok: false,
        msg: 'Erro ao conectar com Google Sheets',
        details: initError.message
      });
    }
    
    // Busca a aba Quadro
    const sheetQuadro = doc.sheetsByTitle['Quadro'];
    if (!sheetQuadro) {
      console.error('[QLP/QUADRO] ✗ Aba Quadro não encontrada');
      console.log('[QLP/QUADRO] Abas disponíveis:', Object.keys(doc.sheetsByTitle));
      return res.status(404).json({
        ok: false,
        msg: 'Aba Quadro não encontrada na planilha',
        abasDisponiveis: Object.keys(doc.sheetsByTitle)
      });
    }
    
    console.log(`[QLP/QUADRO] ✓ Aba Quadro encontrada`);
    
    // Carrega todos os registros
    let rows;
    try {
      rows = await sheetQuadro.getRows();
      console.log(`[QLP/QUADRO] ✓ ${rows.length} registros carregados`);
    } catch (rowsError) {
      console.error('[QLP/QUADRO] ✗ Erro ao carregar linhas:', rowsError);
      return res.status(500).json({
        ok: false,
        msg: 'Erro ao carregar dados da planilha',
        details: rowsError.message
      });
    }
    
    // Debug: mostra os headers disponíveis
    if (rows.length > 0) {
      const headers = rows[0]._sheet.headerValues;
      console.log('[QLP/QUADRO] Headers disponíveis:', headers);
      console.log('[QLP/QUADRO] Total de colunas:', headers.length);
    } else {
      console.warn('[QLP/QUADRO] ⚠ Nenhuma linha encontrada na planilha');
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
    
    let linhasProcessadas = 0;
    let linhasComErro = 0;
    let linhasIgnoradas = 0;
    
    rows.forEach((row, index) => {
      try {
        // Função helper para ler coluna com segurança
        const getCol = (colName) => {
          try {
            return String(row.get(colName) || '').trim();
          } catch (e) {
            return '';
          }
        };
        
        // Lê as colunas
        const filial = getCol('FILIAL');
        const bandeira = getCol('BANDEIRA');
        const chapa = getCol('CHAPA1');
        const dtAdmissao = getCol('DT_ADMISSAO');
        const nome = getCol('NOME');
        const funcao = getCol('FUNCAO');
        const secao = getCol('SECAO');
        const situacao = getCol('SITUACAO');
        const supervisor = getCol('Supervisor');
        const turno = getCol('Turno');
        const gestao = getCol('Gestão');
        
        // Debug para primeira linha
        if (index === 0) {
          console.log('[QLP/QUADRO] Exemplo de dados (linha 1):', {
            chapa, nome, funcao, secao, situacao, supervisor, turno
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
                        !situacaoLower.includes('aviso'));
        
        // Normaliza o turno
        let turnoNormalizado = 'Não definido';
        if (turno && turno !== '') {
          const turnoLower = turno.toLowerCase();
          if (turnoLower.includes('turno a') || turnoLower === 'turno a') {
            turnoNormalizado = 'Turno A';
          } else if (turnoLower.includes('turno b') || turnoLower === 'turno b') {
            turnoNormalizado = 'Turno B';
          } else if (turnoLower.includes('turno c') || turnoLower === 'turno c') {
            turnoNormalizado = 'Turno C';
          } else if (!turnoLower.includes('não') && !turnoLower.includes('definido')) {
            turnoNormalizado = turno; // Mantém o valor original se não for vazio
          }
        }
        
        // Normaliza supervisor
        const supervisorNormalizado = supervisor && supervisor !== '' ? supervisor : 'Sem supervisor';
        
        // Adiciona à lista
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
        console.error(`[QLP/QUADRO] ✗ Erro ao processar linha ${index + 1}:`, rowError.message);
      }
    });
    
    // Adiciona contagens de categorias únicas
    estatisticas.totalSecoes = Object.keys(estatisticas.porSecao).length;
    estatisticas.totalTurnos = Object.keys(estatisticas.porTurno).length;
    estatisticas.totalSupervisores = Object.keys(estatisticas.porSupervisor).length;
    estatisticas.totalFuncoes = Object.keys(estatisticas.porFuncao).length;
    
    console.log('[QLP/QUADRO] ========== RESUMO DO PROCESSAMENTO ==========');
    console.log(`[QLP/QUADRO] ✓ Processados: ${linhasProcessadas} colaboradores`);
    console.log(`[QLP/QUADRO] ⚠ Ignoradas: ${linhasIgnoradas} linhas (sem nome)`);
    console.log(`[QLP/QUADRO] ✗ Com erro: ${linhasComErro} linhas`);
    console.log(`[QLP/QUADRO] Status: ${estatisticas.ativos} ativos, ${estatisticas.inativos} inativos`);
    console.log(`[QLP/QUADRO] Seções: ${estatisticas.totalSecoes} diferentes`);
    console.log(`[QLP/QUADRO] Turnos: ${estatisticas.totalTurnos} diferentes`);
    console.log('[QLP/QUADRO] Distribuição por turno:', JSON.stringify(estatisticas.porTurno, null, 2));
    console.log('[QLP/QUADRO] ========== FIM DA REQUISIÇÃO ==========');
    
    return res.status(200).json({
      ok: true,
      colaboradores: colaboradores,
      estatisticas: estatisticas,
      debug: {
        linhasProcessadas,
        linhasIgnoradas,
        linhasComErro,
        totalLinhas: rows.length
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
