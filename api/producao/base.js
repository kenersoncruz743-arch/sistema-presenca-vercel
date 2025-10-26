// api/producao/resumo-base.js - VERSÃO CORRIGIDA COM DATA
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
    console.log('[RESUMO-BASE] ========== INÍCIO ==========');
    console.log('[RESUMO-BASE] Query params:', req.query);
    
    const doc = await sheetsService.init();
    console.log(`[RESUMO-BASE] ✓ Conectado: ${doc.title}`);
    
    const sheetBase = doc.sheetsByTitle['Base'];
    if (!sheetBase) {
      return res.status(404).json({
        ok: false,
        msg: 'Aba Base não encontrada'
      });
    }
    
    const rows = await sheetBase.getRows();
    console.log(`[RESUMO-BASE] ✓ ${rows.length} registros na Base`);
    
    // ===== CORREÇÃO DA DATA =====
    let dataFiltro;
    
    if (req.query.data) {
      // Se veio data do frontend (formato YYYY-MM-DD)
      const [ano, mes, dia] = req.query.data.split('-');
      dataFiltro = `${dia}/${mes}/${ano}`;
      console.log(`[RESUMO-BASE] Data do query: ${req.query.data} -> ${dataFiltro}`);
    } else {
      // Usa data de hoje
      const hoje = new Date();
      const dia = String(hoje.getDate()).padStart(2, '0');
      const mes = String(hoje.getMonth() + 1).padStart(2, '0');
      const ano = hoje.getFullYear();
      dataFiltro = `${dia}/${mes}/${ano}`;
      console.log(`[RESUMO-BASE] Data de hoje: ${dataFiltro}`);
    }
    
    console.log(`[RESUMO-BASE] Filtrando por data: "${dataFiltro}"`);
    
    // Estruturas para armazenar resumos
    const resumoPorSupervisor = {};
    const resumoPorFuncao = {};
    const resumoGeral = {
      total: 0,
      presente: 0,
      ausente: 0,
      atestado: 0,
      ferias: 0,
      folga: 0,
      afastado: 0,
      desvio: 0,
      outros: 0
    };
    
    let registrosFiltrados = 0;
    let registrosIgnorados = 0;
    
    // Processa cada registro
    rows.forEach((row, index) => {
      const dataRegistro = String(row.get('Data') || '').trim();
      
      // Debug das primeiras 5 datas
      if (index < 5) {
        console.log(`[RESUMO-BASE] Linha ${index + 1} - Data: "${dataRegistro}" | Match: ${dataRegistro === dataFiltro}`);
      }
      
      // COMPARAÇÃO EXATA DE STRINGS
      if (dataRegistro !== dataFiltro) {
        registrosIgnorados++;
        return;
      }
      
      registrosFiltrados++;
      
      const supervisor = String(row.get('Supervisor') || 'Sem supervisor').trim();
      const aba = String(row.get('Aba') || '').trim();
      const funcao = String(row.get('Função') || 'Não informada').trim();
      const turno = String(row.get('Turno') || 'Não informado').trim();
      const status = String(row.get('Status') || 'Outro').trim();
      const desvio = String(row.get('Desvio') || '').trim();
      const nome = String(row.get('Nome') || '').trim();
      const matricula = String(row.get('Matricula') || '').trim();
      
      if (!nome) return;
      
      // ====== RESUMO POR SUPERVISOR ======
      if (!resumoPorSupervisor[supervisor]) {
        resumoPorSupervisor[supervisor] = {
          supervisor,
          total: 0,
          presente: 0,
          ausente: 0,
          atestado: 0,
          ferias: 0,
          folga: 0,
          afastado: 0,
          desvio: 0,
          outros: 0,
          porFuncao: {},
          colaboradores: []
        };
      }
      
      resumoPorSupervisor[supervisor].total++;
      resumoPorSupervisor[supervisor].colaboradores.push({
        nome,
        matricula,
        funcao,
        turno,
        status,
        desvio
      });
      
      // Conta por status no supervisor
      const statusLower = status.toLowerCase();
      if (statusLower === 'presente') {
        resumoPorSupervisor[supervisor].presente++;
      } else if (statusLower === 'ausente') {
        resumoPorSupervisor[supervisor].ausente++;
      } else if (statusLower === 'atestado') {
        resumoPorSupervisor[supervisor].atestado++;
      } else if (statusLower.includes('férias') || statusLower.includes('ferias')) {
        resumoPorSupervisor[supervisor].ferias++;
      } else if (statusLower === 'folga') {
        resumoPorSupervisor[supervisor].folga++;
      } else if (statusLower === 'afastado') {
        resumoPorSupervisor[supervisor].afastado++;
      } else {
        resumoPorSupervisor[supervisor].outros++;
      }
      
      // Conta desvios
      if (desvio && desvio.toLowerCase() === 'desvio') {
        resumoPorSupervisor[supervisor].desvio++;
      }
      
      // Conta por função dentro do supervisor
      if (!resumoPorSupervisor[supervisor].porFuncao[funcao]) {
        resumoPorSupervisor[supervisor].porFuncao[funcao] = {
          total: 0,
          presente: 0,
          ausente: 0
        };
      }
      resumoPorSupervisor[supervisor].porFuncao[funcao].total++;
      if (statusLower === 'presente') {
        resumoPorSupervisor[supervisor].porFuncao[funcao].presente++;
      } else {
        resumoPorSupervisor[supervisor].porFuncao[funcao].ausente++;
      }
      
      // ====== RESUMO POR FUNÇÃO ======
      if (!resumoPorFuncao[funcao]) {
        resumoPorFuncao[funcao] = {
          funcao,
          total: 0,
          presente: 0,
          ausente: 0,
          atestado: 0,
          ferias: 0,
          folga: 0,
          afastado: 0,
          desvio: 0,
          outros: 0,
          porSupervisor: {},
          colaboradores: []
        };
      }
      
      resumoPorFuncao[funcao].total++;
      resumoPorFuncao[funcao].colaboradores.push({
        nome,
        matricula,
        supervisor,
        turno,
        status,
        desvio
      });
      
      // Conta por status na função
      if (statusLower === 'presente') {
        resumoPorFuncao[funcao].presente++;
      } else if (statusLower === 'ausente') {
        resumoPorFuncao[funcao].ausente++;
      } else if (statusLower === 'atestado') {
        resumoPorFuncao[funcao].atestado++;
      } else if (statusLower.includes('férias') || statusLower.includes('ferias')) {
        resumoPorFuncao[funcao].ferias++;
      } else if (statusLower === 'folga') {
        resumoPorFuncao[funcao].folga++;
      } else if (statusLower === 'afastado') {
        resumoPorFuncao[funcao].afastado++;
      } else {
        resumoPorFuncao[funcao].outros++;
      }
      
      // Conta desvios
      if (desvio && desvio.toLowerCase() === 'desvio') {
        resumoPorFuncao[funcao].desvio++;
      }
      
      // Conta por supervisor dentro da função
      if (!resumoPorFuncao[funcao].porSupervisor[supervisor]) {
        resumoPorFuncao[funcao].porSupervisor[supervisor] = {
          total: 0,
          presente: 0,
          ausente: 0
        };
      }
      resumoPorFuncao[funcao].porSupervisor[supervisor].total++;
      if (statusLower === 'presente') {
        resumoPorFuncao[funcao].porSupervisor[supervisor].presente++;
      } else {
        resumoPorFuncao[funcao].porSupervisor[supervisor].ausente++;
      }
      
      // ====== RESUMO GERAL ======
      resumoGeral.total++;
      if (statusLower === 'presente') {
        resumoGeral.presente++;
      } else if (statusLower === 'ausente') {
        resumoGeral.ausente++;
      } else if (statusLower === 'atestado') {
        resumoGeral.atestado++;
      } else if (statusLower.includes('férias') || statusLower.includes('ferias')) {
        resumoGeral.ferias++;
      } else if (statusLower === 'folga') {
        resumoGeral.folga++;
      } else if (statusLower === 'afastado') {
        resumoGeral.afastado++;
      } else {
        resumoGeral.outros++;
      }
      
      // Conta desvios no geral
      if (desvio && desvio.toLowerCase() === 'desvio') {
        resumoGeral.desvio++;
      }
    });
    
    console.log('[RESUMO-BASE] ========== ESTATÍSTICAS ==========');
    console.log(`[RESUMO-BASE] Data filtrada: ${dataFiltro}`);
    console.log(`[RESUMO-BASE] Registros processados: ${registrosFiltrados}`);
    console.log(`[RESUMO-BASE] Registros ignorados: ${registrosIgnorados}`);
    console.log(`[RESUMO-BASE] Total de linhas: ${rows.length}`);
    
    // Converte objetos em arrays e ordena
    const supervisores = Object.values(resumoPorSupervisor)
      .sort((a, b) => a.supervisor.localeCompare(b.supervisor));
    
    const funcoes = Object.values(resumoPorFuncao)
      .sort((a, b) => a.funcao.localeCompare(b.funcao));
    
    // Calcula percentuais no resumo geral
    if (resumoGeral.total > 0) {
      resumoGeral.percentualPresente = ((resumoGeral.presente / resumoGeral.total) * 100).toFixed(1);
      resumoGeral.percentualAusente = (((resumoGeral.total - resumoGeral.presente) / resumoGeral.total) * 100).toFixed(1);
      resumoGeral.percentualDesvio = ((resumoGeral.desvio / resumoGeral.total) * 100).toFixed(1);
    }
    
    console.log('[RESUMO-BASE] ========== RESUMO FINAL ==========');
    console.log(`[RESUMO-BASE] ${supervisores.length} supervisores`);
    console.log(`[RESUMO-BASE] ${funcoes.length} funções`);
    console.log(`[RESUMO-BASE] ${resumoGeral.total} colaboradores no total`);
    console.log(`[RESUMO-BASE] ${resumoGeral.presente} presentes (${resumoGeral.percentualPresente}%)`);
    console.log(`[RESUMO-BASE] ${resumoGeral.desvio} desvios (${resumoGeral.percentualDesvio}%)`);
    console.log('[RESUMO-BASE] ========== FIM ==========');
    
    return res.status(200).json({
      ok: true,
      dataReferencia: dataFiltro,
      resumoGeral,
      porSupervisor: supervisores,
      porFuncao: funcoes,
      totais: {
        supervisores: supervisores.length,
        funcoes: funcoes.length,
        colaboradores: resumoGeral.total
      },
      debug: {
        totalLinhas: rows.length,
        registrosFiltrados,
        registrosIgnorados
      },
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('[RESUMO-BASE] ========== ERRO FATAL ==========');
    console.error('[RESUMO-BASE] Tipo:', error.name);
    console.error('[RESUMO-BASE] Mensagem:', error.message);
    console.error('[RESUMO-BASE] Stack:', error.stack);
    console.error('[RESUMO-BASE] =====================================');
    
    return res.status(500).json({
      ok: false,
      msg: 'Erro ao gerar resumo',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
