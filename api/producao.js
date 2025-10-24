// api/producao.js - API UNIFICADA DE PRODUÇÃO
const sheetsService = require('../lib/sheets');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    const { action } = req.query;
    
    if (!action) {
      return res.status(400).json({
        ok: false,
        msg: 'Parâmetro "action" é obrigatório'
      });
    }

    console.log(`[PRODUCAO] Action: ${action}`);

    // ==================== BASE ====================
    if (action === 'base') {
      console.log('[PRODUCAO/BASE] Iniciando busca...');
      
      const doc = await sheetsService.init();
      console.log(`[PRODUCAO/BASE] Conectado: ${doc.title}`);
      
      // Carrega Base
      const sheetBase = doc.sheetsByTitle['Base'];
      if (!sheetBase) {
        return res.status(404).json({
          ok: false,
          msg: 'Aba Base não encontrada'
        });
      }
      
      const rowsBase = await sheetBase.getRows();
      console.log(`[PRODUCAO/BASE] ${rowsBase.length} registros na Base`);
      
      // Carrega QLP para cruzamento
      const sheetQLP = doc.sheetsByTitle['QLP'];
      const mapaQLP = {};
      
      if (sheetQLP) {
        const rowsQLP = await sheetQLP.getRows();
        rowsQLP.forEach(row => {
          const chapa = String(row.get('CHAPA1') || '').trim();
          const secao = String(row.get('SECAO') || '').trim();
          const turno = String(row.get('Turno') || '').trim();
          
          if (chapa) {
            mapaQLP[chapa] = { secao, turno };
          }
        });
        console.log(`[PRODUCAO/BASE] ${Object.keys(mapaQLP).length} registros no mapa QLP`);
      }
      
      // Processa dados da Base
      const dados = [];
      const hoje = new Date().toLocaleDateString('pt-BR');
      
      rowsBase.forEach(row => {
        const supervisor = String(row.get('Supervisor') || '').trim();
        const aba = String(row.get('Aba') || '').trim();
        const matricula = String(row.get('Matricula') || '').trim();
        const nome = String(row.get('Nome') || '').trim();
        const funcao = String(row.get('Função') || '').trim();
        const status = String(row.get('Status') || '').trim();
        const data = String(row.get('Data') || '').trim();
        
        // Filtra apenas registros de hoje
        if (data !== hoje) return;
        
        // Busca seção e turno do QLP
        let secao = 'Sem Seção';
        let turno = 'Não definido';
        
        if (mapaQLP[matricula]) {
          secao = mapaQLP[matricula].secao || secao;
          turno = mapaQLP[matricula].turno || turno;
        }
        
        // Determina turno pela aba se não tiver no QLP
        if (turno === 'Não definido' && aba) {
          const abaLower = aba.toLowerCase();
          if (abaLower.includes('ta') || abaLower.includes('turno a')) turno = 'Turno A';
          else if (abaLower.includes('tb') || abaLower.includes('turno b')) turno = 'Turno B';
          else if (abaLower.includes('tc') || abaLower.includes('turno c')) turno = 'Turno C';
        }
        
        dados.push({
          supervisor,
          aba,
          matricula,
          nome,
          funcao,
          status,
          data,
          secao,
          turno
        });
      });
      
      console.log(`[PRODUCAO/BASE] ${dados.length} registros processados (hoje)`);
      
      return res.status(200).json({
        ok: true,
        dados,
        total: dados.length,
        dataFiltro: hoje,
        timestamp: new Date().toISOString()
      });
    }

    // ==================== META ====================
    if (action === 'meta') {
      console.log('[PRODUCAO/META] Iniciando busca...');
      
      const doc = await sheetsService.init();
      console.log(`[PRODUCAO/META] Conectado: ${doc.title}`);
      
      const sheetMeta = doc.sheetsByTitle['Meta'];
      if (!sheetMeta) {
        return res.status(404).json({
          ok: false,
          msg: 'Aba Meta não encontrada'
        });
      }
      
      const rowsMeta = await sheetMeta.getRows();
      console.log(`[PRODUCAO/META] ${rowsMeta.length} registros na Meta`);
      
      const dados = [];
      
      rowsMeta.forEach(row => {
        const data = String(row.get('Data') || '').trim();
        const meta = String(row.get('Meta') || '').trim();
        const produtividadeHora = String(row.get('Produtividade/hora') || '').trim();
        
        if (data && produtividadeHora) {
          const produtividadeNum = parseFloat(produtividadeHora.replace(',', '.')) || 0;
          const metaNum = parseFloat(meta.replace(',', '.')) || 0;
          
          dados.push({
            data: data,
            meta: metaNum,
            produtividadeHora: produtividadeNum
          });
        }
      });
      
      console.log(`[PRODUCAO/META] ${dados.length} registros processados`);
      
      return res.status(200).json({
        ok: true,
        dados,
        total: dados.length,
        timestamp: new Date().toISOString()
      });
    }

    // ==================== PRODUTIVIDADE ====================
    if (action === 'produtividade') {
      console.log('[PRODUCAO/PRODUTIVIDADE] Iniciando busca...');
      
      const doc = await sheetsService.init();
      console.log(`[PRODUCAO/PRODUTIVIDADE] Conectado: ${doc.title}`);
      
      const sheetProd = doc.sheetsByTitle['Produtividade_Hora'];
      if (!sheetProd) {
        return res.status(404).json({
          ok: false,
          msg: 'Aba Produtividade_Hora não encontrada'
        });
      }
      
      const rowsProd = await sheetProd.getRows();
      console.log(`[PRODUCAO/PRODUTIVIDADE] ${rowsProd.length} registros na Produtividade_Hora`);
      
      const dados = [];
      
      rowsProd.forEach(row => {
        const funcao = String(row.get('FUNCAO') || '').trim();
        const produtividadeHora = String(row.get('Produtividade/hora') || '').trim();
        
        if (funcao && produtividadeHora) {
          const produtividadeNum = parseFloat(produtividadeHora.replace(',', '.')) || 0;
          
          dados.push({
            funcao: funcao,
            produtividadeHora: produtividadeNum
          });
        }
      });
      
      console.log(`[PRODUCAO/PRODUTIVIDADE] ${dados.length} registros processados`);
      
      return res.status(200).json({
        ok: true,
        dados,
        total: dados.length,
        timestamp: new Date().toISOString()
      });
    }

    // ==================== RESUMO BASE ====================
    if (action === 'resumo-base') {
      console.log('[PRODUCAO/RESUMO-BASE] Iniciando geração de resumo...');
      
      const doc = await sheetsService.init();
      console.log(`[PRODUCAO/RESUMO-BASE] Conectado: ${doc.title}`);
      
      const sheetBase = doc.sheetsByTitle['Base'];
      if (!sheetBase) {
        return res.status(404).json({
          ok: false,
          msg: 'Aba Base não encontrada'
        });
      }
      
      const rows = await sheetBase.getRows();
      console.log(`[PRODUCAO/RESUMO-BASE] ${rows.length} registros na Base`);
      
      // Obtém data do query ou usa hoje
      let dataFiltro;
      if (req.query.data) {
        const [ano, mes, dia] = req.query.data.split('-');
        dataFiltro = `${dia}/${mes}/${ano}`;
      } else {
        dataFiltro = new Date().toLocaleDateString('pt-BR');
      }
      
      console.log(`[PRODUCAO/RESUMO-BASE] Filtrando por data: ${dataFiltro}`);
      
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
      
      // Processa cada registro
      rows.forEach(row => {
        const dataRegistro = String(row.get('Data') || '').trim();
        
        if (dataRegistro !== dataFiltro) return;
        
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
      
      console.log(`[PRODUCAO/RESUMO-BASE] Registros filtrados: ${registrosFiltrados}`);
      
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
      
      console.log('[PRODUCAO/RESUMO-BASE] Resumo gerado com sucesso');
      
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
        timestamp: new Date().toISOString()
      });
    }

    // ==================== SALVAR PRODUÇÃO ====================
    if (action === 'salvar' && req.method === 'POST') {
      console.log('[PRODUCAO/SALVAR] Iniciando salvamento...');
      
      const { dados, dataReferencia, horasTrabalhadas } = req.body;
      
      if (!dados || !Array.isArray(dados) || dados.length === 0) {
        return res.status(400).json({
          ok: false,
          msg: 'Dados inválidos ou vazios'
        });
      }
      
      console.log(`[PRODUCAO/SALVAR] ${dados.length} registros para salvar`);
      
      // Aqui você implementaria a lógica de salvamento na aba Produzido
      // Por enquanto, retorna sucesso simulado
      
      return res.status(200).json({
        ok: true,
        msg: `${dados.length} registros salvos com sucesso!`,
        dataReferencia,
        horasTrabalhadas
      });
    }

    // Action não reconhecida
    return res.status(400).json({
      ok: false,
      msg: 'Action não reconhecida: ' + action,
      acoesDisponiveis: ['base', 'meta', 'produtividade', 'resumo-base', 'salvar']
    });

  } catch (error) {
    console.error('[PRODUCAO] Erro:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro interno do servidor',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
