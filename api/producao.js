// api/producao.js - API UNIFICADA DE PRODUÇÃO - CORRIGIDA PARA TURNO C
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
      
      try {
        const doc = await sheetsService.init();
        console.log(`[PRODUCAO/BASE] Conectado: ${doc.title}`);
        
        const sheetBase = doc.sheetsByTitle['Base'];
        if (!sheetBase) {
          console.error('[PRODUCAO/BASE] Aba Base não encontrada');
          return res.status(404).json({
            ok: false,
            msg: 'Aba "Base" não encontrada na planilha',
            abasDisponiveis: Object.keys(doc.sheetsByTitle)
          });
        }
        
        const rowsBase = await sheetBase.getRows();
        console.log(`[PRODUCAO/BASE] ${rowsBase.length} registros na Base`);
        
        // Carrega QLP para cruzamento
        const sheetQLP = doc.sheetsByTitle['QLP'];
        const mapaQLP = {};
        
        if (sheetQLP) {
          try {
            const rowsQLP = await sheetQLP.getRows();
            rowsQLP.forEach(row => {
              try {
                const chapa = String(row.get('CHAPA1') || '').trim();
                const secao = String(row.get('SECAO') || '').trim();
                const turno = String(row.get('Turno') || '').trim();
                
                if (chapa) {
                  mapaQLP[chapa] = { secao, turno };
                }
              } catch (rowErr) {
                // Ignora erros em linhas individuais
              }
            });
            console.log(`[PRODUCAO/BASE] ${Object.keys(mapaQLP).length} registros no mapa QLP`);
          } catch (qlpErr) {
            console.warn('[PRODUCAO/BASE] Erro ao processar QLP (continuando sem):', qlpErr.message);
          }
        }
        
        // Processa dados da Base
        const dados = [];
        const hoje = new Date().toLocaleDateString('pt-BR');
        
        rowsBase.forEach(row => {
          try {
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
            
            // CORREÇÃO COMPLETA: Lógica melhorada para determinar turno pela aba
            if (!turno || turno === 'Não definido' || turno === '') {
              if (aba) {
                const abaUpper = aba.toUpperCase().trim();
                
                // Remove espaços extras e normaliza
                const abaNormalizada = abaUpper.replace(/\s+/g, ' ');
                
                console.log(`[DEBUG] Processando aba: "${aba}" -> "${abaNormalizada}"`);
                
                // Testa Turno C primeiro (mais específico)
                if (abaNormalizada.includes('TURNO C') || 
                    abaNormalizada === 'C' || 
                    abaNormalizada.endsWith(' C') ||
                    abaNormalizada.startsWith('C ') ||
                    abaNormalizada.includes(' TC') ||
                    abaNormalizada.includes('TC ')) {
                  turno = 'Turno C';
                }
                // Testa Turno B
                else if (abaNormalizada.includes('TURNO B') || 
                         abaNormalizada === 'B' || 
                         abaNormalizada.endsWith(' B') ||
                         abaNormalizada.startsWith('B ') ||
                         abaNormalizada.includes(' TB') ||
                         abaNormalizada.includes('TB ')) {
                  turno = 'Turno B';
                }
                // Testa Turno A
                else if (abaNormalizada.includes('TURNO A') || 
                         abaNormalizada === 'A' || 
                         abaNormalizada.endsWith(' A') ||
                         abaNormalizada.startsWith('A ') ||
                         abaNormalizada.includes(' TA') ||
                         abaNormalizada.includes('TA ')) {
                  turno = 'Turno A';
                }
                
                console.log(`[DEBUG] Turno identificado: "${turno}"`);
              }
            }
            
            // NORMALIZAÇÃO FINAL: Garante formato padrão "Turno X"
            const turnoUpper = turno.toUpperCase().trim();
            
            if (turnoUpper === 'A' || turnoUpper === 'TA' || turnoUpper === 'TURNO A') {
              turno = 'Turno A';
            } else if (turnoUpper === 'B' || turnoUpper === 'TB' || turnoUpper === 'TURNO B') {
              turno = 'Turno B';
            } else if (turnoUpper === 'C' || turnoUpper === 'TC' || turnoUpper === 'TURNO C') {
              turno = 'Turno C';
            } else if (turnoUpper.includes('TURNO A') || turnoUpper.includes('TA')) {
              turno = 'Turno A';
            } else if (turnoUpper.includes('TURNO B') || turnoUpper.includes('TB')) {
              turno = 'Turno B';
            } else if (turnoUpper.includes('TURNO C') || turnoUpper.includes('TC')) {
              turno = 'Turno C';
            } else if (!turno || turno === 'NÃO DEFINIDO' || turno === 'NAO DEFINIDO') {
              turno = 'Não definido';
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
          } catch (rowErr) {
            console.warn('[PRODUCAO/BASE] Erro ao processar linha:', rowErr.message);
          }
        });
        
        console.log(`[PRODUCAO/BASE] ${dados.length} registros processados (hoje)`);
        console.log('[PRODUCAO/BASE] Turnos únicos:', [...new Set(dados.map(d => d.turno))]);
        
        // Debug: Mostra distribuição de turnos
        const contagemTurnos = {};
        dados.forEach(d => {
          contagemTurnos[d.turno] = (contagemTurnos[d.turno] || 0) + 1;
        });
        console.log('[PRODUCAO/BASE] Distribuição de turnos:', contagemTurnos);
        
        return res.status(200).json({
          ok: true,
          dados,
          total: dados.length,
          dataFiltro: hoje,
          timestamp: new Date().toISOString()
        });
      } catch (error) {
        console.error('[PRODUCAO/BASE] Erro:', error);
        return res.status(500).json({
          ok: false,
          msg: 'Erro ao buscar dados da Base',
          error: error.message
        });
      }
    }

    // ==================== META ====================
    if (action === 'meta') {
      console.log('[PRODUCAO/META] Iniciando busca...');
      
      try {
        const doc = await sheetsService.init();
        console.log(`[PRODUCAO/META] Conectado: ${doc.title}`);
        
        const sheetMeta = doc.sheetsByTitle['Meta'];
        if (!sheetMeta) {
          console.warn('[PRODUCAO/META] Aba Meta não encontrada - retornando array vazio');
          return res.status(200).json({
            ok: true,
            dados: [],
            total: 0,
            msg: 'Aba Meta não encontrada (opcional)',
            timestamp: new Date().toISOString()
          });
        }
        
        const rowsMeta = await sheetMeta.getRows();
        console.log(`[PRODUCAO/META] ${rowsMeta.length} registros na Meta`);
        
        const dados = [];
        
        rowsMeta.forEach(row => {
          try {
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
          } catch (rowErr) {
            console.warn('[PRODUCAO/META] Erro ao processar linha:', rowErr.message);
          }
        });
        
        console.log(`[PRODUCAO/META] ${dados.length} registros processados`);
        
        return res.status(200).json({
          ok: true,
          dados,
          total: dados.length,
          timestamp: new Date().toISOString()
        });
      } catch (error) {
        console.error('[PRODUCAO/META] Erro:', error);
        return res.status(200).json({
          ok: true,
          dados: [],
          total: 0,
          msg: 'Erro ao buscar Meta (continuando sem): ' + error.message,
          timestamp: new Date().toISOString()
        });
      }
    }

    // ==================== PRODUTIVIDADE ====================
    if (action === 'produtividade') {
      console.log('[PRODUCAO/PRODUTIVIDADE] Iniciando busca...');
      
      try {
        const doc = await sheetsService.init();
        console.log(`[PRODUCAO/PRODUTIVIDADE] Conectado: ${doc.title}`);
        
        const sheetProd = doc.sheetsByTitle['Produtividade_Hora'];
        if (!sheetProd) {
          console.error('[PRODUCAO/PRODUTIVIDADE] Aba Produtividade_Hora não encontrada');
          return res.status(404).json({
            ok: false,
            msg: 'Aba "Produtividade_Hora" não encontrada na planilha',
            abasDisponiveis: Object.keys(doc.sheetsByTitle),
            instrucoes: 'Crie uma aba chamada "Produtividade_Hora" com as colunas: FUNCAO | Produtividade/hora'
          });
        }
        
        const rowsProd = await sheetProd.getRows();
        console.log(`[PRODUCAO/PRODUTIVIDADE] ${rowsProd.length} registros na Produtividade_Hora`);
        
        const dados = [];
        
        rowsProd.forEach(row => {
          try {
            const funcao = String(row.get('FUNCAO') || '').trim();
            const produtividadeHora = String(row.get('Produtividade/hora') || '').trim();
            
            if (funcao && produtividadeHora) {
              const produtividadeNum = parseFloat(produtividadeHora.replace(',', '.')) || 0;
              
              dados.push({
                funcao: funcao,
                produtividadeHora: produtividadeNum
              });
            }
          } catch (rowErr) {
            console.warn('[PRODUCAO/PRODUTIVIDADE] Erro ao processar linha:', rowErr.message);
          }
        });
        
        console.log(`[PRODUCAO/PRODUTIVIDADE] ${dados.length} registros processados`);
        
        if (dados.length === 0) {
          console.warn('[PRODUCAO/PRODUTIVIDADE] Nenhum registro válido encontrado');
        }
        
        return res.status(200).json({
          ok: true,
          dados,
          total: dados.length,
          timestamp: new Date().toISOString()
        });
      } catch (error) {
        console.error('[PRODUCAO/PRODUTIVIDADE] Erro:', error);
        return res.status(500).json({
          ok: false,
          msg: 'Erro ao buscar Produtividade_Hora',
          error: error.message,
          stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
        });
      }
    }

    // ==================== RESUMO BASE ====================
    if (action === 'resumo-base') {
      console.log('[PRODUCAO/RESUMO-BASE] Iniciando geração de resumo...');
      
      try {
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
          try {
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
            
            if (desvio && desvio.toLowerCase() === 'desvio') {
              resumoGeral.desvio++;
            }
          } catch (rowErr) {
            console.warn('[PRODUCAO/RESUMO-BASE] Erro ao processar linha:', rowErr.message);
          }
        });
        
        // Calcula percentuais
        resumoGeral.percentualPresente = resumoGeral.total > 0 
          ? ((resumoGeral.presente / resumoGeral.total) * 100).toFixed(1) 
          : 0;
        resumoGeral.percentualAusente = resumoGeral.total > 0 
          ? (((resumoGeral.total - resumoGeral.presente) / resumoGeral.total) * 100).toFixed(1) 
          : 0;
        
        console.log(`[PRODUCAO/RESUMO-BASE] Processados ${registrosFiltrados} registros da data ${dataFiltro}`);
        console.log(`[PRODUCAO/RESUMO-BASE] Total: ${resumoGeral.total}, Presente: ${resumoGeral.presente}, Ausente: ${resumoGeral.ausente}`);
        
        return res.status(200).json({
          ok: true,
          dataReferencia: dataFiltro,
          resumoGeral,
          porSupervisor: Object.values(resumoPorSupervisor),
          porFuncao: Object.values(resumoPorFuncao),
          totalSupervisores: Object.keys(resumoPorSupervisor).length,
          totalFuncoes: Object.keys(resumoPorFuncao).length,
          timestamp: new Date().toISOString()
        });
      } catch (error) {
        console.error('[PRODUCAO/RESUMO-BASE] Erro:', error);
        return res.status(500).json({
          ok: false,
          msg: 'Erro ao gerar resumo',
          error: error.message
        });
      }
    }

    // ==================== SALVAR ====================
    if (action === 'salvar') {
      console.log('[PRODUCAO/SALVAR] Iniciando salvamento...');
      
      if (req.method !== 'POST') {
        return res.status(405).json({
          ok: false,
          msg: 'Método não permitido. Use POST.'
        });
      }
      
      const { dados, dataReferencia } = req.body;
      
      if (!dados || !Array.isArray(dados)) {
        return res.status(400).json({
          ok: false,
          msg: 'Dados inválidos'
        });
      }
      
      console.log(`[PRODUCAO/SALVAR] Salvando ${dados.length} registros...`);
      
      // TODO: Implementar lógica de salvamento na aba Produzido
      
      return res.status(200).json({
        ok: true,
        msg: `${dados.length} registros prontos para salvar (implementar salvamento na aba Produzido)`,
        dataReferencia
      });
    }

    // Ação não reconhecida
    return res.status(400).json({
      ok: false,
      msg: `Ação não reconhecida: ${action}`,
      acoesDisponiveis: ['base', 'meta', 'produtividade', 'resumo-base', 'salvar']
    });

  } catch (error) {
    console.error('[PRODUCAO] Erro geral:', error);
    console.error('[PRODUCAO] Stack:', error.stack);
    
    return res.status(500).json({
      ok: false,
      msg: 'Erro interno do servidor',
      error: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
