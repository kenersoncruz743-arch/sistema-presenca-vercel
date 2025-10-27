// lib/sheets_2.js - Serviço para controle de coletores (RFID)
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

class SheetsColetorService {
  constructor() {
    this.doc = null;
    this.initialized = false;
  }

  async init() {
    if (this.initialized) return this.doc;
    
    try {
      console.log('[SHEETS_COLETOR] Iniciando conexão...');
      
      const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SHEETS_CLIENT_EMAIL,
        key: process.env.GOOGLE_SHEETS_PRIVATE_KEY?.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });
      
      // Usa a variável de ambiente específica para coletores
      const sheetId = process.env.GOOGLE_SHEETS_ID_COLETOR;
      
      if (!sheetId) {
        throw new Error('GOOGLE_SHEETS_ID_COLETOR não configurado nas variáveis de ambiente');
      }
      
      this.doc = new GoogleSpreadsheet(sheetId, serviceAccountAuth);
      await this.doc.loadInfo();
      
      this.initialized = true;
      console.log(`[SHEETS_COLETOR] ✓ Conectado: ${this.doc.title}`);
      
      return this.doc;
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro na conexão:', error);
      throw new Error('Falha na conexão com planilha de coletores: ' + error.message);
    }
  }

  // ==================== FUNÇÕES PARA ABA "quadro" ====================
  
  /**
   * Busca dados dos colaboradores na aba "quadro"
   * Retorna lista de colaboradores com chapa, nome e função
   */
  async obterDados() {
    try {
      await this.init();
      console.log('[SHEETS_COLETOR] Buscando dados da aba Quadro...');
      
      const sheet = this.doc.sheetsByTitle['quadro'];
      if (!sheet) {
        console.error('[SHEETS_COLETOR] Aba "quadro" não encontrada');
        return [];
      }
      
      const rows = await sheet.getRows();
      console.log(`[SHEETS_COLETOR] ${rows.length} linhas encontradas na aba Quadro`);
      
      const dados = [];
      
      rows.forEach(row => {
        const chapa = String(row.get('Coluna 1') || '').trim();
        const nome = String(row.get('NOME') || '').trim();
        const funcao = String(row.get('Função que atua') || row.get('FUNÇÃO NO RM') || '').trim();
        
        if (chapa && nome) {
          dados.push({
            chapa: chapa,
            nome: nome,
            funcao: funcao
          });
        }
      });
      
      console.log(`[SHEETS_COLETOR] ✓ ${dados.length} colaboradores carregados`);
      return dados;
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao obter dados:', error);
      return [];
    }
  }

  // ==================== FUNÇÕES PARA ABA "Base" ====================
  
  /**
   * Salva registro de movimentação de coletor na aba "Base"
   * Inclui validação anti-duplicação e atualização de dados da aba Presença
   */
  async salvarRegistro(chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes) {
    try {
      await this.init();
      console.log('[SHEETS_COLETOR] Salvando registro:', {
        chapa, nome, numeroColetor, tipoOperacao, situacoes
      });
      
      // Validação de campos obrigatórios
      if (!chapa || !numeroColetor || !situacoes || situacoes.length === 0) {
        return { ok: false, msg: 'Preencha todos os campos obrigatórios.' };
      }
      
      const sheet = this.doc.sheetsByTitle['Base'];
      if (!sheet) {
        return { ok: false, msg: 'Aba Base não encontrada' };
      }
      
      const lastRow = sheet.getLastRow();
      const agora = new Date();
      const situacoesTexto = situacoes.join(', ');
      
      // Validação anti-duplicação (120 segundos)
      if (lastRow >= 2) {
        const ultimaEntrada = await sheet.getRows({ offset: lastRow - 1, limit: 1 });
        
        if (ultimaEntrada.length > 0) {
          const row = ultimaEntrada[0];
          const ultimaData = row.get('Data');
          const ultimaChapa = String(row.get('Chapa') || '').trim();
          const ultimoColetor = String(row.get('NumeroColetor') || '').trim();
          const ultimoTipo = String(row.get('TipoOpe') || '').trim();
          
          const mesmaChapa = ultimaChapa === chapa;
          const mesmoColetor = ultimoColetor === String(numeroColetor);
          const mesmoTipo = ultimoTipo === tipoOperacao;
          
          if (ultimaData) {
            const tempoDecorrido = (agora - new Date(ultimaData)) / 1000; // segundos
            
            if (mesmaChapa && mesmoColetor && mesmoTipo && tempoDecorrido < 120) {
              console.log('[SHEETS_COLETOR] Registro duplicado detectado');
              return { 
                ok: false, 
                msg: 'Registro já enviado recentemente. Aguarde alguns minutos antes de reenviar.' 
              };
            }
          }
        }
      }
      
      // Salva o registro
      await sheet.addRow({
        'Data': agora,
        'Chapa': chapa,
        'Nome': nome,
        'Funcao': funcao,
        'NumeroColetor': numeroColetor,
        'TipoOpe': tipoOperacao,
        'Situacao': situacoesTexto
      });
      
      console.log('[SHEETS_COLETOR] ✓ Registro salvo com sucesso');
      
      // Atualiza dados da aba Presença
      await this.atualizarDadosPresencaNaBase();
      
      return { ok: true, msg: 'Dados salvos com sucesso!' };
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao salvar registro:', error);
      return { ok: false, msg: 'Erro ao salvar: ' + error.message };
    }
  }

  /**
   * Atualiza colunas J e K da aba Base com dados da aba Presença
   * Busca último registro de cada matrícula
   */
  async atualizarDadosPresencaNaBase() {
    try {
      await this.init();
      console.log('[SHEETS_COLETOR] Atualizando dados da Presença na Base...');
      
      const baseSheet = this.doc.sheetsByTitle['Base'];
      const presencaSheet = this.doc.sheetsByTitle['Presença'];
      
      if (!baseSheet) {
        console.error('[SHEETS_COLETOR] Aba Base não encontrada');
        return;
      }
      
      const baseLastRow = baseSheet.getLastRow();
      if (baseLastRow < 2) {
        console.log('[SHEETS_COLETOR] Base vazia, nada para atualizar');
        return;
      }
      
      // Carrega matrículas da Base (coluna B)
      const baseDataRange = baseSheet.getRange(`B2:B${baseLastRow}`);
      const baseMatriculas = (await baseDataRange.getValues()).flat();
      
      if (!presencaSheet) {
        console.log('[SHEETS_COLETOR] Aba Presença não encontrada, limpando colunas J e K');
        const clearRange = baseSheet.getRange(2, 10, baseMatriculas.length, 2);
        await clearRange.clearContent();
        return;
      }
      
      const presencaLastRow = presencaSheet.getLastRow();
      if (presencaLastRow < 2) {
        console.log('[SHEETS_COLETOR] Presença vazia, limpando colunas J e K');
        const clearRange = baseSheet.getRange(2, 10, baseMatriculas.length, 2);
        await clearRange.clearContent();
        return;
      }
      
      // Carrega dados da Presença (colunas B, E, H)
      const presencaData = await presencaSheet.getRows();
      
      // Mapa para último registro de cada matrícula (de baixo para cima)
      const presencaMap = new Map();
      
      for (let i = presencaData.length - 1; i >= 0; i--) {
        const row = presencaData[i];
        const matricula = String(row.get('Matricula') || row.get('CHAPA1') || '').trim();
        const colunaE = String(row.get('E') || '').trim();
        const colunaH = String(row.get('H') || '').trim();
        
        if (matricula && !presencaMap.has(matricula)) {
          presencaMap.set(matricula, [colunaE, colunaH]);
        }
      }
      
      console.log(`[SHEETS_COLETOR] Mapa Presença criado: ${presencaMap.size} matrículas`);
      
      // Prepara dados para atualização
      const output = baseMatriculas.map(matricula => {
        const dados = presencaMap.get(String(matricula).trim());
        return dados || ['', ''];
      });
      
      // Atualiza colunas J e K da Base
      const updateRange = baseSheet.getRange(2, 10, output.length, 2);
      await updateRange.setValues(output);
      
      console.log('[SHEETS_COLETOR] ✓ Dados da Presença atualizados na Base');
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao atualizar dados da Presença:', error);
    }
  }

  // ==================== FUNÇÕES PARA STATUS DOS COLETORES ====================
  
  /**
   * Retorna o último status de cada coletor
   * Lê as 10 primeiras colunas (incluindo supervisor na coluna J)
   */
  async obterColetorStatus() {
    try {
      await this.init();
      console.log('[SHEETS_COLETOR] Obtendo status dos coletores...');
      
      const sheet = this.doc.sheetsByTitle['Base'];
      if (!sheet) {
        console.error('[SHEETS_COLETOR] Aba Base não encontrada');
        return {};
      }
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        console.log('[SHEETS_COLETOR] Base vazia');
        return {};
      }
      
      // Lê 10 colunas (A até J) para incluir supervisor
      const rows = await sheet.getRows();
      
      const mapa = {};
      
      // Processa de baixo para cima para pegar último status
      for (let i = rows.length - 1; i >= 0; i--) {
        const row = rows[i];
        
        const data = row.get('Data');
        const chapa = String(row.get('Chapa') || '').trim();
        const nome = String(row.get('Nome') || '').trim();
        const funcao = String(row.get('Funcao') || '').trim();
        const coletor = String(row.get('NumeroColetor') || '').trim();
        const tipo = String(row.get('TipoOpe') || '').trim();
        const situacao = String(row.get('Situacao') || '').trim();
        const supervisor = String(row.get('Supervisor') || '').trim();
        
        if (coletor && !mapa[coletor]) {
          mapa[coletor] = {
            data: this.formatarData(data),
            tipo: tipo,
            situacao: situacao,
            chapa: chapa,
            nome: nome,
            supervisor: supervisor
          };
        }
      }
      
      console.log(`[SHEETS_COLETOR] ✓ Status de ${Object.keys(mapa).length} coletores obtido`);
      return mapa;
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao obter status dos coletores:', error);
      return {};
    }
  }

  /**
   * Gera resumo de coletores: disponíveis, indisponíveis e quebrados
   */
  async gerarResumoColetores() {
    try {
      const statusPorColetor = await this.obterColetorStatus();
      
      let disponiveis = 0;
      let indisponiveis = 0;
      let quebrados = 0;
      
      for (const coletor in statusPorColetor) {
        const status = statusPorColetor[coletor];
        
        // Disponível: tipo "Entrega" e situação "OK"
        if (status.tipo === "Entrega" && status.situacao === "OK") {
          disponiveis++;
        } 
        // Indisponível: tipo "Retirada"
        else if (status.tipo === "Retirada") {
          indisponiveis++;
        }
        
        // Quebrado: situação diferente de "OK"
        if (status.situacao !== "OK") {
          quebrados++;
        }
      }
      
      console.log('[SHEETS_COLETOR] Resumo:', { disponiveis, indisponiveis, quebrados });
      
      return {
        disponiveis,
        indisponiveis,
        quebrados,
        total: Object.keys(statusPorColetor).length
      };
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao gerar resumo:', error);
      return {
        disponiveis: 0,
        indisponiveis: 0,
        quebrados: 0,
        total: 0
      };
    }
  }

  /**
   * Gera resumo por supervisor
   * Regra: conta apenas 1 vez cada colaborador que fez retirada no dia
   * Considera apenas registros do dia atual
   */
  async gerarResumoPorSupervisor() {
    try {
      await this.init();
      console.log('[SHEETS_COLETOR] Gerando resumo por supervisor...');
      
      const sheet = this.doc.sheetsByTitle['Base'];
      if (!sheet) {
        console.error('[SHEETS_COLETOR] Aba Base não encontrada');
        return {};
      }
      
      const rows = await sheet.getRows();
      const timezone = 'America/Campo_Grande'; // Ajuste conforme necessário
      
      const resumoSupervisor = {};
      const colaboradoresContados = {};
      
      // Data de hoje sem hora
      const hoje = new Date();
      hoje.setHours(0, 0, 0, 0);
      
      // Agrupa registros por supervisor e colaborador (nome) para o dia atual
      const registrosPorSupCol = {};
      
      for (const row of rows) {
        const data = row.get('Data');
        if (!(data instanceof Date)) continue;
        
        const dataSemHora = new Date(data);
        dataSemHora.setHours(0, 0, 0, 0);
        
        // Ignora registros de outros dias
        if (dataSemHora.getTime() !== hoje.getTime()) continue;
        
        const supervisor = String(row.get('Supervisor') || 'Sem Supervisor').trim();
        const colaborador = String(row.get('Nome') || '').trim();
        const tipo = String(row.get('TipoOpe') || '').trim();
        
        if (!colaborador) continue;
        
        if (!registrosPorSupCol[supervisor]) {
          registrosPorSupCol[supervisor] = {};
        }
        
        if (!registrosPorSupCol[supervisor][colaborador]) {
          registrosPorSupCol[supervisor][colaborador] = [];
        }
        
        registrosPorSupCol[supervisor][colaborador].push({
          data: data,
          tipo: tipo
        });
      }
      
      // Processa cada supervisor e colaborador aplicando regra de contagem
      for (const sup in registrosPorSupCol) {
        if (!resumoSupervisor[sup]) {
          resumoSupervisor[sup] = { retiradaContada: 0 };
        }
        
        colaboradoresContados[sup] = {};
        
        const colaboradores = registrosPorSupCol[sup];
        
        for (const colab in colaboradores) {
          const registros = colaboradores[colab]
            .slice()
            .sort((a, b) => b.data - a.data); // Ordem decrescente (mais recente primeiro)
          
          // Filtra apenas Retirada e Entrega
          const seq = registros
            .filter(r => r.tipo === "Retirada" || r.tipo === "Entrega")
            .map(r => r.tipo);
          
          // Inverte para ordem crescente (cronológica)
          const seqCresc = [...seq].reverse();
          
          // Conta trocas de tipo
          let trocas = 0;
          for (let i = 1; i < seqCresc.length; i++) {
            if (seqCresc[i] !== seqCresc[i - 1]) {
              trocas++;
            }
          }
          
          // Conta quantas retiradas existem
          const countRetirada = seqCresc.filter(x => x === "Retirada").length;
          
          // Regra: conta se tem pelo menos 1 retirada E no máximo 1 troca
          let conta = 0;
          if (countRetirada > 0) {
            if (trocas <= 1) {
              conta = 1;
            } else {
              conta = 0;
            }
          }
          
          // Adiciona ao total se não foi contado ainda
          if (conta === 1 && !colaboradoresContados[sup][colab]) {
            resumoSupervisor[sup].retiradaContada++;
            colaboradoresContados[sup][colab] = true;
          }
        }
      }
      
      console.log('[SHEETS_COLETOR] ✓ Resumo por supervisor gerado:', resumoSupervisor);
      return resumoSupervisor;
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao gerar resumo por supervisor:', error);
      return {};
    }
  }

  // ==================== FUNÇÕES AUXILIARES ====================
  
  /**
   * Formata data para exibição (dd/MM/yyyy HH:mm)
   */
  formatarData(data) {
    if (!data || !(data instanceof Date)) return '';
    
    try {
      const dia = String(data.getDate()).padStart(2, '0');
      const mes = String(data.getMonth() + 1).padStart(2, '0');
      const ano = data.getFullYear();
      const hora = String(data.getHours()).padStart(2, '0');
      const min = String(data.getMinutes()).padStart(2, '0');
      
      return `${dia}/${mes}/${ano} ${hora}:${min}`;
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao formatar data:', error);
      return '';
    }
  }
}

// Exporta instância única (Singleton)
module.exports = new SheetsColetorService();
