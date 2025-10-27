// lib/sheets_2.js - CORREÇÃO: Nome correto da aba
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
      console.log(`[SHEETS_COLETOR] Abas disponíveis:`, Object.keys(this.doc.sheetsByTitle));
      
      return this.doc;
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro na conexão:', error);
      throw new Error('Falha na conexão com planilha de coletores: ' + error.message);
    }
  }

  // ==================== FUNÇÕES PARA ABA "Quadro.1" ====================
  
  /**
   * Busca dados dos colaboradores na aba "Quadro.1"
   * CORREÇÃO: Nome exato da aba com letra maiúscula e .1
   */
  async obterDados() {
    try {
      await this.init();
      console.log('[SHEETS_COLETOR] Buscando dados da aba Quadro.1...');
      
      // ===== CORREÇÃO: Nome exato da aba =====
      const sheet = this.doc.sheetsByTitle['Quadro.1'];
      if (!sheet) {
        console.error('[SHEETS_COLETOR] Aba "Quadro.1" não encontrada');
        console.error('[SHEETS_COLETOR] Abas disponíveis:', Object.keys(this.doc.sheetsByTitle));
        return [];
      }
      
      const rows = await sheet.getRows();
      console.log(`[SHEETS_COLETOR] ${rows.length} linhas encontradas na aba Quadro.1`);
      
      const dados = [];
      
      rows.forEach(row => {
        // ===== CORREÇÃO: Nomes das colunas conforme a imagem =====
        // CHAPA1, NOME, FUNCAO (conforme visto na imagem)
        const chapa = String(row.get('CHAPA1') || '').trim();
        const nome = String(row.get('NOME') || '').trim();
        const funcao = String(row.get('FUNCAO') || '').trim();
        
        if (chapa && nome) {
          dados.push({
            chapa: chapa,
            nome: nome,
            funcao: funcao
          });
        }
      });
      
      console.log(`[SHEETS_COLETOR] ✓ ${dados.length} colaboradores carregados`);
      
      // Debug: Mostra primeiros 3 registros
      if (dados.length > 0) {
        console.log('[SHEETS_COLETOR] Exemplo de dados:', dados.slice(0, 3));
      }
      
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
      
      const agora = new Date();
      const situacoesTexto = situacoes.join(', ');
      
      // Validação anti-duplicação (120 segundos)
      const rows = await sheet.getRows();
      
      if (rows.length > 0) {
        const ultimaEntrada = rows[rows.length - 1];
        const ultimaData = ultimaEntrada.get('Data');
        const ultimaChapa = String(ultimaEntrada.get('Chapa') || '').trim();
        const ultimoColetor = String(ultimaEntrada.get('NumeroColetor') || '').trim();
        const ultimoTipo = String(ultimaEntrada.get('TipoOpe') || '').trim();
        
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
      
      const baseRows = await baseSheet.getRows();
      if (baseRows.length === 0) {
        console.log('[SHEETS_COLETOR] Base vazia, nada para atualizar');
        return;
      }
      
      if (!presencaSheet) {
        console.log('[SHEETS_COLETOR] Aba Presença não encontrada');
        return;
      }
      
      const presencaRows = await presencaSheet.getRows();
      if (presencaRows.length === 0) {
        console.log('[SHEETS_COLETOR] Presença vazia');
        return;
      }
      
      // Mapa para último registro de cada matrícula (de baixo para cima)
      const presencaMap = new Map();
      
      for (let i = presencaRows.length - 1; i >= 0; i--) {
        const row = presencaRows[i];
        const matricula = String(row.get('Matricula') || row.get('CHAPA1') || '').trim();
        const colunaE = String(row.get('E') || '').trim();
        const colunaH = String(row.get('H') || '').trim();
        
        if (matricula && !presencaMap.has(matricula)) {
          presencaMap.set(matricula, { colunaE, colunaH });
        }
      }
      
      console.log(`[SHEETS_COLETOR] Mapa Presença criado: ${presencaMap.size} matrículas`);
      
      // Atualiza cada linha da Base
      for (const baseRow of baseRows) {
        const matricula = String(baseRow.get('Chapa') || '').trim();
        
        if (matricula && presencaMap.has(matricula)) {
          const dados = presencaMap.get(matricula);
          baseRow.set('J', dados.colunaE);
          baseRow.set('K', dados.colunaH);
          await baseRow.save();
        }
      }
      
      console.log('[SHEETS_COLETOR] ✓ Dados da Presença atualizados na Base');
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao atualizar dados da Presença:', error);
    }
  }

  // ==================== FUNÇÕES PARA STATUS DOS COLETORES ====================
  
  /**
   * Retorna o último status de cada coletor
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
      
      const rows = await sheet.getRows();
      if (rows.length === 0) {
        console.log('[SHEETS_COLETOR] Base vazia');
        return {};
      }
      
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
      const resumoSupervisor = {};
      const colaboradoresContados = {};
      
      // Data de hoje sem hora
      const hoje = new Date();
      hoje.setHours(0, 0, 0, 0);
      
      const registrosPorSupCol = {};
      
      for (const row of rows) {
        const data = row.get('Data');
        if (!(data instanceof Date)) continue;
        
        const dataSemHora = new Date(data);
        dataSemHora.setHours(0, 0, 0, 0);
        
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
      
      for (const sup in registrosPorSupCol) {
        if (!resumoSupervisor[sup]) {
          resumoSupervisor[sup] = { retiradaContada: 0 };
        }
        
        colaboradoresContados[sup] = {};
        
        const colaboradores = registrosPorSupCol[sup];
        
        for (const colab in colaboradores) {
          const registros = colaboradores[colab]
            .slice()
            .sort((a, b) => b.data - a.data);
          
          const seq = registros
            .filter(r => r.tipo === "Retirada" || r.tipo === "Entrega")
            .map(r => r.tipo);
          
          const seqCresc = [...seq].reverse();
          
          let trocas = 0;
          for (let i = 1; i < seqCresc.length; i++) {
            if (seqCresc[i] !== seqCresc[i - 1]) {
              trocas++;
            }
          }
          
          const countRetirada = seqCresc.filter(x => x === "Retirada").length;
          
          let conta = 0;
          if (countRetirada > 0) {
            if (trocas <= 1) {
              conta = 1;
            } else {
              conta = 0;
            }
          }
          
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

module.exports = new SheetsColetorService();
