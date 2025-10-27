// lib/sheets_2.js - VERSÃO CORRIGIDA: DUAS PLANILHAS COM SINCRONIZAÇÃO
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

class SheetsColetorService {
  constructor() {
    this.docHistorico = null; // Planilha com histórico completo
    this.docAtual = null;     // Planilha com status atual
    this.initialized = false;
  }

  async init() {
    if (this.initialized) return { docHistorico: this.docHistorico, docAtual: this.docAtual };
    
    try {
      console.log('[SHEETS_COLETOR] Iniciando conexão com ambas as planilhas...');
      
      const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SHEETS_CLIENT_EMAIL,
        key: process.env.GOOGLE_SHEETS_PRIVATE_KEY?.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });
      
      // ===== PLANILHA 1: HISTÓRICO COMPLETO =====
      const sheetIdHistorico = process.env.GOOGLE_SHEETS_ID_COLETOR;
      if (!sheetIdHistorico) {
        throw new Error('GOOGLE_SHEETS_ID_COLETOR não configurado');
      }
      
      this.docHistorico = new GoogleSpreadsheet(sheetIdHistorico, serviceAccountAuth);
      await this.docHistorico.loadInfo();
      console.log(`[SHEETS_COLETOR] ✓ Planilha HISTÓRICO: ${this.docHistorico.title}`);
      
      // ===== PLANILHA 2: STATUS ATUAL =====
      const sheetIdAtual = process.env.GOOGLE_SHEETS_ID;
      if (!sheetIdAtual) {
        throw new Error('GOOGLE_SHEETS_ID não configurado');
      }
      
      this.docAtual = new GoogleSpreadsheet(sheetIdAtual, serviceAccountAuth);
      await this.docAtual.loadInfo();
      console.log(`[SHEETS_COLETOR] ✓ Planilha ATUAL: ${this.docAtual.title}`);
      
      this.initialized = true;
      return { docHistorico: this.docHistorico, docAtual: this.docAtual };
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro na conexão:', error);
      throw new Error('Falha na conexão: ' + error.message);
    }
  }

  // ==================== BUSCAR COLABORADORES (DA PLANILHA ATUAL) ====================
  
  async obterDados() {
    try {
      await this.init();
      console.log('[SHEETS_COLETOR] Buscando dados da aba Quadro (planilha atual)...');
      
      const sheet = this.docAtual.sheetsByTitle['Quadro'];
      if (!sheet) {
        console.error('[SHEETS_COLETOR] Aba "Quadro" não encontrada na planilha atual');
        console.error('[SHEETS_COLETOR] Abas disponíveis:', Object.keys(this.docAtual.sheetsByTitle));
        return [];
      }
      
      const rows = await sheet.getRows();
      console.log(`[SHEETS_COLETOR] ${rows.length} linhas encontradas`);
      
      const dados = [];
      
      rows.forEach(row => {
        const chapa = String(row.get('Coluna 1') || '').trim();
        const nome = String(row.get('NOME') || '').trim();
        const funcao = String(row.get('Função que atua') || row.get('FUNÇÃO NO RM') || '').trim();
        
        if (chapa && nome) {
          dados.push({ chapa, nome, funcao });
        }
      });
      
      console.log(`[SHEETS_COLETOR] ✓ ${dados.length} colaboradores carregados`);
      return dados;
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao obter dados:', error);
      return [];
    }
  }

  // ==================== SALVAR REGISTRO (DUPLO: HISTÓRICO + ATUAL) ====================
  
  async salvarRegistro(chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes) {
    try {
      await this.init();
      console.log('[SHEETS_COLETOR] Salvando registro:', {
        chapa, nome, numeroColetor, tipoOperacao, situacoes
      });
      
      // Validação
      if (!chapa || !numeroColetor || !situacoes || situacoes.length === 0) {
        return { ok: false, msg: 'Preencha todos os campos obrigatórios.' };
      }
      
      const agora = new Date();
      const situacoesTexto = situacoes.join(', ');
      const dataFormatada = this.formatarDataBR(agora);
      const horaFormatada = this.formatarHora(agora);
      
      // Busca supervisor na Base (planilha atual)
      const supervisor = await this.buscarSupervisorNaBase(chapa);
      
      // ===== 1. SALVA NO HISTÓRICO (PLANILHA HISTÓRICO) =====
      await this.salvarNoHistorico({
        data: dataFormatada,
        hora: horaFormatada,
        chapa,
        nome,
        funcao,
        numeroColetor,
        tipoOperacao,
        situacao: situacoesTexto,
        supervisor
      });
      
      // ===== 2. ATUALIZA STATUS ATUAL (PLANILHA ATUAL) =====
      await this.atualizarStatusAtual({
        chapa,
        nome,
        funcao,
        numeroColetor,
        tipoOperacao,
        situacao: situacoesTexto,
        supervisor,
        data: dataFormatada,
        hora: horaFormatada
      });
      
      console.log('[SHEETS_COLETOR] ✓ Registro salvo em ambas as planilhas');
      return { ok: true, msg: 'Dados salvos com sucesso!' };
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao salvar registro:', error);
      return { ok: false, msg: 'Erro ao salvar: ' + error.message };
    }
  }

  // ==================== SALVAR NO HISTÓRICO (APPEND ONLY) ====================
  
  async salvarNoHistorico(dados) {
    try {
      console.log('[SHEETS_COLETOR] Salvando no histórico...');
      
      let sheetHistorico = this.docHistorico.sheetsByTitle['Coletor'];
      
      // Cria aba se não existir
      if (!sheetHistorico) {
        console.log('[SHEETS_COLETOR] Criando aba Coletor no histórico...');
        sheetHistorico = await this.docHistorico.addSheet({
          title: 'Coletor',
          headerValues: ['Data', 'Hora', 'Chapa', 'Nome', 'Funcao', 'NumeroColetor', 'TipoOpe', 'Situacao', 'Supervisor']
        });
      }
      
      // Validação anti-duplicação (120 segundos)
      const rows = await sheetHistorico.getRows();
      
      if (rows.length > 0) {
        const ultimaEntrada = rows[rows.length - 1];
        const ultimaData = ultimaEntrada.get('Data');
        const ultimaHora = ultimaEntrada.get('Hora');
        const ultimaChapa = String(ultimaEntrada.get('Chapa') || '').trim();
        const ultimoColetor = String(ultimaEntrada.get('NumeroColetor') || '').trim();
        const ultimoTipo = String(ultimaEntrada.get('TipoOpe') || '').trim();
        
        if (ultimaChapa === dados.chapa && 
            ultimoColetor === String(dados.numeroColetor) && 
            ultimoTipo === dados.tipoOperacao &&
            ultimaData === dados.data) {
          
          // Verifica se foi há menos de 2 minutos
          const [ultHora, ultMin] = ultimaHora.split(':').map(Number);
          const [novaHora, novaMin] = dados.hora.split(':').map(Number);
          
          const ultMinutos = ultHora * 60 + ultMin;
          const novaMinutos = novaHora * 60 + novaMin;
          
          if (Math.abs(novaMinutos - ultMinutos) < 2) {
            console.log('[SHEETS_COLETOR] Registro duplicado detectado');
            throw new Error('Registro já enviado recentemente. Aguarde alguns minutos.');
          }
        }
      }
      
      // Adiciona ao histórico
      await sheetHistorico.addRow({
        'Data': dados.data,
        'Hora': dados.hora,
        'Chapa': dados.chapa,
        'Nome': dados.nome,
        'Funcao': dados.funcao,
        'NumeroColetor': String(dados.numeroColetor),
        'TipoOpe': dados.tipoOperacao,
        'Situacao': dados.situacao,
        'Supervisor': dados.supervisor
      });
      
      console.log('[SHEETS_COLETOR] ✓ Salvo no histórico');
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao salvar no histórico:', error);
      throw error;
    }
  }

  // ==================== ATUALIZAR STATUS ATUAL (REPLACE) ====================
  
  async atualizarStatusAtual(dados) {
    try {
      console.log('[SHEETS_COLETOR] Atualizando status atual...');
      
      let sheetAtual = this.docAtual.sheetsByTitle['Coletor'];
      
      // Cria aba se não existir
      if (!sheetAtual) {
        console.log('[SHEETS_COLETOR] Criando aba Coletor na planilha atual...');
        sheetAtual = await this.docAtual.addSheet({
          title: 'Coletor',
          headerValues: ['Chapa', 'Nome', 'Funcao', 'NumeroColetor', 'TipoOpe', 'Situacao', 'Supervisor', 'UltimaAtualizacao']
        });
      }
      
      const rows = await sheetAtual.getRows();
      
      // ===== BUSCA SE JÁ EXISTE REGISTRO DESTA CHAPA =====
      let registroExistente = null;
      
      for (const row of rows) {
        const chapaRow = String(row.get('Chapa') || '').trim();
        
        if (chapaRow === dados.chapa) {
          registroExistente = row;
          break;
        }
      }
      
      const ultimaAtualizacao = `${dados.data} ${dados.hora}`;
      
      if (registroExistente) {
        // ===== ATUALIZA REGISTRO EXISTENTE =====
        console.log(`[SHEETS_COLETOR] Atualizando registro existente da chapa ${dados.chapa}`);
        
        registroExistente.set('Nome', dados.nome);
        registroExistente.set('Funcao', dados.funcao);
        registroExistente.set('NumeroColetor', String(dados.numeroColetor));
        registroExistente.set('TipoOpe', dados.tipoOperacao);
        registroExistente.set('Situacao', dados.situacao);
        registroExistente.set('Supervisor', dados.supervisor);
        registroExistente.set('UltimaAtualizacao', ultimaAtualizacao);
        
        await registroExistente.save();
        console.log('[SHEETS_COLETOR] ✓ Status atualizado');
        
      } else {
        // ===== CRIA NOVO REGISTRO =====
        console.log(`[SHEETS_COLETOR] Criando novo registro para chapa ${dados.chapa}`);
        
        await sheetAtual.addRow({
          'Chapa': dados.chapa,
          'Nome': dados.nome,
          'Funcao': dados.funcao,
          'NumeroColetor': String(dados.numeroColetor),
          'TipoOpe': dados.tipoOperacao,
          'Situacao': dados.situacao,
          'Supervisor': dados.supervisor,
          'UltimaAtualizacao': ultimaAtualizacao
        });
        
        console.log('[SHEETS_COLETOR] ✓ Novo status criado');
      }
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao atualizar status atual:', error);
      throw error;
    }
  }

  // ==================== BUSCAR SUPERVISOR NA BASE ====================
  
  async buscarSupervisorNaBase(chapa) {
    try {
      const sheetBase = this.docAtual.sheetsByTitle['Base'];
      if (!sheetBase) {
        console.log('[SHEETS_COLETOR] Aba Base não encontrada');
        return 'Sem Supervisor';
      }
      
      const rows = await sheetBase.getRows();
      
      // Data de hoje
      const hoje = new Date();
      const diaHoje = String(hoje.getDate()).padStart(2, '0');
      const mesHoje = String(hoje.getMonth() + 1).padStart(2, '0');
      const anoHoje = hoje.getFullYear();
      const dataHojeBR = `${diaHoje}/${mesHoje}/${anoHoje}`;
      
      // Busca último registro DO DIA de trás para frente
      for (let i = rows.length - 1; i >= 0; i--) {
        const row = rows[i];
        const matricula = String(row.get('Matricula') || '').trim();
        const dataRegistro = String(row.get('Data') || '').trim();
        
        // Busca apenas registros do dia atual
        if (matricula === chapa && dataRegistro === dataHojeBR) {
          const supervisor = String(row.get('Supervisor') || '').trim();
          console.log(`[SHEETS_COLETOR] Supervisor encontrado: ${supervisor}`);
          return supervisor || 'Sem Supervisor';
        }
      }
      
      console.log('[SHEETS_COLETOR] Supervisor não encontrado para hoje');
      return 'Sem Supervisor';
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao buscar supervisor:', error);
      return 'Sem Supervisor';
    }
  }

  // ==================== OBTER STATUS DOS COLETORES (DA PLANILHA ATUAL) ====================
  
  async obterColetorStatus() {
    try {
      await this.init();
      console.log('[SHEETS_COLETOR] Obtendo status dos coletores (planilha atual)...');
      
      const sheetAtual = this.docAtual.sheetsByTitle['Coletor'];
      if (!sheetAtual) {
        console.log('[SHEETS_COLETOR] Aba Coletor não encontrada na planilha atual');
        return {};
      }
      
      const rows = await sheetAtual.getRows();
      console.log(`[SHEETS_COLETOR] ${rows.length} registros encontrados`);
      
      const mapa = {};
      
      rows.forEach(row => {
        const chapa = String(row.get('Chapa') || '').trim();
        const nome = String(row.get('Nome') || '').trim();
        const funcao = String(row.get('Funcao') || '').trim();
        const coletor = String(row.get('NumeroColetor') || '').trim();
        const tipo = String(row.get('TipoOpe') || '').trim();
        const situacao = String(row.get('Situacao') || '').trim();
        const supervisor = String(row.get('Supervisor') || '').trim();
        const ultimaAtualizacao = String(row.get('UltimaAtualizacao') || '').trim();
        
        if (coletor) {
          mapa[coletor] = {
            chapa,
            nome,
            funcao,
            tipo,
            situacao,
            supervisor,
            ultimaAtualizacao
          };
        }
      });
      
      console.log(`[SHEETS_COLETOR] ✓ Status de ${Object.keys(mapa).length} coletores obtido`);
      return mapa;
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao obter status:', error);
      return {};
    }
  }

  // ==================== GERAR RESUMOS ====================
  
  async gerarResumoColetores() {
    try {
      const statusPorColetor = await this.obterColetorStatus();
      
      let disponiveis = 0;
      let indisponiveis = 0;
      let quebrados = 0;
      
      for (const coletor in statusPorColetor) {
        const status = statusPorColetor[coletor];
        
        if (status.tipo === "Entrega" && status.situacao === "OK") {
          disponiveis++;
        } else if (status.tipo === "Retirada") {
          indisponiveis++;
        }
        
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
      return { disponiveis: 0, indisponiveis: 0, quebrados: 0, total: 0 };
    }
  }

  async gerarResumoPorSupervisor() {
    try {
      await this.init();
      console.log('[SHEETS_COLETOR] Gerando resumo por supervisor...');
      
      const sheetAtual = this.docAtual.sheetsByTitle['Coletor'];
      if (!sheetAtual) {
        console.error('[SHEETS_COLETOR] Aba Coletor não encontrada');
        return {};
      }
      
      const rows = await sheetAtual.getRows();
      const resumoSupervisor = {};
      
      rows.forEach(row => {
        const supervisor = String(row.get('Supervisor') || 'Sem Supervisor').trim();
        const tipo = String(row.get('TipoOpe') || '').trim();
        
        if (!resumoSupervisor[supervisor]) {
          resumoSupervisor[supervisor] = { retiradaContada: 0 };
        }
        
        if (tipo === 'Retirada') {
          resumoSupervisor[supervisor].retiradaContada++;
        }
      });
      
      console.log('[SHEETS_COLETOR] ✓ Resumo por supervisor gerado:', resumoSupervisor);
      return resumoSupervisor;
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao gerar resumo por supervisor:', error);
      return {};
    }
  }

  // ==================== HELPERS ====================
  
  formatarDataBR(data) {
    if (!data || !(data instanceof Date)) return '';
    
    try {
      const dia = String(data.getDate()).padStart(2, '0');
      const mes = String(data.getMonth() + 1).padStart(2, '0');
      const ano = data.getFullYear();
      return `${dia}/${mes}/${ano}`;
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao formatar data:', error);
      return '';
    }
  }
  
  formatarHora(data) {
    if (!data || !(data instanceof Date)) return '';
    
    try {
      const hora = String(data.getHours()).padStart(2, '0');
      const min = String(data.getMinutes()).padStart(2, '0');
      return `${hora}:${min}`;
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao formatar hora:', error);
      return '';
    }
  }
}

module.exports = new SheetsColetorService();
