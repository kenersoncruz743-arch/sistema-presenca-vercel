// lib/sheets_2.js - VERSÃO CORRIGIDA: SALVA DATA E TIPOOPERAÇÃO CORRETAMENTE
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

class SheetsColetorService {
  constructor() {
    this.docHistorico = null;
    this.docAtual = null;
    this.initialized = false;
  }

  async init() {
    if (this.initialized) return { docHistorico: this.docHistorico, docAtual: this.docAtual };
    
    try {
      console.log('[SHEETS_COLETOR] Iniciando conexão...');
      
      const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SHEETS_CLIENT_EMAIL,
        key: process.env.GOOGLE_SHEETS_PRIVATE_KEY?.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });
      
      // PLANILHA HISTÓRICO
      const sheetIdHistorico = process.env.GOOGLE_SHEETS_ID_COLETOR;
      if (!sheetIdHistorico) {
        throw new Error('GOOGLE_SHEETS_ID_COLETOR não configurado');
      }
      
      this.docHistorico = new GoogleSpreadsheet(sheetIdHistorico, serviceAccountAuth);
      await this.docHistorico.loadInfo();
      console.log(`[SHEETS_COLETOR] ✓ Histórico: ${this.docHistorico.title}`);
      
      // PLANILHA ATUAL
      const sheetIdAtual = process.env.GOOGLE_SHEETS_ID;
      if (!sheetIdAtual) {
        throw new Error('GOOGLE_SHEETS_ID não configurado');
      }
      
      this.docAtual = new GoogleSpreadsheet(sheetIdAtual, serviceAccountAuth);
      await this.docAtual.loadInfo();
      console.log(`[SHEETS_COLETOR] ✓ Atual: ${this.docAtual.title}`);
      
      this.initialized = true;
      return { docHistorico: this.docHistorico, docAtual: this.docAtual };
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro:', error);
      throw new Error('Falha na conexão: ' + error.message);
    }
  }

  async obterDados() {
    try {
      await this.init();
      console.log('[SHEETS_COLETOR] Buscando Quadro...');
      
      const sheet = this.docAtual.sheetsByTitle['Quadro'];
      if (!sheet) {
        console.error('[SHEETS_COLETOR] Aba Quadro não encontrada');
        return [];
      }
      
      const rows = await sheet.getRows();
      const dados = [];
      
      rows.forEach(row => {
        const chapa = String(row.get('Coluna 1') || '').trim();
        const nome = String(row.get('NOME') || '').trim();
        const funcao = String(row.get('Função que atua') || row.get('FUNÇÃO NO RM') || '').trim();
        
        if (chapa && nome) {
          dados.push({ chapa, nome, funcao });
        }
      });
      
      console.log(`[SHEETS_COLETOR] ✓ ${dados.length} colaboradores`);
      return dados;
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro:', error);
      return [];
    }
  }

  async salvarRegistro(chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes) {
    try {
      await this.init();
      console.log('[SHEETS_COLETOR] Salvando:', { chapa, numeroColetor, tipoOperacao });
      
      if (!chapa || !numeroColetor || !situacoes || situacoes.length === 0) {
        return { ok: false, msg: 'Campos obrigatórios faltando' };
      }
      
      const agora = new Date();
      const situacoesTexto = situacoes.join(', ');
      const dataFormatada = this.formatarDataBR(agora);
      const horaFormatada = this.formatarHora(agora);
      
      const supervisor = await this.buscarSupervisorNaBase(chapa);
      
      // ===== 1. SALVA NO HISTÓRICO =====
      try {
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
        console.log('[SHEETS_COLETOR] ✓ Salvo no histórico');
      } catch (errorHistorico) {
        console.error('[SHEETS_COLETOR] Erro no histórico:', errorHistorico);
        throw new Error('Erro ao salvar histórico: ' + errorHistorico.message);
      }
      
      // ===== 2. ATUALIZA NA PLANILHA ATUAL =====
      try {
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
        console.log('[SHEETS_COLETOR] ✓ Atualizado na planilha atual');
      } catch (errorAtual) {
        console.error('[SHEETS_COLETOR] Erro na planilha atual:', errorAtual);
        console.log('[SHEETS_COLETOR] ⚠ Continuando apesar do erro na planilha atual');
      }
      
      return { ok: true, msg: 'Dados salvos com sucesso!' };
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro geral:', error);
      return { ok: false, msg: error.message };
    }
  }

  async salvarNoHistorico(dados) {
    try {
      console.log('[SHEETS_COLETOR] Salvando no histórico...');
      
      let sheetHistorico = this.docHistorico.sheetsByTitle['Coletor'];
      
      if (!sheetHistorico) {
        console.log('[SHEETS_COLETOR] Criando aba Coletor no histórico...');
        sheetHistorico = await this.docHistorico.addSheet({
          title: 'Coletor',
          headerValues: ['Data', 'Hora', 'Chapa', 'Nome', 'Funcao', 'NumeroColetor', 'TipoOperacao', 'Situacao', 'Supervisor']
        });
      }
      
      const rows = await sheetHistorico.getRows();
      
      if (rows.length > 0) {
        const ultima = rows[rows.length - 1];
        const ultimaData = ultima.get('Data');
        const ultimaHora = ultima.get('Hora');
        const ultimaChapa = String(ultima.get('Chapa') || '').trim();
        const ultimoColetor = String(ultima.get('NumeroColetor') || '').trim();
        const ultimoTipo = String(ultima.get('TipoOperacao') || '').trim();
        
        if (ultimaChapa === dados.chapa && 
            ultimoColetor === String(dados.numeroColetor) && 
            ultimoTipo === dados.tipoOperacao &&
            ultimaData === dados.data) {
          
          try {
            const [ultH, ultM] = ultimaHora.split(':').map(Number);
            const [novoH, novoM] = dados.hora.split(':').map(Number);
            
            const ultMinutos = ultH * 60 + ultM;
            const novoMinutos = novoH * 60 + novoM;
            
            if (Math.abs(novoMinutos - ultMinutos) < 2) {
              throw new Error('Registro duplicado (menos de 2 minutos)');
            }
          } catch (timeError) {
            console.error('[SHEETS_COLETOR] Erro ao validar tempo:', timeError);
          }
        }
      }
      
      await sheetHistorico.addRow({
        'Data': dados.data,
        'Hora': dados.hora,
        'Chapa': dados.chapa,
        'Nome': dados.nome,
        'Funcao': dados.funcao,
        'NumeroColetor': String(dados.numeroColetor),
        'TipoOperacao': dados.tipoOperacao,
        'Situacao': dados.situacao,
        'Supervisor': dados.supervisor
      });
      
      console.log('[SHEETS_COLETOR] ✓ Adicionado ao histórico');
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao salvar no histórico:', error);
      throw error;
    }
  }

  async atualizarStatusAtual(dados) {
    try {
      console.log('[SHEETS_COLETOR] Atualizando planilha atual...');
      
      // ===== CARREGA OU CRIA ABA COM HEADERS CORRETOS =====
      let sheetAtual = this.docAtual.sheetsByTitle['Coletor'];
      
      if (!sheetAtual) {
        console.log('[SHEETS_COLETOR] Criando aba Coletor na planilha atual...');
        
        try {
          sheetAtual = await this.docAtual.addSheet({
            title: 'Coletor',
            headerValues: ['Chapa', 'Nome', 'Funcao', 'NumeroColetor', 'TipoOperacao', 'Situacao', 'Supervisor', 'Data', 'Hora']
          });
          console.log('[SHEETS_COLETOR] ✓ Aba criada com headers: Chapa, Nome, Funcao, NumeroColetor, TipoOperacao, Situacao, Supervisor, Data, Hora');
        } catch (createError) {
          console.error('[SHEETS_COLETOR] Erro ao criar aba:', createError);
          throw new Error('Falha ao criar aba Coletor: ' + createError.message);
        }
      }
      
      // ===== CARREGA LINHAS EXISTENTES =====
      let rows;
      try {
        await sheetAtual.loadHeaderRow();
        rows = await sheetAtual.getRows();
        console.log(`[SHEETS_COLETOR] ${rows.length} registros existentes`);
      } catch (loadError) {
        console.error('[SHEETS_COLETOR] Erro ao carregar linhas:', loadError);
        throw new Error('Falha ao carregar dados: ' + loadError.message);
      }
      
      // ===== BUSCA REGISTRO EXISTENTE =====
      let registroExistente = null;
      
      for (const row of rows) {
        const chapaRow = String(row.get('Chapa') || '').trim();
        if (chapaRow === dados.chapa) {
          registroExistente = row;
          break;
        }
      }
      
      // ===== ATUALIZA OU CRIA =====
      if (registroExistente) {
        console.log(`[SHEETS_COLETOR] Atualizando chapa ${dados.chapa}...`);
        
        try {
          registroExistente.set('Nome', dados.nome);
          registroExistente.set('Funcao', dados.funcao);
          registroExistente.set('NumeroColetor', String(dados.numeroColetor));
          registroExistente.set('TipoOperacao', dados.tipoOperacao);
          registroExistente.set('Situacao', dados.situacao);
          registroExistente.set('Supervisor', dados.supervisor);
          registroExistente.set('Data', dados.data);
          registroExistente.set('Hora', dados.hora);
          
          await registroExistente.save();
          console.log('[SHEETS_COLETOR] ✓ Registro atualizado com Data e Hora');
        } catch (updateError) {
          console.error('[SHEETS_COLETOR] Erro ao atualizar:', updateError);
          throw new Error('Falha ao atualizar: ' + updateError.message);
        }
        
      } else {
        console.log(`[SHEETS_COLETOR] Criando novo registro para chapa ${dados.chapa}...`);
        
        try {
          await sheetAtual.addRow({
            'Chapa': dados.chapa,
            'Nome': dados.nome,
            'Funcao': dados.funcao,
            'NumeroColetor': String(dados.numeroColetor),
            'TipoOperacao': dados.tipoOperacao,
            'Situacao': dados.situacao,
            'Supervisor': dados.supervisor,
            'Data': dados.data,
            'Hora': dados.hora
          });
          console.log('[SHEETS_COLETOR] ✓ Novo registro criado com Data e Hora');
        } catch (addError) {
          console.error('[SHEETS_COLETOR] Erro ao adicionar:', addError);
          throw new Error('Falha ao adicionar: ' + addError.message);
        }
      }
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro em atualizarStatusAtual:', error);
      throw error;
    }
  }

  async buscarSupervisorNaBase(chapa) {
    try {
      const sheetBase = this.docAtual.sheetsByTitle['Base'];
      if (!sheetBase) {
        console.log('[SHEETS_COLETOR] Aba Base não encontrada');
        return 'Sem Supervisor';
      }
      
      const rows = await sheetBase.getRows();
      
      const hoje = new Date();
      const dia = String(hoje.getDate()).padStart(2, '0');
      const mes = String(hoje.getMonth() + 1).padStart(2, '0');
      const ano = hoje.getFullYear();
      const dataHojeBR = `${dia}/${mes}/${ano}`;
      
      for (let i = rows.length - 1; i >= 0; i--) {
        const row = rows[i];
        const matricula = String(row.get('Matricula') || '').trim();
        const dataRegistro = String(row.get('Data') || '').trim();
        
        if (matricula === chapa && dataRegistro === dataHojeBR) {
          const supervisor = String(row.get('Supervisor') || '').trim();
          console.log(`[SHEETS_COLETOR] Supervisor: ${supervisor}`);
          return supervisor || 'Sem Supervisor';
        }
      }
      
      console.log('[SHEETS_COLETOR] Supervisor não encontrado');
      return 'Sem Supervisor';
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro ao buscar supervisor:', error);
      return 'Sem Supervisor';
    }
  }

  async obterColetorStatus() {
    try {
      await this.init();
      console.log('[SHEETS_COLETOR] Obtendo status...');
      
      const sheetAtual = this.docAtual.sheetsByTitle['Coletor'];
      if (!sheetAtual) {
        console.log('[SHEETS_COLETOR] Aba Coletor não existe ainda');
        return {};
      }
      
      const rows = await sheetAtual.getRows();
      const mapa = {};
      
      rows.forEach(row => {
        const chapa = String(row.get('Chapa') || '').trim();
        const nome = String(row.get('Nome') || '').trim();
        const funcao = String(row.get('Funcao') || '').trim();
        const coletor = String(row.get('NumeroColetor') || '').trim();
        const tipo = String(row.get('TipoOperacao') || '').trim();
        const situacao = String(row.get('Situacao') || '').trim();
        const supervisor = String(row.get('Supervisor') || '').trim();
        const data = String(row.get('Data') || '').trim();
        const hora = String(row.get('Hora') || '').trim();
        
        if (coletor) {
          mapa[coletor] = {
            chapa, nome, funcao, tipo, situacao, supervisor,
            data: data,
            hora: hora
          };
        }
      });
      
      console.log(`[SHEETS_COLETOR] ✓ ${Object.keys(mapa).length} coletores`);
      return mapa;
      
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro:', error);
      return {};
    }
  }

  async gerarResumoColetores() {
    try {
      const statusMap = await this.obterColetorStatus();
      
      let disponiveis = 0, indisponiveis = 0, quebrados = 0;
      
      for (const coletor in statusMap) {
        const s = statusMap[coletor];
        if (s.tipo === "Entrega" && s.situacao === "OK") disponiveis++;
        else if (s.tipo === "Retirada") indisponiveis++;
        if (s.situacao !== "OK") quebrados++;
      }
      
      return { disponiveis, indisponiveis, quebrados, total: Object.keys(statusMap).length };
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro:', error);
      return { disponiveis: 0, indisponiveis: 0, quebrados: 0, total: 0 };
    }
  }

  async gerarResumoPorSupervisor() {
    try {
      await this.init();
      
      const sheetAtual = this.docAtual.sheetsByTitle['Coletor'];
      if (!sheetAtual) return {};
      
      const rows = await sheetAtual.getRows();
      const resumo = {};
      
      rows.forEach(row => {
        const sup = String(row.get('Supervisor') || 'Sem Supervisor').trim();
        const tipo = String(row.get('TipoOperacao') || '').trim();
        
        if (!resumo[sup]) resumo[sup] = { retiradaContada: 0 };
        if (tipo === 'Retirada') resumo[sup].retiradaContada++;
      });
      
      return resumo;
    } catch (error) {
      console.error('[SHEETS_COLETOR] Erro:', error);
      return {};
    }
  }

  formatarDataBR(data) {
    if (!data || !(data instanceof Date)) return '';
    try {
      const d = String(data.getDate()).padStart(2, '0');
      const m = String(data.getMonth() + 1).padStart(2, '0');
      const a = data.getFullYear();
      return `${d}/${m}/${a}`;
    } catch (e) {
      return '';
    }
  }
  
  formatarHora(data) {
    if (!data || !(data instanceof Date)) return '';
    try {
      const h = String(data.getHours()).padStart(2, '0');
      const m = String(data.getMinutes()).padStart(2, '0');
      return `${h}:${m}`;
    } catch (e) {
      return '';
    }
  }
}

module.exports = new SheetsColetorService();
