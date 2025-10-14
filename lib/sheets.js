// lib/sheets.js - VERSÃO FUNCIONAL COM salvarNaBase MELHORADO
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

class SheetsService {
  constructor() {
    this.doc = null;
    this.initialized = false;
  }

  async init() {
    if (this.initialized) return this.doc;
    try {
      const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SHEETS_CLIENT_EMAIL,
        key: process.env.GOOGLE_SHEETS_PRIVATE_KEY?.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });
      this.doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEETS_ID, serviceAccountAuth);
      await this.doc.loadInfo();
      this.initialized = true;
      return this.doc;
    } catch (error) {
      throw new Error('Falha na conexão: ' + error.message);
    }
  }

  async validarLogin(usuario, senha) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Usuarios'];
      if (!sheet) return { ok: false, msg: 'Aba Usuarios não encontrada' };
      
      const rows = await sheet.getRows();
      const abas = [];

      for (const row of rows) {
        const u = String(row.get('Usuario') || '').trim();
        const s = String(row.get('Senha') || '').trim();
        const aba = String(row.get('Aba') || '').trim();
        
        if (u === usuario && s === senha && aba) {
          abas.push(aba);
        }
      }

      const unicas = [...new Set(abas)];
      return unicas.length > 0 ? { ok: true, usuario, abas: unicas } : { ok: false, msg: 'Login inválido' };
    } catch (error) {
      return { ok: false, msg: error.message };
    }
  }

  async buscarColaboradores(filtro = '') {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Quadro'];
      if (!sheet) return [];
      
      const rows = await sheet.getRows();
      const lista = [];

      for (const row of rows) {
        const matricula = String(row.get('Coluna 1') || '').trim();
        const nome = String(row.get('NOME') || '').trim();
        const funcao = String(row.get('Função que atua') || row.get('FUNÇÃO NO RM') || '').trim();
        
        if (nome && (!filtro || nome.toLowerCase().includes(filtro.toLowerCase()) || matricula.toLowerCase().includes(filtro.toLowerCase()))) {
          lista.push({ matricula, nome, funcao });
        }
      }

      return lista;
    } catch (error) {
      console.error('Erro busca:', error);
      return [];
    }
  }

  async getBuffer(supervisor, aba) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) return [];
      
      const rows = await sheet.getRows();
      const buffer = [];

      for (const row of rows) {
        const sup = String(row.get('Supervisor') || '').toLowerCase();
        const rowAba = String(row.get('Grupo') || '').toLowerCase();
        
        if (sup === supervisor.toLowerCase() && (!aba || rowAba === aba.toLowerCase())) {
          buffer.push({
            matricula: row.get('matricula'),
            nome: row.get('Nome'),
            funcao: row.get('Função'),
            status: row.get('status') || ''
          });
        }
      }

      return buffer;
    } catch (error) {
      console.error('Erro buffer:', error);
      return [];
    }
  }

  async adicionarBuffer(supervisor, aba, colaborador) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) return { ok: false, msg: 'Aba Lista não encontrada' };
      
      await sheet.addRow({
        'Supervisor': supervisor,
        'Grupo': aba,
        'matricula': colaborador.matricula,
        'Nome': colaborador.nome,
        'Função': colaborador.funcao,
        'status': ''
      });

      return { ok: true };
    } catch (error) {
      console.error('Erro adicionar:', error);
      return { ok: false, msg: error.message };
    }
  }

  async removerBuffer(supervisor, matricula) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) return { ok: false };
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        if (row.get('Supervisor') === supervisor && String(row.get('matricula')) === String(matricula)) {
          await row.delete();
          return { ok: true };
        }
      }
      
      return { ok: false };
    } catch (error) {
      console.error('Erro remover:', error);
      return { ok: false };
    }
  }

  async atualizarStatusBuffer(supervisor, matricula, status) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) return { ok: false };
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        if (row.get('Supervisor') === supervisor && String(row.get('matricula')) === String(matricula)) {
          row.set('status', status);
          await row.save();
          return { ok: true };
        }
      }
      
      return { ok: false };
    } catch (error) {
      console.error('Erro atualizar status:', error);
      return { ok: false };
    }
  }

  async salvarNaBase(dados) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Base'];
      if (!sheet) return { ok: false, msg: 'Aba Base não encontrada' };
      
      const hoje = new Date().toLocaleDateString('pt-BR');
      
      console.log('[SHEETS] Iniciando salvamento na Base...');
      console.log('[SHEETS] Total de linhas para processar:', dados.length);
      
      // Carrega todas as linhas existentes da Base
      const rowsExistentes = await sheet.getRows();
      console.log('[SHEETS] Linhas existentes na Base:', rowsExistentes.length);
      
      let totalAtualizados = 0;
      let totalNovos = 0;
      
      for (const linha of dados) {
        const [sup, aba, matricula, nome, funcao, status] = linha;
        
        // Ignora linhas vazias
        if (!matricula && !nome) continue;
        
        // Procura por registro existente com mesmo Supervisor e Data
        let registroExistente = null;
        
        for (const row of rowsExistentes) {
          const rowSup = String(row.get('Supervisor') || '').trim();
          const rowData = String(row.get('Data') || '').trim();
          const rowMat = String(row.get('Matricula') || '').trim();
          
          // CONDIÇÃO: Mesmo supervisor, mesma data e mesma matrícula
          if (rowSup === sup && rowData === hoje && rowMat === matricula) {
            registroExistente = row;
            break;
          }
        }
        
        if (registroExistente) {
          // ATUALIZA apenas o Status
          console.log(`[SHEETS] Atualizando status do registro: ${nome} (${matricula})`);
          registroExistente.set('Status', status);
          await registroExistente.save();
          totalAtualizados++;
        } else {
          // INSERE novo registro
          console.log(`[SHEETS] Inserindo novo registro: ${nome} (${matricula})`);
          await sheet.addRow({
            'Supervisor': sup,
            'Aba': aba,
            'Matricula': matricula,
            'Nome': nome,
            'Função': funcao,
            'Status': status,
            'Data': hoje
          });
          totalNovos++;
        }
      }
      
      console.log('[SHEETS] Salvamento concluído!');
      console.log(`[SHEETS] ${totalNovos} novos registros inseridos`);
      console.log(`[SHEETS] ${totalAtualizados} registros atualizados`);
      
      return { 
        ok: true, 
        msg: `Dados salvos com sucesso! ${totalNovos} novos, ${totalAtualizados} atualizados.`,
        totais: {
          novos: totalNovos,
          atualizados: totalAtualizados,
          total: totalNovos + totalAtualizados
        }
      };
    } catch (error) {
      console.error('[SHEETS] Erro ao salvar:', error);
      return { ok: false, msg: error.message };
    }
  }
}

module.exports = new SheetsService();
