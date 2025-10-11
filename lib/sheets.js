// lib/sheets.js - VERSÃO COM SUPORTE A QLP
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
      
      for (const linha of dados) {
        const [sup, aba, matricula, nome, funcao, status] = linha;
        if (matricula || nome) {
          await sheet.addRow({
            'Supervisor': sup,
            'Aba': aba,
            'Matricula': matricula,
            'Nome': nome,
            'Função': funcao,
            'Status': status,
            'Data': hoje
          });
        }
      }
      return { ok: true, msg: 'Dados salvos com sucesso' };
    } catch (error) {
      console.error('Erro salvar:', error);
      return { ok: false, msg: error.message };
    }
  }

  // NOVOS MÉTODOS PARA QLP
  async getDadosQLP() {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['QLP'];
      if (!sheet) return { ok: false, msg: 'Aba QLP não encontrada' };
      
      const rows = await sheet.getRows();
      const dados = {
        funcoes: {}
      };

      for (const row of rows) {
        const funcao = String(row.get('Função') || row.get('FUNÇÃO') || '').trim();
        const supervisor = String(row.get('Supervisor') || '').trim();
        const nome = String(row.get('Nome') || '').trim();
        const turno = String(row.get('Turno') || '').trim();
        const tipoOperacional = String(row.get('Seção') || row.get('Tipo') || 'Não operacional').trim();

        if (!funcao || !nome) continue;

        // Organizar por função
        if (!dados.funcoes[funcao]) {
          dados.funcoes[funcao] = {
            supervisores: {},
            totalContadores: {
              'Turno A': 0,
              'Turno B': 0,
              'Turno C': 0,
              'Situacao': 0
            }
          };
        }

        // Organizar por supervisor dentro da função
        if (!dados.funcoes[funcao].supervisores[supervisor]) {
          dados.funcoes[funcao].supervisores[supervisor] = {
            colaboradores: [],
            totalContadores: {
              'Turno A': 0,
              'Turno B': 0,
              'Turno C': 0,
              'Situacao': 0
            }
          };
        }

        // Adicionar colaborador
        dados.funcoes[funcao].supervisores[supervisor].colaboradores.push({
          nome,
          turno,
          tipoOperacional
        });

        // Atualizar contadores
        const contadorKey = ['Turno A', 'Turno B', 'Turno C'].includes(turno) ? turno : 'Situacao';
        dados.funcoes[funcao].supervisores[supervisor].totalContadores[contadorKey]++;
        dados.funcoes[funcao].totalContadores[contadorKey]++;
      }

      return { ok: true, dados: dados.funcoes };
    } catch (error) {
      console.error('Erro ao buscar dados QLP:', error);
      return { ok: false, msg: error.message };
    }
  }

  async exportarQLP(formato = 'json') {
    try {
      const resultado = await this.getDadosQLP();
      if (!resultado.ok) return resultado;

      if (formato === 'json') {
        return { ok: true, dados: resultado.dados };
      }

      // Formato CSV
      if (formato === 'csv') {
        let csv = 'Função,Supervisor,Nome,Turno,Seção\n';
        
        Object.entries(resultado.dados).forEach(([funcao, dadosFuncao]) => {
          Object.entries(dadosFuncao.supervisores).forEach(([supervisor, dadosSup]) => {
            dadosSup.colaboradores.forEach(colab => {
              csv += `"${funcao}","${supervisor}","${colab.nome}","${colab.turno}","${colab.tipoOperacional}"\n`;
            });
          });
        });

        return { ok: true, csv };
      }

      return { ok: false, msg: 'Formato não suportado' };
    } catch (error) {
      console.error('Erro ao exportar QLP:', error);
      return { ok: false, msg: error.message };
    }
  }
}

module.exports = new SheetsService();
